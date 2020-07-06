/**
 * Configures the service.
 */
function getService() {

  var sandbox = PropertiesService.getScriptProperties().getProperty("SANDBOX");
  var clientId = PropertiesService.getScriptProperties().getProperty("CLIENT_ID");
  var clientSecret = PropertiesService.getScriptProperties().getProperty("CLIENT_SECRET");
  var authUrl = PropertiesService.getScriptProperties().getProperty("LOGIN_URL");
  var tokenUrl = PropertiesService.getScriptProperties().getProperty("TOKEN_URL");
  
  if(sandbox) {
    authUrl = "https://test.salesforce.com/services/oauth2/authorize";
    tokenUrl = "https://test.salesforce.com/services/oauth2/token";
    clientId = 'do something';
    clientSecret = 'do something';
  }
  return OAuth2.createService('Saleforce')
      // Set the endpoint URLs.
      .setAuthorizationBaseUrl(authUrl)
      .setTokenUrl(tokenUrl)

      // Set the client ID and secret.
      .setClientId(clientId)
      .setClientSecret(clientSecret)

      // Set the name of the callback function that should be invoked to
      // complete the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())
      // use to cache to avoid exhausting your PropertiesService quotas
      .setCache(CacheService.getUserCache())
      // avoid a possible lock on race condition 
      .setLock(LockService.getUserLock())
      // Set the scopes to be requested.
      .setScope('api refresh_token');
}

/**
 * Handles the OAuth callback.
 */
function authCallback(request) {
  var service = getService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    var id_url = service.storage_.memory_['oauth2.Saleforce'].id; // this is not a typo here, but a typo in the OAuth lib
    var user_id = id_url.substring(id_url.lastIndexOf('/')+1, id_url.length);
    CacheService.getUserCache().put('sforce_user_id', user_id);
    
    return HtmlService.createHtmlOutput('Success!');
  } else {
    return HtmlService.createHtmlOutput('Denied.');
  }
}

/**
 * Wrapper function that detects an expired session, refreshes the access token,
 * and retries the request again.
 * @param {OAuth2.Service_} service The service to refresh.
 * @param {Function} func The function that makes the UrlFetchApp request
                          and returns the response.
 * @return {UrlFetchApp.HTTPResponse} The HTTP response.
 */
function withRetry(service, func) {
  var response;
  var content;
  try {
    response = func();
    content = response.getContentText();
  } catch (e) {
    content = e.toString();
  }
  if (content.indexOf('INVALID_SESSION_ID') !== -1) {
    service.refresh();
    return func();
  }
  return response;
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  getService().reset();
}

function haveWeLoggedThisTimeAlready(calendarEventId) {
  var service = getService();
  var taskRayTimeSOQLQuery = `select id from TASKRAY__trTaskTime__c where gcal_event_id__c = \'${calendarEventId}\'`;
  var url = `${service.getToken().instance_url}/services/data/v47.0/query?q=${encodeURI(taskRayTimeSOQLQuery)}`;
  
  var response = withRetry(service, function() {
    console.log(`return list of existing TASKRAY time entries for calendar entry ${calendarEventId}`);
    return UrlFetchApp.fetch(url, {
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken(),
      }
    });
  });
  if(response.getResponseCode() == 200) {
    // we don't want to cache this, we want to check everytime
    return JSON.parse(response.getContentText());
  } else {
    console.log(JSON.stringify(response));
    // if we don't get 200, let it crash
    console.log(`ERROR checking if calendar event already exists : ${response.getResponseCode()}`);
  }
}

/**
* returns the ID for this user
* will attempt to get from a cache, if not will query Salesforce
**/
function getUserId() {
  var sandbox = PropertiesService.getScriptProperties().getProperty("SANDBOX");
  var service = getService();
  
  var username = sandbox ? Session.getActiveUser().getEmail() + "." + sandbox : Session.getActiveUser().getEmail();
  var userIdQuery = `select id from user where username = \'${username}\'`;
  var url = `${service.getToken().instance_url}/services/data/v47.0/query?q=${encodeURI(userIdQuery)}`;
  
  var userId = CacheService.getUserCache().get("user_id");
  //console.log(userIdQuery);
  
  if(userId) {
    return userId;
  } else {
  
    var response = withRetry(service, function() {
      return UrlFetchApp.fetch(url, {
        headers: {
          Authorization: 'Bearer ' + service.getAccessToken(),
        }
      });                  
    });
    if(response && response.getResponseCode() == 200) {
      userId = JSON.parse(response.getContentText())['records'][0]['Id'];
      CacheService.getUserCache().put("user_id", userId);

      return userId;
    }
  }
}

/**
*
* Fetch all the TaskRay Projects and associated tasks 
* will look to see if these are cached before fetching from Salesforce
*/
function fetchProjectsAndTasks() {
  var sandbox = PropertiesService.getScriptProperties().getProperty("SANDBOX");
  var settings = getUserSettings();
  var service = getService();
  var userId = getUserId();
  
  // SOQL for getting project Id's user is a contributor in. We assume email address of google user will be the salesforce username.
  var username = sandbox ? Session.getActiveUser().getEmail() + "." + sandbox : Session.getActiveUser().getEmail();
  
  // build the projects part of the query according to the user settings
  var projectSOQLQuery = `TASKRAY__Project__c where TASKRAY__trTemplate__c = false and TASKRAY__trStatus__c in ('Assigned','In Progress','On Hold') and id in (SELECT TASKRAY__Project__c FROM TASKRAY__trContributor__c where TASKRAY__User__r.username = '${username }')`;
                                     
  // build the tasks part of the query according to user settings
  var tasksSOQLQuery = `select id, name from TASKRAY__Tasks__r where taskray__trCompleted__c = false`;
  if(settings.myTasksOnly) {
      tasksSOQLQuery += ` and ownerId = '${userId}'`;
  }    
  
  var taskRaySOQLQuery = `select id, name, (${tasksSOQLQuery}) from ${projectSOQLQuery}`; 
  //console.log(taskRaySOQLQuery);
  
  var url = `${service.getToken().instance_url}/services/data/v47.0/query?q=${encodeURI(taskRaySOQLQuery)}`;
  
  var cache = CacheService.getUserCache();
  var cachedData = cache.get("projects-and-tasks");
  
  // see if we have this in the cache first... unless our data is over 24hrs old.. then refresh.
  if(cachedData != null) {
    console.log('returning cached list of projects');
    return JSON.parse(cachedData); 
  } 
  // Make the HTTP request using a wrapper function that handles expired sessions.
  var response = withRetry(service, function() {
    console.log('fetching new list of projects from Salesforce');
    return UrlFetchApp.fetch(url, {
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken(),
      }
    });
  });

  if(response && response.getResponseCode() == 200) {
    
    var cacheData = response.getContentText();
    cache.put('projects-and-tasks', response.getContentText(), 21600); // longest time we can cache is 6hrs
    
    return JSON.parse(response.getContentText());
  } else {
    console.log(JSON.stringify(response));
    // if we didn't get 200 don't return anything, let it crash... this is unrecoverable.
    console.log(`ERROR FETCHING SALESFORCE RESOURCES : ${response.getResponseCode()}`);
  }
}
/**
* expecting data object with members
*  projectId (required)
*  taskId (required)
*  gcalEventId (required)
*  eventDate (required)
*  hours (required)
*  description (can be null - required member though)
*/
function upsertTime(data) {
  
  var service = getService();
  var url = `${service.getToken().instance_url}/services/data/v47.0/sobjects/TASKRAY__trTaskTime__c/gcal_event_id__c/${data.gcalEventId}`;

  var formData = {
    'TASKRAY__Project__c': data.projectId,
    'TASKRAY__Task__c': data.taskId,
    'TASKRAY__Date__c':  data.eventDate,
    'TASKRAY__Hours__c': data.hours,
    'TASKRAY__Notes__c': data.description,
    'TASKRAY__Owner__c': getUserId()
  };
  console.log(formData);
  
  var response = withRetry(service, function() {
    // patch is an UPSERT operation on Salesforce
    var options = {
      'method' : 'patch',
      'payload' : JSON.stringify(formData),
      'headers': {
        'Authorization': 'Bearer ' + service.getAccessToken(),
        'Content-Type': 'application/json',
      }
    };
    return UrlFetchApp.fetch(url, options);
  });

  console.log(`Response is ${response}`);
  return JSON.parse(response);
}

function refreshProjectList() {
  CacheService.getUserCache().remove('projects-and-tasks');
  fetchProjectsAndTasks();
}

function getUserSettings() {
   var settingsStr = PropertiesService.getUserProperties().getProperty('user_settings');
   var settings = settings = {};
   if(!settingsStr) { // set defaults before returning if not found.
     settings.myTasksOnly = true;
     settings.colourEventsOnLogged = true;
     settings.eventColour = CalendarApp.EventColor.GREEN;
     settings.updateEventNamesOnLogged = false;
     saveUserSettings(settings);
   } else {
     settings = JSON.parse(settingsStr);
   }

   return settings;
}

function saveUserSettings(settingsObj) {
  if (settingsObj)
    PropertiesService.getUserProperties().setProperty('user_settings', JSON.stringify(settingsObj));
}