
/**
* Timesheeter G-Suite Calendar addon
*/

/**
* Header and subtitles never change
*/
const headerTitle = "SA Timesheet Logger";
const headerSubTitle = "Easily log your hours to TaskRay"
// the header is common accross the entire app
var header = CardService.newCardHeader().setTitle(headerTitle).setSubtitle(headerSubTitle);

// if we want to add a logout button when building cards, use this one
var authorisedFooter = CardService.newFixedFooter().setPrimaryButton(CardService.newTextButton().setText("Logout of Salesforce").setOnClickAction(CardService.newAction().setFunctionName('logout')));
var service = getService();

// if we want to add a login button, well... 
var unauthorisedFooter = CardService.newFixedFooter().setPrimaryButton(CardService.newTextButton().setText("Login to Salesforce").setAuthorizationAction(CardService.newAuthorizationAction().setAuthorizationUrl(service.getAuthorizationUrl())));
 
/**
 * Callback for rendering the card for a specific Calendar event.
 * @param {Object} e The event object.
 * @return {CardService.Card} The card to show to the user.
 */
function onCalendarEventOpen(e) {
  
  console.log(e);
  var calendar = CalendarApp.getCalendarById(e.calendar.calendarId);
  var event = calendar.getEventById(e.calendar.id);
  // this variable is used to identify calendar events that are recurring. 
  // Google appends the ISO timestamp of the starttime onto the event.getId() value to distinguish between events on different days
  // we want to store this event Id in Salesforce, rather than the recurring event Id so we get consistent behaviour across days.
  var uniqueEventId = e.calendar.id;

  // array to store the card sections we create dynamically.
  var sections = [];
  // we are going to add a footer depending on whether the user is authenticated or not.. 
  var fixedFooter = undefined;
  
  // instantiate the card builder and set the common header
  var builder = CardService.newCardBuilder();
  builder.setHeader(header);
  
  if (!event) { // a new event is being created.. not for us to worry about
     sections.push(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText('A new event needs to be saved before we can log this as time. Focus on your planning at the moment, come back here later :)'))); 
  } 
   // we have an event we want to log.. 
  else {
    // lets check we have a credential for Salesforce. if so, let's build out the form fields for logging time
    if (service.hasAccess()) { 
      var eventSalesforceInformation = getEventSalesforceInformation(event, uniqueEventId);
      
      sections.push(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText(eventSalesforceInformation.message)));

      var buildTaskSuggestionsAction = CardService.newAction().setFunctionName('createTaskSuggestions');
      var logTimeAction =  CardService.newAction().setFunctionName('upsertInformation').setParameters(eventSalesforceInformation);
      var logTimeActionMessage = eventSalesforceInformation.isUpdate == "true" ? 'Update Time' : 'Log Time';
      var eventStartTime = event.getStartTime().getTime();
      
      var timeInputWidget = CardService.newTextInput().setFieldName("hours").setTitle("Hours Logged").setValue(eventSalesforceInformation.hours);
      
      var eventDescription = event.getTitle();
      
      sections.push(CardService.newCardSection().addWidget(CardService.newTextInput().setSuggestions(createProjectSuggestions()).setFieldName("project").setTitle("Project"))
                  .addWidget(CardService.newTextInput().setFieldName("task").setTitle("Task").setHint('Select your project first and see a list of tasks').setSuggestionsAction(buildTaskSuggestionsAction)) 
                   .addWidget(CardService.newTextInput().setFieldName("description").setTitle("Description").setMultiline(true).setValue(eventDescription))
                    .addWidget(timeInputWidget)
                   .addWidget(CardService.newTextButton().setText(logTimeActionMessage).setOnClickAction(logTimeAction).setTextButtonStyle(CardService.TextButtonStyle.FILLED).setBackgroundColor('#12424a')));
                   // .addWidget(CardService.newImageButton().setIconUrl('https://segment.com/site-assets/images/favicon/favicon-16x16.png').setOnClickAction(logTimeAction).setAltText(logTimeActionMessage)));
      builder.setFixedFooter(authorisedFooter);
    } 
    // otherwise, we probably want to set the login button in the footer.
    else { 
      sections.push(new CardService.newCardSection().addWidget(new CardService.newTextParagraph().setText('You will need to be logged into Salesforce to fetch TaskRay projects and log time. Look for the authorization button in the footer')));
      builder.setFixedFooter(unauthorisedFooter);
    } 
  }
  // add the sections to the builder to create our card.. 
  for(var i=0 ; i<sections.length ; i++) {
    builder.addSection(sections[i]); 
  }
  // and build
  return builder.build();
}

function updateTasks(e, taskSelectionItem) {
   
  var project = e && e.formInput['project'];
  var records = fetchProjectsAndTasks().records;
  
  var suggestions = [];
  for (r in records) {
    if(records[r].Name == project) {
      var tasks = records[r].TASKRAY__Tasks__r.records;
      for(var t in tasks) {
        taskSelectionInput.addItem(tasks[t].Name);  
      }
      //suggestions.sort();
    }
  }
  
}

/**
* Callback that is rendered when no calendar object is selected or opened.
* This could be used as a settings / configuration page if we need it
*/
function onHomepageOpen(e) {
  console.log(e);
  var section = undefined;
  var fixedFooter = undefined;
  var action = undefined;
  var sections = [];
  var settings = getUserSettings();
  
  var builder = CardService.newCardBuilder();
  builder.setHeader(header);
  
  var versionText = CardService.newTextParagraph()
            .setText(`You are currently using the ${getUserSettings().version} version of Timesheeter. If this is not you (i.e you need to use Employee instead of Community) you can change your configuration in the drop down menu at the top of this addon`);
  
  var handleOptionsAction = CardService.newAction()
            .setFunctionName("handleSettingsChange");
  
  var myTasksToggle = CardService.newKeyValue()
    .setTopLabel('When loading project tasks')
    .setContent("Only show my project tasks")
    .setSwitch(CardService.newSwitch()
        .setFieldName("settings_myTasks")
        .setValue('my_project_tasks')
        .setSelected(settings.myTasksOnly)
        .setOnChangeAction(handleOptionsAction));
  
  var colourEventsToggle = CardService.newKeyValue()
     .setTopLabel('When creating time entry')
     .setContent("Change colour of calendar")
     .setSwitch(CardService.newSwitch()
                .setFieldName("settings_colourEventsOnLogged")
                .setValue('set_colour')
                .setSelected(settings.colourEventsOnLogged)
                .setOnChangeAction(handleOptionsAction));
  
  var colourEventsColourSelector = CardService.newTextInput()
     .setSuggestions(CardService.newSuggestions().addSuggestions(Object.keys(CalendarApp.EventColor)))
     .setFieldName("settings_eventColour")
     .setTitle("Event Colour")
     .setOnChangeAction(handleOptionsAction)
     .setValue(getColourKey(settings.eventColour))
     .setHint('Clear the default value to see all options available');
  
  var updateEventNameWhenLogged = CardService.newKeyValue()
  .setTopLabel('When creating time entry')
     .setContent("Append 'logged' to event")
     .setSwitch(CardService.newSwitch()
                .setFieldName("settings_updateEventNameOnLogged")
                .setValue('set_append')
                .setSelected(settings.updateEventNamesOnLogged)
                .setOnChangeAction(handleOptionsAction));
  
  // if we have access to Salesforce, get project list and display the logout footer
  if(service.hasAccess()) {   
    fetchProjectsAndTasks();
    sections.push(new CardService.newCardSection().setHeader('User Settings')
                  .addWidget(versionText)
                  .addWidget(myTasksToggle)
                  .addWidget(colourEventsToggle)
                  .addWidget(updateEventNameWhenLogged)
                  .addWidget(colourEventsColourSelector));
    sections.push(CardService.newCardSection().setHeader('Instructions')
                  .addWidget(CardService.newTextParagraph()
                             .setText('Click on the calendar event you wish to log as time in TaskRay.<br /><br />When you are in an event detail page this window will show you options to log time against a selected project in TaskRay.<br /><br />You can edit TaskRay time entries you have already logged by selecting the calendar event again, you will see the buttons change to <bold>UPDATE</bold> to reflect this.')));

                  
    builder.setFixedFooter(authorisedFooter);
  } 
  // else display the login footer.
  else {
    section = CardService.newCardSection().addWidget(CardService.newTextParagraph().setText('You must first authorize with your Salesforce credentials'));
    builder.setFixedFooter(unauthorisedFooter);
  }
  
  
    // add the sections to the builder to create our card.. 
  for(var i=0 ; i<sections.length ; i++) {
    builder.addSection(sections[i]); 
  }
  builder.setName('home');
  // and build
  return builder.build();
}

function getColourKey(colourValue) {
 
  var keys = Object.keys(CalendarApp.EventColor);
  for(var i=0 ; i<keys.length ; i++) {
    var key = keys[i];
    if(CalendarApp.EventColor[key] == colourValue) {
      return key; 
    }
  }
  return '';
}

/**
*  Function handles when settings are toggled. 
*
*/
function handleSettingsChange(e) {
  var settings = getUserSettings();
  // if either of these are changed we should refresh project list
  var tasksSettings = settings.myTasksOnly;
  var projectSettings = settings.myProjectsOnly;
  
  if (e && e.formInput['settings_myTasks']) {
    settings.myTasksOnly = true;  
  } else {
    settings.myTasksOnly = false;
  }
  
  if (e && e.formInput['settings_myProjects']) {
    settings.myProjectsOnly = true;  
  } else {
    settings.myProjectsOnly = false;
  }
  
  if (e && e.formInput['settings_colourEventsOnLogged']) {
    settings.colourEventsOnLogged  = true;
  } else {
    settings.colourEventsOnLogged = false;
  }
  
  if (e && e.formInput['settings_updateEventNameOnLogged']){
   settings.updateEventNamesOnLogged = true; 
  } else {
    settings.updateEventNamesOnLogged = false; 
  }
  if (e && e.formInput['settings_eventColour']) {
    settings.eventColour = CalendarApp.EventColor[e.formInput['settings_eventColour']];  
  } else {
    settings.eventColour = CalendarApp.EventColor.GREEN;  
  }
  // update settings for the user.
  saveUserSettings(settings);
  
  // if either settings related to the project list have changed, refresh the project list from Salesforce
  if (!(tasksSettings == settings.myTasksOnly) || !(projectSettings == settings.myProjectsOnly)) {
   refreshProjectList();  
  }
  
  return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText('Updated settings')).setStateChanged(true).build();
}


// check if the calendarId for this event has already been 
// updated into Salesforce.
// The uniqueEventId parameter handles the condition that we might be dealing with 
// a recurring meeting. 
function getEventSalesforceInformation(event, uniqueEventId) {
 
  var returnValue = {};

  //function checks if we have a time record against this calendar id
  var taskRayTime = haveWeLoggedThisTimeAlready(uniqueEventId); 
  
  returnValue.eventDate = event.getStartTime().toISOString();
  returnValue.gcalEventId = uniqueEventId;
  returnValue.hours = ((event.getEndTime() - event.getStartTime()) / (1000 * 60 * 60)).toFixed(2);
  
  if (taskRayTime.totalSize == 0){
    returnValue.isUpdate = "false";
    returnValue.message = `The event \"${event.getTitle()}\" will log ${returnValue.hours} hours against the selected project`;    
  } else {
    returnValue.isUpdate = "true";
    returnValue.message = `The event \"${event.getTitle()}\" will update ${returnValue.hours} hours against an existing task`;    
  }

  return returnValue;
}

/**
* Callback that logs out of Salesforce. Returns the homepage event, this might a bug?
*/
function logout(e) {
  service.reset();
  return onHomepageOpen(e);
}
/**
* Simple function to round numbers 
*/
function round_to_precision(x, precision) {
  var num = Math.round(x * (10));
  return num / precision * 10;
}

/**
* figures out if we are updating an existing time record in task ray or creating a new one.. 
* then inserts the record into Salesforce accordingly
*/
function upsertInformation(e) {

  var notification = 'TODO';
  var project = e && e.formInput['project'];
  var task = e && e.formInput['task'];
  var hours = e && e.formInput['hours'];
  
  var settings = getUserSettings();
  
  if(project && task) {
    // calls Salesforce (or cache) and looks for project list
    var records = fetchProjectsAndTasks().records;
    
    for (r in records) {
      if(records[r].Name == project) {
        e.parameters.projectId = records[r].Id;
        var tasks = records[r].TASKRAY__Tasks__r.records;
        for(var t in tasks) {
          if(tasks[t].Name == task) {
            e.parameters.taskId = tasks[t].Id;  
          }
        }
      }
    }
    if(e.parameters.projectId && e.parameters.taskId && hours) {

      e.parameters.description = e && e.formInput['description'];
      e.parameters.hours = hours;
      
      var response = upsertTime(e.parameters);
      
      if (response != undefined && response.success) {
        var calendar = CalendarApp.getCalendarById(e.calendar.calendarId);
        var event = calendar.getEventById(e.calendar.id);
        
        if(settings.colourEventsOnLogged) {
          event.setColor(settings.eventColour); 
        }
        if (settings.updateEventNamesOnLogged) {
          event.setTitle(event.getTitle() + '-logged'); 
        }
        notification = `You have successfully ${response.created ? "created" : "updated"} a time entry in TaskRay`;
      } else {
        notification = `We could not update the TaskRay Time record. \n${response}`; 
      }
    } else {
      notfication = 'We are unable to find the required Salesforce identifiers for the Project and Task selected... this is unexpected';  
    }
  } else {
    notification = 'Both Project and Task fields are required fields';
  } 
  
  return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText(notification)).setStateChanged(true).build();
}

/**
* loop through projectsAndTasks and pull project names for suggestion list
*/
function createProjectSuggestions() {
  var suggestions = CardService.newSuggestions();
  var projectsAndTasks = fetchProjectsAndTasks(); 
  
  var records = projectsAndTasks.records;
  for(var i in records) {
    suggestions.addSuggestion(records[i].Name);
  }
   return suggestions;
}

/**
* loop through and create task list for selected project
*/
function createTaskSuggestions(e) {

  var project = e && e.formInput['project'];
  var records = fetchProjectsAndTasks().records;
  var suggestions = [];
  
  for (r in records) {
    if(records[r].Name == project) {
      if(records[r] && records[r].TASKRAY__Tasks__r.records) {
        var tasks = records[r].TASKRAY__Tasks__r.records;
        for(var t in tasks) {
          suggestions.push(tasks[t].Name);  
        }
        suggestions.sort();
      }
    }
  }
  return CardService.newSuggestionsResponseBuilder().setSuggestions(CardService.newSuggestions().addSuggestions(suggestions)).build();  
}

