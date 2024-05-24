/**
 * This script automates the creation and notification of a weekly report using Google Apps Script.
 * 
 * Features:
 * 1. Creates a new Google Doc from a specified template each week.
 * 2. Replaces a placeholder text in the template with the current date.
 * 3. Saves the generated report in a specified Google Drive folder.
 * 4. Optionally sends notifications to a Slack channel when the report is created and a reminder on the due date.
 * 
 * Setup Instructions:
 * 1. Fill in the required variables below with your own values.
 * 2. Set up time-based triggers in Google Apps Script to call `createReportAndNotify` and `reminder` functions as needed.
 */

// The Slack webhook URL (optional)
var WEBHOOKURI = "https://hooks.slack.com/services/your/webhook/url";
// What the bot will be called when it posts in Slack (optional)
var BOTNAME = "weekly_update_bot";
// The channel you want to post in (optional)
var SLACKCHANNEL = "#your-channel";
// Set this to 'true' to enable Slack notifications, 'false' to disable them
var ENABLE_SLACK_NOTIFICATIONS = true;

// The day of the week that the report gets filled in (0 = Sunday, 1 = Monday, ..., 6 = Saturday)
var DAYOFWEEK = 1; // Monday
// The Google Doc ID for the file you want to use as a template
var TEMPLATEFILEID = "your-template-file-id";
// The prefix to use for the name of each generated report (it will have the report date added to the name)
var FILENAMEPREFIX = "Seeto Weekly Update - ";
// The text you have in the doc that will be replaced with a generated date string
var DATEPLACEHOLDER = "XDATEX";
// The Google Drive ID for the folder the template is in and that the reports will be added to
var REPORTFOLDER = "your-report-folder-id";
// The message you want to send out when the report is created, a few days ahead of when it needs to be filled out. 
// FILEURL and TIMESTAMP are required unless you want to change the code further down!
var CREATIONMESSAGE = "Hey folks, the <FILEURL|Weekly Update> for TIMESTAMP has been generated, please fill it in by end of day Monday";
// The message to send as a reminder. FILEURL is required unless you want to change the code further down!
var REMINDERMESSAGE = "Last reminder! Remember to fill out the <FILEURL|Weekly Update> before lunch!";

/**
 * Function to calculate the next report date based on the specified day of the week.
 * @return {Date} The next report date.
 */
function nextReportDate() {
  var d = new Date();
  
  if (d.getDay() === DAYOFWEEK) {
    return d;
  }
  
  d.setDate(d.getDate() + (1 + 7 - d.getDay()) % 7);
  return d;
}

/**
 * Function to send a message to Slack.
 * @param {string} message The message to send.
 */
function slack(message) {
  if (!ENABLE_SLACK_NOTIFICATIONS) return;

  var payload = {
     "channel": SLACKCHANNEL,
     "username": BOTNAME,
     "text": message
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };
  
  UrlFetchApp.fetch(WEBHOOKURI, options);
}

/**
 * Function to create a report and send a notification.
 * Set a trigger to call this function ahead of the next report date (e.g. on a Friday for a report due on a Monday).
 */
function createReportAndNotify() {
  // Generate a timestamp
  var timestamp = Utilities.formatDate(nextReportDate(), "GMT", "dd/MM/yyyy");
  var nicedate = Utilities.formatDate(nextReportDate(), "GMT", "d MMMM yyyy");
  var iso8061 = Utilities.formatDate(nextReportDate(), "GMT", "yyyy-MM-dd");
  
  // Copy the template file
  var docid = DriveApp.getFileById(TEMPLATEFILEID).makeCopy().getId();
  var doc = DocumentApp.openById(docid);
  doc.setName(FILENAMEPREFIX + iso8061);

  // Replace the date placeholder with the formatted date
  var body = doc.getActiveSection();
  body.replaceText(DATEPLACEHOLDER, nicedate);
  doc.saveAndClose();
  
  // Add the file to the specified folder
  var file = DriveApp.getFileById(docid);
  var folder = DriveApp.getFolderById(REPORTFOLDER);
  folder.addFile(file);
    
  // Notify people, if Slack notifications are enabled
  if (ENABLE_SLACK_NOTIFICATIONS) {
    var message = CREATIONMESSAGE.replace("FILEURL", doc.getUrl()).replace("TIMESTAMP", timestamp);
    slack(message);
  }
}

/**
 * Function to send a reminder to fill out the report.
 * Set a trigger to call this function on the day the report needs to be completed.
 */
function reminder() {
  var iso8061 = Utilities.formatDate(nextReportDate(), "GMT", "yyyy-MM-dd");
  var name = FILENAMEPREFIX + iso8061;
  var docId = DriveApp.getFilesByName(name).next().getId();
  var doc = DocumentApp.openById(docId);

  // Notify people, if Slack notifications are enabled
  if (ENABLE_SLACK_NOTIFICATIONS) {
    var message = REMINDERMESSAGE.replace("FILEURL", doc.getUrl());
    slack(message);
  }
}
