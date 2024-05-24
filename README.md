# Weekly Report Automation Script

This script automates the creation and notification of a weekly report using Google Apps Script.

## Features

1. **Creates a Weekly Report**: The script generates a new Google Doc from a template each week.
2. **Replaces Placeholder Text**: The script replaces a placeholder text in the template with the current date.
3. **Saves Report**: The generated report is saved in a specified Google Drive folder.
4. **Slack Notifications**: Optionally, the script sends notifications to a Slack channel when the report is created and on the due date as a reminder.

## Setup Instructions

### Required Variables

1. **WEBHOOKURI**: Replace with your Slack webhook URL (optional).
   ```javascript
   var WEBHOOKURI = "https://hooks.slack.com/services/your/webhook/url";
2. **BOTNAME**: Customize the bot name for Slack notifications (optional).
   ```javascript
   var BOTNAME = "weekly_update_bot";
3. **SLACKCHANNEL**: Specify your Slack channel name (optional).
   ```javascript
   var SLACKCHANNEL = "#your-channel";

4. **ENABLE_SLACK_NOTIFICATIONS**: Set to `true` to enable Slack notifications, `false` to disable them.
   ```javascript
   var ENABLE_SLACK_NOTIFICATIONS = true;
5. **DAYOFWEEK**: Set the day of the week for the report (0 = Sunday, 1 = Monday, etc.).
   ```javascript
   var DAYOFWEEK = 1; // Monday
6. **TEMPLATEFILEID**: Replace with your Google Doc template ID.
   ```javascript
   var TEMPLATEFILEID = "your-template-file-id";
7. **FILENAMEPREFIX**: Customize the prefix for generated report names.
    ```javascript
    var FILENAMEPREFIX = "Weekly Update - ";
8. **DATEPLACEHOLDER**: Specify the placeholder text in the template to be replaced with the date.
   ```javascript
   var DATEPLACEHOLDER = "XDATEX";
9. **REPORTFOLDER**: Replace with your Google Drive folder ID.
   ```javascript
   var REPORTFOLDER = "your-report-folder-id";
10. **CREATIONMESSAGE**: Customize the creation message for Slack notifications.
    ```javascript
    var CREATIONMESSAGE = "Hey folks, the <FILEURL|Weekly Update> for TIMESTAMP has been generated, please fill it in by end of day Monday";
    ```
11. **REMINDERMESSAGE**: Customize the reminder message for Slack notifications.
    ```javascript
    var REMINDERMESSAGE = "Last reminder! Remember to fill out the <FILEURL|Weekly Update> before lunch!";
    ```

### Time-based Triggers

Set up time-based triggers in Google Apps Script to call the `createReportAndNotify` and `reminder` functions as needed. 

1. Open your Google Apps Script project.
2. Click on the clock icon in the toolbar to open the Triggers page.
3. Click on "Add Trigger" in the bottom right corner.
4. Set up a trigger for `createReportAndNotify` to run on the desired day and time before the report is due (e.g., every Friday).
5. Set up a trigger for `reminder` to run on the day the report is due (e.g., every Monday morning).

## Script Functions

### `nextReportDate`

Calculates the next report date based on the specified day of the week.

```javascript
function nextReportDate() {
  var d = new Date();
  
  if (d.getDay() === DAYOFWEEK) {
    return d;
  }
  
  d.setDate(d.getDate() + (1 + 7 - d.getDay()) % 7);
  return d;
}
```

### `slack`

Sends a message to Slack (optional).

```javascript
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
```

### `reminder`

Sends a reminder to fill out the report.

```javascript
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
```


## License

This project is licensed under the MIT License.