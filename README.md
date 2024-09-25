# Automated Report Generation and Notification Script

This script automates the creation, scheduling, and notification of reports using Google Apps Script. It includes enhanced error handling and customizable Slack notifications.

## Features

1. **Flexible Scheduling**: Supports various scheduling frequencies, including daily, weekly, every X days, multi-weekly, and monthly.
2. **Customizable Reminders**: Allows multiple, user-configurable reminders with custom messages and lead times.
3. **Automated Report Creation**: Generates reports from a Google Docs template, replaces placeholders with dynamic content, and saves them in a specified Google Drive folder.
4. **Slack Notifications**: Sends notifications and reminders to a Slack channel, and optionally sends error notifications in case of failures.
5. **Error Handling**: Integrated error handling via `try...catch` with logging and optional error notifications through Slack.

## Prerequisites

- A Google account with access to Google Drive and Google Apps Script.
- A Google Docs template to use for report generation.
- Slack Incoming Webhook URL for sending messages.
- Create a Slack channel for notifications (optional).

## Setup Instructions

### 1. **Copy the Script**

- Open [Google Apps Script](https://script.google.com/) and create a new project.
- Copy the script into the project editor.

### 2. **Configure Script Settings**

- **Slack Integration Settings**:  
   Set the following variables in the script to configure Slack notifications:
   ```javascript
   var ENABLE_SLACK_NOTIFICATIONS = true; // Enable or disable Slack notifications.
   var WEBHOOKURI = "https://hooks.slack.com/services/your/webhook/url"; // Your Slack Incoming Webhook URL.
   var BOTNAME = "weekly_update_bot"; // Name displayed for the bot in Slack.
   var SLACKCHANNEL = "#your-channel"; // Slack channel to post messages in.
   var ENABLE_ERROR_NOTIFICATIONS = true; // Enable or disable error notifications via Slack.
   ```

- **Report Settings**:  
   Set up the following variables to configure report creation:
   ```javascript
   var TEMPLATEFILEID = "your-template-file-id"; // Google Doc ID of your report template.
   var REPORTFOLDER = "your-report-folder-id"; // Google Drive folder ID to save reports.
   var FILENAMEPREFIX = "Weekly Update - "; // Prefix for generated report names.
   var DATEPLACEHOLDER = "XDATEX"; // Placeholder text in the template to be replaced with the date.
   var LEAD_TIME_DAYS = 2; // Number of days before due date to create the report.
   var CREATIONMESSAGE = "A new report has been created for DUE_DATE. Please start contributing: <FILEURL|Report Link>"; // Slack message when report is created.
   ```

- **Scheduling Settings**:  
   Set the report scheduling frequency and parameters:
   ```javascript
   var SCHEDULE = {
     frequency: 'multiWeekly', // Options: 'daily', 'weekly', 'everyXDays', 'multiWeekly', 'monthly'.
     dayOfWeek: 1, // For 'weekly' and 'multiWeekly' (0 = Sunday, 1 = Monday, ..., 6 = Saturday).
     intervalDays: 7, // For 'everyXDays'.
     intervalWeeks: 2, // For 'multiWeekly'.
     dayOfMonth: 15, // For 'monthly' on a specific date (1-31).
     weekOfMonth: 2, // For 'monthly' on a specific week (1-5).
     dayOfWeekInMonth: 4, // For 'monthly' on a specific day of the week (0 = Sunday, ..., 6 = Saturday).
     timeZone: 'GMT' // Time zone for date calculations.
   };
   ```

- **Reminder Settings**:  
   Configure reminder messages and lead times:
   ```javascript
   var REMINDERS = [
     {
       daysBeforeDue: 7, // Days before the report due date.
       timeOfDay: '09:00', // Time of day to send the reminder (24-hour format).
       message: "The report is due in a week on DUE_DATE. Please start preparing."
     },
     {
       daysBeforeDue: 1,
       timeOfDay: '09:00',
       message: "Reminder: The report is due tomorrow (DUE_DATE)."
     },
     {
       daysBeforeDue: 0,
       timeOfDay: '09:00',
       message: "Today is the due date for the report (DUE_DATE). Please finalize your inputs."
     }
   ];
   ```

### 3. **Set Up Triggers**

- **Trigger for Report Creation**:  
   Go to **Edit > Current project's triggers** and set a time-driven trigger for the `createReportAndNotify()` function to run at the desired time.

- **Trigger for Reminders**:  
   Add another time-driven trigger for the `checkAndSendReminders()` function. It’s recommended to run this at least hourly to ensure timely reminders.

## Error Handling and Notifications

- The script includes detailed error handling using `try...catch` blocks. If an error occurs, it will be logged using `Logger.log()`, and optionally, an error message will be sent to Slack.

- **Enable Error Notifications**:  
   To receive error notifications via Slack, set `ENABLE_ERROR_NOTIFICATIONS` to `true` in the script:
   ```javascript
   var ENABLE_ERROR_NOTIFICATIONS = true; // Set to 'true' to receive error notifications.
   ```

- **Error Logging**:  
   Errors are logged to the Google Apps Script log (`View > Logs`), and you can add custom error logging mechanisms (e.g., email, external service) as needed.

## Customization and Extensibility

- **Customizing Slack Messages**:  
   You can modify the Slack messages for report creation, reminders, and errors by editing the corresponding message templates in the script.

- **Adding More Reminders**:  
   You can add more reminders to the `REMINDERS` array, specifying different lead times and custom messages as needed.

- **Extending Error Logging**:  
   If you need more advanced error logging (e.g., storing errors in Google Sheets or sending emails), you can extend the `logError()` function in the script to include these capabilities.

## License

This project is licensed under the MIT License.

---

**Note**: Ensure that all IDs (template file ID, folder ID) and URLs are correctly set. Test the script thoroughly to confirm that scheduling, reminders, and error handling work as expected.
