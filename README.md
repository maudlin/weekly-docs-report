# Automated Report Generation and Notification Script

This script automates the creation, scheduling, and notification of reports using Google Apps Script. It includes options for Slack and email notifications, flexible scheduling, and automatic reminders.

## Features

1. **Flexible Scheduling**: Supports various scheduling frequencies, including daily, weekly, and monthly reports.
2. **Customizable Notifications**: Choose between Slack and email notifications, or use both. Notifications can be sent to a specific person or distribution group.
3. **Automated Report Creation**: Generates reports from a Google Docs template, replaces placeholders with dynamic content, and saves them in a specified Google Drive folder.
4. **Duplicate Prevention**: Skips report creation if a report for the same date already exists.
5. **Reminder System**: Sends reminders based on days before the report is due. You can use Google Apps Script triggers to set when the reminders should be sent.
6. **Error Handling**: Logs errors and optionally sends error notifications via Slack.

## Prerequisites

1. A Google account with access to Google Drive and Google Apps Script.
2. A Google Docs template to use for report generation.
3. Slack Incoming Webhook URL (optional) for sending messages.
4. Email addresses for notifications (optional).

## Setup Instructions

### 1. **Copy the Script**

- Open [Google Apps Script](https://script.google.com/) and create a new project.
- Copy the script into the project editor.

### 2. **Configure Script Settings**

- **Notification Settings**:  
  Set the following variables to configure Slack and email notifications:
  ```javascript
  var NOTIFICATIONS = {
    slack: true,  // Enable or disable Slack notifications
    email: true,  // Enable or disable Email notifications
    emailRecipient: "team@example.com"  // Email address for notifications
  };
  ```

- **Slack Integration Settings** (if Slack is enabled):  
  ```javascript
  var WEBHOOKURI = "https://hooks.slack.com/services/your/webhook/url";  // Your Slack Incoming Webhook URL
  var BOTNAME = "weekly_update_bot";  // Name displayed for the bot in Slack
  var SLACKCHANNEL = "#your-channel";  // Slack channel to post messages in
  ```

- **Report Settings**:  
  Configure the following variables to specify how reports are generated:
  ```javascript
  var TEMPLATEFILEID = "your-template-file-id";  // Google Doc ID of your report template
  var REPORTFOLDER = "your-report-folder-id";  // Google Drive folder ID to save reports
  var FILENAMEPREFIX = "Weekly Update - ";  // Prefix for generated report names
  var DATEPLACEHOLDER = "XDATEX";  // Placeholder text in the template to be replaced with the date
  var LEAD_TIME_DAYS = 3;  // Number of days before the due date to create the report
  var CREATIONMESSAGE = "A new report has been created for DUE_DATE. Please start contributing: <FILEURL|Report Link>";  // Message when report is created
  ```

- **Scheduling Settings**:  
  Configure scheduling based on the user's desired repeat frequency and the start date for the first report:
  ```javascript
  var SCHEDULE = {
    startDate: new Date('2024-10-18'),  // Absolute start date (when the first report is due)
    repeatType: 'weekly',  // Options: 'daily', 'weekly', 'monthly'
    interval: 2,  // Interval for weekly or daily frequency (e.g., every 2 days or 2 weeks)
    timeZone: 'America/New_York'  // The user's time zone
  };
  ```

- **Reminder Settings**:  
  Configure the reminders that are sent days before the report is due:
  ```javascript
  var REMINDERS = [
    {
      daysBeforeDue: 2,  // Days before the report is due
      message: "The report is due in a week on DUE_DATE. Please start preparing."
    },
    {
      daysBeforeDue: 1,
      message: "Reminder: The report is due tomorrow (DUE_DATE). Don't forget to complete your sections."
    },
    {
      daysBeforeDue: 0,
      message: "Today is the due date for the report (DUE_DATE). Please finalize your inputs."
    }
  ];
  ```

### 3. **Set Up Triggers**

- **Trigger for Report Creation**:  
  Go to **Edit > Current project's triggers** and set a time-driven trigger for the `createReportAndNotify()` function to run at the desired time (e.g., daily).

- **Trigger for Reminders**:  
  Set up a time-driven trigger for the `checkAndSendReminders()` function to run as frequently as needed (e.g., hourly or daily). This allows reminders to be sent based on the days before the report is due.

## Duplicate Prevention

- The script checks whether a report with the same name (based on the date) already exists in the target folder before creating a new report.
- If a report already exists, the creation is skipped, and no Slack or email notifications are sent for that report.

## Error Handling and Notifications

- **Error Logging**:  
  Errors are logged using `Logger.log()` and can optionally be sent as Slack messages if `ENABLE_ERROR_NOTIFICATIONS` is set to `true`.

- **Enable Error Notifications**:  
  To receive error notifications via Slack, set `ENABLE_ERROR_NOTIFICATIONS` to `true` in the script:
  ```javascript
  var ENABLE_ERROR_NOTIFICATIONS = true;  // Set to 'true' to receive error notifications.
  ```

- **Customizing Notifications**:  
  You can modify the Slack and email notification messages by editing the corresponding message templates in the script.

## Customization and Extensibility

- **Adding More Reminders**:  
  You can add more reminders to the `REMINDERS` array by specifying different lead times and custom messages.

- **Extending Error Logging**:  
  You can extend the `Utils.logError()` function to log errors in additional ways, such as sending emails, logging to a Google Sheet, or using an external service.

## License

This project is licensed under the MIT License.

---

**Note**: Ensure that all IDs (template file ID, folder ID) and URLs are correctly set. Test the script thoroughly to confirm that scheduling, reminders, and error handling work as expected.

