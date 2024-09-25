
# Automated Report Generation Script 
 
This script automates the creation, scheduling, and notification of reports using Google Apps Script. It provides flexible scheduling options and customizable reminder notifications to streamline your reporting workflow. 
 
## Table of Contents 
 
- Features 
- Getting Started 
 - Prerequisites 
 - Setup Instructions 
- Configuration 
 - Scheduling Settings 
 - Reminder Settings 
 - Slack Integration Settings 
 - Report Settings 
- Script Functions 
- Time-based Triggers 
- Customization and Extensibility 
- License 
 
--- 
 
## Features 
 
1. Flexible Scheduling: Supports various scheduling frequencies, including daily, weekly, every X days, multi-weekly, and monthly reports on specific dates or days. 
2. Customizable Reminders: Allows multiple, user-configurable reminders with custom messages and lead times before the report due date. 
3. Automated Report Creation: Generates reports from a Google Docs template, replaces placeholders with dynamic content, and saves them in a specified Google Drive folder. 
4. Slack Notifications: Sends notifications and reminders to a Slack channel to keep team members informed about report statuses. 
5. Modular and Reusable Code: Refactored codebase with modular design for better maintainability and extensibility. 
 
--- 
 
## Getting Started 
 
### Prerequisites 
 
- A Google account with access to Google Drive and Google Apps Script. 
- A Google Docs template to use for report generation. 
- A Slack workspace with a channel for notifications. 
- Slack Incoming Webhook URL for sending messages to Slack. 
 
### Setup Instructions 
 
1. Copy the Script: 
 - Open Google Apps Script and create a new project. 
 - Copy and paste the script code into the editor. 
 
2. Configure Script Settings: 
 - Update the configuration variables as described in the Configuration section. 
 
3. Set Up Time-based Triggers: 
 - Go to Edit > Current project's triggers or click the clock icon. 
 - Add a trigger for createReportAndNotify function as per your scheduling needs. 
 - Add a trigger for checkAndSendReminders function to run daily (e.g., every hour). 
 
4. Authorize the Script: 
 - Run the createReportAndNotify function manually to prompt authorization. 
 - Grant the necessary permissions when prompted. 
 
--- 
 
## Configuration 
 
Customize the following variables in the script to suit your requirements. 
 
### Scheduling Settings 
 
Define the scheduling frequency and related parameters. 
 ``` javascript 
var SCHEDULE = { 
 frequency: 'weekly', // Options: 'daily', 'weekly', 'everyXDays', 'multiWeekly', 'monthly' 
 dayOfWeek: 1, // For 'weekly' and 'multiWeekly' (0 = Sunday, 1 = Monday, ..., 6 = Saturday) 
 intervalDays: 7, // For 'everyXDays' 
 intervalWeeks: 2, // For 'multiWeekly' 
 dayOfMonth: 15, // For 'monthly' on a specific date (1-31) 
 weekOfMonth: 1, // For 'monthly' on a specific week (1-5) 
 dayOfWeekInMonth: 1, // For 'monthly' on a specific day of the week (0 = Sunday, ..., 6 = Saturday) 
 timeZone: 'GMT', // Time zone for date calculations 
}; 
```
 
- Examples: 
 - Daily: frequency: 'daily' 
 - Weekly on Monday: frequency: 'weekly', dayOfWeek: 1 
 - Every 3 Days: frequency: 'everyXDays', intervalDays: 3 
 - Every 2 Weeks on Friday: frequency: 'multiWeekly', intervalWeeks: 2, dayOfWeek: 5 
 - Monthly on 15th: frequency: 'monthly', dayOfMonth: 15 
 - Monthly on First Monday: frequency: 'monthly', weekOfMonth: 1, dayOfWeekInMonth: 1 
 
### Reminder Settings 
 
Customize reminder messages and schedules. 
``` javascript 
var REMINDERS = [ 
 { 
 daysBeforeDue: 2, // Days before the report due date 
 timeOfDay: '09:00', // Time to send the reminder (HH:mm in 24-hour format) 
 message: "Reminder: The report is due on DUE_DATE. Please add your updates: <FILEURL|Report Link>", 
 }, 
 { 
 daysBeforeDue: 0, 
 timeOfDay: '12:00', 
 message: "Today is the due date for the report (DUE_DATE). Please finalize your inputs: <FILEURL|Report Link>", 
 }, 
]; 
```
 
- Placeholders in Messages: 
 - DUE_DATE: Replaced with the report due date. 
 - FILEURL: Replaced with the URL of the report document. 
 
### Slack Integration Settings 
 
Configure Slack notifications. 
``` javascript 
var ENABLE_SLACK_NOTIFICATIONS = true; // Set to 'false' to disable Slack notifications 
var WEBHOOKURI = "https://hooks.slack.com/services/your/webhook/url"; // Your Slack Incoming Webhook URL 
var BOTNAME = "report_bot"; // Name displayed for the bot in Slack 
var SLACKCHANNEL = "#your-channel"; // Slack channel to post messages in 
```
 
### Report Settings 
 
Set up report generation details. 
``` javascript 
var TEMPLATEFILEID = "your-template-file-id"; // Google Doc ID of your report template 
var REPORTFOLDER = "your-report-folder-id"; // Google Drive folder ID to save reports 
var FILENAMEPREFIX = "Automated Report - "; // Prefix for generated report names 
var DATEPLACEHOLDER = "XDATEX"; // Placeholder text in the template to be replaced with the date 
var LEAD_TIME_DAYS = 2; // Days before due date to create the report 
var CREATIONMESSAGE = "A new report has been created for DUE_DATE. Please start contributing: <FILEURL|Report Link>"; // Message when report is created 
```
 
--- 
 
## Script Functions 
 
### nextReportDate() 
 
Calculates the next report date based on the scheduling settings. 
``` javascript 
function nextReportDate() { 
 // Implementation based on SCHEDULE settings 
} 
```
 
### createReportAndNotify() 
 
Creates the report and sends an initial notification. 
``` javascript 
function createReportAndNotify() { 
 // Checks if it's time to create the report 
 // Generates the report 
 // Sends a Slack message 
} 
```
 
### createReport(reportDate) 
 
Generates the report document from the template. 
``` javascript 
function createReport(reportDate) { 
 // Copies the template 
 // Replaces placeholders 
 // Saves the report in the specified folder 
} 
```
 
### checkAndSendReminders() 
 
Checks if reminders need to be sent and sends them. 
``` javascript 
function checkAndSendReminders() { 
 // Calculates days until due date 
 // Sends reminders based on REMINDERS configuration 
} 
```
 
### sendSlackMessage(message) 
 
Sends a message to Slack. 
``` javascript 
function sendSlackMessage(message) { 
 // Constructs and sends the Slack message 
} 
```
 
--- 
 
## Time-based Triggers 
 
Set up triggers to automate script execution. 
 
### Creating Triggers 
 
1. Open your Google Apps Script project. 
2. Click on the Triggers icon (clock symbol) or navigate to Edit > Current project's triggers. 
3. Click Add Trigger. 
 
### Trigger for Report Creation 
 
- Function: createReportAndNotify 
- Select event source: Time-driven 
- Type of time-based trigger: Day timer 
- Select time of day: Choose a time that aligns with your LEAD_TIME_DAYS 
 
### Trigger for Reminders 
 
- Function: checkAndSendReminders 
- Select event source: Time-driven 
- Type of time-based trigger: Hour timer 
- Select hour interval: Every hour or as needed based on your earliest timeOfDay in REMINDERS 
 
--- 
 
## Customization and Extensibility 
 
- Adding More Reminders: Append additional objects to the REMINDERS array with custom daysBeforeDue, timeOfDay, and message. 
- Changing the Time Zone: Update SCHEDULE.timeZone to match your preferred time zone (e.g., "America/New_York"). 
- Modifying Placeholders: Ensure consistency in placeholders (DUE_DATE, FILEURL) across messages and replace them accordingly in the script functions. 
- Extending Functionality: Use the modular structure to add new features, such as email notifications or integration with other services. 

--- 
 
Note: Ensure that all IDs (template file ID, folder ID) and URLs are correctly set. Test the script thoroughly to confirm that scheduling and reminders work as expected. 

--- 
 
## License 
 
This project is licensed under the MIT License.
