/**
 * This script automates the creation and notification of reports using Google Apps Script.
 */

// Notification Settings
var NOTIFICATIONS = {
  slack: false,  // Enable or disable Slack notifications
  email: false,  // Enable or disable Email notifications
  emailRecipient: "team@example.com"  // Email address for notifications
};

// Slack Integration Settings
var WEBHOOKURI = "https://hooks.slack.com/services/your/webhook/url"; // Your Slack Incoming Webhook URL
var BOTNAME = "weekly_update_bot"; // Name displayed for the bot in Slack
var SLACKCHANNEL = "#your-channel"; // Slack channel to post messages in
var ENABLE_ERROR_NOTIFICATIONS = false; // Set to 'true' to receive error notifications via Slack

// Report Settings
var TEMPLATEFILEID = "1RyLEFZUG5SFQxWed2y1hqaPKwl_MO3eDKYpw5JF9lTs"; // Google Doc ID of your report template
var REPORTFOLDER = "1yjvaaJkD77a_1JuZ0f58uTrYY2Avjo7K"; // Google Drive folder ID to save reports
var FILENAMEPREFIX = "Weekly Update - "; // Prefix for generated report names
var DATEPLACEHOLDER = "XDATEX"; // Placeholder text in the template to be replaced with the date
var LEAD_TIME_DAYS = 2; // Create the report X days before it's due
var CREATIONMESSAGE = "A new report has been created for DUE_DATE. Please start contributing: <FILEURL|Report Link>"; // Message when report is created

// Scheduling Settings
var SCHEDULE = {
  startDate: new Date('2024-09-27'),  // Absolute start date (when the first report is due)
  repeatType: 'weekly',                // Options: 'daily', 'weekly', 'monthly', 'multiDaily', 'multiWeekly'
  interval: 2,                          // Interval in days or weeks (used for multiDaily and multiWeekly)
  timeZone: 'GMT' // Adjust as needed
};


var MULTIWEEK_SETTINGS = {
  startDate: new Date('2024-09-26'), // Absolute start date (e.g., first report date)
  weekInterval: 2, // Report every X weeks (e.g., every 2 weeks)
};


// Reminder Settings
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
  },
];

// Utility Functions
var Utils = (function() {
  function formatDate(date, format) {
    try {
      return Utilities.formatDate(date, SCHEDULE.timeZone, format);
    } catch (error) {
      logError("Error in formatDate: " + error);
      throw error;
    }
  }

  function parseTime(timeString) {
    try {
      var parts = timeString.split(':');
      return {
        hours: parseInt(parts[0], 10),
        minutes: parseInt(parts[1], 10)
      };
    } catch (error) {
      logError("Error in parseTime: " + error);
      throw error;
    }
  }

  return {
    formatDate: formatDate,
    parseTime: parseTime,
  };
})();

// Scheduler Module
var Scheduler = (function() {
  function getNextReportDate() {
    try {
      var today = new Date();
      var startDate = SCHEDULE.startDate;
      var repeatType = SCHEDULE.repeatType;
      var interval = SCHEDULE.interval || 1; // Default interval is 1 for daily, weekly, monthly

      var nextDate;

      switch (repeatType.toLowerCase()) {
        case 'daily':
          nextDate = getNextDateFromStart(startDate, today, 1);
          break;

        case 'weekly':
          nextDate = getNextDateFromStart(startDate, today, 7);
          break;

        case 'monthly':
          nextDate = getNextMonthlyDate(startDate, today);
          break;

        case 'multidaily':
          nextDate = getNextDateFromStart(startDate, today, interval); // n days
          break;

        case 'multiweekly':
          nextDate = getNextDateFromStart(startDate, today, interval * 7); // n weeks
          break;

        default:
          throw new Error('Invalid repeat type specified.');
      }

      return nextDate;
    } catch (error) {
      logError("Error in getNextReportDate: " + error);
      throw error;
    }
  }

  // Helper function to calculate the next date based on the start date and an interval
  function getNextDateFromStart(startDate, currentDate, dayInterval) {
    var daysBetween = Math.floor((currentDate - startDate) / (1000 * 60 * 60 * 24)); // Difference in days
    var intervalsPassed = Math.floor(daysBetween / dayInterval);
    var nextDate = new Date(startDate);
    nextDate.setDate(startDate.getDate() + (intervalsPassed + 1) * dayInterval);
    return nextDate;
  }

  // Helper function to calculate the next monthly date
  function getNextMonthlyDate(startDate, currentDate) {
    var nextDate = new Date(startDate);
    nextDate.setMonth(currentDate.getMonth());
    if (nextDate < currentDate) {
      nextDate.setMonth(nextDate.getMonth() + 1); // Move to next month
    }
    return nextDate;
  }

  return {
    getNextReportDate: getNextReportDate,
  };
})();


// ReportManager Module
var ReportManager = (function() {
  function createReport(reportDate) {
    try {
      var isoDate = Utils.formatDate(reportDate, "yyyy-MM-dd");
      var niceDate = Utils.formatDate(reportDate, "EEEE, MMMM d, yyyy");
      var reportName = FILENAMEPREFIX + isoDate;

      // Check if the report already exists in the folder
      var existingFiles = DriveApp.getFilesByName(reportName);
      if (existingFiles.hasNext()) {
        Logger.log("Report already exists: " + reportName);
        return; // Skip report creation if it already exists
      }

      Logger.log("Creating new report: " + reportName);

      // Copy the template file
      var docId = DriveApp.getFileById(TEMPLATEFILEID).makeCopy().getId();
      var doc = DocumentApp.openById(docId);
      doc.setName(reportName);

      // Replace the date placeholder
      var body = doc.getBody();
      body.replaceText(DATEPLACEHOLDER, niceDate);
      doc.saveAndClose();

      // Move the file to the specified folder
      var folder = DriveApp.getFolderById(REPORTFOLDER);
      folder.addFile(DriveApp.getFileById(docId));
      DriveApp.getRootFolder().removeFile(DriveApp.getFileById(docId)); // Remove from root if necessary

      Logger.log("Report successfully created: " + reportName);

    } catch (error) {
      logError("Error in createReport: " + error);
      throw error;
    }
  }

  function getReportUrl(reportDate) {
    try {
      var isoDate = Utils.formatDate(reportDate, "yyyy-MM-dd");
      var reportName = FILENAMEPREFIX + isoDate;
      var files = DriveApp.getFilesByName(reportName);
      if (files.hasNext()) {
        return files.next().getUrl();
      } else {
        return "Report URL not available.";
      }
    } catch (error) {
      logError("Error in getReportUrl: " + error);
      return "Report URL not available.";
    }
  }

  return {
    createReport: createReport,
    getReportUrl: getReportUrl,
  };
})();


var Notifier = (function() {

  // Function to send Slack messages
  function sendSlackMessage(message) {
    if (!NOTIFICATIONS.slack) return;  // Slack notifications are disabled

    try {
      var payload = {
        channel: SLACKCHANNEL,
        username: BOTNAME,
        text: message
      };

      var options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload)
      };

      UrlFetchApp.fetch(WEBHOOKURI, options);
    } catch (error) {
      logError("Error in sendSlackMessage: " + error);
    }
  }

  // Function to send email notifications
  function sendEmailNotification(message) {
    if (!NOTIFICATIONS.email) return;  // Email notifications are disabled

    try {
      MailApp.sendEmail({
        to: NOTIFICATIONS.emailRecipient,
        subject: "Report Notification",
        body: message
      });
    } catch (error) {
      logError("Error in sendEmailNotification: " + error);
    }
  }

  // Function to send reminders
  function sendReminders() {
    try {
      var today = new Date();
      var reportDate = Scheduler.getNextReportDate();
      var daysUntilDue = Math.ceil((reportDate - today) / (1000 * 60 * 60 * 24));

      REMINDERS.forEach(function(reminder) {
        if (daysUntilDue === reminder.daysBeforeDue) {
          var message = reminder.message
            .replace("DUE_DATE", Utils.formatDate(reportDate, "EEEE, MMMM d"));

          // Send Slack and/or Email notifications
          sendSlackMessage(message);
          sendEmailNotification(message);
        }
      });
    } catch (error) {
      logError("Error in sendReminders: " + error);
      sendErrorNotification("Error in sendReminders: " + error);
    }
  }

  return {
    sendSlackMessage: sendSlackMessage,
    sendEmailNotification: sendEmailNotification,
    sendReminders: sendReminders,
  };
})();


// Error Logging Function
function logError(error) {
  Logger.log(error);
  // Optionally, you can implement additional error logging mechanisms here, such as:
  // - Sending an email notification
  // - Logging to a spreadsheet or external logging service
  // - Sending an error message via Slack
  if (NOTIFICATIONS.slack && ENABLE_ERROR_NOTIFICATIONS) {
    Notifier.sendSlackMessage("Error: " + error);
  }
  else {
    Logger.log("Error: " + error);
  }
}

// Main Functions
function createReportAndNotify() {
  try {
    Logger.log("Starting createReportAndNotify");

    var reportDate = Scheduler.getNextReportDate();
    var creationDate = new Date(reportDate);
    creationDate.setDate(creationDate.getDate() - LEAD_TIME_DAYS);

    var today = new Date();
    var todayDateString = Utilities.formatDate(today, SCHEDULE.timeZone, 'yyyy-MM-dd');
    var creationDateString = Utilities.formatDate(creationDate, SCHEDULE.timeZone, 'yyyy-MM-dd');

    if (todayDateString === creationDateString) {
      Logger.log("Creating report for date: " + reportDate);

      // Attempt to create the report
      ReportManager.createReport(reportDate);

      // Get the report URL if created, and send notifications
      var reportUrl = ReportManager.getReportUrl(reportDate);
      if (reportUrl !== "Report URL not available.") {
        var message = CREATIONMESSAGE
          .replace("DUE_DATE", Utils.formatDate(reportDate, "EEEE, MMMM d"))
          .replace("FILEURL", reportUrl);
        Notifier.sendSlackMessage(message);
        Notifier.sendEmailNotification(message);
      } else {
        Logger.log("Report already exists, skipping notification.");
      }

    } else {
      Logger.log("Report not created, date mismatch. Today: " + todayDateString + ", Creation date: " + creationDateString);
    }
  } catch (error) {
    logError("Error in createReportAndNotify: " + error);
    Notifier.sendErrorNotification("Error in createReportAndNotify: " + error);
  }
}



function checkAndSendReminders() {
  try {
    Notifier.sendReminders();
  } catch (error) {
    logError("Error in checkAndSendReminders: " + error);
    Notifier.sendErrorNotification("Error in checkAndSendReminders: " + error);
  }
}

