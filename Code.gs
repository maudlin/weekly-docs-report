/**
 * This script automates the creation and notification of reports using Google Apps Script.
 */
// Notification Settings
// Configure whether to enable Slack and Email notifications for different stages of the report workflow.
// Set emailRecipient to the desired recipient for email notifications.
var NOTIFICATIONS = {
  slack: false,  // Enable or disable Slack notifications
  email: false,  // Enable or disable Email notifications
  emailRecipient: "team@example.com"  // Email address for notifications
};

// Slack Integration Settings
// Configure the Slack webhook URL, bot name, and channel for notifications. Update WEBHOOKURI with your actual Slack webhook URL.
var WEBHOOKURI = "[CHANGEME]"; // Your Slack Incoming Webhook URL
var BOTNAME = "status_update_bot"; // Name displayed for the bot in Slack
var SLACKCHANNEL = "#project_status_report_update"; // Slack channel to post messages in
var ENABLE_ERROR_NOTIFICATIONS = false; // Set to 'true' to receive error notifications via Slack

// Report Settings
// Configure the Google Doc template and Drive folder to save reports.
// TEMPLATEFILEID: Google Doc ID for the report template.
// REPORTFOLDER: Google Drive folder ID where reports should be saved.
// DATEPLACEHOLDER: Placeholder text in the template that should be replaced with the date.
// LEAD_TIME_DAYS: Number of days before the report is due that it should be created.
var TEMPLATEFILEID = "[CHANGEME]"; // Google Doc ID of your report template
var REPORTFOLDER = "[CHANGEME]"; // Google Drive folder ID to save reports
var FILENAMEPREFIX = "Weekly Update - "; // Prefix for generated report names
var DATEPLACEHOLDER = "XDATEX"; // Placeholder text in the template to be replaced with the date
var LEAD_TIME_DAYS = 3; // Create the report X days before it's due
var CREATIONMESSAGE = "Hey folks, the bi-weekly report for DUE_DATE has been generated. Please fill in by lunch on Wednesday: <FILEURL|Report Link>"; // Message when report is created

// Scheduling Settings
// SCHEDULE: Define the schedule for generating reports.
// Make sure to adjust the timeZone field according to your location (e.g., 'America/New_York', 'Europe/London').
// startDate: The first date when the report is due.
// repeatType: Frequency of report generation - 'daily', 'weekly', 'monthly'.
// interval: The interval between report generations (used for weekly or daily frequency).
// timeZone: The user's time zone.
var SCHEDULE = {
  startDate: new Date('2024-10-19'),  // Absolute start date (when the first report is due)
  repeatType: 'weekly',                // Options: 'daily', 'weekly', 'monthly'
  interval: 2,                          // Interval in weeks or days (used for weekly/daily)
  timeZone: 'America/New_York' // Adjust as needed
};

// Reminder Settings
// Configure reminders to be sent before the report is due. The reminders array defines how many days before the report is due each reminder should be sent.
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
  // Utility function to format a date in the specified format and time zone
  function formatDate(date, format) {
    try {
      return Utilities.formatDate(date, SCHEDULE.timeZone, format);
    } catch (error) {
      Utils.logError("Error in formatDate: " + error);
      throw error;
    }
  }

  // Utility function to parse time from a given string (HH:MM)
  function parseTime(timeString) {
    try {
      var parts = timeString.split(':');
      return {
        hours: parseInt(parts[0], 10),
        minutes: parseInt(parts[1], 10)
      };
    } catch (error) {
      Utils.logError("Error in parseTime: " + error);
      throw error;
    }
  }

  // Utility function to get today's date normalized to midnight in the user's time zone
  function getNormalizedToday() {
    try {
      var today = new Date();
      var localizedTodayString = Utilities.formatDate(today, SCHEDULE.timeZone, 'yyyy-MM-dd');
      return new Date(localizedTodayString + 'T00:00:00');
    } catch (error) {
      Utils.logError("Error in getNormalizedToday: " + error);
      throw error;
    }
  }

  // Utility function for logging errors
  function logError(error) {
    Logger.log(error);
    // Optionally, implement additional error logging mechanisms:
    if (NOTIFICATIONS.slack && ENABLE_ERROR_NOTIFICATIONS) {
      Notifier.sendSlackMessage("Error: " + error);
    } else {
      Logger.log("Error: " + error);
    }
  }

  return {
    formatDate: formatDate,
    parseTime: parseTime,
    getNormalizedToday: getNormalizedToday,
    logError: logError
  };
})();


// Scheduler Module
var Scheduler = (function() {
  function getNextReportDate() {
    try {
      var today = Utils.getNormalizedToday();
      return getNextReportDateAfter(today);
    } catch (error) {
      Utils.logError("Error in getNextReportDate: " + error);
      throw error;
    }
  }

  function getNextReportDateAfter(referenceDate) {
    try {
      var startDate = new Date(SCHEDULE.startDate);
      var repeatType = SCHEDULE.repeatType.toLowerCase();
      var interval = SCHEDULE.interval || 1; // Default interval is 1 for daily, weekly, monthly
      var nextDate;

      switch (repeatType) {
        case 'daily':
          nextDate = getNextDateFromStart(startDate, referenceDate, interval);
          break;
        case 'weekly':
          nextDate = getNextDateFromStart(startDate, referenceDate, interval * 7);
          break;
        case 'monthly':
          nextDate = getNextMonthlyDate(startDate, referenceDate);
          break;
        default:
          throw new Error('Invalid repeat type specified.');
      }

      // Ensure next date is not in the past
      if (nextDate <= referenceDate) {
        switch (repeatType) {
          case 'daily':
            nextDate.setDate(nextDate.getDate() + interval);
            break;
          case 'weekly':
            nextDate.setDate(nextDate.getDate() + interval * 7);
            break;
          case 'monthly':
            nextDate.setMonth(nextDate.getMonth() + 1);
            break;
        }
      }

      return nextDate;
    } catch (error) {
      Utils.logError("Error in getNextReportDateAfter: " + error);
      throw error;
    }
  }

  // Helper function to calculate the next date based on the start date and an interval
  function getNextDateFromStart(startDate, referenceDate, dayInterval) {
    startDate.setHours(0, 0, 0, 0);
    referenceDate.setHours(0, 0, 0, 0);

    if (referenceDate < startDate) {
      return startDate;
    }

    var daysBetween = Math.floor((referenceDate - startDate) / (1000 * 60 * 60 * 24)); // Difference in days
    var intervalsPassed = Math.floor(daysBetween / dayInterval);
    var nextDate = new Date(startDate);
    nextDate.setDate(startDate.getDate() + (intervalsPassed + 1) * dayInterval);

    return nextDate;
  }

  // Helper function to calculate the next monthly date
  function getNextMonthlyDate(startDate, referenceDate) {
    var nextDate = new Date(startDate);
    nextDate.setMonth(referenceDate.getMonth());

    if (nextDate <= referenceDate) {
      nextDate.setMonth(nextDate.getMonth() + 1); // Move to next month if already passed
    }

    return nextDate;
  }

  return {
    getNextReportDate: getNextReportDate,
    getNextReportDateAfter: getNextReportDateAfter
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
      Utils.logError("Error in createReport: " + error);
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
      Utils.logError("Error in getReportUrl: " + error);
      return "Report URL not available.";
    }
  }

  return {
    createReport: createReport,
    getReportUrl: getReportUrl,
  };
})();

// Notifier Module
var Notifier = (function() {
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
      Logger.log("Slack message sent: " + message);
    } catch (error) {
      Utils.logError("Error in sendSlackMessage: " + error);
    }
  }

  function sendEmailNotification(message) {
    if (!NOTIFICATIONS.email) return;  // Email notifications are disabled
    try {
      MailApp.sendEmail({
        to: NOTIFICATIONS.emailRecipient,
        subject: "Report Notification",
        body: message
      });
      Logger.log("Email notification sent to: " + NOTIFICATIONS.emailRecipient);
    } catch (error) {
      Utils.logError("Error in sendEmailNotification: " + error);
    }
  }

  function sendReminders() {
    try {
      var today = Utils.getNormalizedToday();
      var reportDate = Scheduler.getNextReportDate();
      var daysUntilDue = Math.ceil((reportDate - today) / (1000 * 60 * 60 * 24));

      Logger.log("Next report due date: " + Utils.formatDate(reportDate, "yyyy-MM-dd"));

      REMINDERS.forEach(function(reminder) {
        if (daysUntilDue === reminder.daysBeforeDue) {
          var message = reminder.message
            .replace("DUE_DATE", Utils.formatDate(reportDate, "EEEE, MMMM d"));
          // Send Slack and/or Email notifications
          sendSlackMessage(message);
          sendEmailNotification(message);
          Logger.log("Reminder sent: " + message);
        } else {
          Logger.log("No reminder today for report due date: " + Utils.formatDate(reportDate, "yyyy-MM-dd"));
        }
      });
    } catch (error) {
      Utils.logError("Error in sendReminders: " + error);
      sendSlackMessage("Error in sendReminders: " + error);
    }
  }

  return {
    sendSlackMessage: sendSlackMessage,
    sendEmailNotification: sendEmailNotification,
    sendReminders: sendReminders,
  };
})();

// Main Functions
function createReportAndNotify() {
  try {
    Logger.log("Starting createReportAndNotify");

    // Get today and normalize to midnight
    var today = Utils.getNormalizedToday();

    // Get the next report due date
    var reportDate = Scheduler.getNextReportDate();

    // Calculate the creation date for this report
    var creationDate = new Date(reportDate);
    creationDate.setDate(creationDate.getDate() - LEAD_TIME_DAYS);

    // Log today's date, creation date, and report date
    Logger.log("Today (User Time Zone): " + Utils.formatDate(today, 'yyyy-MM-dd'));
    Logger.log("Time Zone: " + SCHEDULE.timeZone);

    var todayDateString = Utils.formatDate(today, 'yyyy-MM-dd');
    var creationDateString = Utils.formatDate(creationDate, 'yyyy-MM-dd');
    var reportDateString = Utils.formatDate(reportDate, 'yyyy-MM-dd');

    Logger.log("Today: " + todayDateString);
    Logger.log("Current report creation date (" + LEAD_TIME_DAYS + " days before due date): " + creationDateString);
    Logger.log("Current report due date: " + reportDateString);

    // If today matches the creation date, create the report
    if (Utils.formatDate(today, 'yyyy-MM-dd') === Utils.formatDate(creationDate, 'yyyy-MM-dd')) {
      Logger.log("Creating report for due date: " + reportDateString);

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
        Logger.log("Report successfully created and notifications sent.");
      } else {
        Logger.log("Report already exists, skipping notification.");
      }
    } else {
      Logger.log("Report not created, date mismatch. Today: " + todayDateString + ", Creation date: " + creationDateString);
    }

    // Calculate and log the next report dates after the current one
    var nextReportDate = Scheduler.getNextReportDateAfter(reportDate);
        var nextCreationDate = new Date(nextReportDate);
    nextCreationDate.setDate(nextReportDate.getDate() - LEAD_TIME_DAYS);

    Logger.log("Next report creation date: " + Utils.formatDate(nextCreationDate, "yyyy-MM-dd"));
    Logger.log("Next report due date: " + Utils.formatDate(nextReportDate, "yyyy-MM-dd"));

  } catch (error) {
    Utils.logError("Error in createReportAndNotify: " + error);
    Notifier.sendSlackMessage("Error in createReportAndNotify: " + error);
  }
}

function checkAndSendReminders() {
  try {
    Notifier.sendReminders();
  } catch (error) {
    Utils.logError("Error in checkAndSendReminders: " + error);
    Notifier.sendSlackMessage("Error in checkAndSendReminders: " + error);
  }
}

