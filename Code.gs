/**
 * This script automates the creation and notification of reports using Google Apps Script.
 */

// Slack Integration Settings
var ENABLE_SLACK_NOTIFICATIONS = true; // Set to 'false' to disable Slack notifications
var WEBHOOKURI = "https://hooks.slack.com/services/your/webhook/url"; // Your Slack Incoming Webhook URL
var BOTNAME = "weekly_update_bot"; // Name displayed for the bot in Slack
var SLACKCHANNEL = "#your-channel"; // Slack channel to post messages in
var ENABLE_ERROR_NOTIFICATIONS = true; // Set to 'true' to receive error notifications via Slack

// Report Settings
var TEMPLATEFILEID = "your-template-file-id"; // Google Doc ID of your report template
var REPORTFOLDER = "your-report-folder-id"; // Google Drive folder ID to save reports
var FILENAMEPREFIX = "Weekly Update - "; // Prefix for generated report names
var DATEPLACEHOLDER = "XDATEX"; // Placeholder text in the template to be replaced with the date
var LEAD_TIME_DAYS = 2; // Create the report X days before it's due
var CREATIONMESSAGE = "A new report has been created for DUE_DATE. Please start contributing: <FILEURL|Report Link>"; // Message when report is created

// Scheduling Settings
var SCHEDULE = {
  frequency: 'multiWeekly', // Options: 'daily', 'weekly', 'everyXDays', 'multiWeekly', 'monthly'
  dayOfWeek: 1, // For 'weekly' and 'multiWeekly' (0 = Sunday, 1 = Monday, ..., 6 = Saturday)
  intervalDays: 7, // For 'everyXDays'
  intervalWeeks: 2, // For 'multiWeekly'
  dayOfMonth: 15, // For 'monthly' on a specific date (1-31)
  weekOfMonth: 2, // For 'monthly' on a specific week (1-5)
  dayOfWeekInMonth: 4, // For 'monthly' on a specific day of the week (0 = Sunday, ..., 6 = Saturday)
  timeZone: 'GMT' // Adjust as needed
};

// Reminder Settings
var REMINDERS = [
  {
    daysBeforeDue: 7,
    timeOfDay: '09:00', // HH:mm in 24-hour format
    message: "The report is due in a week on DUE_DATE. Please start preparing.",
  },
  {
    daysBeforeDue: 1,
    timeOfDay: '09:00',
    message: "Reminder: The report is due tomorrow (DUE_DATE). Don't forget to complete your sections.",
  },
  {
    daysBeforeDue: 0,
    timeOfDay: '09:00',
    message: "Today is the due date for the report (DUE_DATE). Please finalize your inputs.",
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
      var nextDate;

      switch (SCHEDULE.frequency.toLowerCase()) {
        case 'daily':
          nextDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1);
          break;

        case 'weekly':
          nextDate = getNextWeeklyDate(today, SCHEDULE.dayOfWeek);
          break;

        case 'everyxdays':
          nextDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() + SCHEDULE.intervalDays);
          break;

        case 'multiweekly':
          nextDate = getNextMultiWeeklyDate(today, SCHEDULE.intervalWeeks, SCHEDULE.dayOfWeek);
          break;

        case 'monthly':
          if (SCHEDULE.dayOfMonth) {
            nextDate = getNextMonthlyDateByDayOfMonth(today, SCHEDULE.dayOfMonth);
          } else {
            nextDate = getNextMonthlyDateByWeekday(today, SCHEDULE.weekOfMonth, SCHEDULE.dayOfWeekInMonth);
          }
          break;

        default:
          throw new Error('Invalid scheduling frequency specified.');
      }

      return nextDate;
    } catch (error) {
      logError("Error in getNextReportDate: " + error);
      throw error;
    }
  }

  // Include the helper functions here
  function getNextWeeklyDate(currentDate, dayOfWeek) {
    try {
      var daysUntilNext = (dayOfWeek + 7 - currentDate.getDay()) % 7 || 7;
      return new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() + daysUntilNext);
    } catch (error) {
      logError("Error in getNextWeeklyDate: " + error);
      throw error;
    }
  }

  function getNextMultiWeeklyDate(currentDate, intervalWeeks, dayOfWeek) {
    try {
      var nextDate = getNextWeeklyDate(currentDate, dayOfWeek);
      var weeksSinceEpoch = Math.floor(nextDate.getTime() / (7 * 24 * 60 * 60 * 1000));
      if (weeksSinceEpoch % intervalWeeks !== 0) {
        var weeksToAdd = intervalWeeks - (weeksSinceEpoch % intervalWeeks);
        nextDate.setDate(nextDate.getDate() + weeksToAdd * 7);
      }
      return nextDate;
    } catch (error) {
      logError("Error in getNextMultiWeeklyDate: " + error);
      throw error;
    }
  }

  function getNextMonthlyDateByDayOfMonth(currentDate, dayOfMonth) {
    try {
      var nextDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), dayOfMonth);
      if (nextDate <= currentDate) {
        nextDate.setMonth(nextDate.getMonth() + 1);
      }
      return nextDate;
    } catch (error) {
      logError("Error in getNextMonthlyDateByDayOfMonth: " + error);
      throw error;
    }
  }

  function getNextMonthlyDateByWeekday(currentDate, weekOfMonth, dayOfWeek) {
    try {
      var firstDay = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
      var dayOffset = (dayOfWeek - firstDay.getDay() + 7) % 7;
      var day = 1 + dayOffset + 7 * (weekOfMonth - 1);
      var nextDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), day);
      if (nextDate <= currentDate) {
        firstDay.setMonth(firstDay.getMonth() + 1);
        dayOffset = (dayOfWeek - firstDay.getDay() + 7) % 7;
        day = 1 + dayOffset + 7 * (weekOfMonth - 1);
        nextDate = new Date(firstDay.getFullYear(), firstDay.getMonth(), day);
      }
      return nextDate;
    } catch (error) {
      logError("Error in getNextMonthlyDateByWeekday: " + error);
      throw error;
    }
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

      // Copy the template file
      var docId = DriveApp.getFileById(TEMPLATEFILEID).makeCopy().getId();
      var doc = DocumentApp.openById(docId);
      doc.setName(FILENAMEPREFIX + isoDate);

      // Replace the date placeholder
      var body = doc.getBody();
      body.replaceText(DATEPLACEHOLDER, niceDate);
      doc.saveAndClose();

      // Move the file to the specified folder
      var file = DriveApp.getFileById(docId);
      var folder = DriveApp.getFolderById(REPORTFOLDER);
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file); // Remove from root if necessary
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

// Notifier Module
var Notifier = (function() {
  function sendSlackMessage(message) {
    try {
      if (!ENABLE_SLACK_NOTIFICATIONS) return;

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

  function sendErrorNotification(error) {
    if (ENABLE_SLACK_NOTIFICATIONS && ENABLE_ERROR_NOTIFICATIONS) {
      var message = "An error occurred: " + error;
      sendSlackMessage(message);
    }
  }

  function sendReminders() {
    try {
      var today = new Date();
      var reportDate = Scheduler.getNextReportDate();
      var daysUntilDue = Math.ceil((reportDate - today) / (1000 * 60 * 60 * 24));

      REMINDERS.forEach(function(reminder) {
        if (daysUntilDue === reminder.daysBeforeDue) {
          var currentTimeString = Utilities.formatDate(today, SCHEDULE.timeZone, 'HH:mm');
          if (currentTimeString === reminder.timeOfDay) {
            var message = reminder.message
              .replace("DUE_DATE", Utils.formatDate(reportDate, "EEEE, MMMM d"))
              .replace("FILEURL", ReportManager.getReportUrl(reportDate));
            sendSlackMessage(message);
          }
        }
      });
    } catch (error) {
      logError("Error in sendReminders: " + error);
      sendErrorNotification("Error in sendReminders: " + error);
    }
  }

  return {
    sendSlackMessage: sendSlackMessage,
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
  if (ENABLE_SLACK_NOTIFICATIONS && ENABLE_ERROR_NOTIFICATIONS) {
    Notifier.sendSlackMessage("Error: " + error);
  }
  else {
    console.log("Error: " + error);
  }
}

// Main Functions
function createReportAndNotify() {
  try {
    var reportDate = Scheduler.getNextReportDate();
    var creationDate = new Date(reportDate);
    creationDate.setDate(creationDate.getDate() - LEAD_TIME_DAYS);

    var today = new Date();
    var todayDateString = Utilities.formatDate(today, SCHEDULE.timeZone, 'yyyy-MM-dd');
    var creationDateString = Utilities.formatDate(creationDate, SCHEDULE.timeZone, 'yyyy-MM-dd');

    if (todayDateString === creationDateString) {
      ReportManager.createReport(reportDate);
      var message = CREATIONMESSAGE
        .replace("DUE_DATE", Utils.formatDate(reportDate, "EEEE, MMMM d"))
        .replace("FILEURL", ReportManager.getReportUrl(reportDate));
      Notifier.sendSlackMessage(message);
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
