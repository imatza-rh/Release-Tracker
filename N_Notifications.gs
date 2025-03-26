/**
 * N_Notifications.gs - Phase notification system
 * 
 * This file contains functions for sending notifications when phases
 * start or end, as defined in a "Phases and Dates" sheet.
 * 
 * @requires log, handleError, safeExecute from B_Logging.gs
 * @requires safeParseJson from C_Utils.gs
 */

// Property key for storing notification settings
const NOTIFICATION_SETTINGS_KEY = 'RELEASE_TRACKER_NOTIFICATIONS';
const PHASES_SHEET_NAME = "Phases and Dates";

/**
 * Retrieves notification settings
 * 
 * @returns {Object} Notification settings
 */
function getNotificationSettings() {
  return safeExecute('getNotificationSettings', () => {
    const props = PropertiesService.getScriptProperties();
    const stored = props.getProperty(NOTIFICATION_SETTINGS_KEY);
    
    if (stored) {
      return safeParseJson(stored, getDefaultNotificationSettings());
    }
    
    return getDefaultNotificationSettings();
  }, false);
}

/**
 * Returns default notification settings
 * 
 * @returns {Object} Default notification settings
 */
function getDefaultNotificationSettings() {
  return {
    slackWebhookUrls: [],
    notificationEmails: [],
    enableNotifications: false,
    lastRunTime: null,
    lastRunStatus: null
  };
}

/**
 * Saves notification settings
 * 
 * @param {Object} settings - Notification settings to save
 * @returns {Object} Result indicating success or failure
 */
function saveNotificationSettings(settings) {
  return safeExecute('saveNotificationSettings', () => {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(NOTIFICATION_SETTINGS_KEY, JSON.stringify(settings));
    
    // Update trigger based on enableNotifications setting
    if (settings.enableNotifications) {
      setupDailyTrigger();
    } else {
      removeDailyTrigger();
    }
    
    log('Notification settings saved', LOG_LEVELS.INFO, settings);
    return { 
      status: 'success', 
      message: 'Notification settings saved successfully' 
    };
  }, false);
}

/**
 * Sets up a daily trigger to run notifications
 * 
 * @param {number} [hour=8] - Hour of the day to run the trigger (0-23)
 * @returns {boolean} Success state
 */
function setupDailyTrigger(hour = 8) {
  return safeExecute('setupDailyTrigger', () => {
    removeDailyTrigger(); // Remove existing triggers first
    
    // Create a new trigger to run at specified hour every day
    ScriptApp.newTrigger("sendPhaseNotifications")
      .timeBased()
      .atHour(hour)
      .everyDays(1)
      .create();
    
    log('Daily notification trigger created', LOG_LEVELS.INFO);
    return true;
  }, false);
}

/**
 * Removes the daily notification trigger
 * 
 * @returns {number} Number of triggers removed
 */
function removeDailyTrigger() {
  return safeExecute('removeDailyTrigger', () => {
    const triggers = ScriptApp.getProjectTriggers();
    let removed = 0;
    
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === "sendPhaseNotifications") {
        ScriptApp.deleteTrigger(trigger);
        removed++;
      }
    }
    
    if (removed > 0) {
      log(`Removed ${removed} notification trigger(s)`, LOG_LEVELS.INFO);
    }
    
    return removed;
  }, false);
}

/**
 * Directly checks if a specified sheet exists
 * 
 * @param {string} sheetName - Name of the sheet to check
 * @returns {boolean} Whether the sheet exists
 */
function checkSheetExists(sheetName) {
  return safeExecute('checkSheetExists', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    return sheet !== null;
  }, false);
}

/**
 * Checks for phases that are starting or ending today
 * 
 * @returns {Object} Object containing information about phases today
 */
function checkPhasesToday() {
  return safeExecute('checkPhasesToday', () => {
    // Get the spreadsheet and find the sheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    log("Looking for sheet: " + PHASES_SHEET_NAME, LOG_LEVELS.INFO);
    
    const sheet = spreadsheet.getSheetByName(PHASES_SHEET_NAME);
    if (!sheet) {
      log("Sheet not found: " + PHASES_SHEET_NAME, LOG_LEVELS.WARN);
      return { found: false, count: 0, details: "Sheet not found" };
    }
    
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    
    if (data.length <= 1) { // Just header or empty
      log("Sheet is empty or only contains headers.", LOG_LEVELS.INFO);
      return { found: false, count: 0, details: "No phase data found" };
    }
    
    const headers = data[0].map(String).map(h => h.trim());
    log("Detected Headers: " + headers.join(", "), LOG_LEVELS.DEBUG);
    
    const startDateIndex = headers.indexOf("Start Date");
    const endDateIndex = headers.indexOf("End Date");
    const phaseIndex = headers.indexOf("Phases");
    
    if (startDateIndex === -1 || endDateIndex === -1 || phaseIndex === -1) {
      log("Required columns not found in the sheet.", LOG_LEVELS.WARN);
      return { 
        found: false, 
        count: 0, 
        details: "Required columns (Phases, Start Date, End Date) not found in the sheet"
      };
    }
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayFormatted = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
    log("Today's Date: " + todayFormatted, LOG_LEVELS.INFO);
    
    let phasesToday = findPhasesForToday(data, headers, todayFormatted);
    
    if (phasesToday.length > 0) {
      let details = "";
      phasesToday.forEach(p => {
        details += `• "${p.phase}" ${p.eventType} today (${p.eventType === "starts" ? p.startDate : p.endDate})\n`;
      });
      
      return {
        found: true,
        count: phasesToday.length,
        details: details,
        phases: phasesToday
      };
    } else {
      return {
        found: false,
        count: 0,
        details: "No phases starting or ending today"
      };
    }
  }, false);
}

/**
 * Helper to find phases for today's date
 * 
 * @param {Array} data - Sheet data array
 * @param {Array} headers - Headers array
 * @param {string} todayFormatted - Today's date formatted as yyyy-MM-dd
 * @returns {Array} Array of phase objects for today
 */
function findPhasesForToday(data, headers, todayFormatted) {
  const startDateIndex = headers.indexOf("Start Date");
  const endDateIndex = headers.indexOf("End Date");
  const phaseIndex = headers.indexOf("Phases");
  let phasesToday = [];
  
  // Check each row for phases starting or ending today
  for (let i = 1; i < data.length; i++) {
    const phase = data[i][phaseIndex];
    if (!phase) continue; // Skip rows without a phase name
    
    let startDate = new Date(data[i][startDateIndex]);
    let endDate = new Date(data[i][endDateIndex]);
    
    // Handle date parsing issues
    if (isNaN(startDate.getTime())) {
      try {
        startDate = new Date(Date.parse(data[i][startDateIndex]));
      } catch (e) {
        log(`Invalid start date for phase "${phase}"`, LOG_LEVELS.WARN);
        continue;
      }
    }
    
    if (isNaN(endDate.getTime())) {
      try {
        endDate = new Date(Date.parse(data[i][endDateIndex]));
      } catch (e) {
        log(`Invalid end date for phase "${phase}"`, LOG_LEVELS.WARN);
        continue;
      }
    }
    
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      log("Skipping row due to invalid date format: " + phase, LOG_LEVELS.WARN);
      continue;
    }
    
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(0, 0, 0, 0);
    
    const startDateFormatted = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const endDateFormatted = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    log(`Checking phase: ${phase} | Start Date: ${startDateFormatted} | End Date: ${endDateFormatted}`, LOG_LEVELS.DEBUG);
    
    if (startDateFormatted === todayFormatted || endDateFormatted === todayFormatted) {
      const eventType = startDateFormatted === todayFormatted ? "starts" : "ends";
      
      phasesToday.push({
        phase: phase,
        eventType: eventType,
        startDate: startDateFormatted,
        endDate: endDateFormatted,
        rowData: data[i]
      });
    }
  }
  
  return phasesToday;
}

/**
 * Sends notifications for phases that are starting or ending today
 */
function sendPhaseNotifications() {
  return safeExecute('sendPhaseNotifications', () => {
    log('Running phase notifications check', LOG_LEVELS.INFO);
    
    // Get notification settings
    const settings = getNotificationSettings();
    
    // Update last run time
    settings.lastRunTime = new Date().toISOString();
    
    if (!settings.enableNotifications) {
      log('Notifications are disabled. Skipping execution.', LOG_LEVELS.INFO);
      settings.lastRunStatus = "Skipped - Notifications disabled";
      saveNotificationSettings(settings);
      return {
        status: 'skipped',
        message: 'Notifications are disabled'
      };
    }
    
    const emailRecipients = settings.notificationEmails || [];
    const slackWebhookUrls = settings.slackWebhookUrls || [];
    
    if (emailRecipients.length === 0 && slackWebhookUrls.length === 0) {
      log('No recipients configured. Skipping execution.', LOG_LEVELS.INFO);
      settings.lastRunStatus = "Skipped - No recipients configured";
      saveNotificationSettings(settings);
      return {
        status: 'skipped',
        message: 'No recipients configured'
      };
    }
    
    // Check for phases today
    const phaseCheck = checkPhasesToday();
    
    if (!phaseCheck.found) {
      log('No phases starting or ending today. Skipping notifications.', LOG_LEVELS.INFO);
      settings.lastRunStatus = "Completed - No phases today";
      saveNotificationSettings(settings);
      return {
        status: 'success',
        message: 'No phases today',
        phasesCount: 0
      };
    }
    
    // Process and send notifications based on phases found
    const result = processAndSendNotifications(phaseCheck, emailRecipients, slackWebhookUrls);
    
    // Update settings with status
    settings.lastRunStatus = `Completed - Sent ${result.emailsSent} emails and ${result.slackSent} Slack notifications for ${result.phasesCount} phase(s)`;
    saveNotificationSettings(settings);
    
    return result;
  }, false);
}

/**
 * Process phases and send notifications to all configured channels
 * 
 * @param {Object} phaseCheck - Phase check result
 * @param {Array} emailRecipients - List of email recipients
 * @param {Array} slackWebhookUrls - List of Slack webhook URLs
 * @returns {Object} Result of notification sending
 */
function processAndSendNotifications(phaseCheck, emailRecipients, slackWebhookUrls) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const fileName = spreadsheet.getName();
  const fileUrl = spreadsheet.getUrl();
  
  // Use the phases from checkPhasesToday
  const phasesToday = phaseCheck.phases;
  const phasesCount = phasesToday.length;
  
  // Get email content
  const { subject, emailBody } = createEmailContent(phasesToday, phasesCount, fileName, fileUrl);
  
  // Get Slack content
  const slackBlocks = createSlackContent(phasesToday, phasesCount, fileName, fileUrl);
  
  // Send notifications
  let emailsSent = 0;
  let slackSent = 0;
  
  // Send email notifications to all recipients
  if (emailBody && emailRecipients.length > 0) {
    emailRecipients.forEach(email => {
      try {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: emailBody
        });
        emailsSent++;
      } catch (e) {
        log(`Failed to send email to ${email}: ${e}`, LOG_LEVELS.ERROR);
      }
    });
    log(`${emailsSent} emails sent successfully.`, LOG_LEVELS.INFO);
  }
  
  // Send Slack notifications to all configured webhook URLs
  if (slackBlocks.length > 0 && slackWebhookUrls.length > 0) {
    const payload = JSON.stringify({ "blocks": slackBlocks });
    const options = {
      "method": "post",
      "contentType": "application/json",
      "payload": payload,
      "muteHttpExceptions": true
    };
    
    slackWebhookUrls.forEach(url => {
      try {
        const response = UrlFetchApp.fetch(url, options);
        if (response.getResponseCode() === 200) {
          slackSent++;
        } else {
          log(`Failed to send Slack message, got response code: ${response.getResponseCode()}`, LOG_LEVELS.ERROR);
        }
      } catch (e) {
        log(`Failed to send Slack notification to webhook: ${e}`, LOG_LEVELS.ERROR);
      }
    });
    
    log(`${slackSent} Slack notifications sent successfully.`, LOG_LEVELS.INFO);
  }
  
  return {
    status: 'success',
    message: `Sent ${emailsSent} emails and ${slackSent} Slack notifications`,
    emailsSent: emailsSent,
    slackSent: slackSent,
    phasesCount: phasesCount
  };
}

/**
 * Creates email content for notifications
 * 
 * @param {Array} phasesToday - Phases happening today
 * @param {number} phasesCount - Count of phases
 * @param {string} fileName - Spreadsheet filename
 * @param {string} fileUrl - Spreadsheet URL
 * @returns {Object} Email subject and body
 */
function createEmailContent(phasesToday, phasesCount, fileName, fileUrl) {
  let subject = "";
  let emailBody = "";
  
  // Get sheet for header info
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PHASES_SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const objectiveIndex = headers.indexOf("Objective");
  const criteriaIndex = headers.indexOf("Criteria for Completion");
  
  // Create the subject line
  if (phasesCount === 1) {
    // Single phase subject line
    const phase = phasesToday[0];
    subject = `🚀 Phase Alert: '${phase.phase}' ${phase.eventType} today!`;
  } else {
    // Multiple phases subject line
    subject = `🚀 Phase Alert: ${phasesCount} phases today!`;
  }
  
  // Create the heading for the email
  if (phasesCount > 1) {
    emailBody += `<div style="font-family: Arial, sans-serif; padding: 15px; border: 2px solid #ddd; border-radius: 8px; background-color: #f9f9f9; margin-bottom: 20px;">
      <h2 style="color: #2a7ae2;">🚀 ${phasesCount} phases are active today!</h2>
      <p style="font-size: 14px;">The following phases are starting or ending today:</p>
    </div>`;
  }
  
  // Process each phase
  for (const phase of phasesToday) {
    const objective = objectiveIndex !== -1 ? phase.rowData[objectiveIndex] || "N/A" : "N/A";
    const criteria = criteriaIndex !== -1 ? phase.rowData[criteriaIndex] || "N/A" : "N/A";
    
    // Build email content for this phase
    emailBody += `<div style="font-family: Arial, sans-serif; padding: 15px; border: 2px solid #ddd; border-radius: 8px; background-color: #f9f9f9; margin-bottom: 15px;">
      <h2 style="color: #2a7ae2;">🚀 '${phase.phase}' ${phase.eventType} today!</h2>
      <p style="font-size: 14px;"><strong>📌 Objective:</strong> ${objective}</p>
      <p style="font-size: 14px;"><strong>✅ Completion Criteria:</strong> ${criteria}</p>
      <p style="font-size: 14px;">🗓 <strong>Start Date:</strong> ${phase.startDate} | <strong>End Date:</strong> ${phase.endDate}</p>
      <hr>
      <p style="font-size: 14px;"><a href="${fileUrl}" style="color: #2a7ae2; text-decoration: none;">📎 Open '${fileName}'</a></p>
    </div>`;
  }
  
  // Add notification count to the email footer
  emailBody += `<div style="font-size: 12px; color: #777; margin-top: 15px; text-align: center;">
    This is an automated notification from your Release Tracker.
    <br>To configure these notifications, open the spreadsheet and use the "🚀 Release Tracker" menu.
  </div>`;
  
  return { subject, emailBody };
}

/**
 * Creates Slack message blocks for notifications
 * 
 * @param {Array} phasesToday - Phases happening today
 * @param {number} phasesCount - Count of phases
 * @param {string} fileName - Spreadsheet filename
 * @param {string} fileUrl - Spreadsheet URL
 * @returns {Array} Slack message blocks
 */
function createSlackContent(phasesToday, phasesCount, fileName, fileUrl) {
  const slackBlocks = [];
  
  // Get sheet for header info
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PHASES_SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const objectiveIndex = headers.indexOf("Objective");
  const criteriaIndex = headers.indexOf("Criteria for Completion");
  
  // Add header for multiple phases
  if (phasesCount > 1) {
    slackBlocks.push(
      { 
        "type": "header", 
        "text": { 
          "type": "plain_text", 
          "text": `🚀 ${phasesCount} phases are active today!`, 
          "emoji": true 
        } 
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "The following phases are starting or ending today:"
        }
      },
      { "type": "divider" }
    );
  }
  
  // Process each phase
  for (const phase of phasesToday) {
    const objective = objectiveIndex !== -1 ? phase.rowData[objectiveIndex] || "N/A" : "N/A";
    const criteria = criteriaIndex !== -1 ? phase.rowData[criteriaIndex] || "N/A" : "N/A";
    
    // Add phase blocks
    slackBlocks.push(
      { "type": "header", "text": { "type": "plain_text", "text": `🚀 '${phase.phase}' ${phase.eventType}!`, "emoji": true } },
      { "type": "section", "text": { "type": "mrkdwn", "text": `🗓 *Start Date:* ${phase.startDate} | *End Date:* ${phase.endDate}` } },
      { "type": "divider" },
      { "type": "section", "text": { "type": "mrkdwn", "text": `*Objective:* ${objective}` } },
      { "type": "section", "text": { "type": "mrkdwn", "text": `*Completion Criteria:* ${criteria}` } },
      { "type": "section", "text": { "type": "mrkdwn", "text": `*More Details:* <${fileUrl}|${fileName}>` } },
      { "type": "divider" }
    );
  }
  
  return slackBlocks;
}

/**
 * Sends a bulk test notification to multiple recipients
 * 
 * @param {Object} testConfig - Configuration for the test
 * @returns {Object} Result of the test
 */
function sendBulkTestNotification(testConfig) {
  return safeExecute('sendBulkTestNotification', () => {
    log('Sending bulk test notification', LOG_LEVELS.INFO, testConfig);
    
    // Validate test config
    if (!testConfig || 
        (!Array.isArray(testConfig.emails) && !Array.isArray(testConfig.webhooks)) ||
        (Array.isArray(testConfig.emails) && testConfig.emails.length === 0 && 
         Array.isArray(testConfig.webhooks) && testConfig.webhooks.length === 0)) {
      return {
        status: 'error',
        message: 'No test recipients specified'
      };
    }
    
    // Get the spreadsheet info
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const fileName = spreadsheet.getName();
    const fileUrl = spreadsheet.getUrl();
    
    // Create test content
    const subject = `🧪 Test Notification from Release Tracker`;
    
    const emailBody = `<div style="font-family: Arial, sans-serif; padding: 15px; border: 2px solid #ddd; border-radius: 8px; background-color: #f9f9f9;">
      <h2 style="color: #2a7ae2;">🧪 This is a test notification</h2>
      <p style="font-size: 14px;">If you're seeing this, your email notifications are working correctly!</p>
      <p style="font-size: 14px;">This test was sent at: ${new Date().toLocaleString()}</p>
      <hr>
      <p style="font-size: 14px;"><a href="${fileUrl}" style="color: #2a7ae2; text-decoration: none;">📎 Open '${fileName}'</a></p>
    </div>`;
    
    const slackBlocks = [
      { 
        "type": "header", 
        "text": { 
          "type": "plain_text", 
          "text": "🧪 Test Notification from Release Tracker", 
          "emoji": true 
        } 
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "If you're seeing this, your Slack notifications are working correctly!"
        }
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": `This test was sent at: ${new Date().toLocaleString()}`
        }
      },
      { 
        "type": "section", 
        "text": { 
          "type": "mrkdwn", 
          "text": `*Source:* <${fileUrl}|${fileName}>` 
        } 
      }
    ];
    
    let emailsSent = 0;
    let slackSent = 0;
    let errors = [];
    
    // Send email tests
    if (Array.isArray(testConfig.emails) && testConfig.emails.length > 0) {
      testConfig.emails.forEach(email => {
        try {
          MailApp.sendEmail({
            to: email,
            subject: subject,
            htmlBody: emailBody
          });
          emailsSent++;
          log(`Test email sent to ${email}`, LOG_LEVELS.INFO);
        } catch (e) {
          const errorMsg = `Failed to send test email to ${email}: ${e.message || e}`;
          log(errorMsg, LOG_LEVELS.ERROR);
          errors.push(errorMsg);
        }
      });
    }
    
    // Send Slack tests
    if (Array.isArray(testConfig.webhooks) && testConfig.webhooks.length > 0) {
      const payload = JSON.stringify({ "blocks": slackBlocks });
      const options = {
        "method": "post",
        "contentType": "application/json",
        "payload": payload,
        "muteHttpExceptions": true
      };
      
      testConfig.webhooks.forEach(webhook => {
        try {
          const response = UrlFetchApp.fetch(webhook, options);
          const responseCode = response.getResponseCode();
          
          if (responseCode === 200) {
            slackSent++;
            log(`Test Slack notification sent to webhook`, LOG_LEVELS.INFO);
          } else {
            const errorMsg = `Failed to send Slack notification: Received status code ${responseCode}`;
            log(errorMsg, LOG_LEVELS.ERROR);
            errors.push(errorMsg);
          }
        } catch (e) {
          const errorMsg = `Failed to send test Slack notification: ${e.message || e}`;
          log(errorMsg, LOG_LEVELS.ERROR);
          errors.push(errorMsg);
        }
      });
    }
    
    // Return result based on successes and failures
    if (emailsSent > 0 || slackSent > 0) {
      return {
        status: 'success',
        message: `Test notifications sent successfully`,
        emailsSent: emailsSent,
        slackSent: slackSent,
        errors: errors
      };
    } else if (errors.length > 0) {
      return {
        status: 'error',
        message: errors[0], // Just show the first error
        errors: errors,
        emailsSent: 0,
        slackSent: 0
      };
    } else {
      return {
        status: 'error',
        message: 'Failed to send any test notifications',
        errors: ['Unknown error occurred'],
        emailsSent: 0,
        slackSent: 0
      };
    }
  }, false);
}

/**
 * Shows the notification settings sidebar
 */
function showNotificationSettings() {
  return safeExecute('showNotificationSettings', () => {
    const html = HtmlService.createTemplateFromFile('NotificationSettings')
      .evaluate()
      .setTitle('Phase Notification Settings')
      .setWidth(400);
    
    SpreadsheetApp.getUi().showSidebar(html);
    log('Notification settings sidebar opened', LOG_LEVELS.INFO);
  }, true);
}

/**
 * Gets the status of triggers for the UI
 * 
 * @returns {Object} Trigger status information
 */
function getTriggersStatus() {
  return safeExecute('getTriggersStatus', () => {
    const triggers = ScriptApp.getProjectTriggers();
    let found = false;
    let triggerTime = "";
    
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === "sendPhaseNotifications") {
        found = true;
        
        // Get the trigger time if available
        if (trigger.getEventType() === ScriptApp.EventType.CLOCK) {
          const hour = trigger.getAtHour();
          triggerTime = `${hour}:00`;
        }
        
        break;
      }
    }
    
    return {
      exists: found,
      time: triggerTime
    };
  }, false) || { exists: false, time: "" };
}

/**
 * Creates a Phases and Dates sheet with the proper columns
 * 
 * @returns {Object} Result object indicating success or failure
 */
function createPhasesSheet() {
  return safeExecute('createPhasesSheet', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(PHASES_SHEET_NAME);
    
    // If sheet already exists, do nothing
    if (sheet) {
      return {
        success: true,
        message: "Sheet already exists"
      };
    }
    
    // Create new sheet
    sheet = ss.insertSheet(PHASES_SHEET_NAME);
    
    // Define headers
    const headers = [
      "Phases", 
      "Start Date", 
      "End Date", 
      "Objective", 
      "Criteria for Completion"
    ];
    
    // Add header row
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Add example data
    const today = new Date();
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    
    const nextWeek = new Date(today);
    nextWeek.setDate(nextWeek.getDate() + 7);
    
    const twoWeeks = new Date(today);
    twoWeeks.setDate(twoWeeks.getDate() + 14);
    
    const exampleData = [
      [
        "Planning Phase", 
        today, 
        nextWeek, 
        "Define project scope and requirements", 
        "Approved project plan document"
      ],
      [
        "Development Phase", 
        nextWeek, 
        twoWeeks, 
        "Build core functionality", 
        "Passing all unit tests and QA approval"
      ]
    ];
    
    // Add example rows
    sheet.getRange(2, 1, exampleData.length, headers.length).setValues(exampleData);
    
    // Format date columns
    sheet.getRange(2, 2, exampleData.length, 2).setNumberFormat("MM/dd/yyyy");
    
    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
    
    log('Created Phases and Dates sheet successfully', LOG_LEVELS.INFO);
    
    return {
      success: true,
      message: "Sheet created successfully"
    };
  }, false);
}