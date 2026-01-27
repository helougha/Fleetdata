// ============================================================
// MACHINE PAPERWORK EXPIRY TRACKER - Google Apps Script
// ============================================================
// Tracks expiry dates for machine licenses, mulkiyas, insurance,
// and other construction site paperwork. Calculates days until
// expiry, color-codes rows by urgency, and sends email/WhatsApp
// alerts at 30/14/7/1 day thresholds plus a daily digest.
// ============================================================

// ----- CONFIGURATION -----
var CONFIG = {
  SHEET_NAME: "Machine Tracker",
  CONFIG_SHEET_NAME: "Config",
  LOG_SHEET_NAME: "Notification Log",

  // Column indices (0-based) in Machine Tracker sheet
  COL_MACHINE_ID: 0,       // A
  COL_MACHINE_NAME: 1,     // B
  COL_DOCUMENT_TYPE: 2,    // C
  COL_EXPIRY_DATE: 3,      // D
  COL_DAYS_REMAINING: 4,   // E
  COL_STATUS: 5,           // F
  COL_LAST_NOTIFIED: 6,    // G
  COL_RENEWAL_REQS: 7,     // H

  HEADER_ROW: 1,
  DATA_START_ROW: 2,

  // Alert thresholds in days
  ALERT_THRESHOLDS: [30, 14, 7, 1],

  // Status labels
  STATUS: {
    OK: "OK",
    WARNING: "WARNING",
    URGENT: "URGENT",
    CRITICAL: "CRITICAL",
    EXPIRED: "EXPIRED"
  },

  // Status colors (hex)
  COLORS: {
    OK: "#C6EFCE",
    WARNING: "#FFEB9C",
    URGENT: "#FFC7CE",
    CRITICAL: "#FF4444",
    EXPIRED: "#333333",
    HEADER: "#1F4E79",
    HEADER_FONT: "#FFFFFF"
  }
};


// ============================================================
// MENU & SETUP
// ============================================================

/**
 * Creates a custom menu when the spreadsheet opens.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Paperwork Tracker")
    .addItem("Run Daily Update Now", "dailyUpdate")
    .addItem("Send Test Email", "sendTestEmail")
    .addSeparator()
    .addItem("Setup Tracker (First Time)", "setupTracker")
    .addItem("Create Daily Trigger", "createDailyTrigger")
    .addItem("Remove All Triggers", "removeAllTriggers")
    .addToUi();
}

/**
 * First-time setup: creates sheets, headers, formatting, and Config.
 */
function setupTracker() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Machine Tracker sheet ---
  var tracker = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
  var headers = [
    "Machine ID", "Machine Name", "Document Type",
    "Expiry Date", "Days Remaining", "Status",
    "Last Notified", "Renewal Requirements"
  ];

  var headerRange = tracker.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontColor(CONFIG.COLORS.HEADER_FONT)
    .setHorizontalAlignment("center");

  // Column widths
  tracker.setColumnWidth(1, 120);  // Machine ID
  tracker.setColumnWidth(2, 200);  // Machine Name
  tracker.setColumnWidth(3, 180);  // Document Type
  tracker.setColumnWidth(4, 140);  // Expiry Date
  tracker.setColumnWidth(5, 130);  // Days Remaining
  tracker.setColumnWidth(6, 120);  // Status
  tracker.setColumnWidth(7, 140);  // Last Notified
  tracker.setColumnWidth(8, 300);  // Renewal Requirements

  // Format Expiry Date column as dates
  tracker.getRange("D:D").setNumberFormat("dd/MM/yyyy");
  tracker.getRange("G:G").setNumberFormat("dd/MM/yyyy");

  // Freeze header row
  tracker.setFrozenRows(1);

  // --- Config sheet ---
  var config = getOrCreateSheet(ss, CONFIG.CONFIG_SHEET_NAME);
  var configData = [
    ["Setting", "Value", "Description"],
    ["Email Address", "", "Your email address for notifications"],
    ["Send Daily Summary", "TRUE", "Send daily digest of upcoming expiries"],
    ["Alert at 30 Days", "TRUE", "Send alert when 30 days remain"],
    ["Alert at 14 Days", "TRUE", "Send alert when 14 days remain"],
    ["Alert at 7 Days", "TRUE", "Send alert when 7 days remain"],
    ["Alert at 1 Day", "TRUE", "Send alert when 1 day remains"],
    ["WhatsApp Phone", "", "Phone number with country code (e.g. +971501234567)"],
    ["WhatsApp API Key", "", "CallMeBot API key (see SETUP.md)"],
    ["WhatsApp Method", "callmebot", "callmebot or twilio"],
    ["Twilio Account SID", "", "Twilio Account SID (if using Twilio)"],
    ["Twilio Auth Token", "", "Twilio Auth Token (if using Twilio)"],
    ["Twilio From Number", "", "Twilio WhatsApp sender (e.g. whatsapp:+14155238886)"]
  ];

  config.getRange(1, 1, configData.length, 3).setValues(configData);
  config.getRange(1, 1, 1, 3)
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontColor(CONFIG.COLORS.HEADER_FONT);

  config.setColumnWidth(1, 200);
  config.setColumnWidth(2, 300);
  config.setColumnWidth(3, 400);
  config.setFrozenRows(1);

  // --- Notification Log sheet ---
  var log = getOrCreateSheet(ss, CONFIG.LOG_SHEET_NAME);
  var logHeaders = ["Timestamp", "Machine", "Document", "Days Left", "Method", "Status"];
  log.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);
  log.getRange(1, 1, 1, logHeaders.length)
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER)
    .setFontColor(CONFIG.COLORS.HEADER_FONT);
  log.setFrozenRows(1);

  // Add sample data to Machine Tracker
  var sampleData = [
    ["EQ-001", "CAT 320 Excavator", "Mulkiya", addDays(new Date(), 28), "", "", "", "Mulkiya card, insurance certificate, inspection report"],
    ["EQ-001", "CAT 320 Excavator", "Insurance", addDays(new Date(), 60), "", "", "", "Insurance renewal form, vehicle details"],
    ["EQ-002", "Komatsu PC200", "Operating License", addDays(new Date(), 7), "", "", "", "Operator certificate, machine inspection, safety check"],
    ["EQ-003", "Liebherr Crane LTM", "Mulkiya", addDays(new Date(), 2), "", "", "", "Mulkiya card, crane load test certificate"],
    ["EQ-003", "Liebherr Crane LTM", "Third Party Inspection", addDays(new Date(), 14), "", "", "", "Third party inspection form, previous certificate"],
    ["EQ-004", "JCB Backhoe 3CX", "Insurance", addDays(new Date(), 45), "", "", "", "Insurance policy, machine registration"]
  ];
  tracker.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);

  // Run the first update
  updateDaysRemaining();

  SpreadsheetApp.getUi().alert(
    "Setup Complete!\n\n" +
    "1. Go to the Config sheet and enter your email address in cell B2.\n" +
    "2. Replace sample data with your actual machines.\n" +
    "3. Use Paperwork Tracker > Create Daily Trigger to automate.\n" +
    "4. Use Paperwork Tracker > Send Test Email to verify notifications."
  );
}


// ============================================================
// DAILY UPDATE - CORE LOGIC
// ============================================================

/**
 * Main function: updates days remaining, statuses, and sends notifications.
 * Called daily by a time-based trigger.
 */
function dailyUpdate() {
  updateDaysRemaining();
  sendThresholdAlerts();
  sendDailySummary();
}

/**
 * Recalculates days remaining and status for every row.
 */
function updateDaysRemaining() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return;

  var dataRange = sheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.HEADER_ROW, 8);
  var data = dataRange.getValues();
  var today = stripTime(new Date());

  for (var i = 0; i < data.length; i++) {
    var expiryDate = data[i][CONFIG.COL_EXPIRY_DATE];
    if (!expiryDate || !(expiryDate instanceof Date)) continue;

    var expiry = stripTime(new Date(expiryDate));
    var daysLeft = Math.ceil((expiry - today) / (1000 * 60 * 60 * 24));

    // Update Days Remaining
    data[i][CONFIG.COL_DAYS_REMAINING] = daysLeft;

    // Update Status
    var status;
    if (daysLeft <= 0) {
      status = CONFIG.STATUS.EXPIRED;
    } else if (daysLeft <= 7) {
      status = CONFIG.STATUS.CRITICAL;
    } else if (daysLeft <= 14) {
      status = CONFIG.STATUS.URGENT;
    } else if (daysLeft <= 30) {
      status = CONFIG.STATUS.WARNING;
    } else {
      status = CONFIG.STATUS.OK;
    }
    data[i][CONFIG.COL_STATUS] = status;
  }

  // Write updated data back
  dataRange.setValues(data);

  // Apply color coding
  applyColorCoding(sheet, CONFIG.DATA_START_ROW, lastRow);
}

/**
 * Color-codes rows based on their status.
 */
function applyColorCoding(sheet, startRow, lastRow) {
  for (var row = startRow; row <= lastRow; row++) {
    var status = sheet.getRange(row, CONFIG.COL_STATUS + 1).getValue();
    var rowRange = sheet.getRange(row, 1, 1, 8);
    var fontColor = "#000000";

    switch (status) {
      case CONFIG.STATUS.OK:
        rowRange.setBackground(CONFIG.COLORS.OK);
        break;
      case CONFIG.STATUS.WARNING:
        rowRange.setBackground(CONFIG.COLORS.WARNING);
        break;
      case CONFIG.STATUS.URGENT:
        rowRange.setBackground(CONFIG.COLORS.URGENT);
        break;
      case CONFIG.STATUS.CRITICAL:
        rowRange.setBackground(CONFIG.COLORS.CRITICAL);
        fontColor = "#FFFFFF";
        break;
      case CONFIG.STATUS.EXPIRED:
        rowRange.setBackground(CONFIG.COLORS.EXPIRED);
        fontColor = "#FFFFFF";
        break;
      default:
        rowRange.setBackground("#FFFFFF");
    }
    rowRange.setFontColor(fontColor);
  }
}


// ============================================================
// THRESHOLD ALERTS (30 / 14 / 7 / 1 days)
// ============================================================

/**
 * Sends email/WhatsApp alerts when a document hits an exact threshold.
 * Uses the "Last Notified" column to avoid duplicate alerts.
 */
function sendThresholdAlerts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var configValues = getConfigValues(ss);

  if (!configValues.email) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return;

  var data = sheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.HEADER_ROW, 8).getValues();
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  for (var i = 0; i < data.length; i++) {
    var daysLeft = data[i][CONFIG.COL_DAYS_REMAINING];
    var lastNotified = data[i][CONFIG.COL_LAST_NOTIFIED];
    var lastNotifiedStr = "";

    if (lastNotified instanceof Date) {
      lastNotifiedStr = Utilities.formatDate(lastNotified, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    // Skip if already notified today
    if (lastNotifiedStr === today) continue;

    // Check each threshold
    var thresholdMap = {
      30: configValues.alert30,
      14: configValues.alert14,
      7: configValues.alert7,
      1: configValues.alert1
    };

    for (var t = 0; t < CONFIG.ALERT_THRESHOLDS.length; t++) {
      var threshold = CONFIG.ALERT_THRESHOLDS[t];
      if (daysLeft === threshold && thresholdMap[threshold]) {
        var machineId = data[i][CONFIG.COL_MACHINE_ID];
        var machineName = data[i][CONFIG.COL_MACHINE_NAME];
        var docType = data[i][CONFIG.COL_DOCUMENT_TYPE];
        var expiryDate = data[i][CONFIG.COL_EXPIRY_DATE];
        var renewalReqs = data[i][CONFIG.COL_RENEWAL_REQS];

        var subject = getAlertEmoji(daysLeft) + " ALERT: " + docType + " for " + machineName + " expires in " + daysLeft + " day(s)";
        var body = buildThresholdEmailBody(machineId, machineName, docType, expiryDate, daysLeft, renewalReqs);

        // Send email
        sendEmail(configValues.email, subject, body);

        // Send WhatsApp if configured
        if (configValues.whatsappPhone && configValues.whatsappApiKey) {
          var whatsappMsg = getAlertEmoji(daysLeft) + " " + docType + " for " + machineName +
            " (" + machineId + ") expires in " + daysLeft + " day(s)! Renew ASAP.";
          sendWhatsApp(configValues, whatsappMsg);
        }

        // Mark as notified
        var notifyCell = sheet.getRange(CONFIG.DATA_START_ROW + i, CONFIG.COL_LAST_NOTIFIED + 1);
        notifyCell.setValue(new Date());

        // Log notification
        logNotification(ss, machineName, docType, daysLeft, "Email");

        break; // Only send one alert per document per day
      }
    }
  }
}

/**
 * Builds HTML email body for a threshold alert.
 */
function buildThresholdEmailBody(machineId, machineName, docType, expiryDate, daysLeft, renewalReqs) {
  var expiryStr = Utilities.formatDate(new Date(expiryDate), Session.getScriptTimeZone(), "dd/MM/yyyy");
  var urgency = getUrgencyLevel(daysLeft);

  var html = '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">';
  html += '<div style="background-color: ' + getUrgencyColor(daysLeft) + '; color: white; padding: 20px; border-radius: 8px 8px 0 0;">';
  html += '<h2 style="margin: 0;">' + getAlertEmoji(daysLeft) + ' ' + urgency + ' - Document Expiry Alert</h2>';
  html += '</div>';
  html += '<div style="padding: 20px; border: 1px solid #ddd; border-top: none; border-radius: 0 0 8px 8px;">';
  html += '<table style="width: 100%; border-collapse: collapse;">';
  html += '<tr><td style="padding: 8px; font-weight: bold; width: 40%;">Machine ID:</td><td style="padding: 8px;">' + machineId + '</td></tr>';
  html += '<tr><td style="padding: 8px; font-weight: bold;">Machine Name:</td><td style="padding: 8px;">' + machineName + '</td></tr>';
  html += '<tr><td style="padding: 8px; font-weight: bold;">Document Type:</td><td style="padding: 8px;">' + docType + '</td></tr>';
  html += '<tr><td style="padding: 8px; font-weight: bold;">Expiry Date:</td><td style="padding: 8px;">' + expiryStr + '</td></tr>';
  html += '<tr><td style="padding: 8px; font-weight: bold;">Days Remaining:</td><td style="padding: 8px; font-size: 24px; font-weight: bold; color: ' + getUrgencyColor(daysLeft) + ';">' + daysLeft + '</td></tr>';
  html += '</table>';

  if (renewalReqs) {
    html += '<div style="margin-top: 16px; padding: 12px; background: #f5f5f5; border-left: 4px solid ' + getUrgencyColor(daysLeft) + '; border-radius: 4px;">';
    html += '<strong>Renewal Requirements:</strong><br>' + renewalReqs;
    html += '</div>';
  }

  html += '<p style="margin-top: 20px; color: #666; font-size: 12px;">This is an automated alert from your Machine Paperwork Tracker.</p>';
  html += '</div></div>';

  return html;
}


// ============================================================
// DAILY SUMMARY EMAIL
// ============================================================

/**
 * Sends a daily digest email with all documents expiring within 30 days.
 */
function sendDailySummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configValues = getConfigValues(ss);

  if (!configValues.email || !configValues.dailySummary) return;

  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return;

  var data = sheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.HEADER_ROW, 8).getValues();
  var today = stripTime(new Date());
  var todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd/MM/yyyy");

  var expired = [];
  var upcoming = [];

  for (var i = 0; i < data.length; i++) {
    var expiryDate = data[i][CONFIG.COL_EXPIRY_DATE];
    if (!expiryDate || !(expiryDate instanceof Date)) continue;

    var daysLeft = data[i][CONFIG.COL_DAYS_REMAINING];

    var item = {
      machineId: data[i][CONFIG.COL_MACHINE_ID],
      machineName: data[i][CONFIG.COL_MACHINE_NAME],
      docType: data[i][CONFIG.COL_DOCUMENT_TYPE],
      expiryDate: Utilities.formatDate(new Date(expiryDate), Session.getScriptTimeZone(), "dd/MM/yyyy"),
      daysLeft: daysLeft,
      renewalReqs: data[i][CONFIG.COL_RENEWAL_REQS],
      status: data[i][CONFIG.COL_STATUS]
    };

    if (daysLeft <= 0) {
      expired.push(item);
    } else if (daysLeft <= 30) {
      upcoming.push(item);
    }
  }

  // Only send if there are items to report
  if (expired.length === 0 && upcoming.length === 0) return;

  // Sort upcoming by days remaining (most urgent first)
  upcoming.sort(function (a, b) { return a.daysLeft - b.daysLeft; });

  var subject = "Daily Paperwork Summary - " + todayStr + " (" + (expired.length + upcoming.length) + " items need attention)";
  var body = buildDailySummaryEmail(todayStr, expired, upcoming);

  sendEmail(configValues.email, subject, body);

  // Send WhatsApp summary if configured
  if (configValues.whatsappPhone && configValues.whatsappApiKey) {
    var waMsg = "Daily Summary " + todayStr + ": ";
    if (expired.length > 0) waMsg += expired.length + " EXPIRED. ";
    waMsg += upcoming.length + " expiring within 30 days.";
    if (upcoming.length > 0) {
      waMsg += " Most urgent: " + upcoming[0].docType + " for " + upcoming[0].machineName + " (" + upcoming[0].daysLeft + " days).";
    }
    sendWhatsApp(configValues, waMsg);
  }
}

/**
 * Builds HTML email body for the daily summary.
 */
function buildDailySummaryEmail(todayStr, expired, upcoming) {
  var html = '<div style="font-family: Arial, sans-serif; max-width: 700px; margin: 0 auto;">';

  // Header
  html += '<div style="background: linear-gradient(135deg, #1F4E79, #2980B9); color: white; padding: 20px; border-radius: 8px 8px 0 0;">';
  html += '<h2 style="margin: 0;">Fleet Paperwork Daily Summary</h2>';
  html += '<p style="margin: 5px 0 0; opacity: 0.9;">' + todayStr + '</p>';
  html += '</div>';

  html += '<div style="padding: 20px; border: 1px solid #ddd; border-top: none; border-radius: 0 0 8px 8px;">';

  // Expired section
  if (expired.length > 0) {
    html += '<h3 style="color: #FF4444; border-bottom: 2px solid #FF4444; padding-bottom: 8px;">EXPIRED (' + expired.length + ')</h3>';
    html += buildSummaryTable(expired, "#FF4444");
  }

  // Upcoming section
  if (upcoming.length > 0) {
    html += '<h3 style="color: #E67E22; border-bottom: 2px solid #E67E22; padding-bottom: 8px; margin-top: 24px;">EXPIRING WITHIN 30 DAYS (' + upcoming.length + ')</h3>';
    html += buildSummaryTable(upcoming, "#E67E22");
  }

  html += '<p style="margin-top: 20px; color: #666; font-size: 12px;">This is an automated daily summary from your Machine Paperwork Tracker.</p>';
  html += '</div></div>';

  return html;
}

/**
 * Builds an HTML table for the daily summary email.
 */
function buildSummaryTable(items, accentColor) {
  var html = '<table style="width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 14px;">';
  html += '<tr style="background: #f8f8f8;">';
  html += '<th style="padding: 10px; text-align: left; border-bottom: 2px solid #ddd;">Machine</th>';
  html += '<th style="padding: 10px; text-align: left; border-bottom: 2px solid #ddd;">Document</th>';
  html += '<th style="padding: 10px; text-align: center; border-bottom: 2px solid #ddd;">Expiry</th>';
  html += '<th style="padding: 10px; text-align: center; border-bottom: 2px solid #ddd;">Days Left</th>';
  html += '<th style="padding: 10px; text-align: left; border-bottom: 2px solid #ddd;">Renewal Requirements</th>';
  html += '</tr>';

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var rowBg = i % 2 === 0 ? "#ffffff" : "#f9f9f9";
    var daysColor = getUrgencyColor(item.daysLeft);

    html += '<tr style="background: ' + rowBg + ';">';
    html += '<td style="padding: 8px; border-bottom: 1px solid #eee;">' + item.machineId + '<br><small>' + item.machineName + '</small></td>';
    html += '<td style="padding: 8px; border-bottom: 1px solid #eee;">' + item.docType + '</td>';
    html += '<td style="padding: 8px; text-align: center; border-bottom: 1px solid #eee;">' + item.expiryDate + '</td>';
    html += '<td style="padding: 8px; text-align: center; border-bottom: 1px solid #eee; font-weight: bold; color: ' + daysColor + ';">' + item.daysLeft + '</td>';
    html += '<td style="padding: 8px; border-bottom: 1px solid #eee; font-size: 12px;">' + (item.renewalReqs || "-") + '</td>';
    html += '</tr>';
  }

  html += '</table>';
  return html;
}


// ============================================================
// EMAIL & WHATSAPP SENDING
// ============================================================

/**
 * Sends an HTML email using Gmail.
 */
function sendEmail(to, subject, htmlBody) {
  try {
    GmailApp.sendEmail(to, subject, "", { htmlBody: htmlBody });
  } catch (e) {
    Logger.log("Email send failed: " + e.message);
  }
}

/**
 * Sends a WhatsApp message using CallMeBot or Twilio.
 */
function sendWhatsApp(configValues, message) {
  try {
    var method = (configValues.whatsappMethod || "callmebot").toLowerCase();

    if (method === "callmebot") {
      sendWhatsAppCallMeBot(configValues.whatsappPhone, configValues.whatsappApiKey, message);
    } else if (method === "twilio") {
      sendWhatsAppTwilio(configValues, message);
    }
  } catch (e) {
    Logger.log("WhatsApp send failed: " + e.message);
  }
}

/**
 * Sends WhatsApp via CallMeBot free API.
 */
function sendWhatsAppCallMeBot(phone, apiKey, message) {
  var encodedMsg = encodeURIComponent(message);
  var url = "https://api.callmebot.com/whatsapp.php?phone=" + phone +
    "&text=" + encodedMsg + "&apikey=" + apiKey;

  var options = {
    method: "get",
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  Logger.log("CallMeBot response: " + response.getContentText());
}

/**
 * Sends WhatsApp via Twilio API.
 */
function sendWhatsAppTwilio(configValues, message) {
  var accountSid = configValues.twilioSid;
  var authToken = configValues.twilioToken;
  var fromNumber = configValues.twilioFrom;
  var toNumber = "whatsapp:" + configValues.whatsappPhone;

  var url = "https://api.twilio.com/2010-04-01/Accounts/" + accountSid + "/Messages.json";

  var options = {
    method: "post",
    headers: {
      "Authorization": "Basic " + Utilities.base64Encode(accountSid + ":" + authToken)
    },
    payload: {
      "To": toNumber,
      "From": fromNumber,
      "Body": message
    },
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  Logger.log("Twilio response: " + response.getContentText());
}

/**
 * Sends a test email so the user can verify notifications work.
 */
function sendTestEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configValues = getConfigValues(ss);

  if (!configValues.email) {
    SpreadsheetApp.getUi().alert("Please set your email address in the Config sheet (cell B2) first.");
    return;
  }

  var subject = "Test - Machine Paperwork Tracker";
  var body = '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">';
  body += '<div style="background: #27AE60; color: white; padding: 20px; border-radius: 8px 8px 0 0;">';
  body += '<h2 style="margin: 0;">Test Email Successful!</h2>';
  body += '</div>';
  body += '<div style="padding: 20px; border: 1px solid #ddd; border-top: none; border-radius: 0 0 8px 8px;">';
  body += '<p>Your Machine Paperwork Tracker email notifications are working correctly.</p>';
  body += '<p>You will receive:</p>';
  body += '<ul>';
  body += '<li><strong>Threshold alerts</strong> at 30, 14, 7, and 1 day(s) before expiry</li>';
  body += '<li><strong>Daily summary</strong> of all documents expiring within 30 days</li>';
  body += '</ul>';
  body += '<p style="color: #666; font-size: 12px;">Sent from your Machine Paperwork Tracker.</p>';
  body += '</div></div>';

  sendEmail(configValues.email, subject, body);

  SpreadsheetApp.getUi().alert("Test email sent to " + configValues.email + ".\nCheck your inbox (and spam folder).");
}


// ============================================================
// TRIGGERS
// ============================================================

/**
 * Creates a daily trigger to run at 7:00 AM.
 */
function createDailyTrigger() {
  // Remove existing triggers first
  removeAllTriggers();

  ScriptApp.newTrigger("dailyUpdate")
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  SpreadsheetApp.getUi().alert(
    "Daily trigger created!\n\nThe tracker will run every day at approximately 7:00 AM.\n" +
    "It will:\n" +
    "1. Update all days remaining and statuses\n" +
    "2. Send threshold alerts (30/14/7/1 day)\n" +
    "3. Send daily summary email"
  );
}

/**
 * Removes all triggers for this project.
 */
function removeAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}


// ============================================================
// HELPERS
// ============================================================

/**
 * Returns a sheet by name, creating it if it doesn't exist.
 */
function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * Reads all configuration values from the Config sheet.
 */
function getConfigValues(ss) {
  var configSheet = ss.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
  if (!configSheet) {
    return { email: "", dailySummary: true, alert30: true, alert14: true, alert7: true, alert1: true };
  }

  var data = configSheet.getRange("A2:B13").getValues();
  var configMap = {};
  for (var i = 0; i < data.length; i++) {
    configMap[data[i][0]] = data[i][1];
  }

  return {
    email: configMap["Email Address"] || "",
    dailySummary: configMap["Send Daily Summary"] === true || configMap["Send Daily Summary"] === "TRUE",
    alert30: configMap["Alert at 30 Days"] === true || configMap["Alert at 30 Days"] === "TRUE",
    alert14: configMap["Alert at 14 Days"] === true || configMap["Alert at 14 Days"] === "TRUE",
    alert7: configMap["Alert at 7 Days"] === true || configMap["Alert at 7 Days"] === "TRUE",
    alert1: configMap["Alert at 1 Day"] === true || configMap["Alert at 1 Day"] === "TRUE",
    whatsappPhone: configMap["WhatsApp Phone"] || "",
    whatsappApiKey: configMap["WhatsApp API Key"] || "",
    whatsappMethod: configMap["WhatsApp Method"] || "callmebot",
    twilioSid: configMap["Twilio Account SID"] || "",
    twilioToken: configMap["Twilio Auth Token"] || "",
    twilioFrom: configMap["Twilio From Number"] || ""
  };
}

/**
 * Logs a notification to the Notification Log sheet.
 */
function logNotification(ss, machine, document, daysLeft, method) {
  var logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!logSheet) return;

  logSheet.appendRow([
    new Date(),
    machine,
    document,
    daysLeft,
    method,
    "Sent"
  ]);
}

/**
 * Returns urgency level text based on days remaining.
 */
function getUrgencyLevel(daysLeft) {
  if (daysLeft <= 0) return "EXPIRED";
  if (daysLeft <= 1) return "CRITICAL";
  if (daysLeft <= 7) return "CRITICAL";
  if (daysLeft <= 14) return "URGENT";
  if (daysLeft <= 30) return "WARNING";
  return "OK";
}

/**
 * Returns a color hex code based on urgency.
 */
function getUrgencyColor(daysLeft) {
  if (daysLeft <= 0) return "#333333";
  if (daysLeft <= 7) return "#FF4444";
  if (daysLeft <= 14) return "#E67E22";
  if (daysLeft <= 30) return "#F39C12";
  return "#27AE60";
}

/**
 * Returns an alert emoji based on days remaining.
 */
function getAlertEmoji(daysLeft) {
  if (daysLeft <= 1) return "\u{1F6A8}";  // rotating light
  if (daysLeft <= 7) return "\u{26A0}\u{FE0F}";   // warning
  if (daysLeft <= 14) return "\u{1F536}";  // orange diamond
  return "\u{1F514}";                       // bell
}

/**
 * Strips the time component from a Date, returning midnight.
 */
function stripTime(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

/**
 * Returns a new Date that is `days` days from the given date.
 */
function addDays(date, days) {
  var result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}
