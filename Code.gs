/**
 * Fleet paperwork reminder script for Google Sheets.
 *
 * Rules:
 * 1) If the DATE CELL's FONT color is red -> EXCLUDE that document from notifications.
 * 2) If the DATE CELL's BACKGROUND is yellow -> still notify, but add a warning: "under inspection".
 *
 * IMPORTANT NOTE:
 * If your red/yellow is applied ONLY via Conditional Formatting,
 * Apps Script may NOT reliably read those displayed colors.
 * This works reliably when the color is actually set on the cell.
 *
 * Email sections:
 * - PAST GRACE PERIOD — expired more than 30 days ago
 * - IN GRACE PERIOD   — expired within last 30 days
 * - UNDER INSPECTION  — yellow-highlighted date cells
 * - CLOSE TO EXPIRY   — expires within next 30 days
 *
 * Each line: Reg# — Model — Document Type — days left/overdue
 */

// ============================================================
// MENU & TRIGGER SETUP
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Fleet Reminders")
    .addItem("Send Reminders Now", "sendExpiryReminders")
    .addItem("Send Test Email", "sendTestEmail")
    .addSeparator()
    .addItem("Create Daily Trigger (7 AM)", "createDailyTrigger")
    .addItem("Remove All Triggers", "removeAllTriggers")
    .addToUi();
}

function createDailyTrigger() {
  removeAllTriggers();
  ScriptApp.newTrigger("sendExpiryReminders")
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  SpreadsheetApp.getUi().alert(
    "Daily trigger created. Reminders will run every day at ~7:00 AM."
  );
}

function removeAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function sendTestEmail() {
  var recipients = ["elhelou.ghazi@sarooj.com"]; // must match DEFAULT_RECIPIENTS
  var body = "This is a test email from the Fleet Paperwork Reminder script.\n\n"
    + "If you see this, email notifications are working correctly.\n\n"
    + "You will receive:\n"
    + "- Threshold alerts at 30, 14, 7, and 1 day(s) before expiry\n"
    + "- Overdue alerts for any expired documents\n"
    + "- Daily summary when run by the trigger\n\n"
    + "— Fleet Reminder Bot";
  try {
    GmailApp.sendEmail(recipients[0], "Fleet Reminder — Test Email", body);
    SpreadsheetApp.getUi().alert("Test email sent to " + recipients[0] + ". Check your inbox (and spam folder).");
  } catch (e) {
    SpreadsheetApp.getUi().alert("Failed to send test email: " + e.message);
  }
}

// ============================================================
// HELPERS — Color detection
// ============================================================

function normHex(c) {
  if (!c) return "";
  c = String(c).trim().toLowerCase();
  if (c[0] !== "#") return c;
  if (c.length === 4) c = "#" + c[1] + c[1] + c[2] + c[2] + c[3] + c[3];
  return c;
}

function hexToRgb(hex) {
  hex = normHex(hex);
  if (!hex || hex[0] !== "#" || hex.length !== 7) return null;
  return {
    r: parseInt(hex.slice(1, 3), 16),
    g: parseInt(hex.slice(3, 5), 16),
    b: parseInt(hex.slice(5, 7), 16),
  };
}

function isFontRed(fontHex) {
  const h = normHex(fontHex);
  if (["#ff0000", "#ea4335", "#d93025"].includes(h)) return true;
  const rgb = hexToRgb(h);
  return rgb ? (rgb.r >= 180 && rgb.g <= 80 && rgb.b <= 80) : false;
}

function isBgYellow(bgHex) {
  const h = normHex(bgHex);
  if (["#ffff00", "#fff2cc", "#ffe599", "#fce8b2", "#fff9c4"].includes(h)) return true;
  const rgb = hexToRgb(h);
  return rgb ? (rgb.r >= 200 && rgb.g >= 200 && rgb.b <= 160) : false;
}

// ============================================================
// NOTIFICATION LOG
// ============================================================

/**
 * Appends a row to a "Notification Log" sheet for audit purposes.
 * Creates the sheet + header if it doesn't exist yet.
 */
function logNotification(ss, reg, doc, daysLeft, recipient, status) {
  var LOG_SHEET = "Notification Log";
  var logSheet = ss.getSheetByName(LOG_SHEET);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET);
    logSheet.getRange(1, 1, 1, 6).setValues([
      ["Timestamp", "Reg #", "Document", "Days Left", "Recipient", "Status"]
    ]);
    logSheet.getRange(1, 1, 1, 6).setFontWeight("bold");
    logSheet.setFrozenRows(1);
  }
  logSheet.appendRow([new Date(), reg, doc, daysLeft, recipient, status]);
}

// ============================================================
// MAIN — sendExpiryReminders
// ============================================================

function sendExpiryReminders() {
  // --- Duplicate-run guard ---
  // Prevents double emails if the trigger fires more than once per day.
  // UNCOMMENT the block below when done testing to re-enable once-per-day limit.
  /*
  var props = PropertiesService.getScriptProperties();
  var todayKey = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd");
  var lastRun = props.getProperty("lastRunDate");
  if (lastRun === todayKey) {
    Logger.log("Already ran today (" + todayKey + "). Skipping.");
    return;
  }
  */

  const SHEET_NAME = "SCC"; // <-- CHANGE to your exact tab name
  const GRACE_PERIOD_DAYS = 30; // days after expiry before "past grace period"

  // Emails that receive reminders:
  const DEFAULT_RECIPIENTS = ["elhelou.ghazi@sarooj.com"]; // <-- CHANGE / add more
  const ESCALATION_EMAIL = ""; // optional manager email

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Sheet not found: " + SHEET_NAME);

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  // Normalize headers
  const headers = data[0].map(function(h) { return String(h).replace(/\s+/g, " ").trim(); });
  var idx = function(name) { return headers.indexOf(name); };

  // Your headers
  const iReg = idx("Reg #");
  const iExpiry = idx("Expiry Date");
  const iInspection = idx("Inspection Exp.Date");
  const iModel = idx("Model");

  ["Reg #", "Expiry Date"].forEach(function(col) {
    if (idx(col) === -1) throw new Error("Missing required column: " + col);
  });

  const tz = ss.getSpreadsheetTimeZone();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Pull formatting for the date columns
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const expBg = sheet.getRange(2, iExpiry + 1, lastRow - 1, 1).getBackgrounds();
  const expFont = sheet.getRange(2, iExpiry + 1, lastRow - 1, 1).getFontColors();

  var inspBg = null, inspFont = null;
  if (iInspection !== -1) {
    inspBg = sheet.getRange(2, iInspection + 1, lastRow - 1, 1).getBackgrounds();
    inspFont = sheet.getRange(2, iInspection + 1, lastRow - 1, 1).getFontColors();
  }

  // Collect all items per recipient
  var perRecipient = new Map();
  var addItem = function(email, item) {
    var e = String(email || "").trim();
    if (!e) return;
    if (!perRecipient.has(e)) perRecipient.set(e, []);
    perRecipient.get(e).push(item);
  };

  var evaluateDateField = function(dateRaw, docLabel, reg, model, isUnderInspection) {
    if (!dateRaw) return null;

    var d;
    if (dateRaw instanceof Date) {
      d = dateRaw;
    } else {
      d = new Date(dateRaw);
    }
    if (isNaN(d.getTime())) return null;
    d.setHours(0, 0, 0, 0);

    var daysLeft = Math.floor((d - today) / (1000 * 60 * 60 * 24));

    // Include anything expired or within 30 days of expiry
    if (daysLeft > 30) return null;

    return {
      reg: reg,
      model: model,
      doc: docLabel,
      daysLeft: daysLeft,
      underInspection: isUnderInspection || false
    };
  };

  for (var r = 1; r < data.length; r++) {
    var row = data[r];

    var reg = String(row[iReg] || "").trim();
    if (!reg) continue;

    var model = (iModel !== -1 && row[iModel]) ? String(row[iModel]).trim() : "";

    var rowIdx = r - 1; // formatting arrays start at row 2

    // Rule: red font on the date cell => exclude that document
    var expiryFontIsRed = isFontRed(expFont[rowIdx][0]);
    var expiryBgIsYellow = isBgYellow(expBg[rowIdx][0]);

    var inspFontIsRed = false;
    var inspBgIsYellow = false;
    if (iInspection !== -1) {
      inspFontIsRed = isFontRed(inspFont[rowIdx][0]);
      inspBgIsYellow = isBgYellow(inspBg[rowIdx][0]);
    }

    var expItem = expiryFontIsRed
      ? null
      : evaluateDateField(row[iExpiry], "Expiry Date", reg, model, expiryBgIsYellow);

    var inspItem = (iInspection !== -1 && !inspFontIsRed)
      ? evaluateDateField(row[iInspection], "Inspection Exp.Date", reg, model, inspBgIsYellow)
      : null;

    for (var ri = 0; ri < DEFAULT_RECIPIENTS.length; ri++) {
      var email = DEFAULT_RECIPIENTS[ri];
      if (expItem) addItem(email, expItem);
      if (inspItem) addItem(email, inspItem);
    }
  }

  // --- Build & send emails ---
  for (var entry of perRecipient.entries()) {
    var recipientEmail = entry[0];
    var items = entry[1];

    // Sort into 4 buckets:
    //   1. PAST GRACE PERIOD  — expired more than 30 days ago (daysLeft < -30)
    //   2. IN GRACE PERIOD    — expired within last 30 days (-30 <= daysLeft <= 0)
    //   3. UNDER INSPECTION   — yellow-highlighted cells (any daysLeft)
    //   4. CLOSE TO EXPIRY    — not yet expired, within 30 days (1 <= daysLeft <= 30)
    var pastGrace = [];
    var inGrace = [];
    var underInspection = [];
    var closeToExpiry = [];

    for (var j = 0; j < items.length; j++) {
      var it = items[j];
      if (it.underInspection) {
        underInspection.push(it);
      } else if (it.daysLeft < -30) {
        pastGrace.push(it);
      } else if (it.daysLeft <= 0) {
        inGrace.push(it);
      } else {
        closeToExpiry.push(it);
      }
    }

    // Sort each bucket by urgency (most urgent first)
    pastGrace.sort(function(a, b) { return a.daysLeft - b.daysLeft; });
    inGrace.sort(function(a, b) { return a.daysLeft - b.daysLeft; });
    underInspection.sort(function(a, b) { return a.daysLeft - b.daysLeft; });
    closeToExpiry.sort(function(a, b) { return a.daysLeft - b.daysLeft; });

    var totalItems = items.length;
    var subject = "Fleet Paperwork Reminder (" + totalItems + " items)";

    // Clean format: Reg# — Model — DocType — days info
    var formatLine = function(item) {
      var prefix = "\u2022 " + item.reg + " \u2014 " + (item.model || "N/A") + " \u2014 " + item.doc + " \u2014 ";
      if (item.daysLeft < 0) {
        return prefix + Math.abs(item.daysLeft) + " day(s) overdue";
      } else if (item.daysLeft === 0) {
        return prefix + "Expires today";
      } else {
        return prefix + item.daysLeft + " day(s) left";
      }
    };

    var body = "Fleet Paperwork Status\n";
    body += "======================\n";

    if (pastGrace.length > 0) {
      body += "\n\u274C PAST GRACE PERIOD (>" + GRACE_PERIOD_DAYS + " days overdue) \u2014 " + pastGrace.length + " item(s):\n";
      body += "----------------------------------------------\n";
      for (var p = 0; p < pastGrace.length; p++) body += formatLine(pastGrace[p]) + "\n";
    }

    if (inGrace.length > 0) {
      body += "\n\u26A0\uFE0F IN GRACE PERIOD (0\u2013" + GRACE_PERIOD_DAYS + " days overdue) \u2014 " + inGrace.length + " item(s):\n";
      body += "----------------------------------------------\n";
      for (var g = 0; g < inGrace.length; g++) body += formatLine(inGrace[g]) + "\n";
    }

    if (underInspection.length > 0) {
      body += "\n\uD83D\uDD0D UNDER INSPECTION \u2014 " + underInspection.length + " item(s):\n";
      body += "----------------------------------------------\n";
      for (var u = 0; u < underInspection.length; u++) body += formatLine(underInspection[u]) + "\n";
    }

    if (closeToExpiry.length > 0) {
      body += "\n\uD83D\uDCC5 CLOSE TO EXPIRY (within 30 days) \u2014 " + closeToExpiry.length + " item(s):\n";
      body += "----------------------------------------------\n";
      for (var c = 0; c < closeToExpiry.length; c++) body += formatLine(closeToExpiry[c]) + "\n";
    }

    body += "\n\u2014 Automated Fleet Reminder";

    try {
      GmailApp.sendEmail(recipientEmail, subject, body);
      for (var li = 0; li < items.length; li++) {
        logNotification(ss, items[li].reg, items[li].doc, items[li].daysLeft, recipientEmail, "Sent");
      }
    } catch (e) {
      Logger.log("Failed to send to " + recipientEmail + ": " + e.message);
      for (var li2 = 0; li2 < items.length; li2++) {
        logNotification(ss, items[li2].reg, items[li2].doc, items[li2].daysLeft, recipientEmail, "FAILED: " + e.message);
      }
    }

    // Optional escalation for anything past grace period
    if (ESCALATION_EMAIL && pastGrace.length > 0) {
      var escSubject = "PAST GRACE PERIOD \u2014 " + pastGrace.length + " fleet item(s)";
      var escBody = "These items are past the " + GRACE_PERIOD_DAYS + "-day grace period:\n\n";
      for (var ep = 0; ep < pastGrace.length; ep++) escBody += formatLine(pastGrace[ep]) + "\n";
      try {
        GmailApp.sendEmail(ESCALATION_EMAIL, escSubject, escBody);
      } catch (e) {
        Logger.log("Failed to send escalation to " + ESCALATION_EMAIL + ": " + e.message);
      }
    }
  }

  // Mark today as completed so duplicate runs are skipped
  // UNCOMMENT when done testing:
  // props.setProperty("lastRunDate", todayKey);
}
