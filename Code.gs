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
 * Adjustments applied:
 * - Window-based threshold checking (catches alerts even if trigger skips a day)
 * - Math.floor instead of Math.round for DST-safe day counting
 * - try/catch around each email send so one failure doesn't abort the rest
 * - Duplicate-run guard using ScriptProperties (prevents double emails if trigger fires twice)
 * - Notification log sheet for audit trail
 * - onOpen menu + trigger helper functions
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
  var props = PropertiesService.getScriptProperties();
  var todayKey = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd");
  var lastRun = props.getProperty("lastRunDate");
  if (lastRun === todayKey) {
    Logger.log("Already ran today (" + todayKey + "). Skipping.");
    return;
  }

  const SHEET_NAME = "SCC"; // <-- CHANGE to your exact tab name

  // --- ADJUSTMENT 1: Window-based thresholds ---
  // Instead of only alerting on exact days [30, 14, 7, 1], we alert when
  // daysLeft falls within a window. This catches alerts even if the trigger
  // misses a day. The windows are: <=30 & >14, <=14 & >7, <=7 & >1, <=1 & >0.
  // Each document is tagged with the highest threshold it crossed today.
  const THRESHOLDS = [30, 14, 7, 1]; // sorted descending

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
  const iPlant = idx("Plant Code :");
  const iLocation = idx("Location");

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

  // Group reminders per recipient
  var perRecipient = new Map();
  var addItem = function(email, item) {
    var e = String(email || "").trim();
    if (!e) return;
    if (!perRecipient.has(e)) perRecipient.set(e, []);
    perRecipient.get(e).push(item);
  };

  // Track "under inspection" warnings per recipient
  var underInspectionByRecipient = new Map();
  var addUnderInspection = function(email, msg) {
    var e = String(email || "").trim();
    if (!e) return;
    if (!underInspectionByRecipient.has(e)) underInspectionByRecipient.set(e, []);
    underInspectionByRecipient.get(e).push(msg);
  };

  /**
   * Determines the active threshold bracket for a given daysLeft value.
   * Returns the threshold number (30, 14, 7, 1) or null if daysLeft > 30.
   * This is the window-based replacement for the old exact-match check.
   */
  var getThresholdBracket = function(daysLeft) {
    if (daysLeft < 0) return null; // overdue handled separately
    for (var t = 0; t < THRESHOLDS.length; t++) {
      if (daysLeft <= THRESHOLDS[t]) {
        // daysLeft is within this threshold window
        // but we want the tightest matching bracket
        continue;
      } else {
        // daysLeft > THRESHOLDS[t], so the previous threshold was the bracket
        return t > 0 ? THRESHOLDS[t - 1] : null;
      }
    }
    // daysLeft <= smallest threshold (1)
    return THRESHOLDS[THRESHOLDS.length - 1];
  };

  var evaluateDateField = function(dateRaw, docLabel, reg, model, plant, location) {
    if (!dateRaw) return null;

    var d;
    if (dateRaw instanceof Date) {
      d = dateRaw;
    } else {
      d = new Date(dateRaw);
    }
    if (isNaN(d.getTime())) return null;
    d.setHours(0, 0, 0, 0);

    // ADJUSTMENT 3: Math.floor instead of Math.round for DST safety
    var daysLeft = Math.floor((d - today) / (1000 * 60 * 60 * 24));

    var isOverdue = daysLeft < 0;
    var bracket = getThresholdBracket(daysLeft);

    // ADJUSTMENT 1: notify if within any threshold bracket OR overdue
    if (!bracket && !isOverdue) return null;

    return {
      reg: reg,
      model: model,
      plant: plant,
      location: location,
      doc: docLabel,
      date: Utilities.formatDate(d, tz, "yyyy-MM-dd"),
      daysLeft: daysLeft,
      bracket: bracket // which threshold window this falls in (30, 14, 7, 1, or null if overdue)
    };
  };

  for (var r = 1; r < data.length; r++) {
    var row = data[r];

    var reg = String(row[iReg] || "").trim();
    if (!reg) continue;

    var model = (iModel !== -1 && row[iModel]) ? String(row[iModel]).trim() : "";
    var plant = (iPlant !== -1 && row[iPlant]) ? String(row[iPlant]).trim() : "";
    var location = (iLocation !== -1 && row[iLocation]) ? String(row[iLocation]).trim() : "";

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
      : evaluateDateField(row[iExpiry], "Expiry Date", reg, model, plant, location);

    var inspItem = (iInspection !== -1 && !inspFontIsRed)
      ? evaluateDateField(row[iInspection], "Inspection Exp.Date", reg, model, plant, location)
      : null;

    // Send everything to the default recipients list
    for (var ri = 0; ri < DEFAULT_RECIPIENTS.length; ri++) {
      var email = DEFAULT_RECIPIENTS[ri];
      if (expItem) {
        addItem(email, expItem);
        if (expiryBgIsYellow) {
          addUnderInspection(email, "Reg: " + reg + " \u2014 Expiry Date (" + expItem.date + ") \u2014 UNDER INSPECTION");
        }
      }
      if (inspItem) {
        addItem(email, inspItem);
        if (inspBgIsYellow) {
          addUnderInspection(email, "Reg: " + reg + " \u2014 Inspection Exp.Date (" + inspItem.date + ") \u2014 UNDER INSPECTION");
        }
      }
    }
  }

  // --- Send emails ---
  for (var entry of perRecipient.entries()) {
    var recipientEmail = entry[0];
    var items = entry[1];

    items.sort(function(a, b) { return a.daysLeft - b.daysLeft; }); // overdue first

    var subject = "Fleet paperwork reminders (" + items.length + ")";

    var lines = items.map(function(it) {
      var tag = it.daysLeft < 0
        ? "OVERDUE by " + Math.abs(it.daysLeft) + " day(s)"
        : it.daysLeft + " day(s) left";

      var meta = [
        it.model ? "Model: " + it.model : null,
        it.plant ? "Plant: " + it.plant : null,
        it.location ? "Location: " + it.location : null
      ].filter(Boolean).join(" | ");

      return "\u2022 Reg: " + it.reg + " \u2014 " + it.doc + " \u2014 Date: " + it.date + " \u2014 " + tag + (meta ? " \u2014 " + meta : "");
    });

    var warnings = underInspectionByRecipient.get(recipientEmail) || [];
    var warningBlock = warnings.length
      ? "\n\n\u26A0 UNDER INSPECTION (yellow-highlighted):\n" + warnings.map(function(w) { return "- " + w; }).join("\n")
      : "";

    var body =
      "These fleet items need action:\n\n" +
      lines.join("\n") +
      "\n\nReminder windows: 30 / 14 / 7 / 1 days + overdue." +
      "\n\u2014 Automated reminder" + warningBlock;

    // ADJUSTMENT 2: try/catch so one failure doesn't abort all emails
    try {
      GmailApp.sendEmail(recipientEmail, subject, body);
      // Log each item that was notified
      for (var li = 0; li < items.length; li++) {
        logNotification(ss, items[li].reg, items[li].doc, items[li].daysLeft, recipientEmail, "Sent");
      }
    } catch (e) {
      Logger.log("Failed to send to " + recipientEmail + ": " + e.message);
      for (var li2 = 0; li2 < items.length; li2++) {
        logNotification(ss, items[li2].reg, items[li2].doc, items[li2].daysLeft, recipientEmail, "FAILED: " + e.message);
      }
    }

    // Optional escalation to manager if overdue items exist
    if (ESCALATION_EMAIL) {
      var overdue = items.filter(function(x) { return x.daysLeft < 0; });
      if (overdue.length) {
        var escSubject = "OVERDUE fleet paperwork (" + overdue.length + ") \u2014 " + recipientEmail;
        var escBody =
          "Overdue items included in reminder to " + recipientEmail + ":\n\n" +
          overdue.map(function(it) {
            return "\u2022 Reg: " + it.reg + " \u2014 " + it.doc + " \u2014 Date: " + it.date + " (" + Math.abs(it.daysLeft) + " day(s) overdue)";
          }).join("\n");

        try {
          GmailApp.sendEmail(ESCALATION_EMAIL, escSubject, escBody);
        } catch (e) {
          Logger.log("Failed to send escalation to " + ESCALATION_EMAIL + ": " + e.message);
        }
      }
    }
  }

  // Mark today as completed so duplicate runs are skipped
  props.setProperty("lastRunDate", todayKey);
}
