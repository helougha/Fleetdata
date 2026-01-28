/**
 * Fleet paperwork reminder script for Google Sheets.
 *
 * Rules:
 * 1) If the DATE CELL's FONT color is red -> EXCLUDE that document from notifications.
 * 2) If the DATE CELL's BACKGROUND is yellow OR Inspection Exp.Date has a date -> "under inspection".
 *
 * IMPORTANT NOTE:
 * If your red/yellow is applied ONLY via Conditional Formatting,
 * Apps Script may NOT reliably read those displayed colors.
 * This works reliably when the color is actually set on the cell.
 *
 * Email sections:
 * - PAST GRACE PERIOD — expired more than 30 days ago
 * - IN GRACE PERIOD   — expired within last 30 days
 * - UNDER INSPECTION  — yellow-highlighted date cells OR has Inspection Exp.Date
 * - CLOSE TO EXPIRY   — expires within next 30 days
 *
 * Each line: Reg# — Model — Document Type — Expiry Date — days left/overdue
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
    .addItem("Create Renewal Form", "createRenewalForm")
    .addItem("Refresh Form Dropdown", "refreshRenewalFormDropdown")
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
      date: Utilities.formatDate(d, tz, "dd/MM/yyyy"),
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

    // Under inspection if: yellow bg on expiry cell OR Inspection Exp.Date has a date
    var hasInspectionDate = (iInspection !== -1 && row[iInspection] instanceof Date);
    var isUnderInspection = expiryBgIsYellow || hasInspectionDate;

    var expItem = expiryFontIsRed
      ? null
      : evaluateDateField(row[iExpiry], "Expiry Date", reg, model, isUnderInspection);

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
    //   3. UNDER INSPECTION   — yellow-highlighted cells OR has Inspection Exp.Date
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

    // Clean format: Reg# — Model — DocType — Expiry Date — days info
    var formatLine = function(item) {
      var prefix = "\u2022 " + item.reg + " \u2014 " + (item.model || "N/A") + " \u2014 " + item.doc + " \u2014 " + item.date + " \u2014 ";
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


// ============================================================
// GOOGLE FORM — Renewal Submissions
// ============================================================

/**
 * Creates a Google Form for submitting registration renewals.
 * Populates the Reg # dropdown from the sheet, adds Document Type
 * dropdown, a date field for new expiry, and optional notes.
 * Also installs the onFormSubmit trigger automatically.
 */
function createRenewalForm() {
  var SHEET_NAME = "SCC";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Sheet not found: " + SHEET_NAME);

  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).replace(/\s+/g, " ").trim(); });
  var iReg = headers.indexOf("Reg #");
  var iModel = headers.indexOf("Model");

  // Collect unique Reg # values with model names for the dropdown
  var regOptions = [];
  var seen = {};
  for (var r = 1; r < data.length; r++) {
    var reg = String(data[r][iReg] || "").trim();
    if (!reg || seen[reg]) continue;
    seen[reg] = true;
    var model = (iModel !== -1 && data[r][iModel]) ? String(data[r][iModel]).trim() : "";
    regOptions.push(model ? reg + " (" + model + ")" : reg);
  }

  if (regOptions.length === 0) {
    SpreadsheetApp.getUi().alert("No machines found in the sheet. Add some data first.");
    return;
  }

  // Create the form
  var form = FormApp.create("Fleet Renewal Form");
  form.setDescription(
    "Submit a renewal for a machine registration, mulkiya, insurance, or inspection.\n" +
    "The spreadsheet will be updated automatically when you submit."
  );

  // Reg # dropdown
  var regItem = form.addListItem();
  regItem.setTitle("Machine (Reg #)")
    .setHelpText("Select the machine to renew")
    .setChoiceValues(regOptions)
    .setRequired(true);

  // Document type dropdown
  var docItem = form.addListItem();
  docItem.setTitle("Document Type")
    .setHelpText("Which document is being renewed?")
    .setChoiceValues(["Expiry Date", "Inspection Exp.Date"])
    .setRequired(true);

  // New expiry date
  var dateItem = form.addDateItem();
  dateItem.setTitle("New Expiry Date")
    .setHelpText("The new expiry date after renewal (dd/MM/yyyy)")
    .setRequired(true);

  // Notes
  var notesItem = form.addParagraphTextItem();
  notesItem.setTitle("Notes")
    .setHelpText("Optional: any notes about the renewal");

  // Link form responses to this spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // Install the form submit trigger
  ScriptApp.newTrigger("onRenewalFormSubmit")
    .forForm(form)
    .onFormSubmit()
    .create();

  var formUrl = form.getPublishedUrl();
  var editUrl = form.getEditUrl();

  SpreadsheetApp.getUi().alert(
    "Renewal Form Created!\n\n" +
    "Form URL (share this):\n" + formUrl + "\n\n" +
    "Edit URL:\n" + editUrl + "\n\n" +
    "When someone submits the form, the matching row in SCC will be updated automatically."
  );

  Logger.log("Form URL: " + formUrl);
  Logger.log("Edit URL: " + editUrl);
}

/**
 * Handles form submissions — finds the matching row and updates the expiry date.
 * Triggered automatically when the renewal form is submitted.
 */
function onRenewalFormSubmit(e) {
  var SHEET_NAME = "SCC";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;

  var responses = e.response.getItemResponses();
  var regRaw = "";
  var docType = "";
  var newDate = null;
  var notes = "";

  for (var i = 0; i < responses.length; i++) {
    var title = responses[i].getItem().getTitle();
    var answer = responses[i].getResponse();

    if (title === "Machine (Reg #)") {
      // Extract just the Reg # (before the parentheses if model was appended)
      regRaw = String(answer).replace(/\s*\(.*\)\s*$/, "").trim();
    } else if (title === "Document Type") {
      docType = String(answer).trim();
    } else if (title === "New Expiry Date") {
      newDate = new Date(answer);
    } else if (title === "Notes") {
      notes = String(answer || "").trim();
    }
  }

  if (!regRaw || !docType || !newDate || isNaN(newDate.getTime())) {
    Logger.log("Invalid form submission: reg=" + regRaw + ", doc=" + docType + ", date=" + newDate);
    return;
  }

  // Find the matching row
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).replace(/\s+/g, " ").trim(); });
  var iReg = headers.indexOf("Reg #");
  var iExpiry = headers.indexOf("Expiry Date");
  var iInspection = headers.indexOf("Inspection Exp.Date");

  var targetCol = -1;
  if (docType === "Expiry Date") {
    targetCol = iExpiry;
  } else if (docType === "Inspection Exp.Date") {
    targetCol = iInspection;
  }

  if (targetCol === -1) {
    Logger.log("Could not find column for document type: " + docType);
    return;
  }

  var updated = false;
  for (var r = 1; r < data.length; r++) {
    var rowReg = String(data[r][iReg] || "").trim();
    if (rowReg === regRaw) {
      // Update the expiry date cell
      sheet.getRange(r + 1, targetCol + 1).setValue(newDate);

      // Clear yellow background if it was set (machine is no longer under inspection)
      sheet.getRange(r + 1, targetCol + 1).setBackground(null);

      // Log the renewal
      logNotification(ss, regRaw, docType + " renewed", 0, "Form", "Renewed to " + Utilities.formatDate(newDate, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") + (notes ? " | " + notes : ""));

      updated = true;
      Logger.log("Updated " + regRaw + " " + docType + " to " + newDate);
      break;
    }
  }

  if (!updated) {
    Logger.log("No matching row found for Reg #: " + regRaw);
  }
}

/**
 * Refreshes the Reg # dropdown in the renewal form with current sheet data.
 * Run this after adding new machines to the sheet.
 */
function refreshRenewalFormDropdown() {
  var SHEET_NAME = "SCC";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).replace(/\s+/g, " ").trim(); });
  var iReg = headers.indexOf("Reg #");
  var iModel = headers.indexOf("Model");

  var regOptions = [];
  var seen = {};
  for (var r = 1; r < data.length; r++) {
    var reg = String(data[r][iReg] || "").trim();
    if (!reg || seen[reg]) continue;
    seen[reg] = true;
    var model = (iModel !== -1 && data[r][iModel]) ? String(data[r][iModel]).trim() : "";
    regOptions.push(model ? reg + " (" + model + ")" : reg);
  }

  // Find the form linked to this spreadsheet
  var triggers = ScriptApp.getProjectTriggers();
  var formId = null;
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === "onRenewalFormSubmit") {
      formId = triggers[t].getTriggerSourceId();
      break;
    }
  }

  if (!formId) {
    SpreadsheetApp.getUi().alert("No renewal form found. Create one first using the menu.");
    return;
  }

  var form = FormApp.openById(formId);
  var items = form.getItems();
  for (var i = 0; i < items.length; i++) {
    if (items[i].getTitle() === "Machine (Reg #)") {
      items[i].asListItem().setChoiceValues(regOptions);
      SpreadsheetApp.getUi().alert("Dropdown updated with " + regOptions.length + " machines.");
      return;
    }
  }
}
