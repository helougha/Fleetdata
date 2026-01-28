/**
 * Fleet Paperwork Reminder Script
 * 
 * Reads the SCC sheet, checks expiry dates, and sends a formatted email.
 * 
 * Rules:
 * - Red font on date cell = exclude from email
 * - Yellow background on date cell OR date in Inspection Exp.Date column = under inspection
 * 
 * Email sections:
 * - EXPIRED: more than 30 days past expiry
 * - GRACE PERIOD: 0-30 days past expiry  
 * - CLOSE TO EXPIRY: 1-30 days until expiry
 * - UNDER INSPECTION: yellow bg or has inspection date
 */

// ============================================================
// CONFIGURATION — Change these values
// ============================================================

var CONFIG = {
  SHEET_NAME: "SCC",
  RECIPIENTS: ["elhelou.ghazi@sarooj.com"],
  GRACE_PERIOD_DAYS: 30
};

// ============================================================
// MENU
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
  SpreadsheetApp.getUi().alert("Daily trigger created. Emails will send at ~7 AM.");
}

function removeAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function sendTestEmail() {
  var body = "This is a test email from Fleet Reminders.\n\nIf you see this, it's working.";
  GmailApp.sendEmail(CONFIG.RECIPIENTS[0], "Fleet Reminder - Test", body);
  SpreadsheetApp.getUi().alert("Test email sent to " + CONFIG.RECIPIENTS[0]);
}

// ============================================================
// COLOR HELPERS — Simple checks
// ============================================================

function isRed(hex) {
  if (!hex) return false;
  var h = String(hex).toLowerCase().trim();
  // Common red colors in Google Sheets
  return (h === "#ff0000" || h === "#ea4335" || h === "#d93025" || h === "#cc0000");
}

function isYellow(hex) {
  if (!hex) return false;
  var h = String(hex).toLowerCase().trim();
  // Common yellow colors in Google Sheets
  return (h === "#ffff00" || h === "#fff2cc" || h === "#ffe599" || h === "#fce8b2" || h === "#ffeb3b");
}

// ============================================================
// MAIN FUNCTION
// ============================================================

function sendExpiryReminders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    Logger.log("Sheet not found: " + CONFIG.SHEET_NAME);
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  
  var tz = ss.getSpreadsheetTimeZone();
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  
  // Find column indices from headers
  var headers = data[0];
  var iReg = -1, iModel = -1, iExpiry = -1, iInspection = -1, iPlant = -1, iLocation = -1;
  
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c]).trim();
    if (h === "Reg #") iReg = c;
    else if (h === "Model") iModel = c;
    else if (h === "Expiry Date") iExpiry = c;
    else if (h === "Inspection Exp.Date") iInspection = c;
    else if (h === "Plant") iPlant = c;
    else if (h === "Location") iLocation = c;
  }
  
  if (iReg === -1 || iExpiry === -1) {
    Logger.log("Missing required columns: Reg # or Expiry Date");
    return;
  }
  
  // Get cell colors for the Expiry Date column
  var lastRow = sheet.getLastRow();
  var expiryRange = sheet.getRange(2, iExpiry + 1, lastRow - 1, 1);
  var expiryBg = expiryRange.getBackgrounds();
  var expiryFont = expiryRange.getFontColors();
  
  // Collect items into 4 buckets
  var expired = [];
  var gracePeriod = [];
  var closeToExpiry = [];
  var underInspection = [];
  
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var rowIdx = r - 1;
    
    var reg = String(row[iReg] || "").trim();
    if (!reg) continue;
    
    var model = (iModel !== -1) ? String(row[iModel] || "").trim() : "";
    var plant = (iPlant !== -1) ? String(row[iPlant] || "").trim() : "";
    var location = (iLocation !== -1) ? String(row[iLocation] || "").trim() : "";
    
    // Check colors
    var fontColor = expiryFont[rowIdx][0];
    var bgColor = expiryBg[rowIdx][0];
    
    // Red font = skip this row
    if (isRed(fontColor)) continue;
    
    // Get expiry date
    var expiryRaw = row[iExpiry];
    if (!expiryRaw) continue;
    
    var expiryDate;
    if (expiryRaw instanceof Date) {
      expiryDate = expiryRaw;
    } else {
      expiryDate = new Date(expiryRaw);
    }
    if (isNaN(expiryDate.getTime())) continue;
    
    expiryDate.setHours(0, 0, 0, 0);
    var daysLeft = Math.floor((expiryDate - today) / (1000 * 60 * 60 * 24));
    
    // Only include if within range (expired or ≤30 days until expiry)
    if (daysLeft > 30) continue;
    
    // Check if under inspection: yellow bg OR has inspection date
    var hasInspectionDate = (iInspection !== -1 && row[iInspection] instanceof Date);
    var isInspection = isYellow(bgColor) || hasInspectionDate;
    
    var item = {
      reg: reg,
      model: model,
      plant: plant,
      location: location,
      date: Utilities.formatDate(expiryDate, tz, "dd/MM/yyyy"),
      daysLeft: daysLeft
    };
    
    // Sort into bucket
    if (isInspection) {
      underInspection.push(item);
    } else if (daysLeft < -30) {
      expired.push(item);
    } else if (daysLeft <= 0) {
      gracePeriod.push(item);
    } else {
      closeToExpiry.push(item);
    }
  }
  
  // Sort each bucket (most overdue/urgent first)
  expired.sort(function(a, b) { return a.daysLeft - b.daysLeft; });
  gracePeriod.sort(function(a, b) { return a.daysLeft - b.daysLeft; });
  closeToExpiry.sort(function(a, b) { return a.daysLeft - b.daysLeft; });
  underInspection.sort(function(a, b) { return a.daysLeft - b.daysLeft; });
  
  var totalItems = expired.length + gracePeriod.length + closeToExpiry.length + underInspection.length;
  
  if (totalItems === 0) {
    Logger.log("No items to report.");
    return;
  }
  
  // Build email
  var body = "";
  body += "================================================\n";
  body += "         FLEET PAPERWORK STATUS\n";
  body += "================================================\n";
  
  if (expired.length > 0) {
    body += "\n";
    body += "EXPIRED (" + expired.length + ")\n";
    body += "More than 30 days past expiry\n";
    body += "------------------------------------------------\n";
    for (var i = 0; i < expired.length; i++) {
      body += formatLine(expired[i]) + "\n";
    }
  }
  
  if (gracePeriod.length > 0) {
    body += "\n";
    body += "GRACE PERIOD (" + gracePeriod.length + ")\n";
    body += "Less than 30 days past expiry\n";
    body += "------------------------------------------------\n";
    for (var i = 0; i < gracePeriod.length; i++) {
      body += formatLine(gracePeriod[i]) + "\n";
    }
  }
  
  if (closeToExpiry.length > 0) {
    body += "\n";
    body += "CLOSE TO EXPIRY (" + closeToExpiry.length + ")\n";
    body += "30 days or less until expiry\n";
    body += "------------------------------------------------\n";
    for (var i = 0; i < closeToExpiry.length; i++) {
      body += formatLine(closeToExpiry[i]) + "\n";
    }
  }
  
  if (underInspection.length > 0) {
    body += "\n";
    body += "UNDER INSPECTION (" + underInspection.length + ")\n";
    body += "Currently going through inspection\n";
    body += "------------------------------------------------\n";
    for (var i = 0; i < underInspection.length; i++) {
      body += formatLine(underInspection[i]) + "\n";
    }
  }
  
  body += "\n================================================\n";
  body += "Automated Fleet Reminder\n";
  body += "================================================\n";
  
  // Send email
  var subject = "Fleet Paperwork Reminder (" + totalItems + " items)";
  
  for (var i = 0; i < CONFIG.RECIPIENTS.length; i++) {
    try {
      GmailApp.sendEmail(CONFIG.RECIPIENTS[i], subject, body);
      Logger.log("Email sent to " + CONFIG.RECIPIENTS[i]);
    } catch (e) {
      Logger.log("Failed to send to " + CONFIG.RECIPIENTS[i] + ": " + e.message);
    }
  }
}

// ============================================================
// FORMAT LINE
// ============================================================

function formatLine(item) {
  var line = "  " + item.reg;
  line += " - " + (item.model || "N/A");
  line += " - Exp: " + item.date;
  
  if (item.daysLeft < 0) {
    line += " - OVERDUE by " + Math.abs(item.daysLeft) + " day(s)";
  } else if (item.daysLeft === 0) {
    line += " - Expires TODAY";
  } else {
    line += " - " + item.daysLeft + " day(s) left";
  }
  
  if (item.plant) line += " | Plant: " + item.plant;
  if (item.location) line += " | Location: " + item.location;
  
  return line;
}
