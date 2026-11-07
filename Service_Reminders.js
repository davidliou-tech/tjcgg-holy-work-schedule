// =============================================================================
// TRUE JESUS CHURCH — GARDEN GROVE
// Holy Work Assignment Email Reminder (2027 Master Integration)
// =============================================================================
//
// OVERVIEW
// --------
// sendAssignmentReminders() is triggered automatically and sends HTML reminder
// emails to the church group 5 days before each scheduled service at 8:00 AM.
// This version dynamically references the MappingConfig sheet (with fallback to
// corrected column mappings) and pulls data directly from Q1-Q4 sheets.
//
// =============================================================================

var REMINDER_LEAD_DAYS  = 5;

/**
 * Returns the URL of the active Master Holy Work Schedule spreadsheet.
 */
function getActiveSheetUrl() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss.getUrl();
  } catch (e) { console.error(e); }
  return "";
}
var SHEET_URL = "";

/**
 * Retrieves the linked 2027 Sermon Topics & Hymns Spreadsheet URL with multi-tiered fallback.
 */
function getSermonTopicsSheetUrl() {
  function isValidSheetUrl(u) {
    return u && typeof u === 'string' && u.indexOf("docs.google.com/spreadsheets/d/") !== -1;
  }
  
  // 1. Check Script Properties
  try {
    var url = PropertiesService.getScriptProperties().getProperty('SERMON_TOPICS_SHEET_URL');
    if (isValidSheetUrl(url)) return url.trim();
  } catch (e) { console.error(e); }
  
  // 2. Check metadata cell in internal Unavailability tab (cell E2)
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      var tab = ss.getSheetByName('Unavailability');
      if (tab) {
        var cellUrl = String(tab.getRange("E2").getValue() || "").trim();
        if (isValidSheetUrl(cellUrl)) return cellUrl;
      }
      
      // 3. Search within current Google Drive folder for file containing "sermon topic"
      var currentFile = DriveApp.getFileById(ss.getId());
      var currentParents = currentFile.getParents();
      if (currentParents.hasNext()) {
        var folder = currentParents.next();
        var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
        while (files.hasNext()) {
          var f = files.next();
          if (f.getName().toLowerCase().indexOf('sermon topic') !== -1 && f.getId() !== ss.getId()) {
            return f.getUrl();
          }
        }
      }
    }
  } catch (e) { console.error(e); }
  
  // 4. Default fallback: check APP_CONFIG.DEFAULT_SERMON_TOPICS_ID or return SHEET_URL
  if (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.DEFAULT_SERMON_TOPICS_ID) {
    return "https://docs.google.com/spreadsheets/d/" + APP_CONFIG.DEFAULT_SERMON_TOPICS_ID + "/edit";
  }
  console.warn("Configuration resolution notice: Sermon Topics Sheet URL not found.");
  return SHEET_URL || "";
}
var SERMON_TOPICS_SHEET_URL = "";

/**
 * Resolves standard email addresses dynamically from ChurchProfile / APP_CONFIG.
 */
function getEmailConfig_() {
  var cc = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.CC_EMAIL) ? APP_CONFIG.CC_EMAIL : 'garden_grove@tjc.org';
  var alert = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.ADMIN_ALERT_EMAIL) ? APP_CONFIG.ADMIN_ALERT_EMAIL : 'gg.info@tjc.org';
  if (typeof getChurchConfig === 'function') {
    try {
      var cfg = getChurchConfig(false);
      if (cfg && cfg.profile) {
        if (cfg.profile.cc_email) cc = String(cfg.profile.cc_email).trim();
        if (cfg.profile.contact_email) alert = String(cfg.profile.contact_email).trim();
      }
    } catch(e) {}
  }
  return { ccRecipient: cc, adminAlertEmail: alert };
}

var CC_RECIPIENT        = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.CC_EMAIL) ? APP_CONFIG.CC_EMAIL : "garden_grove@tjc.org";
var COLOR_NAVY          = "#1F4E79";
var COLOR_BLUE          = "#2E75B6";
var COLOR_ALT_ROW       = "#EBF3FB";
var COLOR_CLEAN_BG      = "#E2F0D9";
var COLOR_REMINDER_BG   = "#FFF8DC";
var COLOR_BORDER        = "#D0D7DE";

function getLogoBlob() {
  var fileId = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.TJC_LOGO_FILE_ID) ? APP_CONFIG.TJC_LOGO_FILE_ID : "15-OcvnePRkYVBg5yz63ILCcHxvVYLQIr";
  try {
    return DriveApp.getFileById(fileId).getBlob().setName("tjcLogo.png");
  } catch (e) {
    Logger.log("DriveApp logo fetch failed for " + fileId + ": " + e.toString());
    return null;
  }
}

// ---------------------------------------------------------------------------
// COLUMN MAPPINGS (Q1-Q4 Sheets)
// 0-indexed columns. Used as robust fallbacks if MappingConfig is not found.
// ---------------------------------------------------------------------------
var COL = {
  DATE: 0,            // Col A
  SPEAKER: 1,         // Col B
  INTERPRETER: 2,     // Col C
  HYMN: 3,            // Col D
  PIANO: 4,           // Col E
  USHER1: 5,          // Col F
  USHER2: 6,          // Col G
  AV1: 7,             // Col H
  AV2: 8,             // Col I
  GRACE: 9,           // Col J
  MORN_PRAYER: 10,    // Col K
  
  RE_K: 11,           // Col L
  RE_K_AST: 12,       // Col M
  RE_E1: 13,          // Col N
  RE_E1_AST: 14,      // Col O
  RE_E2: 15,          // Col P
  RE_E2_AST: 16,      // Col Q
  RE_J1: 17,          // Col R
  RE_J1_AST: 18,      // Col S
  RE_J2: 19,          // Col T
  RE_J2_AST: 20,      // Col U
  RE_SNR: 21,         // Col V
  
  LUNCH_PRAYER: 23,   // Col X
  BIBLE_FELLOW: 26,   // Col AA
  SIS_FELLOW: 27,     // Col AB
  VESSEL_FELLOW: 28,  // Col AC
  
  // Cleaning (Friday)
  CLN_FRI_MAIN1: 29,  // Col AD
  CLN_FRI_MAIN2: 30,  // Col AE
  CLN_FRI_SIDE:  31,  // Col AF
  FRI_CLN_CHAPEL: [29, 30, 31],
  
  // Cleaning (Saturday Windows) - Cols AG, AH (0-indexed 32, 33)
  CLN_SAT_WIN1:  32,  // Col AG (Window Cleaning 1)
  CLN_SAT_WIN2:  33,  // Col AH (Window Cleaning 2)
  SAT_CLN_WINDOWS: [32, 33],
  
  // Cleaning (Saturday) - Aligned with Q4 and MappingConfig
  CLN_SAT_MAIN1: 34,  // Col AI (Chapel 1)
  CLN_SAT_MAIN2: 35,  // Col AJ (Chapel 2)
  CLN_SAT_MAIN3: 36,  // Col AK (Chapel 3)
  CLN_TRASH:     37,  // Col AL (Trash / Dining Hall)
  CLN_BRO_DOWN1: 38,  // Col AM (Bros Downstairs 1)
  CLN_BRO_DOWN2: 39,  // Col AN (Bros Downstairs 2)
  CLN_SIS_DOWN1: 40,  // Col AO (Sis Downstairs 1)
  CLN_SIS_DOWN2: 41,  // Col AP (Sis Downstairs 2)
  CLN_BRO_UP:    42,  // Col AQ (Bros Upstairs)
  CLN_SIS_UP:    43,  // Col AR (Sis Upstairs)
  
  SAT_CLN_CHAPEL:    [34, 35, 36],
  SAT_CLN_TRASH:     [37],
  SAT_CLN_BROS_DOWN: [38, 39],
  SAT_CLN_BROS_UP:   [42],
  SAT_CLN_SIS_DOWN:  [40, 41],
  SAT_CLN_SIS_UP:    [43],

  TUE_JOY:       44,  // Col AS (0-indexed 44)
  WED_BIBLE:     45,  // Col AT (0-indexed 45)
  WED_LOCATIONS: [46, 47, 48, 49], // Cols AU, AV, AW, AX
  SPEAKER_FULL:  50   // Col AY (0-indexed 50)
};

// Helper: Convert column letter to 0-indexed column number
function letterToColIdx0_(letter) {
  var col = 0;
  var len = letter.length;
  for (var i = 0; i < len; i++) {
    col += (letter.charCodeAt(i) - 64) * Math.pow(26, len - i - 1);
  }
  return col - 1;
}

// Dynamically load & override column mappings from MappingConfig tab if present
function loadColMappingsFromConfig() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var configSheet = ss.getSheetByName("MappingConfig");
    if (!configSheet) return;

    var data = configSheet.getDataRange().getValues();
    var satWindows = [];
    var satChapel = [];
    var satTrash = [];
    var satBrosDown = [];
    var satBrosUp = [];
    var satSisDown = [];
    var satSisUp = [];
    var friChapel = [];

    for (var i = 1; i < data.length; i++) {
      var dayType = String(data[i][0] || "").trim().toUpperCase();
      var roleName = String(data[i][1] || "").trim().toUpperCase();
      var section = String(data[i][2] || "").trim().toUpperCase();
      var masterColStr = String(data[i][3] || "").trim().toUpperCase();
      if (!dayType || !masterColStr) continue;

      var colIdx = letterToColIdx0_(masterColStr);
      if (colIdx < 0) continue;

      // Saturday Cleaning
      if (dayType === "SAT" && section === "CLEANING") {
        if (roleName.indexOf("WINDOW") !== -1) {
          if (satWindows.indexOf(colIdx) === -1) satWindows.push(colIdx);
        } else if (roleName.indexOf("CHAPEL") !== -1) {
          if (satChapel.indexOf(colIdx) === -1) satChapel.push(colIdx);
        } else if (roleName.indexOf("TRASH") !== -1) {
          if (satTrash.indexOf(colIdx) === -1) satTrash.push(colIdx);
          COL.CLN_TRASH = colIdx;
        } else if (roleName.indexOf("BROS DOWN") !== -1 || roleName.indexOf("BROTHERS DOWN") !== -1) {
          if (satBrosDown.indexOf(colIdx) === -1) satBrosDown.push(colIdx);
        } else if (roleName.indexOf("BROS UP") !== -1 || roleName.indexOf("BROTHERS UP") !== -1) {
          if (satBrosUp.indexOf(colIdx) === -1) satBrosUp.push(colIdx);
          COL.CLN_BRO_UP = colIdx;
        } else if (roleName.indexOf("SIS DOWN") !== -1 || roleName.indexOf("SISTERS DOWN") !== -1) {
          if (satSisDown.indexOf(colIdx) === -1) satSisDown.push(colIdx);
        } else if (roleName.indexOf("SIS UP") !== -1 || roleName.indexOf("SISTERS UP") !== -1) {
          if (satSisUp.indexOf(colIdx) === -1) satSisUp.push(colIdx);
          COL.CLN_SIS_UP = colIdx;
        }
      } 
      // Friday Cleaning
      else if (dayType === "FRI" && section === "CLEANING") {
        if (roleName.indexOf("CHAPEL") !== -1) {
          if (friChapel.indexOf(colIdx) === -1) friChapel.push(colIdx);
        }
      }
      // General Ministry Roles (Synchronize dynamic overrides if defined in MappingConfig)
      else if (section === "AM" || section === "PM") {
        if (roleName.indexOf("INTERPRETER") !== -1) {
          COL.INTERPRETER = colIdx;
        } else if (roleName.indexOf("PRAISE") !== -1 || roleName.indexOf("HYMN") !== -1) {
          COL.HYMN = colIdx;
        } else if (roleName.indexOf("ACCOMPANIST") !== -1 || roleName.indexOf("PIANO") !== -1) {
          COL.PIANO = colIdx;
        } else if (roleName.indexOf("GRACE") !== -1) {
          COL.GRACE = colIdx;
        } else if (roleName.indexOf("PRAYER") !== -1 && section === "AM") {
          COL.MORN_PRAYER = colIdx;
        } else if (roleName.indexOf("PRAYER") !== -1 && section === "PM") {
          COL.LUNCH_PRAYER = colIdx;
        } else if (roleName.indexOf("ENGLISH BIBLE") !== -1) {
          COL.BIBLE_FELLOW = colIdx;
        } else if (roleName.indexOf("SISTERS") !== -1) {
          COL.SIS_FELLOW = colIdx;
        } else if (roleName.indexOf("VESSEL") !== -1) {
          COL.VESSEL_FELLOW = colIdx;
        }
      }
    }

    if (satWindows.length > 0) {
      COL.SAT_CLN_WINDOWS = satWindows;
      COL.CLN_SAT_WIN1 = satWindows[0];
      COL.CLN_SAT_WIN2 = satWindows.length > 1 ? satWindows[1] : null;
    }
    if (satChapel.length > 0) {
      COL.CLN_SAT_MAIN1 = satChapel[0];
      COL.CLN_SAT_MAIN2 = satChapel.length > 1 ? satChapel[1] : null;
      COL.CLN_SAT_MAIN3 = satChapel.length > 2 ? satChapel[2] : null;
      COL.SAT_CLN_CHAPEL = satChapel;
    }
    if (satTrash.length > 0) {
      COL.SAT_CLN_TRASH = satTrash;
    }
    if (satBrosDown.length > 0) {
      COL.CLN_BRO_DOWN1 = satBrosDown[0];
      COL.CLN_BRO_DOWN2 = satBrosDown.length > 1 ? satBrosDown[1] : null;
      COL.SAT_CLN_BROS_DOWN = satBrosDown;
    }
    if (satBrosUp.length > 0) {
      COL.SAT_CLN_BROS_UP = satBrosUp;
    }
    if (satSisDown.length > 0) {
      COL.CLN_SIS_DOWN1 = satSisDown[0];
      COL.CLN_SIS_DOWN2 = satSisDown.length > 1 ? satSisDown[1] : null;
      COL.SAT_CLN_SIS_DOWN = satSisDown;
    }
    if (satSisUp.length > 0) {
      COL.SAT_CLN_SIS_UP = satSisUp;
    }
    if (friChapel.length > 0) {
      COL.CLN_FRI_MAIN1 = friChapel[0];
      COL.CLN_FRI_MAIN2 = friChapel.length > 1 ? friChapel[1] : null;
      COL.CLN_FRI_SIDE = friChapel.length > 2 ? friChapel[2] : null;
      COL.FRI_CLN_CHAPEL = friChapel;
    }
  } catch (e) {
    Logger.log("loadColMappingsFromConfig warning: " + e.message);
  }
}


// =============================================================================
// ENTRY POINT
// =============================================================================

function sendAssignmentReminders(isManualForce) {
  return executeWithLockAndRetry_('sendAssignmentReminders', function() {
    try {
      SHEET_URL = getActiveSheetUrl();
      SERMON_TOPICS_SHEET_URL = getSermonTopicsSheetUrl();
      loadColMappingsFromConfig();
      emailLog("INFO", "sendAssignmentReminders() started", "");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var parts = todayStr.split('-');

  // Compute target date anchored in spreadsheet's timezone
  var targetDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10) + REMINDER_LEAD_DAYS, 12, 0, 0);

  // Guard Clause: Only process assignment reminders for target schedule services (e.g. Jan 1, 2027 service reminder sent Dec 27, 2026)
  var targetYear = getScheduleTargetYear();
  if (!isManualForce && targetDate.getFullYear() < targetYear) {
    emailLog("SKIP", "Automated service assignment reminders suspended until " + targetYear + " services", "Target Service Date Year: " + targetDate.getFullYear());
    return;
  }
  
  var dayOfWeek = targetDate.getDay();
  // Only process on Mon(1), Fri(5), Sat(6)
  if (dayOfWeek !== 1 && dayOfWeek !== 5 && dayOfWeek !== 6) {
    emailLog("SKIP", "Target date is not Mon, Fri, or Sat", Utilities.formatDate(targetDate, tz, "yyyy-MM-dd"));
    return;
  }
  
  var qSheet = getQuarterSheet(targetDate);
  if (!qSheet) {
    emailLog("ERROR", "Target quarter sheet not found", Utilities.formatDate(targetDate, tz, "yyyy-MM-dd"));
    return;
  }
  
  var qData = qSheet.getDataRange().getValues();
  var matchingRows = [];
  
  // Find all rows matching the target date (usually 1 for Mon/Fri, 2 for Sat)
  for (var r = 5; r < qData.length; r++) { // Data starts at Row 6
    var cellDate = qData[r][COL.DATE];
    if (!cellDate) continue;
    var d = (cellDate instanceof Date) ? cellDate : new Date(cellDate);
    
    if (!isNaN(d.getTime()) && isSameDay(d, targetDate)) {
      matchingRows.push(qData[r]);
    }
  }
  
  var formattedDate = Utilities.formatDate(targetDate, tz, "MMMM d, yyyy");
  
  if (matchingRows.length === 0) {
    emailLog("WARN", "No matching rows found in Q sheet", formattedDate);
    return;
  }

  emailLog("INFO", "Matched service date", formattedDate + " (Found " + matchingRows.length + " rows)");

  // Build and send email
  try {
    var emailConfig = buildEmailContent(dayOfWeek, formattedDate, matchingRows);

    var emailCfg = getEmailConfig_();
    if (emailConfig) {
      var workerEmails = getWorkerEmails(matchingRows);
      var toLine = workerEmails.join(",");
      if (!toLine) {
        toLine = emailCfg.ccRecipient; // Google Apps Script requires a 'to' field
      }
      
      var targetDateString = Utilities.formatDate(targetDate, tz, 'yyyy-MM-dd');
      var sendKey = 'SENT_REMINDER_' + targetDateString;
      
      if (!isManualForce) {
        var existing = PropertiesService.getScriptProperties().getProperty(sendKey);
        if (existing) {
          emailLog("SKIP", "Duplicate send prevented", sendKey);
          return;
        }
      }

      if (MailApp.getRemainingDailyQuota() <= 10) {
        emailLog("ERROR", "Insufficient email quota", "Remaining: " + MailApp.getRemainingDailyQuota());
        return;
      }
      
      var mailOptions = {
        to:       toLine,
        cc:       emailCfg.ccRecipient,
        subject:  emailConfig.subject,
        htmlBody: emailConfig.body
      };
      var logoBlob = getLogoBlob();
      if (logoBlob) {
        mailOptions.inlineImages = { tjcLogo: logoBlob };
      }
      MailApp.sendEmail(mailOptions);
      PropertiesService.getScriptProperties().setProperty(sendKey, new Date().toISOString());
      emailLog("SENT", "Email sent", emailConfig.subject);
    }
  } catch (e) {
    console.error(e);
    emailLog("ERROR", "Failed to build or send email: " + e.message, formattedDate);
    var alertEmail = getEmailConfig_().adminAlertEmail;
    MailApp.sendEmail({
      to: alertEmail,
      subject: "Error in Holy Work Assignment Reminders",
      body: "An unhandled exception occurred during scheduled send: " + e.toString() + "\n\nStack:\n" + e.stack
    });
  }

      emailLog("INFO", "sendAssignmentReminders() completed", "");
    } finally {
      flushEmailLogs();
    }
  }, { timeoutMs: 30000 });
}


// =============================================================================
// LOGGER
// =============================================================================

const logBuffer = [];

function emailLog(level, message, detail) {
  try {
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    logBuffer.push([timestamp, level, message, detail || ""]);
  } catch (logErr) {
    Logger.log("emailLog error: " + logErr.message);
  }
}

function flushEmailLogs() {
  if (logBuffer.length === 0) return;
  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet  = ss.getSheetByName("Email Log");

    if (!logSheet) {
      logSheet = ss.insertSheet("Email Log");
      logSheet.appendRow(["Timestamp", "Level", "Message", "Detail"]);
      logSheet.getRange("1:1").setFontWeight("bold").setBackground(typeof APP_CONFIG !== 'undefined' && APP_CONFIG.THEME ? APP_CONFIG.THEME.PRIMARY_COLOR : "#1F4E79").setFontColor("#FFFFFF");
      logSheet.setFrozenRows(1);
      logSheet.setColumnWidth(1, 180);
      logSheet.setColumnWidth(2, 70);
      logSheet.setColumnWidth(3, 300);
      logSheet.setColumnWidth(4, 400);
      logSheet.hideSheet();
    }

    var lastRow = logSheet.getLastRow();
    
    // Auto-maintenance: Prevent log from growing indefinitely by keeping the last 1000 rows
    if (lastRow > 1000) {
      logSheet.deleteRows(2, 100);
      lastRow = logSheet.getLastRow();
    }

    var startRow = lastRow + 1;
    var numRows = logBuffer.length;
    var range = logSheet.getRange(startRow, 1, numRows, 4);
    range.setValues(logBuffer);

    var bgDanger = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.THEME) ? APP_CONFIG.THEME.ACCENT_DANGER : "#FFDCE0";
    var bgWarn = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.THEME) ? APP_CONFIG.THEME.ACCENT_WARNING : "#FFF8DC";
    var bgSuccess = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.THEME) ? APP_CONFIG.THEME.ACCENT_SUCCESS : "#E2F0D9";
    var bgMatrix = [];
    for (var i = 0; i < numRows; i++) {
      var level = logBuffer[i][1];
      var color = "#FFFFFF";
      if (level === "ERROR") color = bgDanger;
      else if (level === "WARN") color = bgWarn;
      else if (level === "SENT") color = bgSuccess;
      bgMatrix.push([color, color, color, color]);
    }
    range.setBackgrounds(bgMatrix);
    
    logBuffer.length = 0;
  } catch (err) {
    Logger.log("flushEmailLogs error: " + err.message);
  }
}


// =============================================================================
// HELPER FUNCTIONS
// =============================================================================

function isSameDay(d1, d2) {
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var str1 = Utilities.formatDate(d1, tz, "yyyy-MM-dd");
  var str2 = Utilities.formatDate(d2, tz, "yyyy-MM-dd");
  return str1 === str2;
}

function getQuarterSheet(serviceDate) {
  var month       = serviceDate.getMonth() + 1;
  var quarterName = month <= 3 ? "Q1" : month <= 6 ? "Q2" : month <= 9 ? "Q3" : "Q4";
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(quarterName);
}

function filterAssigned(arr) {
  return arr.filter(function(x) { 
    return x && x !== "X" && String(x).trim() !== ""; 
  });
}

function getWorkerEmails(matchingRows) {
  var assignedWorkers = [];
  
  // 1. Extract all unique names from the matchingRows (supports multiple names per cell)
  matchingRows.forEach(function(row) {
    for (var c = 1; c < row.length; c++) {
      var cellVal = String(row[c]).trim();
      if (cellVal !== "" && cellVal.toUpperCase() !== "X") {
        // Split by common delimiters: slash, comma, or newline
        var names = cellVal.split(/[/,\n\r]+/);
        names.forEach(function(rawName) {
          var val = rawName.trim();
          // Remove notes in parentheses e.g. "John Doe (Youth)" -> "John Doe"
          var cleanName = val.replace(/\s*\(.*?\)\s*/g, "").trim().toLowerCase();
          if (cleanName && cleanName !== "x" && assignedWorkers.indexOf(cleanName) === -1) {
            assignedWorkers.push(cleanName);
          }
        });
      }
    }
  });
  
  if (assignedWorkers.length === 0) return [];
  
  // 2. Lookup emails dynamically from the Workers sheet (Header-based dynamic detection)
  var emailList = [];
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Workers");
    if (!sheet) {
      Logger.log("Workers sheet not found");
      return [];
    }
    
    var allData = sheet.getDataRange().getValues();
    if (allData.length < 2) return [];

    var schema = null;
    if (typeof getWorkerSchema === "function") {
      schema = getWorkerSchema(sheet);
    } else {
      throw new Error("getWorkerSchema not defined. Please include Admin_Tools.js.");
    }

    var startDataRow = (schema && schema.dataStartIdx !== undefined) ? schema.dataStartIdx : 2;
    for (var j = startDataRow; j < allData.length; j++) {
      var rowVals = allData[j];
      var email = String(rowVals[schema.emailCol] || "").trim();
      if (!email || email.indexOf("@") === -1) continue;

      var isAssigned = false;
      for (var k = 0; k < schema.nameCols.length; k++) {
        var nIdx = schema.nameCols[k];
        var workerName = String(rowVals[nIdx] || "").trim().toLowerCase();
        if (workerName && assignedWorkers.indexOf(workerName) !== -1) {
          isAssigned = true;
          break;
        }
      }

      if (isAssigned && emailList.indexOf(email) === -1) {
        emailList.push(email);
      }
    }
  } catch (e) {
    Logger.log("Could not extract emails: " + e.message);
  }
  
  return emailList;
}

// =============================================================================
// EMAIL BUILDER — HTML COMPONENT FUNCTIONS
// =============================================================================

function buildHeader(serviceLabel, dateStr, dayLabel) {
  return (
    "<div style='background:" + COLOR_NAVY + "; border-radius:8px 8px 0 0; padding:30px 36px; text-align:center;'>" +
      "<img src='cid:tjcLogo' " +
           "alt='True Jesus Church' style='max-width:280px; width:100%; height:auto; margin-bottom:16px; display:block; margin-left:auto; margin-right:auto;' />" +
      "<p style='margin:0 0 16px; color:#FFFFFF; font-size:18px; font-weight:bold; font-family:Arial,sans-serif;'>Holy Work Assignment Reminder</p>" +
      "<div style='background:rgba(255,255,255,0.12); border-radius:6px; display:inline-block; padding:10px 24px;'>" +
        "<p style='margin:0; color:#FFFFFF; font-size:15px; font-weight:bold; font-family:Arial,sans-serif;'>" + serviceLabel + "</p>" +
        "<p style='margin:4px 0 0; color:#C9DAEA; font-size:13px; font-family:Arial,sans-serif;'>" + dayLabel + "</p>" +
      "</div>" +
    "</div>"
  );
}

function buildGreeting(serviceName, serviceNameZh, dateStr, sermonTopicsUrl, scheduleUrl) {
  var sermonLink = sermonTopicsUrl || scheduleUrl;
  return (
    "<div style='padding:24px 32px 0; font-family:Arial,sans-serif; font-size:14px; line-height:1.8; color:#333;'>" +
      "<p>Dear Brothers and Sisters,<br>" +
      "親愛的弟兄姊妹們平安，</p>" +
      "<p>This is a friendly reminder of your holy work assignments for the upcoming " +
      "<b>" + serviceName + " Service</b> on <b>" + dateStr + "</b>." +
      "<br>" +
      "提醒您查看以下您在 <b>" + dateStr + "</b> " + serviceNameZh + "聚會中所安排的事工。</p>" +
      "<p>Please review your assignments below. We kindly ask the sermon speaker and hymn leader to complete the " +
      "<a href='" + sermonLink + "' style='color:" + COLOR_BLUE + "; font-weight:bold;'>Sermon Topic &amp; Hymns Google Sheet</a> " +
      "prior to the service so that the AV team and pianist may prepare in advance." +
      "<br>" +
      "請講道者及領詩者於聚會前完成 " +
      "<a href='" + sermonLink + "' style='color:" + COLOR_BLUE + "; font-weight:bold;'>Sermon Topic &amp; Hymns Google Sheet</a>，" +
      "以便負責視訊和司琴的同工能夠提前做好準備。</p>" +
    "</div>"
  );
}

function buildSectionTitle(icon, title) {
  return (
    "<div style='background:" + COLOR_BLUE + "; padding:10px 32px; margin:24px 0 0;'>" +
      "<p style='margin:0; color:#FFFFFF; font-size:14px; font-weight:bold; font-family:Arial,sans-serif;'>" +
        icon + "&nbsp;&nbsp;" + title +
      "</p>" +
    "</div>"
  );
}

function buildScheduleTable(rows) {
  var html = "<div style='padding:0 32px; margin-top:16px;'>" +
    "<table cellspacing='0' cellpadding='0' style='width:100%; border-collapse:collapse; border:1px solid " + COLOR_BORDER + "; border-radius:6px; overflow:hidden; font-family:Arial,sans-serif; font-size:13px;'>" +
    "<tr style='background:" + COLOR_BLUE + ";'>" +
      "<th style='padding:10px 14px; color:#fff; text-align:left; width:38%;'>Time</th>" +
      "<th style='padding:10px 14px; color:#fff; text-align:left;'>Service</th>" +
    "</tr>";

  rows.forEach(function(r, idx) {
    var bg = idx % 2 === 0 ? "#FFFFFF" : COLOR_ALT_ROW;
    html += "<tr style='background:" + bg + ";'>" +
              "<td style='padding:9px 14px; border-top:1px solid " + COLOR_BORDER + ";'>" + r[0] + "</td>" +
              "<td style='padding:9px 14px; border-top:1px solid " + COLOR_BORDER + ";'>" + r[1] + "</td>" +
            "</tr>";
  });

  html += "</table></div>";
  return html;
}

/**
 * roles = [{ title: "Speaker", worker: "David" }, { title: "A/V", worker: ["Peter", "John"] }]
 */
function buildAssignmentTable(tableTitle, roles) {
  var html = "<div style='padding:0 32px; margin-top:16px;'>" +
    "<table cellspacing='0' cellpadding='0' style='width:100%; border-collapse:collapse; border:1px solid " + COLOR_BORDER + "; border-radius:6px; overflow:hidden; font-family:Arial,sans-serif; font-size:13px;'>" +
      "<tr style='background:" + COLOR_BLUE + ";'>" +
        "<th style='padding:10px 14px; color:#fff; text-align:left; width:45%;'>Ministry</th>" +
        "<th style='padding:10px 14px; color:#fff; text-align:left;'>Assigned Worker</th>" +
      "</tr>";

  roles.forEach(function(r) {
    var workersHtml = "";
    if (Array.isArray(r.worker)) {
      var wList = filterAssigned(r.worker);
      if (wList.length > 0) {
        workersHtml = "<span style='color:" + COLOR_NAVY + "; font-weight:bold;'>" + 
                      wList.join("<br>") + 
                      "</span>";
      } else {
        workersHtml = "<i style='color:#888;'>Unassigned</i>";
      }
    } else {
      workersHtml = (r.worker && r.worker !== "X" && String(r.worker).trim() !== "") 
        ? "<b style='color:" + COLOR_NAVY + ";'>" + r.worker + "</b>" 
        : "<i style='color:#888;'>Unassigned</i>";
    }
    
    html += "<tr>" +
              "<td style='padding:10px 14px; border-bottom:1px solid " + COLOR_BORDER + ";'>" + r.title + "</td>" +
              "<td style='padding:10px 14px; border-bottom:1px solid " + COLOR_BORDER + ";'>" + workersHtml + "</td>" +
            "</tr>";
  });

  html += "</table></div>";
  return html;
}

function buildCleaningTable(rows) {
  var html = "<div style='padding:0 32px; margin-top:16px;'>" +
    "<table cellspacing='0' cellpadding='0' style='width:100%; border-collapse:collapse; border:1px solid " + COLOR_BORDER + "; border-radius:6px; overflow:hidden; font-family:Arial,sans-serif; font-size:13px;'>" +
    "<tr style='background:#3A7A3A;'>" +
      "<th style='padding:10px 14px; color:#fff; text-align:left; width:50%;'>Area</th>" +
      "<th style='padding:10px 14px; color:#fff; text-align:left;'>Worker(s)</th>" +
    "</tr>";

  rows.forEach(function(r, idx) {
    var bg = idx % 2 === 0 ? COLOR_CLEAN_BG : "#FFFFFF";
    var workerHtml = (r.workers && r.workers.length > 0)
      ? r.workers.map(function(w) { return "<b style='color:" + COLOR_NAVY + ";'>" + w + "</b>"; }).join("<br>")
      : "<i style='color:#888;'>Unassigned</i>";
    html += "<tr style='background:" + bg + ";'>" +
              "<td style='padding:9px 14px; border-top:1px solid " + COLOR_BORDER + ";'>" + r.icon + " " + r.area + "</td>" +
              "<td style='padding:9px 14px; border-top:1px solid " + COLOR_BORDER + ";'>" + workerHtml + "</td>" +
            "</tr>";
  });

  html += "</table></div>";
  return html;
}

function buildRETable(classes) {
  var validClasses = classes.filter(function(c) {
    var hasTeacher = c.teacher && c.teacher !== "X" && String(c.teacher).trim() !== "";
    var hasAsst = c.assistant && c.assistant !== "X" && String(c.assistant).trim() !== "";
    return hasTeacher || hasAsst;
  });

  if (validClasses.length === 0) return ""; 

  var html = "<div style='padding:0 32px; margin-top:16px;'>" +
    "<table cellspacing='0' cellpadding='0' style='width:100%; border-collapse:collapse; border:1px solid " + COLOR_BORDER + "; border-radius:6px; overflow:hidden; font-family:Arial,sans-serif; font-size:13px;'>" +
      "<tr style='background:" + COLOR_BLUE + ";'>" +
        "<th style='padding:10px 14px; color:#fff; text-align:left; width:45%;'>Class</th>" +
        "<th style='padding:10px 14px; color:#fff; text-align:left;'>Teacher / Assistant</th>" +
      "</tr>";

  validClasses.forEach(function(c, idx) {
    var bg = idx % 2 === 0 ? "#FFFFFF" : COLOR_ALT_ROW;
    
    var hasTeacher = c.teacher && c.teacher !== "X" && String(c.teacher).trim() !== "";
    var hasAsst = c.assistant && c.assistant !== "X" && String(c.assistant).trim() !== "";
    
    var workersHtml = "";
    if (hasTeacher) workersHtml += "<b style='color:" + COLOR_NAVY + ";'>" + c.teacher + "</b>";
    if (hasAsst) {
      if (hasTeacher) workersHtml += "<br>";
      workersHtml += c.assistant;
    }
    
    html += "<tr style='background:" + bg + ";'>" +
              "<td style='padding:10px 14px; border-bottom:1px solid " + COLOR_BORDER + ";'>" + c.name + "</td>" +
              "<td style='padding:10px 14px; border-bottom:1px solid " + COLOR_BORDER + ";'>" + workersHtml + "</td>" +
            "</tr>";
  });

  html += "</table></div>";
  return html;
}

function buildReminderBox(sermonTopicsUrl, scheduleUrl) {
  var sermonLink = sermonTopicsUrl || scheduleUrl;
  return (
    "<div style='margin:24px 32px; padding:18px 22px; background:" + COLOR_REMINDER_BG + "; border-left:4px solid #DAA520; border-radius:4px; font-family:Arial,sans-serif; font-size:13px; line-height:1.9;'>" +
      "<p style='margin:0 0 10px; font-weight:bold; font-size:14px; color:#7A5F00;'>📋 Reminders</p>" +
      "<p style='margin:0;'>" +
        "✓&nbsp; Please complete the " +
        "<a href='" + sermonLink + "' style='color:" + COLOR_BLUE + "; font-weight:bold;'>Sermon Topic &amp; Hymns Google Sheet</a> " +
        "before the service.<br>" +
        "✓&nbsp; If you are unable to serve, please arrange a replacement and notify the coordinator.<br>" +
        "✓&nbsp; Please arrive at least <b>10–15 minutes</b> before your assigned duty." +
      "</p>" +
    "</div>"
  );
}

function buildFooter() {
  return (
    "<div style='margin:0 32px 32px; padding-top:20px; border-top:1px solid " + COLOR_BORDER + "; font-family:Arial,sans-serif; font-size:13px; color:#555; line-height:1.8;'>" +
      "<p>May the Lord bless your preparation and service.<br>" +
      "If you have any questions regarding your assignment, please contact the coordinating worker.</p>" +
      "<p>In Christ,<br>" +
      "Garden Grove Information Department<br>" +
      "True Jesus Church in Garden Grove</p>" +
    "</div>"
  );
}

function buildEmailShell(contentHtml) {
  return (
    "<div style='background:#F0F2F5; padding:20px; font-family:Arial,sans-serif;'>" +
      "<div style='max-width:620px; margin:0 auto; background:#FFFFFF; border-radius:8px; " +
           "box-shadow:0 2px 8px rgba(0,0,0,0.12); overflow:hidden;'>" +
        contentHtml +
      "</div>" +
      "<p style='text-align:center; font-size:11px; color:#999; margin-top:12px;'>" +
        "True Jesus Church Garden Grove &mdash; Information Department" +
      "</p>" +
    "</div>"
  );
}


// =============================================================================
// EMAIL LOGIC
// =============================================================================

function buildEmailContent(dayOfWeek, formattedDate, matchingRows) {
  var subject = "";
  var body    = "";
  
  var row1 = matchingRows[0]; 

  switch (dayOfWeek) {
    // -------------------------------------------------------------------------
    // MONDAY
    // -------------------------------------------------------------------------
    case 1:
      subject = formattedDate + " Monday Evening Holy Work Assignments Reminder";
      body += buildHeader("Monday Evening Service", formattedDate, "Monday, " + formattedDate);
      body += buildGreeting("Monday Evening", "週一晚間", formattedDate, SERMON_TOPICS_SHEET_URL, SHEET_URL);

      body += buildSectionTitle("📅", "Service Schedule");
      body += buildScheduleTable([
        ["7:30 – 7:50 PM", "Prayer 禱告會"],
        ["7:50 – 8:00 PM", "Hymnal Worship 詩頌"],
        ["8:00 – 9:00 PM", "Worship Service 崇拜聚會<br>(Sermon &amp; Prayer 證道和禱告)"]
      ]);

      body += buildSectionTitle("👥", "Holy Work Assignments");
      body += buildAssignmentTable("Monday Evening Assignments", [
        { title: "講 道 Speaker", worker: row1[COL.SPEAKER] },
        { title: "翻 譯 Interpreter", worker: row1[COL.INTERPRETER] },
        { title: "司 琴 Pianist", worker: row1[COL.PIANO] },
        { title: "領 詩 Hymn Leader", worker: row1[COL.HYMN] },
        { title: "視 聽 A/V", worker: [row1[COL.AV1], row1[COL.AV2]] },
        { title: "接 待 Usher", worker: [row1[COL.USHER1], row1[COL.USHER2]] }
      ]);

      body += buildReminderBox(SERMON_TOPICS_SHEET_URL, SHEET_URL);
      body += buildFooter();
      break;

    // -------------------------------------------------------------------------
    // FRIDAY
    // -------------------------------------------------------------------------
    case 5:
      subject = formattedDate + " Friday Evening Holy Work Assignments Reminder";
      body += buildHeader("Friday Evening Service", formattedDate, "Friday, " + formattedDate);
      body += buildGreeting("Friday Evening", "週五晚間", formattedDate, SERMON_TOPICS_SHEET_URL, SHEET_URL);

      body += buildSectionTitle("📅", "Service Schedule");
      body += buildScheduleTable([
        ["7:30 – 7:50 PM", "Prayer 禱告會"],
        ["7:50 – 8:00 PM", "Hymnal Worship 詩頌"],
        ["8:00 – 9:00 PM", "Worship Service 崇拜聚會<br>(Sermon &amp; Prayer 證道和禱告)"]
      ]);

      body += buildSectionTitle("👥", "Holy Work Assignments");
      body += buildAssignmentTable("Friday Evening Assignments", [
        { title: "講 道 Speaker", worker: row1[COL.SPEAKER] },
        { title: "翻 譯 Interpreter", worker: row1[COL.INTERPRETER] },
        { title: "司 琴 Pianist", worker: row1[COL.PIANO] },
        { title: "領 詩 Hymn Leader", worker: row1[COL.HYMN] },
        { title: "視 聽 A/V", worker: [row1[COL.AV1], row1[COL.AV2]] },
        { title: "接 待 Usher", worker: [row1[COL.USHER1], row1[COL.USHER2]] }
      ]);

      body += buildSectionTitle("🧹", "Cleaning Duties");
      var friChapelCols = COL.FRI_CLN_CHAPEL || [COL.CLN_FRI_MAIN1, COL.CLN_FRI_MAIN2, COL.CLN_FRI_SIDE];
      var friChapelWorkers = filterAssigned(friChapelCols.map(function(c) { return row1[c]; }));
      body += buildCleaningTable([
        { icon: "🏛", area: "Chapel + MPR", workers: friChapelWorkers }
      ]);

      body += buildReminderBox(SERMON_TOPICS_SHEET_URL, SHEET_URL);
      body += buildFooter();
      break;

    // -------------------------------------------------------------------------
    // SATURDAY
    // -------------------------------------------------------------------------
    case 6:
      subject = formattedDate + " Saturday AM/PM Holy Work Assignments Reminder";
      body += buildHeader("Sabbath Service", formattedDate, "Saturday, " + formattedDate);
      body += buildGreeting("Sabbath", "安息日", formattedDate, SERMON_TOPICS_SHEET_URL, SHEET_URL);

      body += buildSectionTitle("📅", "Sabbath Schedule");
      body += buildScheduleTable([
        ["10:30 – 10:50 AM", "Morning Prayer 禱告會"],
        ["10:50 – 11:00 AM", "Hymnal Worship 詩頌"],
        ["11:00 AM – 12:00 PM", "Worship Service 崇拜聚會<br>(Sermon &amp; Prayer 證道和禱告)"],
        ["12:00 – 1:30 PM",  "Lunch 午餐"],
        ["1:10 – 1:30 PM",   "Afternoon Prayer 禱告會"],
        ["1:30 – 1:40 PM",   "Hymnal Worship 詩頌"],
        ["1:40 – 2:40 PM",   "Worship Service 崇拜聚會<br>(Sermon &amp; Prayer 證道和禱告)"]
      ]);

      var amRow = matchingRows[0];
      var pmRow = matchingRows.length > 1 ? matchingRows[1] : matchingRows[0];

      // AM Assignments
      body += buildSectionTitle("☀️", "Morning Assignments");
      body += buildAssignmentTable("Morning Assignments", [
        { title: "禱 告 會 Morning Prayer", worker: amRow[COL.MORN_PRAYER] },
        { title: "講 道 Speaker", worker: amRow[COL.SPEAKER] },
        { title: "翻 譯 Interpreter", worker: amRow[COL.INTERPRETER] },
        { title: "司 琴 Pianist", worker: amRow[COL.PIANO] },
        { title: "領 詩 Hymn Leader", worker: amRow[COL.HYMN] },
        { title: "視 聽 A/V", worker: [amRow[COL.AV1], amRow[COL.AV2]] },
        { title: "接 待 Usher", worker: [amRow[COL.USHER1], amRow[COL.USHER2]] },
        { title: "謝 飯 Grace", worker: amRow[COL.GRACE] }
      ]);

      // PM Assignments
      body += buildSectionTitle("🌤", "Afternoon Assignments");
      body += buildAssignmentTable("Afternoon Assignments", [
        { title: "午間禱告 Lunchtime Prayer", worker: pmRow[COL.LUNCH_PRAYER] },
        { title: "講 道 Speaker", worker: pmRow[COL.SPEAKER] },
        { title: "翻 譯 Interpreter", worker: pmRow[COL.INTERPRETER] },
        { title: "司 琴 Pianist", worker: pmRow[COL.PIANO] },
        { title: "領 詩 Hymn Leader", worker: pmRow[COL.HYMN] },
        { title: "視 聽 A/V", worker: [pmRow[COL.AV1], pmRow[COL.AV2]] },
        { title: "接 待 Usher", worker: [pmRow[COL.USHER1], pmRow[COL.USHER2]] },
        { title: "英文查經 English Bible Fellowship", worker: pmRow[COL.BIBLE_FELLOW] },
        { title: "姐妹團契 Sisters' Fellowship", worker: pmRow[COL.SIS_FELLOW] },
        { title: "社青團契 Vessel Fellowship", worker: pmRow[COL.VESSEL_FELLOW] }
      ]);

      // Cleaning Duties
      body += buildSectionTitle("🧹", "Cleaning Duties");
      
      // Helper to merge names from both rows in case they are scattered
      function getCln(colIndexOrArr) {
        var cols = Array.isArray(colIndexOrArr) ? colIndexOrArr : [colIndexOrArr];
        var vals = [];
        cols.forEach(function(colIndex) {
          if (colIndex === null || colIndex === undefined) return;
          var a = amRow[colIndex];
          var p = pmRow[colIndex];
          if (a && a !== "X" && String(a).trim() !== "" && vals.indexOf(a) === -1) vals.push(a);
          if (p && p !== "X" && String(p).trim() !== "" && p !== a && vals.indexOf(p) === -1) vals.push(p);
        });
        return vals;
      }

      body += buildCleaningTable([
        { icon: "🪟", area: "Windows", workers: getCln(COL.SAT_CLN_WINDOWS || [COL.CLN_SAT_WIN1, COL.CLN_SAT_WIN2]) },
        { icon: "🏛", area: "Chapel + MPR", workers: getCln(COL.SAT_CLN_CHAPEL || [COL.CLN_SAT_MAIN1, COL.CLN_SAT_MAIN2, COL.CLN_SAT_MAIN3]) },
        { icon: "🗑", area: "Trash (Dining Hall)", workers: getCln(COL.SAT_CLN_TRASH || COL.CLN_TRASH) },
        { icon: "🚹", area: "Brothers Restroom (Downstairs)", workers: getCln(COL.SAT_CLN_BROS_DOWN || [COL.CLN_BRO_DOWN1, COL.CLN_BRO_DOWN2]) },
        { icon: "🚹", area: "Brothers Restroom (Upstairs)", workers: getCln(COL.SAT_CLN_BROS_UP || COL.CLN_BRO_UP) },
        { icon: "🚺", area: "Sisters Restroom (Downstairs)", workers: getCln(COL.SAT_CLN_SIS_DOWN || [COL.CLN_SIS_DOWN1, COL.CLN_SIS_DOWN2]) },
        { icon: "🚺", area: "Sisters Restroom (Upstairs)", workers: getCln(COL.SAT_CLN_SIS_UP || COL.CLN_SIS_UP) }
      ]);

      // Religious Education Classes
      var amREClasses = [
        { name: "K", teacher: amRow[COL.RE_K], assistant: amRow[COL.RE_K_AST] },
        { name: "E1", teacher: amRow[COL.RE_E1], assistant: amRow[COL.RE_E1_AST] },
        { name: "E2", teacher: amRow[COL.RE_E2], assistant: amRow[COL.RE_E2_AST] },
        { name: "J1", teacher: amRow[COL.RE_J1], assistant: amRow[COL.RE_J1_AST] },
        { name: "J2", teacher: amRow[COL.RE_J2], assistant: amRow[COL.RE_J2_AST] },
        { name: "Senior", teacher: amRow[COL.RE_SNR], assistant: null }
      ];
      var pmREClasses = [
        { name: "K", teacher: pmRow[COL.RE_K], assistant: pmRow[COL.RE_K_AST] },
        { name: "E1", teacher: pmRow[COL.RE_E1], assistant: pmRow[COL.RE_E1_AST] },
        { name: "E2", teacher: pmRow[COL.RE_E2], assistant: pmRow[COL.RE_E2_AST] },
        { name: "J1", teacher: pmRow[COL.RE_J1], assistant: pmRow[COL.RE_J1_AST] },
        { name: "J2", teacher: pmRow[COL.RE_J2], assistant: pmRow[COL.RE_J2_AST] },
        { name: "Senior", teacher: pmRow[COL.RE_SNR], assistant: null }
      ];

      var amREHtml = buildRETable(amREClasses);
      if (amREHtml !== "") {
        body += buildSectionTitle("📘", "Morning Religious Education");
        body += amREHtml;
      }

      var pmREHtml = buildRETable(pmREClasses);
      if (pmREHtml !== "") {
        body += buildSectionTitle("📘", "Afternoon Religious Education");
        body += pmREHtml;
      }

      body += buildReminderBox(SERMON_TOPICS_SHEET_URL, SHEET_URL);
      body += buildFooter();
      break;

    default:
      return null;
  }

  return { subject: subject, body: buildEmailShell(body) };
}


// =============================================================================
// TEST FUNCTIONS
// =============================================================================

/**
 * Sends a sample email for ANY date to a designated test email.
 * Does not send to real volunteers and does not CC the church group.
 *
 * @param {string|Date} [customDateStr] e.g. "2027-10-02" or Date object. Defaults to "2027-10-02" (Sat).
 * @param {string} [customRecipient] Destination email. Defaults to configured admin alert email.
 */
function testSendAssignmentReminders(customDateStr, customRecipient) {
  var emailCfg = getEmailConfig_();
  var TEST_RECIPIENT = customRecipient || emailCfg.adminAlertEmail;
  var testDateInput  = customDateStr || "2027-10-02";
  var TEST_DATE      = (testDateInput instanceof Date) ? testDateInput : new Date(testDateInput + "T00:00:00");
  
  SHEET_URL = getActiveSheetUrl();
  SERMON_TOPICS_SHEET_URL = getSermonTopicsSheetUrl();
  loadColMappingsFromConfig();
  emailLog("INFO", "testSendAssignmentReminders() started", Utilities.formatDate(TEST_DATE, Session.getScriptTimeZone(), "yyyy-MM-dd"));
  
  var dayOfWeek = TEST_DATE.getDay();
  var qSheet = getQuarterSheet(TEST_DATE);
  
  if (!qSheet) {
    Logger.log("Error: Quarter sheet not found for test date.");
    return;
  }
  
  var qData = qSheet.getDataRange().getValues();
  var matchingRows = [];
  
  for (var r = 5; r < qData.length; r++) { 
    var cellDate = qData[r][COL.DATE];
    if (!cellDate) continue;
    var d = (cellDate instanceof Date) ? cellDate : new Date(cellDate);
    
    if (!isNaN(d.getTime()) && isSameDay(d, TEST_DATE)) {
      matchingRows.push(qData[r]);
    }
  }
  
  var formattedDate = Utilities.formatDate(TEST_DATE, Session.getScriptTimeZone(), "MMMM d, yyyy");
  
  if (matchingRows.length === 0) {
    Logger.log("Test Failed: No rows found in the sheet for " + formattedDate);
    return;
  }

  try {
    var emailConfig = buildEmailContent(dayOfWeek, formattedDate, matchingRows);
    if (emailConfig) {
      var testOptions = {
        to:           TEST_RECIPIENT,
        subject:      "[TEST] " + emailConfig.subject,
        htmlBody:     emailConfig.body
      };
      var logoBlob = getLogoBlob();
      if (logoBlob) {
        testOptions.inlineImages = { tjcLogo: logoBlob };
      }
      MailApp.sendEmail(testOptions);
      Logger.log("Success! Test email sent to " + TEST_RECIPIENT);
      emailLog("SENT", "Test Email sent to " + TEST_RECIPIENT, emailConfig.subject);
      flushEmailLogs();
    }
  } catch (e) {
    Logger.log("Test Error: " + e.message);
  }
}

// -----------------------------------------------------------------------------
// CONVENIENCE PRESET TEST RUNNERS (Select & Run in Apps Script)
// -----------------------------------------------------------------------------

/** Test sending a reminder for a specific date */
function sendTestReminderForDate(targetDateStr, testRecipient) {
  var targetRecipient = testRecipient || getEmailConfig_().adminAlertEmail;
  testSendAssignmentReminders(targetDateStr, targetRecipient);
}


// =============================================================================
// MONTHLY UNAVAILABILITY EMAIL REMINDER (Sent on 1st of each month)
// =============================================================================

/**
 * Returns the target year for the active Holy Work Schedule (e.g. 2027).
 */
function getScheduleTargetYear() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      var q1 = ss.getSheetByName("Q1");
      if (q1) {
        var a2Val = q1.getRange("A2").getValue();
        if (a2Val instanceof Date && !isNaN(a2Val.getTime())) {
          return a2Val.getFullYear();
        }
      }
      var name = ss.getName();
      var m = name.match(/20\d\d/);
      if (m) return parseInt(m[0], 10);
    }
  } catch (e) {
    console.error(e);
  }
  return new Date().getFullYear();
}

/**
 * Retrieves the linked Unavailability Google Sheet URL for the specified year.
 * Multi-tiered detection: Script Properties -> Linked Unavailability ID -> Unavailability Tab -> Drive Search -> 2027 Fallback.
 */
function getUnavailabilitySheetUrl(targetYear) {
  if (!targetYear) targetYear = getScheduleTargetYear();
  
  // 1. Check Script Properties
  try {
    var props = PropertiesService.getScriptProperties();
    var url = props.getProperty('UNAVAILABILITY_SHEET_URL');
    if (url && url.indexOf("docs.google.com/spreadsheets/d/") !== -1) return url.trim();
    
    var propKey = (typeof CONFIG !== 'undefined' && CONFIG.UNAVAILABILITY_SHEET_PROPERTY) ? CONFIG.UNAVAILABILITY_SHEET_PROPERTY : 'UNAVAILABILITY_SPREADSHEET_ID';
    var id = props.getProperty(propKey);
    if (id && id.length > 20) {
      return "https://docs.google.com/spreadsheets/d/" + id + "/edit?usp=sharing";
    }
  } catch (e) {
    console.error(e);
  }
  
  // 2. Check getLinkedUnavailabilityId() from Admin_Tools.js
  try {
    if (typeof getLinkedUnavailabilityId === 'function') {
      var unavailId = getLinkedUnavailabilityId();
      if (unavailId) {
        return "https://docs.google.com/spreadsheets/d/" + unavailId + "/edit?usp=sharing";
      }
    }
  } catch (e) {
    console.error(e);
  }

  // 3. Check metadata cell E1 in internal Unavailability tab
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      var tab = ss.getSheetByName('Unavailability');
      if (tab) {
        var cellVal = String(tab.getRange("E1").getValue() || "").trim();
        if (cellVal.indexOf("docs.google.com/spreadsheets/d/") !== -1) {
          return cellVal;
        } else if (cellVal.length > 20) {
          return "https://docs.google.com/spreadsheets/d/" + cellVal + "/edit?usp=sharing";
        }
      }
      
      // 4. Search Drive folder of current file for spreadsheet with 'unavailability' in name
      var currentFile = DriveApp.getFileById(ss.getId());
      var parents = currentFile.getParents();
      if (parents.hasNext()) {
        var folder = parents.next();
        var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
        var urlFound = null;
        while (files.hasNext()) {
          var f = files.next();
          var fname = f.getName().toLowerCase();
          if (fname.indexOf('unavailability') !== -1 && f.getId() !== ss.getId()) {
            if (fname.indexOf(String(targetYear)) !== -1 || !urlFound) {
              urlFound = f.getUrl();
            }
          }
        }
        if (urlFound) return urlFound;
      }
    }
  } catch (e) {
    console.error(e);
  }

  // 5. Default 2027 fallback URL
  return "https://docs.google.com/spreadsheets/d/1Ge3iKjPUIkJCuAdlRDqhgEWt4DlO7dL3mFyl0TVLuaY/edit?usp=sharing";
}

/**
 * Extracts qualified worker emails from the "Workers" tab based on criteria:
 * 1. Email exists in Column F OR Column N (unavailability reminder) == "email"
 * 2. Column J (Resident Church) is blank
 * 3. Column M (Roster Status) has a value not equal to "not rostered"
 */
function getUnavailabilityReminderRecipients() {
  var emailList = [];
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return emailList;
    var workerSheet = ss.getSheetByName("Workers");
    if (!workerSheet) {
      Logger.log("Workers sheet not found");
      return emailList;
    }

    var allData = workerSheet.getDataRange().getValues();
    if (allData.length < 2) return emailList;

    var detection = (typeof detectWorkersHeaderRow_ === 'function')
      ? detectWorkersHeaderRow_(allData)
      : { headerRowIdx: 0, dataStartIdx: 1, headers: allData[0] };
    var headerRow = detection.headers;
    var emailColIdx = 5;       // Default Col F
    var churchColIdx = 9;      // Default Col J
    var rosterColIdx = 12;     // Default Col M
    var unavailRemColIdx = 13; // Default Col N

    for (var c = 0; c < headerRow.length; c++) {
      var h = String(headerRow[c] || "").trim().toLowerCase();
      if ((h === "email" || h.indexOf("email") !== -1) && emailColIdx === 5) emailColIdx = c;
      if (h.indexOf("resident church") !== -1 || h === "church") churchColIdx = c;
      if (h.indexOf("roster status") !== -1) rosterColIdx = c;
      if (h.indexOf("unavailability reminder") !== -1) unavailRemColIdx = c;
    }

    // Process worker rows using detected data start index
    for (var r = detection.dataStartIdx; r < allData.length; r++) {
      var row = allData[r];
      var email = String(row[emailColIdx] || "").trim();
      var church = String(row[churchColIdx] || "").trim();
      var rosterStatus = String(row[rosterColIdx] || "").trim().toLowerCase();
      var unavailRem = String(row[unavailRemColIdx] || "").trim().toLowerCase();

      // Rule 1: Email exists in col F or col N == "email"
      var hasEmailCriteria = (email !== "" && email.indexOf("@") !== -1) || (unavailRem === "email");
      // Rule 2: Col J (Resident Church) is blank
      var isLocalChurch = (church === "");
      // Rule 3: Col M (Roster Status) has a value not equal to "not rostered"
      var isRosteredWorker = (rosterStatus !== "" && rosterStatus.indexOf("not rostered") === -1);

      if (hasEmailCriteria && isLocalChurch && isRosteredWorker) {
        if (email && email.indexOf("@") !== -1 && emailList.indexOf(email) === -1) {
          emailList.push(email);
        }
      }
    }
  } catch (e) {
    Logger.log("Error fetching unavailability reminder recipients: " + e.message);
  }
  return emailList;
}

/**
 * Builds the HTML email shell and body for the Monthly Unavailability Reminder email.
 */
function buildMonthlyUnavailabilityEmailBody(targetYear, unavailUrl) {
  var engBracketText = "[" + targetYear + " Holy Work Unavailability]";
  var chiBracketText = "【" + targetYear + "聖工不空時間表】";

  var engLinkHtml = "<a href='" + unavailUrl + "' target='_blank' style='color:" + COLOR_BLUE + "; font-weight:bold; text-decoration:underline;'>" + engBracketText + "</a>";
  var chiLinkHtml = "<a href='" + unavailUrl + "' target='_blank' style='color:" + COLOR_BLUE + "; font-weight:bold; text-decoration:underline;'>" + chiBracketText + "</a>";

  var body = "";

  // Header Banner
  body += buildHeader("Monthly Unavailability Reminder", targetYear + " Holy Work Unavailability", "月度【不空時間】填寫提醒");

  // English Section
  body += "<div style='padding:24px 32px 16px; font-family:Arial,sans-serif; font-size:14px; line-height:1.8; color:#333;'>";
  body += "<p style='font-size:16px; font-weight:bold; color:" + COLOR_NAVY + "; margin-top:0;'>Hallelujah!</p>";
  body += "<p>Thank God for everyone’s dedication and service across our church ministries. To help coordinators plan effectively and minimize scheduling conflicts, please take a moment to update your availability on the " + engLinkHtml + " Google Sheet as soon as possible.</p>";
  body += "<p>May the Lord remember your labor of love!</p>";
  body += "<p style='margin-bottom:0;'>In Christ,<br><b>Garden Grove Church Council</b></p>";
  body += "</div>";

  // English CTA Button
  body += "<div style='text-align:center; margin:10px 0 20px;'>";
  body += "<a href='" + unavailUrl + "' target='_blank' style='background-color:" + COLOR_NAVY + "; color:#ffffff; padding:12px 24px; text-decoration:none; border-radius:6px; font-weight:bold; font-size:14px; display:inline-block;'>📋 Open " + targetYear + " Unavailability Sheet</a>";
  body += "</div>";

  // Section Divider
  body += "<div style='margin:20px 32px; border-top:1px solid " + COLOR_BORDER + ";'></div>";

  // Chinese Section
  body += "<div style='padding:16px 32px 24px; font-family:Arial,sans-serif; font-size:14px; line-height:1.8; color:#333;'>";
  body += "<p style='font-size:13px; color:#666; margin-top:0;'><b>主旨：</b> 月度提醒：請更新您的【不空時間】</p>";
  body += "<p style='font-size:16px; font-weight:bold; color:" + COLOR_NAVY + "; margin-top:0;'>哈利路亞！</p>";
  body += "<p>感謝神，看見大家在教會各項事工上的付出與奉獻。為了方便負責人進行有效規劃並盡可能減少排班衝突，請大家盡快抽空至 " + chiLinkHtml + " Google 表單更新您的時間。</p>";
  body += "<p>願主記念您勞苦所付出的愛心！</p>";
  body += "<p style='margin-bottom:0;'>在基督裏,<br><b>園林教會職務會</b></p>";
  body += "</div>";

  // Chinese CTA Button
  body += "<div style='text-align:center; margin:10px 0 32px;'>";
  body += "<a href='" + unavailUrl + "' target='_blank' style='background-color:" + COLOR_BLUE + "; color:#ffffff; padding:12px 24px; text-decoration:none; border-radius:6px; font-weight:bold; font-size:14px; display:inline-block;'>📋 開啟 " + targetYear + " 聖工不空時間表</a>";
  body += "</div>";

  return buildEmailShell(body);
}

/**
 * Main Entry Point: Sends out the monthly unavailability reminder email on the 1st of each month.
 */
function sendMonthlyUnavailabilityReminder(isManualForce) {
  return executeWithLockAndRetry_('sendMonthlyUnavailabilityReminder', function() {
    try {
      SHEET_URL = getActiveSheetUrl();
      loadColMappingsFromConfig();

      var targetYear = getScheduleTargetYear();
      var today = new Date();
      // Start 1 month prior to target schedule year (e.g. December 1, 2026 for 2027 schedule)
      var startDate = new Date(targetYear - 1, 11, 1, 0, 0, 0);

      if (!isManualForce && today < startDate) {
        var startStr = Utilities.formatDate(startDate, Session.getScriptTimeZone() || "GMT", "yyyy-MM-dd");
        emailLog("SKIP", "Automated unavailability reminder suspended until " + startStr, "Target Schedule Year: " + targetYear);
        return;
      }

      var recipients = getUnavailabilityReminderRecipients();

      emailLog("INFO", "sendMonthlyUnavailabilityReminder() started", "Target Year: " + targetYear + ", Recipients found: " + recipients.length);

      if (recipients.length === 0) {
        emailLog("WARN", "No qualified recipients found for monthly unavailability reminder", "Check Workers sheet columns F, J, M, N");
        return;
      }

      var unavailUrl = getUnavailabilitySheetUrl(targetYear);
      var subject = "Monthly Reminder: Please Update Your Unavailability";
      var htmlBody = buildMonthlyUnavailabilityEmailBody(targetYear, unavailUrl);

      try {
        var CHUNK_SIZE = 50;
        var totalSent = 0;
        var logoBlob = getLogoBlob();

        for (var b = 0; b < recipients.length; b += CHUNK_SIZE) {
          var chunk = recipients.slice(b, b + CHUNK_SIZE);
          var remainingQuota = MailApp.getRemainingDailyQuota();
          if (remainingQuota <= 0) {
            emailLog("ERROR", "MailApp daily quota exhausted", "Remaining quota: " + remainingQuota);
            break;
          }

          var mailOptions = {
            to: CC_RECIPIENT,
            bcc: chunk.join(","),
            subject: subject,
            htmlBody: htmlBody
          };

          if (logoBlob) {
            mailOptions.inlineImages = { tjcLogo: logoBlob };
          }

          MailApp.sendEmail(mailOptions);
          totalSent += chunk.length;
        }

        emailLog("SENT", "Monthly Unavailability Reminder sent", "Recipients: " + totalSent + " of " + recipients.length + " workers, Year: " + targetYear);
      } catch (e) {
        emailLog("ERROR", "Failed to send Monthly Unavailability Reminder: " + e.message, "Target Year: " + targetYear);
      }
    } finally {
      flushEmailLogs();
    }
  }, { timeoutMs: 30000 });
}

/**
 * UI Prompt for manually triggering or testing the Monthly Unavailability Reminder.
 */
function promptSendMonthlyUnavailabilityReminder() {
  var ui = SpreadsheetApp.getUi();
  var targetYear = getScheduleTargetYear();
  var recipients = getUnavailabilityReminderRecipients();
  var unavailUrl = getUnavailabilitySheetUrl(targetYear);

  var msg = "Monthly Unavailability Reminder Email Summary:\n\n" +
            "• Target Year: " + targetYear + "\n" +
            "• Qualified Recipients: " + recipients.length + " workers\n" +
            "• Unavailability Link: " + unavailUrl + "\n\n" +
            "Select action:\n" +
            "• YES: Send reminder email to all " + recipients.length + " workers.\n" +
            "• NO: Send a TEST email to yourself.\n" +
            "• CANCEL: Exit without sending.";

  var resp = ui.alert("Send Monthly Unavailability Reminder", msg, ui.ButtonSet.YES_NO_CANCEL);

  if (resp === ui.Button.YES) {
    sendMonthlyUnavailabilityReminder(true);
    ui.alert("Success", "Monthly Unavailability Reminder email sent to " + recipients.length + " workers!", ui.ButtonSet.OK);
  } else if (resp === ui.Button.NO) {
    var userEmail = Session.getActiveUser().getEmail() || CC_RECIPIENT;
    var subject = "[TEST] Monthly Reminder: Please Update Your Unavailability";
    var htmlBody = buildMonthlyUnavailabilityEmailBody(targetYear, unavailUrl);

    var mailOptions = {
      to: userEmail,
      subject: subject,
      htmlBody: htmlBody
    };
    var logoBlob = getLogoBlob();
    if (logoBlob) {
      mailOptions.inlineImages = { tjcLogo: logoBlob };
    }

    MailApp.sendEmail(mailOptions);

    ui.alert("Test Email Sent", "A test reminder email has been sent to: " + userEmail, ui.ButtonSet.OK);
  }
}

/** Test function to send a preview of the monthly unavailability reminder */
function testMonthlyUnavailabilityReminder(recipientEmail) {
  var targetYear = getScheduleTargetYear();
  var unavailUrl = getUnavailabilitySheetUrl(targetYear);
  var userEmail = recipientEmail || Session.getActiveUser().getEmail() || CC_RECIPIENT;
  var subject = "[TEST] Monthly Reminder: Please Update Your Unavailability";
  var htmlBody = buildMonthlyUnavailabilityEmailBody(targetYear, unavailUrl);

  var mailOptions = {
    to: userEmail,
    subject: subject,
    htmlBody: htmlBody
  };
  var logoBlob = getLogoBlob();
  if (logoBlob) {
    mailOptions.inlineImages = { tjcLogo: logoBlob };
  }

  MailApp.sendEmail(mailOptions);
  Logger.log("Test Monthly Unavailability Reminder sent to " + userEmail);
}

