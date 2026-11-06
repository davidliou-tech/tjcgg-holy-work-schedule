/**
 * GenerateESSC.js
 * Automated Generation Engine, Info Dept Date Controls, Annual Provisioning & Protection Suite
 * for Semi-Annual ESSC Events (Spring ESSC & Fall ESSC)
 *
 * HOW IT WORKS:
 * 1. Connects to that year's Master Holy Work Schedule workbook via MASTER_HWS_SPREADSHEET_ID.
 * 2. Info Department sets Spring & Fall Wednesday Start Dates in MasterConfig or UI prompt.
 * 3. Script automatically updates all 45 date rows in Spring_ESSC_Grid and Fall_ESSC_Grid.
 * 4. Coordinators edit the 16-column horizontal data entry grid ('Spring_ESSC_Grid' or 'Fall_ESSC_Grid').
 * 5. This script reads 'ESSC_MappingConfig' to map cells from the grid to the presentation templates.
 * 6. Formatted outputs are written to 'Spring_ESSC_Worship', 'Spring_ESSC_Work', etc.
 * 7. Presentation sheets are automatically LOCKED (protected) against manual edits.
 */

const ESSC_CONFIG = {
  GRID: {
    FIRST_ROW: 2,
    LAST_ROW: 45,  // Expanded to support rows 43–45 from commit b4fb5fd
    FIRST_COL: 5,
    LAST_COL: 16
  },
  SPRING_GRID: 'Spring_ESSC_Grid',
  FALL_GRID: 'Fall_ESSC_Grid',
  MAPPING_SHEET: 'ESSC_MappingConfig',
  VALIDATION_CONFIG: 'ValidationConfig',
  MASTER_CONFIG_SHEET: 'MasterConfig',
  SCRIPT_PROP_MASTER_ID: 'MASTER_HWS_SPREADSHEET_ID',
  SCRIPT_PROP_ESSC_ID: 'ESSC_SPREADSHEET_ID',
  PRESENTATION_SHEETS: {
    SPRING_WORSHIP: 'Spring_ESSC_Worship',
    SPRING_WORK: 'Spring_ESSC_Work',
    FALL_WORSHIP: 'Fall_ESSC_Worship',
    FALL_WORK: 'Fall_ESSC_Work'
  }
};

/**
 * Resolves the target ESSC spreadsheet.
 * 1. If active spreadsheet contains Spring_ESSC_Grid or Fall_ESSC_Grid, returns it.
 * 2. Otherwise, opens the connected standalone ESSC spreadsheet via Script Properties / MasterConfig.
 * 3. If not yet connected, prompts the user in the Master HWS UI.
 */
function getTargetESSCSpreadsheet() {
  var activeSs = SpreadsheetApp.getActiveSpreadsheet();
  if (activeSs.getSheetByName(ESSC_CONFIG.SPRING_GRID) || activeSs.getSheetByName(ESSC_CONFIG.FALL_GRID)) {
    return activeSs;
  }
  
  var esscId = PropertiesService.getScriptProperties().getProperty(ESSC_CONFIG.SCRIPT_PROP_ESSC_ID);
  if (!esscId) {
    var mcSheet = activeSs.getSheetByName(ESSC_CONFIG.MASTER_CONFIG_SHEET);
    if (mcSheet) {
      var data = mcSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === "ESSC_SPREADSHEET_ID") {
          esscId = String(data[i][1]).trim();
          break;
        }
      }
    }
  }
  
  if (esscId) {
    try {
      return SpreadsheetApp.openById(esscId);
    } catch(e) {
      console.warn("Could not open connected ESSC workbook (" + esscId + "): " + e.message);
    }
  }
  
  return connectStandaloneESSCWorkbookPrompt();
}

/**
 * Prompt Admin in Master HWS to Connect the Standalone ESSC Schedule Workbook
 */
function connectStandaloneESSCWorkbookPrompt() {
  var ui = SpreadsheetApp.getUi();
  var masterSs = SpreadsheetApp.getActiveSpreadsheet();
  
  var currentId = PropertiesService.getScriptProperties().getProperty(ESSC_CONFIG.SCRIPT_PROP_ESSC_ID) || "";
  if (!currentId) {
    var mcSheet = masterSs.getSheetByName(ESSC_CONFIG.MASTER_CONFIG_SHEET);
    if (mcSheet) {
      var data = mcSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === "ESSC_SPREADSHEET_ID") {
          currentId = String(data[i][1]).trim();
          break;
        }
      }
    }
  }

  var result = ui.prompt(
    "Connect Standalone ESSC Schedule",
    "Please enter the Google Spreadsheet ID or full URL of the 2027 ESSC Schedule workbook (e.g. 2027 GG ESSC Schedule):\n\n(Current ID: " + (currentId || "None") + ")",
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return null;
  
  var input = result.getResponseText().trim();
  var match = input.match(/[-\w]{25,}/);
  var esscId = match ? match[0] : input;
  if (!esscId) return null;
  
  try {
    var esscSs = SpreadsheetApp.openById(esscId);
    PropertiesService.getScriptProperties().setProperty(ESSC_CONFIG.SCRIPT_PROP_ESSC_ID, esscId);
    
    var mc = masterSs.getSheetByName(ESSC_CONFIG.MASTER_CONFIG_SHEET);
    if (mc) {
      updateOrAppendConfigRow(mc, "ESSC_SPREADSHEET_ID", esscId);
      updateOrAppendConfigRow(mc, "ESSC_SPREADSHEET_NAME", esscSs.getName());
      updateOrAppendConfigRow(mc, "ESSC_LAST_LINKED_TIMESTAMP", new Date().toISOString());
    }
    ui.alert("Connected 🎉", "Successfully connected to Standalone ESSC Workbook:\n'" + esscSs.getName() + "'\nID: " + esscId, ui.ButtonSet.OK);
    return esscSs;
  } catch(e) {
    ui.alert("Connection Error", "Could not open spreadsheet with ID '" + esscId + "'. Please verify permissions:\n" + e.message, ui.ButtonSet.OK);
    return null;
  }
}

/**
 * Automatically builds the ESSC Assistant menu when opening the standalone ESSC workbook.
 */
function onOpenESSC() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🎪 ESSC Assistant')
    .addItem('👥 Open Assignment Sidebar', 'showSidebarESSC')
    .addSeparator()
    .addItem('🚀 Generate Spring ESSC Schedule', 'generateSpringESSC')
    .addItem('🚀 Generate Fall ESSC Schedule', 'generateFallESSC')
    .addSeparator()
    .addItem('🔄 Apply / Sync Dropdowns', 'applyValidationForESSC')
    .addItem('✨ Resolve Speaker Placeholders', 'resolveGridSpeakerPlaceholdersMenu')
    .addItem('📅 Set ESSC Event Dates...', 'promptSetESSCDates')
    .addItem('🎨 Set ESSC Event Themes...', 'promptSetESSCThemes')
    .addItem('🔗 Connect Master HWS Workbook...', 'connectMasterHWSWorkbookPrompt')
    .addToUi();
}

/**
 * Menu handler to manually trigger speaker placeholder resolution
 */
function resolveGridSpeakerPlaceholdersMenu() {
  var targetSs = getTargetESSCSpreadsheet();
  if (!targetSs) return;
  var count = resolveGridSpeakerPlaceholders(targetSs);
  SpreadsheetApp.getUi().alert("Placeholders Resolved 🎉", "Successfully resolved " + count + " speaker placeholder tokens on '" + targetSs.getName() + "'.", SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Info Department UI Prompt to Set ESSC Wednesday Start Dates
 */
function promptSetESSCDates() {
  var ui = SpreadsheetApp.getUi();
  var masterSs = SpreadsheetApp.getActiveSpreadsheet();
  var targetSs = getTargetESSCSpreadsheet();
  if (!targetSs) return;
  
  var defaultSpring = getConfigValue(masterSs, "SPRING_WEDNESDAY_START_DATE") || getConfigValue(targetSs, "SPRING_WEDNESDAY_START_DATE") || "2027-04-14";
  var defaultFall = getConfigValue(masterSs, "FALL_WEDNESDAY_START_DATE") || getConfigValue(targetSs, "FALL_WEDNESDAY_START_DATE") || "2027-10-13";
  
  var springResp = ui.prompt(
    "Set Spring ESSC Start Date",
    "Enter Wednesday start date for Spring ESSC (YYYY-MM-DD):\n(Current / Default: " + defaultSpring + ")",
    ui.ButtonSet.OK_CANCEL
  );
  if (springResp.getSelectedButton() !== ui.Button.OK) return;
  var springDateStr = springResp.getResponseText().trim() || defaultSpring;
  
  var fallResp = ui.prompt(
    "Set Fall ESSC Start Date",
    "Enter Wednesday start date for Fall ESSC (YYYY-MM-DD):\n(Current / Default: " + defaultFall + ")",
    ui.ButtonSet.OK_CANCEL
  );
  if (fallResp.getSelectedButton() !== ui.Button.OK) return;
  var fallDateStr = fallResp.getResponseText().trim() || defaultFall;
  
  // 1. Save to Master HWS MasterConfig
  var masterMc = masterSs.getSheetByName(ESSC_CONFIG.MASTER_CONFIG_SHEET);
  if (masterMc) {
    updateOrAppendConfigRow(masterMc, "SPRING_ESSC_WEDNESDAY_START", springDateStr);
    updateOrAppendConfigRow(masterMc, "SPRING_WEDNESDAY_START_DATE", springDateStr);
    updateOrAppendConfigRow(masterMc, "FALL_ESSC_WEDNESDAY_START", fallDateStr);
    updateOrAppendConfigRow(masterMc, "FALL_WEDNESDAY_START_DATE", fallDateStr);
  }
  
  // 2. Save to target ESSC MasterConfig
  var targetMc = targetSs.getSheetByName(ESSC_CONFIG.MASTER_CONFIG_SHEET);
  if (!targetMc) {
    targetMc = targetSs.insertSheet(ESSC_CONFIG.MASTER_CONFIG_SHEET);
    targetMc.appendRow(["Setting / Parameter", "Value / Description"]);
  }
  updateOrAppendConfigRow(targetMc, "SPRING_ESSC_WEDNESDAY_START", springDateStr);
  updateOrAppendConfigRow(targetMc, "SPRING_WEDNESDAY_START_DATE", springDateStr);
  updateOrAppendConfigRow(targetMc, "FALL_ESSC_WEDNESDAY_START", fallDateStr);
  updateOrAppendConfigRow(targetMc, "FALL_WEDNESDAY_START_DATE", fallDateStr);
  
  // 3. Recalculate grid dates on target ESSC spreadsheet
  recalculateGridDates(targetSs, 'Spring', springDateStr);
  recalculateGridDates(targetSs, 'Fall', fallDateStr);
  
  ui.alert("Success! 🎉", "Updated ESSC Start Dates on '" + targetSs.getName() + "':\nSpring Wednesday: " + springDateStr + "\nFall Wednesday: " + fallDateStr + "\n\nAll 45 grid date rows updated.", ui.ButtonSet.OK);
}

/**
 * Info Department UI Prompt to Set ESSC Convocation and Evangelical Themes
 */
function promptSetESSCThemes() {
  var ui = SpreadsheetApp.getUi();
  var masterSs = SpreadsheetApp.getActiveSpreadsheet();
  var targetSs = getTargetESSCSpreadsheet();
  if (!targetSs) return;
  
  var seasonResp = ui.prompt(
    "Set ESSC Event Themes",
    "Which season would you like to configure themes for? Enter 'Spring' or 'Fall':",
    ui.ButtonSet.OK_CANCEL
  );
  if (seasonResp.getSelectedButton() !== ui.Button.OK) return;
  var season = seasonResp.getResponseText().trim();
  if (season.toLowerCase() !== 'spring' && season.toLowerCase() !== 'fall') {
    ui.alert("Invalid Input", "Please enter 'Spring' or 'Fall'.", ui.ButtonSet.OK);
    return;
  }
  season = season.charAt(0).toUpperCase() + season.slice(1).toLowerCase();
  
  var convThemeResp = ui.prompt(
    "Set " + season + " Convocation Theme",
    "Enter the Spiritual Convocation theme (e.g. 聖經動物園 Biblical Menagerie):",
    ui.ButtonSet.OK_CANCEL
  );
  if (convThemeResp.getSelectedButton() !== ui.Button.OK) return;
  var convTheme = convThemeResp.getResponseText().trim();
  
  var evanThemeResp = ui.prompt(
    "Set " + season + " Evangelical Services Theme",
    "Enter the Evangelical Services theme (e.g. 從聖經看神的教會 Understanding God's Church From The Bible):",
    ui.ButtonSet.OK_CANCEL
  );
  if (evanThemeResp.getSelectedButton() !== ui.Button.OK) return;
  var evanTheme = evanThemeResp.getResponseText().trim();
  
  var prefix = season.toUpperCase();
  var convVal = "Theme: " + convTheme.replace(/^Theme:\s*/i, '');
  var evanVal = "Theme: " + evanTheme.replace(/^Theme:\s*/i, '');
  
  // Save to target ESSC
  var targetMc = targetSs.getSheetByName(ESSC_CONFIG.MASTER_CONFIG_SHEET);
  if (!targetMc) {
    targetMc = targetSs.insertSheet(ESSC_CONFIG.MASTER_CONFIG_SHEET);
    targetMc.appendRow(["Setting / Parameter", "Value / Description"]);
  }
  updateOrAppendConfigRow(targetMc, prefix + "_CONVOCATION_THEME", convVal);
  updateOrAppendConfigRow(targetMc, prefix + "_EVANGELICAL_THEME", evanVal);
  
  // Save to masterSs
  if (masterSs.getId() !== targetSs.getId()) {
    var masterMc = masterSs.getSheetByName(ESSC_CONFIG.MASTER_CONFIG_SHEET);
    if (masterMc) {
      updateOrAppendConfigRow(masterMc, prefix + "_CONVOCATION_THEME", convVal);
      updateOrAppendConfigRow(masterMc, prefix + "_EVANGELICAL_THEME", evanVal);
    }
  }
  
  ui.alert("Success! 🎉", "Updated " + season + " ESSC Themes on '" + targetSs.getName() + "':\nConvocation: " + convTheme + "\nEvangelical: " + evanTheme, ui.ButtonSet.OK);
}

/**
 * Helper to read key-value parameter from MasterConfig tab (supports common aliases)
 */
function getSafeSpreadsheetTimeZone(ss) {
  var tz = "";
  try {
    if (ss && typeof ss.getSpreadsheetTimeZone === 'function') {
      tz = ss.getSpreadsheetTimeZone();
    }
  } catch(e) {}
  if (!tz || typeof tz !== 'string' || !tz.trim()) {
    try {
      tz = Session.getScriptTimeZone();
    } catch(e) {}
  }
  if (!tz || typeof tz !== 'string' || !tz.trim()) {
    tz = "America/Los_Angeles";
  }
  try {
    if (ss && typeof ss.setSpreadsheetTimeZone === 'function') {
      var currentTz = ss.getSpreadsheetTimeZone();
      if (!currentTz || typeof currentTz !== 'string') {
        ss.setSpreadsheetTimeZone(tz);
      }
    }
  } catch(e) {}
  return tz;
}

/**
 * Robust date formatter that converts Date objects to 'YYYY-MM-DD' without failing on null timezones
 */
function formatDateSafe(val, ss) {
  if (!val) return "";
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return "";
    var y = val.getFullYear();
    var m = ('0' + (val.getMonth() + 1)).slice(-2);
    var d = ('0' + val.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  return String(val !== undefined && val !== null ? val : "").trim();
}

/**
 * Helper to read key-value parameter from MasterConfig tab (supports common aliases)
 */
var _masterConfigCache = {};

function clearMasterConfigCache(ssId) {
  if (ssId) {
    delete _masterConfigCache[ssId];
  } else {
    _masterConfigCache = {};
  }
}

function getConfigValue(ss, key) {
  if (!ss) return "";
  var ssId = ss.getId();
  var map = _masterConfigCache[ssId];
  if (!map) {
    var mcSheet = ss.getSheetByName(ESSC_CONFIG.MASTER_CONFIG_SHEET);
    if (!mcSheet) return "";
    var data = mcSheet.getDataRange().getValues();
    map = {};
    for (var i = 1; i < data.length; i++) {
      var k = String(data[i][0] || "").trim();
      if (k) {
        map[k] = formatDateSafe(data[i][1], ss);
      }
    }
    _masterConfigCache[ssId] = map;
  }
  
  if (map.hasOwnProperty(key)) return map[key];
  var aliases = {
    "SPRING_ESSC_WEDNESDAY_START": "SPRING_WEDNESDAY_START_DATE",
    "SPRING_WEDNESDAY_START_DATE": "SPRING_ESSC_WEDNESDAY_START",
    "FALL_ESSC_WEDNESDAY_START": "FALL_WEDNESDAY_START_DATE",
    "FALL_WEDNESDAY_START_DATE": "FALL_ESSC_WEDNESDAY_START"
  };
  if (aliases[key] && map.hasOwnProperty(aliases[key])) {
    return map[aliases[key]];
  }
  return "";
}

/**
 * Updates all 45 date rows in Spring_ESSC_Grid or Fall_ESSC_Grid based on Wednesday Start Date
 */
function recalculateGridDates(ss, season, wednesdayDateStr) {
  var gridName = (season === 'Spring') ? ESSC_CONFIG.SPRING_GRID : ESSC_CONFIG.FALL_GRID;
  var gridSheet = ss.getSheetByName(gridName);
  if (!gridSheet) return;
  
  var baseDate = null;
  if (wednesdayDateStr instanceof Date) {
    baseDate = new Date(wednesdayDateStr.getFullYear(), wednesdayDateStr.getMonth(), wednesdayDateStr.getDate());
  } else {
    var s = String(wednesdayDateStr || "").trim();
    var m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) {
      baseDate = new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
    } else {
      var d = new Date(s);
      if (!isNaN(d.getTime())) {
        baseDate = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      }
    }
  }
  if (!baseDate || isNaN(baseDate.getTime())) return;
  
  var lastRow = gridSheet.getLastRow();
  if (lastRow < 2) return;
  
  var dayOffsets = { 'Wed': 0, 'Thu': 1, 'Fri': 2, 'Sat': 3, 'Sun': 4 };
  var dataRange = gridSheet.getRange(2, 1, lastRow - 1, 2);
  var values = dataRange.getValues();
  
  var newDates = [];
  for (var i = 0; i < values.length; i++) {
    var dayCode = String(values[i][1] || "").trim();
    var offset = dayOffsets[dayCode];
    if (offset !== undefined) {
      var targetDate = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate() + offset);
      var y = targetDate.getFullYear();
      var m = ('0' + (targetDate.getMonth() + 1)).slice(-2);
      var d = ('0' + targetDate.getDate()).slice(-2);
      newDates.push([y + '-' + m + '-' + d]);
    } else {
      newDates.push([values[i][0]]); // keep existing if unknown
    }
  }
  if (newDates.length > 0) {
    gridSheet.getRange(2, 1, newDates.length, 1).setValues(newDates);
  }
}

/**
 * Helper to update key-value in MasterConfig tab
 */
function updateOrAppendConfigRow(sheet, key, value) {
  try { clearMasterConfigCache(sheet.getParent().getId()); } catch(e) {}
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      data[i][1] = value;
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

/**
 * Prompt Coordinator to Connect Active Year's Master Holy Work Schedule Workbook
 */
function connectMasterHWSWorkbookPrompt() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var currentMasterId = PropertiesService.getScriptProperties().getProperty(ESSC_CONFIG.SCRIPT_PROP_MASTER_ID) || "";
  if (!currentMasterId) {
    var mcSheet = ss.getSheetByName(ESSC_CONFIG.MASTER_CONFIG_SHEET);
    if (mcSheet) {
      var data = mcSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === "MASTER_HWS_SPREADSHEET_ID") {
          currentMasterId = String(data[i][1]).trim();
          break;
        }
      }
    }
  }

  var result = ui.prompt(
    "Connect Master Holy Work Schedule",
    "Please enter the Google Spreadsheet ID or full URL of that year's Master Holy Work Schedule (e.g. 2027 Holy Work Schedule):\n\n(Current ID: " + (currentMasterId || "None") + ")",
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    var input = result.getResponseText().trim();
    var masterId = "";
    if (input) {
      var idMatch = input.match(/[-\w]{25,}/);
      masterId = idMatch ? idMatch[0] : input;
    } else if (currentMasterId) {
      masterId = currentMasterId;
    } else {
      masterId = ss.getId();
    }
    
    try {
      var masterSs = SpreadsheetApp.openById(masterId);
      PropertiesService.getScriptProperties().setProperty(ESSC_CONFIG.SCRIPT_PROP_MASTER_ID, masterId);
      
      var targetMcSheet = ss.getSheetByName(ESSC_CONFIG.MASTER_CONFIG_SHEET);
      if (!targetMcSheet) {
        targetMcSheet = ss.insertSheet(ESSC_CONFIG.MASTER_CONFIG_SHEET);
        targetMcSheet.appendRow(["Setting / Parameter", "Value / Description"]);
      }
      updateOrAppendConfigRow(targetMcSheet, "MASTER_HWS_SPREADSHEET_ID", masterId);
      updateOrAppendConfigRow(targetMcSheet, "MASTER_HWS_NAME", masterSs.getName());
      updateOrAppendConfigRow(targetMcSheet, "LAST_SYNCED_TIMESTAMP", new Date().toISOString());
      
      ui.alert("Success! 🎉", "Successfully connected to Master HWS Workbook: '" + masterSs.getName() + "'\nID: " + masterId, ui.ButtonSet.OK);
    } catch(e) {
      ui.alert("Connection Error", "Could not open spreadsheet with ID '" + masterId + "'. Please verify permissions.\nError: " + e.message, ui.ButtonSet.OK);
    }
  }
}

/**
 * Helper to get the connected Master HWS Spreadsheet object
 */
function getMasterHWSSpreadsheet() {
  var masterId = PropertiesService.getScriptProperties().getProperty(ESSC_CONFIG.SCRIPT_PROP_MASTER_ID);
  if (!masterId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mcSheet = ss.getSheetByName(ESSC_CONFIG.MASTER_CONFIG_SHEET);
    if (mcSheet) {
      var data = mcSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === "MASTER_HWS_SPREADSHEET_ID") {
          masterId = String(data[i][1]).trim();
          break;
        }
      }
    }
  }
  if (!masterId) {
    throw new Error("Master HWS Workbook not connected. Please run 'Schedule Assistant > Connect Master HWS Workbook...' first.");
  }
  return SpreadsheetApp.openById(masterId);
}

/**
 * Menu Entrypoint to Generate Spring ESSC Schedule
 */
function generateSpringESSC() {
  var targetSs = getTargetESSCSpreadsheet();
  if (!targetSs) return;
  generateESSCForSeason('Spring', targetSs);
}

/**
 * Menu Entrypoint to Generate Fall ESSC Schedule
 */
function generateFallESSC() {
  var targetSs = getTargetESSCSpreadsheet();
  if (!targetSs) return;
  generateESSCForSeason('Fall', targetSs);
}

/**
 * Main Generator Function for ESSC
 * @param {string} season - 'Spring' or 'Fall'
 * @param {Spreadsheet} [targetSs] - Optional target ESSC spreadsheet
 */
function generateESSCForSeason(season, targetSs) {
  var ss = targetSs || getTargetESSCSpreadsheet();
  if (!ss) return;
  var gridSheetName = (season === 'Spring') ? ESSC_CONFIG.SPRING_GRID : ESSC_CONFIG.FALL_GRID;
  var gridSheet = ss.getSheetByName(gridSheetName);
  var mappingSheet = ss.getSheetByName(ESSC_CONFIG.MAPPING_SHEET);
  
  if (!gridSheet) {
    SpreadsheetApp.getUi().alert("Error", "Grid sheet '" + gridSheetName + "' not found.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  if (!mappingSheet) {
    SpreadsheetApp.getUi().alert("Error", "Mapping config sheet '" + ESSC_CONFIG.MAPPING_SHEET + "' not found.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    SpreadsheetApp.getUi().alert("Busy", "Another schedule generation is in progress. Please try again.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  try {
    // 1. Read Grid Data
    var lastRow = gridSheet.getLastRow();
    var gridValues = gridSheet.getRange(1, 1, lastRow, ESSC_CONFIG.GRID.LAST_COL).getValues(); // Cols A to P
    
    // Column header index mapping
    var headers = gridValues[0].map(function(h) { return String(h || "").trim(); });
    
    // 2. Read Mapping Config
    var mapData = mappingSheet.getDataRange().getValues();
    var rulesCount = 0;
    
    // Group mapping rules by target presentation sheet
    var sheetUpdates = {}; // targetSheetName -> array of { cellA1, value }
    
    for (var i = 1; i < mapData.length; i++) {
      var rowSeason = String(mapData[i][0] || "").trim();
      if (rowSeason.toLowerCase() !== season.toLowerCase()) continue;
      
      var targetSheetName = String(mapData[i][1] || "").trim();
      var targetCellA1 = String(mapData[i][2] || "").trim();
      var sourceColsStr = String(mapData[i][3] || "").trim();
      var sourceRule = String(mapData[i][4] || "").trim();
      var expr = String(mapData[i][5] || "").trim();
      
      if (!targetSheetName || !targetCellA1 || !expr) continue;
      
      // Resolve grid row index from sourceRule
      var sourceRowIdx = parseSourceRowRule(sourceRule, gridValues);
      if (sourceRowIdx === -1) continue;
      
      var rowData = gridValues[sourceRowIdx];
      
      // Substitute placeholders in expr (e.g. {Speaker}, {Interpreter}, {Usher1}, {Usher2})
      var formattedVal = expr;
      for (var c = 0; c < headers.length; c++) {
        var hName = headers[c];
        var val = String(rowData[c] || "").trim();
        formattedVal = formattedVal.replace(new RegExp('\\{' + escapeRegExp(hName) + '\\}', 'g'), val);
      }
      
      // Clean up empty lines or prefixes
      formattedVal = cleanFormattedText(formattedVal);
      
      // Resolve any MasterConfig speaker placeholders (e.g. {VISITING_MINISTER}, {RESIDENT_MINISTER}, etc.)
      formattedVal = resolveSpeakerPlaceholders(ss, season, formattedVal);
      
      if (!sheetUpdates[targetSheetName]) sheetUpdates[targetSheetName] = [];
      sheetUpdates[targetSheetName].push({ a1: targetCellA1, val: formattedVal });
      rulesCount++;
    }
    
    // 3. Write to Target Presentation Sheets & Protect (Lock) Them
    for (var targetName in sheetUpdates) {
      var targetSheet = ss.getSheetByName(targetName);
      if (!targetSheet) continue;
      
      // Temporarily unprotect to allow write
      unprotectSheetTemporarily(targetSheet);
      
      var updates = sheetUpdates[targetName];
      if (updates.length > 0) {
        var minRow = Infinity, maxRow = -1, minCol = Infinity, maxCol = -1;
        var parsedUpdates = [];
        for (var u = 0; u < updates.length; u++) {
          var a1 = updates[u].a1;
          var range = targetSheet.getRange(a1);
          var r = range.getRow();
          var c = range.getColumn();
          minRow = Math.min(minRow, r);
          maxRow = Math.max(maxRow, r);
          minCol = Math.min(minCol, c);
          maxCol = Math.max(maxCol, c);
          parsedUpdates.push({ r: r, c: c, val: updates[u].val });
        }
        var targetRange = targetSheet.getRange(minRow, minCol, maxRow - minRow + 1, maxCol - minCol + 1);
        var values = targetRange.getValues();
        for (var u = 0; u < parsedUpdates.length; u++) {
          values[parsedUpdates[u].r - minRow][parsedUpdates[u].c - minCol] = parsedUpdates[u].val;
        }
        targetRange.setValues(values);
      }
      
      // Ensure official logo is synced to presentation sheet header
      syncLogoToPresentationSheet(targetSheet);
      
      // Re-lock presentation sheet
      lockPresentationSheet(targetSheet);
    }
    
    // 4. Inject Convocation & Evangelical Themes into Worship Presentation Sheet
    var worshipSheetName = (season === 'Spring') ? ESSC_CONFIG.PRESENTATION_SHEETS.SPRING_WORSHIP : ESSC_CONFIG.PRESENTATION_SHEETS.FALL_WORSHIP;
    var worshipSheet = ss.getSheetByName(worshipSheetName);
    if (worshipSheet) {
      unprotectSheetTemporarily(worshipSheet);
      var convTheme = getConfigValue(ss, season.toUpperCase() + "_CONVOCATION_THEME");
      var evanTheme = getConfigValue(ss, season.toUpperCase() + "_EVANGELICAL_THEME");
      if (convTheme) {
        worshipSheet.getRange("B4").setValue(convTheme);
      }
      if (evanTheme) {
        worshipSheet.getRange("B26").setValue("Evangelical Services* 佈道會 " + evanTheme);
      }
      lockPresentationSheet(worshipSheet);
    }
    
    SpreadsheetApp.getUi().alert("Success 🎉", "Successfully generated " + season + " ESSC presentation schedule (" + rulesCount + " mapping rules applied). Target presentation sheets are locked against manual edits.", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {
    SpreadsheetApp.getUi().alert("Generation Error", e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Parses source row rule from ESSC_MappingConfig
 */
function parseSourceRowRule(ruleStr, gridValues) {
  var rowMatch = ruleStr.match(/^(\d+)/);
  if (rowMatch) {
    var r = parseInt(rowMatch[1], 10);
    return r - 1; // 0-indexed
  }
  return -1;
}

/**
 * Resolves speaker placeholder tokens (e.g. {VISITING_MINISTER}, {RESIDENT_MINISTER}, etc.)
 * by looking up the corresponding setting in MasterConfig.
 */
function resolveSpeakerPlaceholders(ss, season, str) {
  if (!str || str.indexOf("{") === -1) return str;
  
  var prefix = (season || "").toUpperCase() + "_";
  var speakerKeys = [
    "VISITING_MINISTER",
    "RESIDENT_MINISTER",
    "LAYMINISTER_1",
    "LAYMINISTER_2",
    "LAYMINISTER_3",
    "LAYMINISTER_4",
    "CHAIRPERSON",
    "WEDNESDAY_CHOIR",
    "THURSDAY_CHOIR",
    "FRIDAY_CHOIR",
    "SATURDAY_CHOIR",
    "SUNDAY_CHOIR"
  ];
  
  for (var i = 0; i < speakerKeys.length; i++) {
    var key = speakerKeys[i];
    if (str.indexOf("{" + key + "}") !== -1 || str.indexOf("{" + prefix + key + "}") !== -1) {
      var masterSs = SpreadsheetApp.getActiveSpreadsheet();
      var val = getConfigValue(ss, prefix + key);
      if (!val && masterSs && masterSs.getId() !== ss.getId()) {
        val = getConfigValue(masterSs, prefix + key);
      }
      if (!val) {
        val = getConfigValue(ss, key);
      }
      if (!val && masterSs && masterSs.getId() !== ss.getId()) {
        val = getConfigValue(masterSs, key);
      }
      if (val) {
        str = str.replace(new RegExp('\\{' + prefix + key + '\\}', 'g'), val);
        str = str.replace(new RegExp('\\{' + key + '\\}', 'g'), val);
      }
    }
  }
  return str;
}

/**
 * Resolves speaker placeholder tokens (e.g. {VISITING_MINISTER}, {RESIDENT_MINISTER}, etc.)
 * directly inside Spring_ESSC_Grid and Fall_ESSC_Grid data entry sheets based on MasterConfig.
 */
function resolveGridSpeakerPlaceholders(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSs = SpreadsheetApp.getActiveSpreadsheet();
  var grids = [
    { name: ESSC_CONFIG.SPRING_GRID, season: 'Spring' },
    { name: ESSC_CONFIG.FALL_GRID, season: 'Fall' }
  ];
  var totalResolved = 0;
  
  var mcSheet = ss.getSheetByName(ESSC_CONFIG.MASTER_CONFIG_SHEET) || (masterSs ? masterSs.getSheetByName(ESSC_CONFIG.MASTER_CONFIG_SHEET) : null);
  
  // Track minister keys to check for changes
  var trackedKeys = [
    'VISITING_MINISTER',
    'SPRING_VISITING_MINISTER',
    'FALL_VISITING_MINISTER',
    'RESIDENT_MINISTER',
    'LAYMINISTER_1',
    'LAYMINISTER_2',
    'LAYMINISTER_3',
    'LAYMINISTER_4',
    'CHAIRPERSON'
  ];
  
  // Build replacement map for previously resolved minister names that have changed in MasterConfig
  var renameMap = {}; // oldName -> newName
  if (mcSheet) {
    for (var k = 0; k < trackedKeys.length; k++) {
      var tKey = trackedKeys[k];
      var currentVal = getConfigValue(ss, tKey) || (masterSs ? getConfigValue(masterSs, tKey) : "");
      var lastVal = getConfigValue(ss, "LAST_" + tKey) || (masterSs ? getConfigValue(masterSs, "LAST_" + tKey) : "");
      if (currentVal && lastVal && currentVal !== lastVal) {
        renameMap[lastVal.trim()] = currentVal.trim();
        // Also support single-line versions if bilingual newline
        if (lastVal.indexOf("\n") !== -1 && currentVal.indexOf("\n") !== -1) {
          renameMap[lastVal.replace(/\n/g, " / ").trim()] = currentVal.replace(/\n/g, " / ").trim();
        }
      }
    }
  }

  for (var g = 0; g < grids.length; g++) {
    var sheet = ss.getSheetByName(grids[g].name);
    if (!sheet) continue;
    var season = grids[g].season;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    
    // Column E is Speaker (col 5)
    var speakerRange = sheet.getRange(2, 5, lastRow - 1, 1);
    var values = speakerRange.getValues();
    var changed = false;
    for (var r = 0; r < values.length; r++) {
      var cellVal = String(values[r][0] || "").trim();
      if (!cellVal) continue;
      
      // 1. Resolve raw bracket tokens ({VISITING_MINISTER}, etc.)
      if (cellVal.indexOf("{") !== -1) {
        var resolved = resolveSpeakerPlaceholders(ss, season, cellVal);
        if (resolved !== cellVal) {
          values[r][0] = resolved;
          changed = true;
          totalResolved++;
        }
      } 
      // 2. Automatically swap previously resolved minister if name changed in MasterConfig
      else if (renameMap[cellVal]) {
        values[r][0] = renameMap[cellVal];
        changed = true;
        totalResolved++;
      }
    }
    if (changed) {
      speakerRange.setValues(values);
    }
  }

  // Update LAST_* tracking keys in MasterConfig to current values
  if (mcSheet) {
    for (var k = 0; k < trackedKeys.length; k++) {
      var tKey = trackedKeys[k];
      var currentVal = getConfigValue(ss, tKey) || (masterSs ? getConfigValue(masterSs, tKey) : "");
      if (currentVal) {
        updateOrAppendConfigRow(mcSheet, "LAST_" + tKey, currentVal);
        if (masterSs && masterSs.getId() !== ss.getId()) {
          var masterMc = masterSs.getSheetByName(ESSC_CONFIG.MASTER_CONFIG_SHEET);
          if (masterMc) updateOrAppendConfigRow(masterMc, "LAST_" + tKey, currentVal);
        }
      }
    }
  }

  return totalResolved;
}

/**
 * Injects dynamic date and speaker XLOOKUP formulas into Spring_ESSC_Grid and Fall_ESSC_Grid
 * so that any edits to MasterConfig dates or ministers update the grids in real time.
 */
function injectDynamicGridFormulas(targetSs) {
  if (!targetSs) targetSs = getTargetESSCSpreadsheet();
  if (!targetSs) return 0;
  
  var seasons = [
    { name: ESSC_CONFIG.SPRING_GRID, season: 'Spring', prefix: 'SPRING_' },
    { name: ESSC_CONFIG.FALL_GRID, season: 'Fall', prefix: 'FALL_' }
  ];
  
  var speakerRoleMap = {
    4: 'VISITING_MINISTER',
    7: 'VISITING_MINISTER',
    8: 'RESIDENT_MINISTER',
    10: 'VISITING_MINISTER',
    12: 'RESIDENT_MINISTER',
    14: 'LAYMINISTER_4',
    16: 'LAYMINISTER_3',
    18: 'VISITING_MINISTER',
    19: 'RESIDENT_MINISTER',
    22: 'VISITING_MINISTER',
    23: 'LAYMINISTER_1',
    24: 'RESIDENT_MINISTER',
    25: 'VISITING_MINISTER',
    27: 'LAYMINISTER_3',
    29: 'LAYMINISTER_2',
    31: 'RESIDENT_MINISTER',
    32: 'LAYMINISTER_2',
    35: 'VISITING_MINISTER',
    36: 'LAYMINISTER_4',
    37: 'RESIDENT_MINISTER',
    39: 'RESIDENT_MINISTER',
    42: 'CHAIRPERSON',
    43: 'RESIDENT_MINISTER',
    45: 'VISITING_MINISTER'
  };

  var totalInjected = 0;

  for (var s = 0; s < seasons.length; s++) {
    var sInfo = seasons[s];
    var sheet = targetSs.getSheetByName(sInfo.name);
    if (!sheet) continue;
    var lastRow = Math.max(sheet.getLastRow(), 45);

    // 1. Column A (Date Formulas)
    // Row 2: Wednesday start date formula
    var dateFormulas = [];
    var dateKey = sInfo.prefix + "WEDNESDAY_START_DATE";
    var altDateKey = sInfo.prefix + "ESSC_WEDNESDAY_START";
    var row2DateFormula = '=IFERROR(DATEVALUE(TO_TEXT(XLOOKUP("' + dateKey + '", MasterConfig!$A:$A, MasterConfig!$B:$B, XLOOKUP("' + altDateKey + '", MasterConfig!$A:$A, MasterConfig!$B:$B, "")))), "")';
    dateFormulas.push([row2DateFormula]);

    for (var r = 3; r <= lastRow; r++) {
      if (r <= 4) {
        dateFormulas.push(['=A$2']);       // Wed
      } else if (r <= 7) {
        dateFormulas.push(['=A$2 + 1']);   // Thu
      } else if (r <= 22) {
        dateFormulas.push(['=A$2 + 2']);   // Fri
      } else if (r <= 35) {
        dateFormulas.push(['=A$2 + 3']);   // Sat
      } else {
        dateFormulas.push(['=A$2 + 4']);   // Sun
      }
    }
    var dateRange = sheet.getRange(2, 1, dateFormulas.length, 1);
    dateRange.setFormulas(dateFormulas);
    dateRange.setNumberFormat('yyyy-mm-dd');

    // 2. Column E (Speaker Formulas)
    var speakerFormulas = sheet.getRange(2, 5, lastRow - 1, 1).getFormulas();

    for (var rowIdx = 2; rowIdx <= lastRow; rowIdx++) {
      var arrIdx = rowIdx - 2;
      var role = speakerRoleMap[rowIdx];

      if (role) {
        var formula = "";
        if (role === 'VISITING_MINISTER') {
          formula = '=IFERROR(XLOOKUP("' + sInfo.prefix + 'VISITING_MINISTER", MasterConfig!$A:$A, MasterConfig!$B:$B, XLOOKUP("VISITING_MINISTER", MasterConfig!$A:$A, MasterConfig!$B:$B, "")), "")';
        } else {
          formula = '=IFERROR(XLOOKUP("' + role + '", MasterConfig!$A:$A, MasterConfig!$B:$B, ""), "")';
        }
        speakerFormulas[arrIdx][0] = formula;
        totalInjected++;
      } else {
        speakerFormulas[arrIdx][0] = "";
      }
    }
    sheet.getRange(2, 5, speakerFormulas.length, 1).setFormulas(speakerFormulas);
  }

  return totalInjected;
}

/**
 * Menu Handler to Inject Dynamic Grid Formulas into Standalone ESSC Workbook
 */
function injectDynamicGridFormulasMenu() {
  var targetSs = getTargetESSCSpreadsheet();
  if (!targetSs) return;
  var count = injectDynamicGridFormulas(targetSs);
  SpreadsheetApp.getUi().alert("Dynamic Formulas Injected 🎉", "Successfully injected " + count + " dynamic date & speaker formulas into '" + targetSs.getName() + "'.\n\nEditing MasterConfig dates or ministers will now update all grids in real time!", SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Escapes regex special characters in a string
 */
function escapeRegExp(str) {
  return String(str || "").replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Ensures the official TJC logo image is synced onto cell B1 (Col 2, Row 1) of the presentation sheet,
 * clears any legacy cell formulas in B1, sets Row 1 background to Theme Primary, adjusts Row 1 height,
 * sets cell D1 font color to WHITE, and sets Columns E & F width to 125.
 */
function syncLogoToPresentationSheet(sheet) {
  try {
    // 1. Clear any default text or legacy cell formula in B1:H1 (e.g. legacy =IMAGE(...) formula)
    sheet.getRange("B1:H1").clearContent();
    
    // 2. Set Row 1 background to Theme Primary for white logo contrast
    var maxCols = sheet.getMaxColumns();
    var colsToFormat = Math.min(maxCols, 8);
    sheet.getRange(1, 1, 1, colsToFormat).setBackground(typeof APP_CONFIG !== 'undefined' && APP_CONFIG.THEME ? APP_CONFIG.THEME.PRIMARY_COLOR : "#1F4E79");

    // 3. Set Row 1 height to 65 to fit logo completely
    sheet.setRowHeight(1, 65);
    
    // 4. Set cell D1 font color to WHITE
    sheet.getRange("D1").setFontColor("#FFFFFF").setFontWeight("bold");
    
    // 5. Adjust default column width of Columns E & F to 125 each
    sheet.setColumnWidth(5, 125);
    sheet.setColumnWidth(6, 125);

    // 6. Insert official TJC logo image if not already present
    var images = sheet.getImages();
    if (images && images.length > 0) {
      return;
    }
    var logoFileId = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.TJC_LOGO_FILE_ID) ? APP_CONFIG.TJC_LOGO_FILE_ID : "15-OcvnePRkYVBg5yz63ILCcHxvVYLQIr";
    var logoBlob = (typeof getLogoBlob === 'function') ? getLogoBlob() : null;
    if (!logoBlob) {
      try {
        logoBlob = DriveApp.getFileById(logoFileId).getBlob().setName("tjcLogo.png");
      } catch (err) {
        var response = UrlFetchApp.fetch("https://drive.google.com/uc?export=view&id=" + logoFileId);
        logoBlob = response.getBlob().setName("tjcLogo.png");
      }
    }
    if (logoBlob) {
      sheet.insertImage(logoBlob, 2, 1);
    }
  } catch (e) {
    Logger.log("syncLogoToPresentationSheet warning for " + sheet.getName() + ": " + e.toString());
  }
}

/**
 * Cleans formatted text expressions with empty placeholders, prefixes, and time headers
 */
function cleanFormattedText(str) {
  var lines = str.split('\n');
  var cleanedLines = [];
  var hasActualWorker = false;
  
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;
    
    // Check if line is just a prefix like 'Int:', 'Lunch:', 'Dinner:'
    var strippedPrefix = line.replace(/^(Int:|Lunch:|Dinner:)\s*/i, '').trim();
    
    // Check if line is a time header or session label (e.g. '9:00 - 11:00 AM', 'Baptism Sermon', etc.)
    var isHeaderOnly = /^(\d{1,2}:\d{2}|Baptism Sermon|Water Baptism|Introductions|Choir Presentation|Holy Communion|Foot Washing|Sermon|Prayer|見證會|洗禮|聖餐禮|洗腳禮)/i.test(strippedPrefix);
    
    if (strippedPrefix && !isHeaderOnly) {
      hasActualWorker = true;
    }
    
    if (strippedPrefix) {
      // Strip trailing slash if second worker placeholder was empty (e.g. 'Person1 /')
      line = line.replace(/\s*\/$/, '').trim();
      if (line) {
        cleanedLines.push(line);
      }
    }
  }
  
  if (!hasActualWorker) {
    return '';
  }
  return cleanedLines.join('\n');
}

/**
 * Protects (locks) a presentation sheet so coordinators can only edit the data grid
 */
function lockPresentationSheet(sheet) {
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  var protection;
  if (protections.length > 0) {
    protection = protections[0];
  } else {
    protection = sheet.protect();
  }
  protection.setDescription("Locked ESSC Presentation Schedule (Generated from Grid)");
  protection.setWarningOnly(false);
  
  // Allow spreadsheet owners, revoke for general coordinators
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
}

/**
 * Temporarily unprotects a presentation sheet for script updates
 */
function unprotectSheetTemporarily(sheet) {
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var i = 0; i < protections.length; i++) {
    protections[i].remove();
  }
}



/**
 * Static column header definitions for ESSC 16-column grid
 */
const ESSC_GRID_HEADERS = [
  'Date', 'Day', 'Time block', 'Session / Title',
  'Speaker', 'Interpreter', 'Hymn Leader', 'Pianist',
  'AV1', 'AV2', 'Usher1', 'Usher2', 'Cleaning',
  'Lunchtime Grace', 'Dinnertime Grace', 'Choir'
];

/**
 * Role name to Rosters column header mapping
 */
const ESSC_ROLE_TO_ROSTER_KEY = {
  'Speaker': 'Speakers',
  'Interpreter': 'Interpreters',
  'Hymn Leader': 'Hymn Leaders',
  'Pianist': 'Pianists',
  'AV1': 'AV',
  'AV2': 'AV',
  'Usher1': 'Ushers',
  'Usher2': 'Ushers',
  'Cleaning': 'Cleaning',
  'Lunchtime Grace': 'Grace',
  'Dinnertime Grace': 'Grace',
  'Choir': 'Choir'
};

/**
 * Opens the interactive ESSC Assignment Sidebar using the unified Sidebar.html
 */
function showSidebarESSC() {
  var html = HtmlService.createTemplateFromFile('Sidebar').evaluate()
    .setTitle('ESSC Schedule Assistant')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Simple trigger that fires whenever selection changes in the spreadsheet.
 * Immediately caches active cell coordinate to reduce RPC latency.
 */
function onSelectionChangeESSC(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    var sName = sheet.getName();
    if (sName !== ESSC_CONFIG.SPRING_GRID && sName !== ESSC_CONFIG.FALL_GRID) return;
    
    var cache = CacheService.getUserCache();
    cache.put('ESSC_ACTIVE_CELL', JSON.stringify({
      sheet: sName,
      row: e.range.getRow(),
      col: e.range.getColumn(),
      a1: e.range.getA1Notation()
    }), 15);
  } catch(err) {
    console.error('Error in onSelectionChange:', err);
  }
}

/**
 * Initial fast payload: reads all rosters and grid definitions once into client memory
 */
function getESSCInitialData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rostersSheet = ss.getSheetByName('Rosters');
    if (!rostersSheet || rostersSheet.getLastRow() <= 1) {
      rostersSheet = ensureRostersSheet(ss, false);
    }
    
    var rosters = {};
    if (rostersSheet && rostersSheet.getLastRow() > 1) {
      var data = rostersSheet.getDataRange().getValues();
      var maxR = data.length;
      var maxC = data[0].length;
      for (var c = 0; c < maxC; c++) {
        var hName = '';
        var sRow = 1;
        var r0 = String(data[0][c] || '').trim();
        var r1 = (maxR > 1) ? String(data[1][c] || '').trim() : '';
        if (r1 && (!r0 || !isNaN(Number(r0)) || r1.match(/[a-zA-Z]/))) {
          hName = r1;
          sRow = 2;
        } else {
          hName = r0;
          sRow = 1;
        }
        if (!hName) continue;
        rosters[hName] = [];
        for (var r = sRow; r < maxR; r++) {
          var name = String(data[r][c] || '').trim();
          if (name && rosters[hName].indexOf(name) === -1) {
            rosters[hName].push(name);
          }
        }
      }
    }
    
    var grids = {};
    var gridNames = [ESSC_CONFIG.SPRING_GRID, ESSC_CONFIG.FALL_GRID];
    for (var g = 0; g < gridNames.length; g++) {
      var gName = gridNames[g];
      var sh = ss.getSheetByName(gName);
      if (sh) {
        var lastR = Math.max(sh.getLastRow(), ESSC_CONFIG.GRID.LAST_ROW);
        var vals = sh.getRange(1, 1, lastR, ESSC_CONFIG.GRID.LAST_COL).getValues();
        grids[gName] = vals.map(function(row) {
          return row.map(function(cellVal) {
            if (cellVal instanceof Date) {
              return formatDateSafe(cellVal, ss);
            }
            return String(cellVal || '').trim();
          });
        });
      }
    }
    
    return { success: true, rosters: rosters, grids: grids };
  } catch(e) {
    return { success: false, error: e.toString(), rosters: {}, grids: {} };
  }
}

/**
 * Ultra-fast active cell monitor (sub-20ms execution)
 * Returns { unchanged: true } when user stays on same cell.
 */
function getESSCActiveCell(lastA1) {
  try {
    var cache = CacheService.getUserCache();
    var cached = cache.get('ESSC_ACTIVE_CELL');
    var sName, r, c, a1;
    
    if (cached) {
      try {
        var parsed = JSON.parse(cached);
        sName = parsed.sheet;
        r = parsed.row;
        c = parsed.col;
        a1 = parsed.a1;
      } catch(ex) {}
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!a1) {
      console.log('Selection cache expired or missing; falling back to getActiveCell()');
      var sheet = ss.getActiveSheet();
      sName = sheet.getName();
      var isGrid = (sName === ESSC_CONFIG.SPRING_GRID || sName === ESSC_CONFIG.FALL_GRID);
      if (!isGrid) {
        return { success: true, isGridCell: false, sheetName: sName, message: 'Select a cell in ' + ESSC_CONFIG.SPRING_GRID + ' or ' + ESSC_CONFIG.FALL_GRID };
      }
      var cell = sheet.getActiveCell();
      a1 = cell.getA1Notation();
      r = cell.getRow();
      c = cell.getColumn();
    }
    
    if (sName !== ESSC_CONFIG.SPRING_GRID && sName !== ESSC_CONFIG.FALL_GRID) {
      return { success: true, isGridCell: false, sheetName: sName, message: 'Select a cell in ' + ESSC_CONFIG.SPRING_GRID + ' or ' + ESSC_CONFIG.FALL_GRID };
    }
    
    if (a1 === lastA1) {
      return { success: true, unchanged: true, a1: a1, sheetName: sName };
    }
    
    if (r < ESSC_CONFIG.GRID.FIRST_ROW || r > ESSC_CONFIG.GRID.LAST_ROW || c < ESSC_CONFIG.GRID.FIRST_COL || c > ESSC_CONFIG.GRID.LAST_COL) {
      return { success: true, isGridCell: false, a1: a1, sheetName: sName, message: 'Select an assignment slot (Cols E–P, Rows ' + ESSC_CONFIG.GRID.FIRST_ROW + '–' + ESSC_CONFIG.GRID.LAST_ROW + ')' };
    }
    
    var targetSheet = ss.getSheetByName(sName);
    var val = String(targetSheet.getRange(r, c).getValue() || '').trim();
    
    return {
      success: true,
      unchanged: false,
      isGridCell: true,
      sheetName: sName,
      row: r,
      col: c,
      a1: a1,
      currentVal: val
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * High-speed context fetcher (legacy fallback)
 */
function getESSCCellContext() {
  return getESSCActiveCell('');
}

/**
 * Ensures the 'Rosters' tab exists and is populated.
 * Only connects to Master HWS when explicitly requested or if Rosters is completely missing.
 */
function ensureRostersSheet(ss, forceSyncFromMaster) {
  var rostersSheet = ss.getSheetByName('Rosters');
  if (!rostersSheet) {
    rostersSheet = ss.insertSheet('Rosters');
  }
  
  // Only connect to Master HWS if explicitly forced or sheet is brand new/empty
  var shouldSyncMaster = forceSyncFromMaster || (rostersSheet.getLastRow() <= 1);
  var pulledFromMaster = false;
  
  if (shouldSyncMaster) {
    try {
      var masterSs = getMasterHWSSpreadsheet();
      if (masterSs) {
        var workersSheet = masterSs.getSheetByName('Workers');
        if (workersSheet) {
          var masterRosters = extractRostersFromMasterWorkers(workersSheet);
          if (masterRosters && Object.keys(masterRosters).length > 0) {
            writeRostersToSheet(rostersSheet, masterRosters);
            pulledFromMaster = true;
          }
        }
      }
    } catch(e) {
      console.log("Master HWS sync skipped: " + e.message);
    }
  }
  
  if (pulledFromMaster || rostersSheet.getLastRow() > 1) {
    return rostersSheet;
  }
  
  // Fallback: populate standard 2027 church roster
  var defaultRosters = {
    "Speakers": ["Alan Wang", "Andrew Wang", "Anthony Chang", "Anthony Tsai", "Benson Kuo", "Bruce Wang", "C. H. Chou", "C.Y. Kuo", "Caroline Huang", "Chad Liu", "Charles Tsai", "David Liou", "Derren Liang", "Enoch Hou", "Ephraim Liu", "Ethan Wang", "F. F. Chong", "George Jeng", "H. H. Ko", "Horn Hsieh", "Jackson Jeng", "James Chen", "James Liu", "Janice Chang", "Jessica Chen", "John Huang", "Johnny Liu", "Jordan Shen", "Joseph Jung", "Joseph Kim", "Josh Yeh", "K. M. Huang", "Lawrence Lee", "Mark Du", "Matthew Shih", "Meishi Tsai", "Micah Lee", "Michael Tang", "Mike Yuan", "N. H. Chen", "P. H. Liu", "Patrick Wong", "Paul Chen", "Paul Lai", "Raymond Chou", "Richard Chang", "Richard Pang", "Samuel Kuo", "Scott Tsai", "Sean Yang", "Shawn Chou", "Solomon Ho", "Stephen Ku", "Steve Hwang", "Steven Lin", "Tim Chen", "Y.C. Lee", "Yin Chu Lee"],
    "Interpreters": ["Adriane Young", "Andrea Liu", "Dan Wang", "Daniel Chong", "Elaine Pang", "Ephraim Liu", "Janice Chang", "Jeremy Chang", "Jessica Chen", "Jessica Wang", "Joanne Shih", "Jordan Shen", "Kevin Liu", "Lawrence Lee", "Linda Suryani", "Matthew Shih", "Philip Chen", "Phylicia Ng", "Sarah Pai", "Tricia Chen"],
    "Hymn Leaders": ["Albert Tan", "Alice Tan", "Anthony Chang", "Anthony Tsai", "Bright Lin", "Caroline Huang", "Cera Han", "Cindy Lin", "Connor Ho", "Darrus Lee", "Elaine Pang", "Enoch Liu", "Ephraim Liu", "Gabriel Shen", "Grant Liu", "Janice Chang", "Jeremy Chang", "Jesriel Shen", "Jessica Chen", "Jonathan Chen", "Jordan Shen", "Joseph Tan", "Joshua Lai", "Justina Lai", "Kevin Liu", "Matt Shih", "Moses Han", "Naomi Lee"],
    "Pianists": ["Albert Tan", "Alice Tan", "Ally Huang", "Annie Chao", "Chloe Huang", "Cindy Lin", "Connor Ho", "Derrick Hwang", "Elaine Pang", "Esther Yang", "Gabriel Shen", "Janice Chang", "Jeremy Chang", "Jessica Chen", "Jocelyn Ho", "Justina Lai", "Lydia Chang", "Matt Shih", "Mei Helms", "Momo Kuo", "Phylicia Ng", "Sarah Pai", "Tricia Chen", "Yolanda Yang"],
    "AV": ["Albert Tan", "Andrea Liu", "Anson Thong", "Anthony Tsai", "Benny  Liu", "Cera Han", "Charlie Shen", "Connor Ho", "Dan Wang", "Daniel Chong", "Darrus Lee", "Elaine Pang", "Enoch Liu", "Gabriel Shen", "Jesriel Shen", "Jordan Shen", "Joshua Lai", "Kevin Liu", "Lawrence Lee", "Matt Shih", "Mei Helms", "Michael Yu", "Moses Han", "Naomi Lee", "Paul Lai", "Phylicia Ng", "Rio Shen", "Sarah Pai"],
    "Ushers": ["Ally Huang", "Amy Huang", "Andrea Liu", "Andrew Chiu", "Annie Chao", "Anson Thong", "Anthony Chang", "Anthony Tsai", "Bright Lin", "Cera Han", "Charlie Shen", "Chloe Huang", "Cindy Lin", "Clementine Diep", "Connor Ho", "Daniel Chong", "Darrus Lee", "Derrick Hwang", "Elaine Pang", "Ephraim Liu", "Grace Chou", "Grace Lin", "Grant Liu", "Janice Chang", "Jeremy Chang", "Jessica Chen", "Jocelyn Ho", "Jonathan Chen", "Jordan Shen", "Joseph Tan", "Judy Wang", "Justina Lai", "Kevin Liu", "Lawrence Lee", "Lydia Chang", "Margaretta Chiu", "Matt Shih", "Michael Tang", "Michael Wei", "Michelle Li", "Moses Han", "Naomi Lee", "Paul Lai", "Peggy Liu", "Rio Shen", "Sandy Huang", "Sarah Pai", "Simon Liu", "Yolanda Yang"],
    "Cleaning": ["Cypress / Buena Park", "Anaheim Hills / Yorba Linda", "Irvine / Lake Forest / Orange", "Agape / EBF / Sisters / Vessel", "R.E. Coordinator"],
    "Grace": ["Albert Tan", "Amy Huang", "Andrea Liu", "Andrew Chiu", "Andrew Young", "Anthony Chang", "Anthony Tsai", "Bright Lin", "Bruce Wang", "Caroline Huang", "Cera Han", "Charlie Shen", "Cindy Lin", "Dan Wang", "Daniel Chong", "Darrus Lee", "David Liou", "Derrick Hwang", "Elaine Pang", "Ephraim Liu", "Esther Yang", "Gabriel Shen", "Grace Chou", "Grace Lin", "Janice Chang", "Jeremy Chang", "Jessica Chen", "Joanne Shih", "Jocelyn Ho", "Jordan Shen", "Joseph Tan", "Justina Lai", "Kevin Liu", "Lawrence Lee", "Linda Suryani", "Lydia Chang", "Margaretta Chiu", "Matt Shih", "Matthew Shih", "Mei Helms", "Michael Tang", "Michael Wei", "Michelle Li", "Momo Kuo", "Moses Han", "Naomi Lee", "Paul Lai", "Paulina Tse", "Peggy Liu", "Pilan Chao", "Rio Shen", "Sandy Huang", "Sarah Pai", "Simon Liu", "Solomon Ho", "Tricia Chen", "Tristina Lam"],
    "Choir": ["Adult Choir", "Youth Choir", "Children's Choir"]
  };
  
  writeRostersToSheet(rostersSheet, defaultRosters);
  return rostersSheet;
}

/**
 * Extracts ministry worker lists from Master HWS Workers tab
 */
function extractRostersFromMasterWorkers(ws) {
  var data = ws.getDataRange().getValues();
  if (data.length < 3) return null;
  
  function getWorkersForCols(colIndices) {
    var workers = [];
    for (var r = 2; r < data.length; r++) {
      var row = data[r];
      var first = String(row[0] || '').trim();
      var last = String(row[1] || '').trim();
      var active = String(row[11] || '').trim();
      if (!first || active !== 'Active') continue;
      
      var fullName = (first + ' ' + last).trim();
      var assigned = false;
      for (var c = 0; c < colIndices.length; c++) {
        var val = row[colIndices[c] - 1];
        if (val === true || val === 'TRUE' || val === 1 || val === '1') {
          assigned = true;
          break;
        }
      }
      if (assigned && workers.indexOf(fullName) === -1) {
        workers.push(fullName);
      }
    }
    workers.sort();
    return workers;
  }
  
  return {
    "Speakers": getWorkersForCols([16, 17]),
    "Interpreters": getWorkersForCols([18, 19]),
    "Hymn Leaders": getWorkersForCols([20]),
    "Pianists": getWorkersForCols([21]),
    "AV": getWorkersForCols([23, 24]),
    "Ushers": getWorkersForCols([22]),
    "Cleaning": ["Cypress / Buena Park", "Anaheim Hills / Yorba Linda", "Irvine / Lake Forest / Orange", "Agape / EBF / Sisters / Vessel", "R.E. Coordinator"],
    "Grace": getWorkersForCols([15, 25, 26, 27]),
    "Choir": ["Adult Choir", "Youth Choir", "Children's Choir"]
  };
}

/**
 * Writes roster map to Rosters sheet
 */
function writeRostersToSheet(sheet, rosterMap) {
  var keys = ["Speakers", "Interpreters", "Hymn Leaders", "Pianists", "AV", "Ushers", "Cleaning", "Grace", "Choir"];
  var maxRows = 0;
  for (var k = 0; k < keys.length; k++) {
    var list = rosterMap[keys[k]] || [];
    if (list.length > maxRows) maxRows = list.length;
  }
  
  var gridData = [];
  gridData.push(keys);
  for (var r = 0; r < maxRows; r++) {
    var row = [];
    for (var k = 0; k < keys.length; k++) {
      var arr = rosterMap[keys[k]] || [];
      row.push(r < arr.length ? arr[r] : '');
    }
    gridData.push(row);
  }
  
  sheet.clear();
  sheet.getRange(1, 1, gridData.length, keys.length).setValues(gridData);
  sheet.getRange(1, 1, 1, keys.length).setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff');
}

/**
 * Assigns a worker to the selected grid cell
 */
function assignESSCWorker(sheetName, row, col, workerName) {
  try {
    const r = parseInt(row, 10);
    const c = parseInt(col, 10);
    const { FIRST_ROW, LAST_ROW, FIRST_COL, LAST_COL } = ESSC_CONFIG.GRID;
    if (isNaN(r) || r < FIRST_ROW || r > LAST_ROW || isNaN(c) || c < FIRST_COL || c > LAST_COL) {
      return { success: false, error: 'Target cell is outside valid grid bounds.' };
    }
    const cleanWorker = String(workerName || '').trim().slice(0, 100);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sheet not found' };

    const lock = LockService.getDocumentLock();
    if (!lock.tryLock(5000)) {
      return { success: false, error: 'Could not obtain lock. Another user may be updating the sheet.' };
    }
    try {
      sheet.getRange(r, c).setValue(cleanWorker);
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

