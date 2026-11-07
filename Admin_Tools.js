/**
 * Google Sheets Schedule Assistant - Server-side Code
 * File: Admin_Tools.gs
 */

const CONFIG = {
  // --- External Integrations ---
  SWR_SOURCE_QUARTERS: ['Q1', 'Q2', 'Q3', 'Q4'],
  HOLY_WORK_SCHEDULES_FOLDER_ID: (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.HOLY_WORK_SCHEDULES_FOLDER_ID) ? APP_CONFIG.HOLY_WORK_SCHEDULES_FOLDER_ID : '1sH-ELkxwRLB78ts2rJeTpPmYtDY8KfzE', // Parent directory in Google Drive
  UNAVAILABILITY_SHEET_PROPERTY: 'UNAVAILABILITY_SPREADSHEET_ID',
  SERMON_TOPICS_SHEET_PROPERTY: 'SERMON_TOPICS_SHEET_ID',
  SERMON_TOPICS_SHEET_URL_PROPERTY: 'SERMON_TOPICS_SHEET_URL',
  DEFAULT_SERMON_TOPICS_ID: (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.DEFAULT_SERMON_TOPICS_ID) ? APP_CONFIG.DEFAULT_SERMON_TOPICS_ID : '17qRYYuZTV3PWf92CnMmqwARaC2bk4BUPxz8DQ_9f3hU', // 2027 Sermon Topic & Hymns workbook
  DEFAULT_UNAVAILABILITY_ID: (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.DEFAULT_UNAVAILABILITY_ID) ? APP_CONFIG.DEFAULT_UNAVAILABILITY_ID : '1Ge3iKjPUIkJCuAdlRDqhgEWt4DlO7dL3mFyl0TVLuaY',
  
  // --- Assignment Settings ---
  // How many days to check backward/forward to warn about scheduling someone too closely
  PROXIMITY_WARNING_DAYS: 14, 
  
  // --- Sheet Names ---
  QUARTERS: ['Q1', 'Q2', 'Q3', 'Q4'],
  TEAMS_SHEET: 'Teams',
  UNAVAILABILITY_SHEET: 'Unavailability',
  WORKERS_SHEET: 'Workers',
  SYNC_LOG_SHEET: 'SWR Sync Log',
  
  // --- Layout Boundaries (Used for memory optimization) ---
  START_ROW: 5,             // Header row; data begins at Row 6
  MAX_COLUMNS: 100,         // Supports up to Column CV (Supports Cleaning roles in Col BJ+)
  MAX_SCHEDULE_ROWS: 250,   // Maximum rows to scan in Quarterly sheets
  MAX_TEAM_ROWS: 1000,      // Maximum rows to scan in Teams sheet
  MAX_TEAM_COLS: 65,        // Maximum columns to scan in Teams sheet (covers all 41+ cols including Cooking & Cleaning)
  MAX_WORKER_ROWS: 1500,    // Maximum rows to scan in Workers sheet
  MAX_WORKER_COLS: 75,      // Maximum columns to scan in Workers sheet (covers all 55+ cols including Cleaning Leaders)
  MAX_UNAVAIL_ROWS: 3000,   // Maximum rows to scan in Unavailability sheet
  
  // --- Performance & Cache Settings (in seconds) ---
  CACHE_POLL_TIME: 1200,    // 20 minutes for active cell polling
  CACHE_DATA_TIME: 1800,    // 30 minutes for teams/workers/unavailability data cache
};

function getAppConfig() {
  return CONFIG;
}

// -----------------------------------------------------------------------------
// ESSC (Evangelical Service & Spiritual Convocation) Block Definitions & Helpers
// -----------------------------------------------------------------------------
const ESSC_BLOCK_DEFINITIONS = [
  { key: 'WED_EVE', dayOffset: 0, label: 'Wednesday Evening', matchSessions: ['dinner', 'evening hymns', 'evangelical service'] },
  { key: 'THU_EVE', dayOffset: 1, label: 'Thursday Evening', matchSessions: ['dinner', 'evening hymns', 'evangelical service'] },
  { key: 'FRI_AM',  dayOffset: 2, label: 'Friday Morning', matchSessions: ['morning prayer', 'hymns 1', 'sermon 1', 'hymns 2', 'sermon 2'] },
  { key: 'FRI_PM1', dayOffset: 2, label: 'Friday Early Afternoon', matchSessions: ['lunch', 'fasting prayer', 'hymns 3', 'sermon 3', 'testimony'] },
  { key: 'FRI_PM2', dayOffset: 2, label: 'Friday Late Afternoon', matchSessions: ['hymns 4', 'sermon 4', 'afternoon prayer'] },
  { key: 'FRI_EVE', dayOffset: 2, label: 'Friday Evening', matchSessions: ['dinner', 'evening hymns', 'evening sermon'] },
  { key: 'SAT_AM',  dayOffset: 3, label: 'Saturday Morning', matchSessions: ['morning prayer', 'sermon 1', 'evangelical service'] },
  { key: 'SAT_PM',  dayOffset: 3, label: 'Saturday Afternoon', matchSessions: ['lunch', 'fasting prayer', 'testimony', 'hymns 2', 'sermon 2', 'afternoon prayer'] },
  { key: 'SAT_EVE', dayOffset: 3, label: 'Saturday Evening', matchSessions: ['dinner', 'evening hymns', 'evening sermon'] },
  { key: 'SUN_AM',  dayOffset: 4, label: 'Sunday Morning (Baptism)', matchSessions: ['baptism sermon', 'water baptism', 'hymns', 'sermon 4', 'fasting prayer'] },
  { key: 'SUN_PM',  dayOffset: 4, label: 'Sunday Afternoon (Communion)', matchSessions: ['lunch', 'introductions', 'foot washing', 'choir presentation', 'holy communion'] }
];

function cleanLabel(str) {
  if (!str) return '';
  var eng = String(str).replace(/[^\x00-\x7F]/g, "").replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
  return eng || String(str).trim();
}

function getMasterConfigStartDate(ss, configKey) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (typeof getConfigValue === 'function') {
    var v = getConfigValue(ss, configKey);
    if (v) return v;
  }
  var mc = ss.getSheetByName('MasterConfig');
  if (mc) {
    var data = mc.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === configKey) {
        var val = data[i][1];
        if (val instanceof Date) {
          return Utilities.formatDate(val, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
        }
        return String(val || '').trim();
      }
    }
  }
  return '';
}

function parseAnchorDate(val) {
  if (!val) return null;
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    return new Date(val.getFullYear(), val.getMonth(), val.getDate(), 12, 0, 0);
  }
  var str = String(val).trim();
  var parts = str.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
  if (parts) {
    return new Date(parseInt(parts[1], 10), parseInt(parts[2], 10) - 1, parseInt(parts[3], 10), 12, 0, 0);
  }
  var d = new Date(str);
  if (!isNaN(d.getTime())) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0);
  }
  return null;
}

function parseTimeRangeMinutes(timeStr) {
  if (!timeStr) return { start: null, end: null };
  var str = String(timeStr).trim().toUpperCase();
  var parts = str.split('-');
  
  function parseSingleTime(p, defaultAmPm) {
    var m = p.match(/(\d+):?(\d+)?\s*(AM|PM)?/i);
    if (!m) return null;
    var h = parseInt(m[1], 10);
    var mins = parseInt(m[2] || '0', 10);
    var ampm = m[3] ? m[3].toUpperCase() : defaultAmPm;
    if (ampm === 'PM' && h < 12) h += 12;
    else if (ampm === 'AM' && h === 12) h = 0;
    return h * 60 + mins;
  }
  
  if (parts.length === 1) {
    var t = parseSingleTime(parts[0], null);
    if (t === null) return { start: null, end: null };
    return { start: t, end: t + 45 };
  } else {
    var p2 = parts[1].trim();
    var ampm2 = (p2.indexOf('PM') !== -1) ? 'PM' : ((p2.indexOf('AM') !== -1) ? 'AM' : null);
    var endMin = parseSingleTime(p2, ampm2);
    
    var p1 = parts[0].trim();
    var ampm1 = (p1.indexOf('PM') !== -1) ? 'PM' : ((p1.indexOf('AM') !== -1) ? 'AM' : null);
    if (!ampm1 && ampm2 === 'PM') {
      var m1 = p1.match(/(\d+):?(\d+)?/);
      if (m1) {
        var h1 = parseInt(m1[1], 10);
        var min1 = parseInt(m1[2] || '0', 10);
        if (h1 < 12 && (h1 + 12) * 60 + min1 > endMin) {
          ampm1 = 'AM';
        } else {
          ampm1 = 'PM';
        }
      }
    }
    var startMin = parseSingleTime(p1, ampm1 || ampm2);
    return { start: startMin, end: endMin };
  }
}

function resolveCookingTeamWorkers(rawTeamStr, teamsMap) {
  if (!rawTeamStr || !teamsMap) return [];
  var cleanStr = cleanLabel(rawTeamStr);
  var rawLower = cleanStr.toLowerCase();
  
  // 1. Direct match in teamsMap
  for (var k in teamsMap) {
    if (k.toLowerCase() === rawLower) {
      return (teamsMap[k] || []).slice();
    }
  }
  
  // 2. Check numbered cooking teams (e.g. "Cooking Team 1", "Team 1")
  var numMatch = rawLower.match(/team\s*(\d+)/i) || rawLower.match(/cooking\s*(\d+)/i);
  if (numMatch) {
    var teamNum = numMatch[1];
    for (var k in teamsMap) {
      var kl = k.toLowerCase();
      if (kl.indexOf('cooking team ' + teamNum) !== -1 || kl.indexOf('team ' + teamNum) !== -1) {
        return (teamsMap[k] || []).slice();
      }
    }
  }
  
  // 3. Multi-group or token matching (e.g. "Agape / EBF / Sisters / Vessel", "West Region (Cypress / Buena Park)")
  var tokens = cleanStr.split(/[\/\n,()&]+/).map(function(t) { return t.trim().toLowerCase(); }).filter(Boolean);
  var matchedWorkers = [];
  var workerSet = {};
  
  var aliasMap = {
    'ebf': 'english bible fellowship',
    'sisters': 'sisters\' fellowship',
    'vessel': 'vessel fellowship'
  };
  
  tokens.forEach(function(token) {
    if (token === 'region' || token === 'west' || token === 'north' || token === 'south' || token === 'east') return;
    var expanded = aliasMap[token] || token;
    
    for (var k in teamsMap) {
      var kl = k.toLowerCase();
      if (kl === expanded || kl.indexOf(expanded) !== -1 || (expanded.length >= 4 && kl.indexOf(token) !== -1)) {
        var members = teamsMap[k] || [];
        for (var m = 0; m < members.length; m++) {
          var name = members[m];
          if (!workerSet[name]) {
            workerSet[name] = true;
            matchedWorkers.push(name);
          }
        }
      }
    }
  });
  
  return matchedWorkers;
}

// -------------------------------------------------------------------------
// Cache helpers: chunked storage (CacheService caps each value at 100KB)
// -------------------------------------------------------------------------
function cachePutLarge(cache, key, value, ttl) {
  try {
    var json = JSON.stringify(value);
    var size = 30000, parts = {}, n = 0;
    for (var i = 0; i < json.length; ) {
      var end = Math.min(i + size, json.length);
      if (end < json.length) {
        var cc = json.charCodeAt(end - 1);
        if (cc >= 0xD800 && cc <= 0xDBFF) end--; // never split a surrogate pair
      }
      parts[key + '_' + n] = json.substring(i, end);
      n++;
      i = end;
    }
    var prev = cache.get(key + '_n');
    parts[key + '_n'] = n + ':' + json.length;
    cache.putAll(parts, ttl);
    if (prev) {
      var pc = parseInt(String(prev).split(':')[0], 10);
      if (pc > n) {
        var stale = [];
        for (var k = n; k < pc; k++) stale.push(key + '_' + k);
        cache.removeAll(stale);
      }
    }
  } catch (e) {
    console.warn('cachePutLarge failed for "' + key + '": ' + e);
  }
}

// Returns the joined JSON string (or null on any miss / length mismatch).
function cacheGetLargeJson(cache, key) {
  try {
    var meta = cache.get(key + '_n');
    if (!meta) return null;
    var p = String(meta).split(':');
    var n = parseInt(p[0], 10), len = parseInt(p[1], 10);
    if (!(n > 0)) return null;
    var keys = [];
    for (var i = 0; i < n; i++) keys.push(key + '_' + i);
    var got = cache.getAll(keys);
    var out = '';
    for (var j = 0; j < n; j++) {
      var s = got[keys[j]];
      if (s === null || s === undefined) return null;
      out += s;
    }
    return out.length === len ? out : null;
  } catch (e) {
    return null;
  }
}

function cacheRemoveLarge(cache, key) {
  try {
    var keys = [key, key + '_n'];
    var meta = cache.get(key + '_n');
    if (meta) {
      var n = parseInt(String(meta).split(':')[0], 10);
      for (var i = 0; i < n; i++) keys.push(key + '_' + i);
    }
    cache.removeAll(keys);
  } catch (e) {
    console.warn('cacheRemoveLarge failed for "' + key + '": ' + e);
  }
}

// -------------------------------------------------------------------------
// Quarter data versioning (shared across all coordinators via script cache)
// -------------------------------------------------------------------------
function getQuarterVersion(sheetName) {
  try {
    return CacheService.getScriptCache().get('q_ver_' + sheetName) || '0';
  } catch (e) { return '0'; }
}

function bumpQuarterVersion(sheetName) {
  var prev = getQuarterVersion(sheetName);
  var next = new Date().getTime() + '-' + Math.floor(Math.random() * 1000000);
  try { CacheService.getScriptCache().put('q_ver_' + sheetName, next, 21600); } catch (e) {}
  return { prev: prev, next: next };
}

function bumpAllQuarterVersions() {
  for (var i = 0; i < CONFIG.QUARTERS.length; i++) bumpQuarterVersion(CONFIG.QUARTERS[i]);
}

// -------------------------------------------------------------------------
// SWR Speaker Sync & Trigger Setup (Using PropertiesService)
// -------------------------------------------------------------------------

function setupDailyTrigger() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.prompt(
    'Setup Daily SWR Sync',
    'Please paste the full URL of the SWR Speaker Google Sheet:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Setup cancelled.');
    return;
  }
  
  var url = response.getResponseText().trim();
  
  try {
    var sourceSs = SpreadsheetApp.openByUrl(url);
    var missingQuarters = [];
    CONFIG.SWR_SOURCE_QUARTERS.forEach(function(q) {
       if (!sourceSs.getSheetByName(q)) {
          missingQuarters.push(q);
       }
    });
    
    if (missingQuarters.length > 0) {
      ui.alert('Error', 'The spreadsheet was found, but it is missing the following tabs: ' + missingQuarters.join(', ') + '. Please check the sheet and try again.', ui.ButtonSet.OK);
      return;
    }
  } catch (e) {
    ui.alert('Invalid URL', 'Could not open the spreadsheet. Please make sure the URL is correct and that you have permission to access it.\n\nError: ' + e.message, ui.ButtonSet.OK);
    return;
  }
  
  PropertiesService.getScriptProperties().setProperty('SWR_SPEAKER_SHEET_URL', url);
  
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncSWRSpeakers') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger('syncSWRSpeakers')
    .timeBased()
    .everyDays(1)
    .atHour(1) 
    .create();
    
  ui.alert('Success!', 'The URL has been validated and saved. The daily sync is now scheduled to run automatically every night at 1 AM.', ui.ButtonSet.OK);
}

// Note: syncSWRSpeakers() is defined in SyncSWR.js

// -------------------------------------------------------------------------
// UI Menus & Assistant Functions
// -------------------------------------------------------------------------

function colNumToLetter(n) {
  var r = '';
  while (n > 0) { var rem=(n-1)%26; r=String.fromCharCode(65+rem)+r; n=Math.floor((n-1)/26); }
  return r;
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. Main Schedule Assistant Menu
  ui.createMenu('📅 Schedule Assistant')
    .addItem('Open Assignment Sidebar', 'showSidebar')
    .addSeparator()
    .addItem('Generate Monthly Schedule', 'promptMonthlyGenerator')
    .addItem('Sync RE Schedules (Current Quarter)', 'syncRESchedules')
    .addItem('🔄 Auto-Sync Active & Next Month', 'promptAutoSyncScheduledMonths')
    .addSeparator()
    .addSubMenu(ui.createMenu('🎪 ESSC Schedule Generator')
      .addItem('🚀 Generate Spring ESSC Schedule', 'generateSpringESSC')
      .addItem('🚀 Generate Fall ESSC Schedule', 'generateFallESSC')
      .addSeparator()
      .addItem('🔄 Push Dropdowns & Resolve Placeholders', 'applyValidationForESSC')
      .addItem('✨ Inject Dynamic Date & Speaker Formulas', 'injectDynamicGridFormulasMenu')
      .addItem('✨ Resolve Speaker Placeholders in Grid', 'resolveGridSpeakerPlaceholdersMenu')
      .addSeparator()
      .addItem('📅 Set ESSC Event Dates...', 'promptSetESSCDates')
      .addItem('🎨 Set ESSC Event Themes...', 'promptSetESSCThemes')
      .addItem('🔗 Connect Standalone ESSC Workbook...', 'connectStandaloneESSCWorkbookPrompt')
      .addSeparator()
      .addItem('🔄 Sync ESSC Unavailability Rosters', 'promptSyncESSCUnavailabilityRosters')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('📧 Email')
      .addItem('Export & Email Next Month PDF', 'promptExportAndEmailNextMonthPdf')
      .addItem('Preview Email (send to me)', 'promptPreviewScheduleEmail')
      .addItem('Send Monthly Unavailability Reminder', 'promptSendMonthlyUnavailabilityReminder')
      .addItem('Check Missing Slots for Next Month', 'auditUpcomingMonthMissingSlots')
      .addSeparator()
      .addItem('Edit Email Template', 'promptEditEmailTemplate')
      .addItem('Set PDF Distribution List', 'promptSetPdfDistributionList')
    )
    .addToUi();
    
  // 2. Admin Tools Menu (one entry per function; ESSC items live only under Schedule Assistant)
  ui.createMenu("⚙️ Admin Tools")
    .addItem("🚀 Open Setup Wizard", "showSetupWizard")
    .addItem("🌱 Seed New Church", "promptSeedNewChurch")
    .addItem("📊 Generate Annual Summary", "generateAnnualSummary")
    .addSeparator()
    .addSubMenu(ui.createMenu('🗂️ Tabs')
      .addItem("Organize & Clean Up Tabs", "organizeAndCleanUpTabs")
      .addItem("Toggle Admin Config & Log Tabs", "toggleAdminConfigTabs")
    )
    .addSubMenu(ui.createMenu('🔗 Connections')
      .addItem("Link Unavailability Sheet", "promptLinkUnavailabilitySheet")
      .addItem("Link Sermon Topics & Hymns Sheet", "promptLinkSermonTopicsSheet")
    )
    .addSubMenu(ui.createMenu('🔄 Sync Data')
      .addItem("⚡ Apply All Worker Roster Changes (1-Click)", "applyWorkerRosterChanges")
      .addSeparator()
      .addItem("Compile Role Configurations", "promptCompileRoleConfig")
      .addItem("Sync Coordinators to Q1-Q4", "promptSyncCoordinatorsToQuarters")
      .addSeparator()
      .addItem("Sync SWR Speakers", "syncSWRSpeakers")
      .addItem("Sync Unavailability Matrix", "syncUnavailabilityFromMatrix")
      .addItem("Sync Workers to Unavailability Matrix", "promptSyncWorkersToUnavailability")
    )
    .addSubMenu(ui.createMenu('⏰ Automation')
      .addItem("Activate All Automation Triggers", "activateAllTriggers")
      .addItem("Setup Daily SWR Sync URL", "setupDailyTrigger")
    )
    .addSubMenu(ui.createMenu('📅 Annual Rollover')
      .addItem("🚀 Provision New Year (Full Rollover & Year Folder)", "promptProvisionNewYear")
      .addItem("Reset Current Holy Work Schedule for New Year", "promptResetForNewYear")
      .addItem("Reset Sermon Topics & Hymns Sheet for New Year", "promptResetSermonTopicsSheet")
      .addItem("Provision Sermon Topics & Hymns for Year", "promptProvisionSermonTopicsOnly")
      .addItem("Update Sermon Topics Speaker Links", "promptUpdateSermonTopicsSpeakerFormulas")
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('🧰 Repair & Maintenance (Advanced)')
      .addItem('Re-apply Data Validation Dropdowns', 'applyAllValidation')
      .addItem('Remove Yellow Cell Highlights (Fix Lag)', 'removeDuplicateWarningFormatting')
      .addItem('Restructure Q1-Q4 (Universal 14-Row)', 'promptRestructureQuarterSheetsToUniversal14')
      .addItem('Remove Column(s) (Interactive)', 'promptRemoveCustomColumns')
      .addItem('Repair 2027 Live Spreadsheets (Formulas & Timezone)', 'repair2027LiveSpreadsheets')
    )
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('Schedule Assistant')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function removeDuplicateWarningFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quarters = CONFIG.QUARTERS;
  var removed = false;
  
  quarters.forEach(function(q) {
    var sheet = ss.getSheetByName(q);
    if (!sheet) return;
    
    var rules = sheet.getConditionalFormatRules();
    var newRules = rules.filter(function(rule) {
       var criteria = rule.getBooleanCondition();
       if (criteria && criteria.getCriteriaType() == SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) {
          var args = criteria.getCriteriaValues();
          if (args && args.length > 0 && String(args[0]).indexOf('COUNTIF') !== -1) {
             removed = true;
             return false;
          }
       }
       return true;
    });
    sheet.setConditionalFormatRules(newRules);
  });
  
  if (removed) {
     SpreadsheetApp.getUi().alert('Success!', 'The heavy Conditional Formatting rule has been removed from all Quarters. The 5-10 second assignment lag is fixed!', SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
     SpreadsheetApp.getUi().alert('Already Clean!', 'The lag-inducing Conditional Formatting rule was not found.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function onEdit(e) {
  if (!e || !e.range) return;
  var s = e.range.getSheet().getName();
  
  // 1. Sidebar Cache Refresh: if a quarter sheet was edited, flag sidebar to refresh
  if (CONFIG.QUARTERS.indexOf(s) !== -1) {
     bumpQuarterVersion(s);
      return;
  }
  
  // 1b. Clear worker availability cache if Workers tab was edited
  if (s === CONFIG.WORKERS_SHEET || s === CONFIG.TEAMS_SHEET || s === CONFIG.UNAVAILABILITY_SHEET) {
      try {
        var sc = CacheService.getScriptCache();
        if (s === CONFIG.WORKERS_SHEET) {
          cacheRemoveLarge(sc, 'workers_day_avail');
          sc.remove('workers_data');
        } else if (s === CONFIG.TEAMS_SHEET) {
          cacheRemoveLarge(sc, 'teams_and_workers');
        } else {
          cacheRemoveLarge(sc, 'unavailability_data');
        }
        bumpAllQuarterVersions();
      } catch(err) {}
   }
   
  // 1c. Clear church configuration cache if configuration tabs were edited
  if (s === 'ChurchProfile' || s === 'ServiceLayoutConfig' || s === 'RoleConfig') {
      try {
        CacheService.getScriptCache().remove('church_config');
      } catch(err) {}
  }
  
  // 2. Audit Log: Track manual overrides made directly on Generated Monthly Schedules
  var ignoredSheets = ["MonthlyTemplate", "MappingConfig", "MonthlyGeneratorLog", "MonthlyOverridesLog", CONFIG.TEAMS_SHEET, CONFIG.WORKERS_SHEET, CONFIG.UNAVAILABILITY_SHEET, CONFIG.SYNC_LOG_SHEET].concat(CONFIG.QUARTERS);
  if (ignoredSheets.indexOf(s) !== -1) return;
  
  var monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var isScheduleSheet = false;
  for (var i = 0; i < monthNames.length; i++) {
    if (s.indexOf(monthNames[i]) === 0) { isScheduleSheet = true; break; }
  }
  if (!isScheduleSheet) return;
  
  var oldValue = e.oldValue === undefined ? "(empty)" : e.oldValue;
  var newValue = e.value === undefined ? "(empty)" : e.value;
  if (oldValue === newValue) return;

  var ss = e.source;
  var tz = ss.getSpreadsheetTimeZone();
  var time = Utilities.formatDate(new Date(), tz, "MM/dd/yyyy hh:mm:ss a");
  var cellRef = e.range.getA1Notation();
  var userEmail = (e.user && e.user.getEmail()) ? e.user.getEmail() : "";
  if (!userEmail) {
    try { userEmail = Session.getActiveUser().getEmail(); } catch (uErr) {}
  }
  if (!userEmail) userEmail = "Authenticated Editor";
  
  var editLock = LockService.getScriptLock();
  var hasEditLock = false;
  try {
    hasEditLock = editLock.tryLock(2000);
    var logSheet = ss.getSheetByName("MonthlyOverridesLog");
    if (!logSheet) {
      logSheet = ss.insertSheet("MonthlyOverridesLog");
      var headers = ["Timestamp", "User Email", "Schedule Month", "Cell Edited", "Old Value", "New Value"];
      logSheet.appendRow(headers);
      logSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#fce5cd");
      logSheet.setFrozenRows(1);
      logSheet.autoResizeColumns(1, headers.length);
      logSheet.hideSheet();
    }
    
    logSheet.appendRow([time, userEmail, s, cellRef, oldValue, newValue]);
    
    // Keep the log bounded: trim oldest rows from top in batches
    var LOG_MAX_ROWS = 5000, LOG_TRIM_SLACK = 500;
    var lastLogRow = logSheet.getLastRow();
    if (lastLogRow > LOG_MAX_ROWS + 1 + LOG_TRIM_SLACK) {
      logSheet.deleteRows(2, lastLogRow - (LOG_MAX_ROWS + 1));
    }
  } catch (logErr) {
    console.warn("Could not write to MonthlyOverridesLog: " + logErr.message);
  } finally {
    if (hasEditLock) editLock.releaseLock();
  }
}

function onSelectionChange(e) {
  if (!e || !e.range) return;
  try {
    var sheetName = e.range.getSheet().getName();
    var isESSC = (sheetName === 'Spring_ESSC_Grid' || sheetName === 'Fall_ESSC_Grid');
    if (CONFIG.QUARTERS.indexOf(sheetName) === -1 && !isESSC) return;
    var r = e.range.getRow();
    var c = e.range.getColumn();
    var esscMaxRow = (typeof ESSC_CONFIG !== 'undefined' && ESSC_CONFIG.GRID && ESSC_CONFIG.GRID.LAST_ROW) ? ESSC_CONFIG.GRID.LAST_ROW : 45;
    if (isESSC) {
      if (r < 2 || r > esscMaxRow || c < 5 || c > 16) return;
    } else {
      if (r < CONFIG.START_ROW || c < 2 || c > CONFIG.MAX_COLUMNS) return;
    }

    var d = { r: r, c: c, s: sheetName };
    CacheService.getUserCache().put('active_cell', JSON.stringify(d), CONFIG.CACHE_POLL_TIME);
  } catch(err) {}
}

function pollContext(loadedSheet, clientVersions) {
  var t0 = new Date().getTime();
  try {
    var cache = CacheService.getUserCache();
    var cached = cache.get('active_cell');
    if (!cached) {
       return { success: true, waitingForClick: true };
    }
    
    var d = JSON.parse(cached);
    var r = d.r, c = d.c, s = d.s;
    
    var isESSC = (s === 'Spring_ESSC_Grid' || s === 'Fall_ESSC_Grid' || s === 'ESSC_Data');
    var isSched = false;
    var esscMaxRow = (typeof ESSC_CONFIG !== 'undefined' && ESSC_CONFIG.GRID && ESSC_CONFIG.GRID.LAST_ROW) ? ESSC_CONFIG.GRID.LAST_ROW : 45;
    if (isESSC) {
      isSched = (r >= 2 && r <= esscMaxRow && c >= 5 && c <= 16);
    } else {
      isSched = (CONFIG.QUARTERS.indexOf(s) !== -1) && r >= 2 && c <= CONFIG.MAX_COLUMNS;
    }
    var result = { success: true, sheetName: s, row: r, col: c, isSchedulerCell: isSched, isAdmin: isAdminUser() };
    
    if (isSched) {
       var ver = getQuarterVersion(s);
       result.version = ver;
       var needsChunk = (clientVersions === null || clientVersions === undefined)
         ? (s !== loadedSheet)
         : (clientVersions[s] !== ver);
       if (needsChunk) {
          var ss2 = SpreadsheetApp.getActiveSpreadsheet();
          var targetSheet = ss2.getSheetByName(s);
          if (targetSheet) {
             var chunk = getSheetChunkData(targetSheet, s, false);
             if (chunk && chunk.success) chunk.version = ver;
             result.chunkData = chunk;
          }
       }
    }
    
    result.serverTime = new Date().getTime() - t0;
    return result;
  } catch(e) {
    return { success: false, error: 'pollContext: ' + e.toString() };
  }
}

function isAdminUser() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var config = getChurchConfig();
    var editors = [];
    var prof = config ? (config.profile || config.churchProfile) : null;
    if (prof && prof.authorized_editors) {
      editors = prof.authorized_editors.split(',').map(function(e) { return e.trim().toLowerCase(); }).filter(function(e) { return e; });
    } else if (typeof APP_CONFIG !== 'undefined' && Array.isArray(APP_CONFIG.AUTHORIZED_EDITORS)) {
      editors = APP_CONFIG.AUTHORIZED_EDITORS.map(function(e) { return String(e).trim().toLowerCase(); }).filter(Boolean);
    }
    var owner = ss ? ss.getOwner() : null;
    var ownerEmail = (owner && typeof owner.getEmail === 'function' && owner.getEmail()) ? owner.getEmail().toLowerCase() : '';
    var currentUser = (Session.getActiveUser() && Session.getActiveUser().getEmail()) ? Session.getActiveUser().getEmail().toLowerCase() : '';
    if (!currentUser) return false;
    return editors.indexOf(currentUser) !== -1 || (ownerEmail !== '' && currentUser === ownerEmail);
  } catch (e) {
    console.warn('isAdminUser fallback error: ' + e.message);
    return false;
  }
}

function getNeighborQuarters(sName) {
  var i = CONFIG.QUARTERS.indexOf(sName), out = [];
  if (i === -1) return out;
  if (i > 0) out.push(CONFIG.QUARTERS[i - 1]);
  if (i < CONFIG.QUARTERS.length - 1) out.push(CONFIG.QUARTERS[i + 1]);
  return out;
}

// Background pre-load of an adjacent quarter (uses cached reference data, never forces a refresh)
function prefetchChunk(sheetName) {
  try {
    if (CONFIG.QUARTERS.indexOf(sheetName) === -1) return { error: 'Not a quarter sheet' };
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return { error: 'Sheet not found' };
    var ver = getQuarterVersion(sheetName);
    var chunk = getSheetChunkData(sheet, sheetName, false);
    if (chunk && chunk.success) chunk.version = ver;
    return { chunkData: chunk };
  } catch(e) {
    return { error: e.toString() };
  }
}

function manualFetchChunk(sheetName, hardRefresh) {
  try {
    var hard = (hardRefresh === undefined) ? true : !!hardRefresh;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { error: 'Sheet not found' };
    var ver = getQuarterVersion(sheetName);
    var chunk = getSheetChunkData(sheet, sheetName, hard);
    if (chunk && chunk.success) chunk.version = ver;
    return { chunkData: chunk };
  } catch(e) {
    return { error: e.toString() };
  }
}

function assignWorkerExact(sheetName, row, col, workerName) {
  var t0 = new Date().getTime();
  var lock = LockService.getScriptLock();
  var locked = false;
  
  try {
    if (!lock.tryLock(5000)) {
      return { success: false, error: 'Could not acquire lock. Another assignment is in progress. Please try again.' };
    }
    locked = true;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var t1 = new Date().getTime();
    
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sheet "' + sheetName + '" not found.' };
    var t2 = new Date().getTime();
    
    var cell = sheet.getRange(row, col);
    var t3 = new Date().getTime();
    
    // Check and enforce Data Validation dropdown rules if configured on cell (unless allowInvalid is true)
    var rule = cell.getDataValidation();
    if (rule && !rule.getAllowInvalid() && workerName !== "" && workerName.toUpperCase() !== "X") {
      var criteriaType = rule.getCriteriaType();
      var args = rule.getCriteriaValues();
      if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
        var srcRange = args[0];
        if (srcRange) {
          var validVals = srcRange.getValues().flat().map(function(v) { return String(v).trim().toLowerCase(); });
          if (validVals.length > 0 && validVals.indexOf(workerName.trim().toLowerCase()) === -1) {
            return { 
              success: false, 
              error: 'Validation Error: "' + workerName + '" is not an authorized option in the dropdown list for cell ' + cell.getA1Notation() + '.' 
            };
          }
        }
      } else if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        var list = args[0];
        if (list && list.map(function(v) { return String(v).trim().toLowerCase(); }).indexOf(workerName.trim().toLowerCase()) === -1) {
          return { 
            success: false, 
            error: 'Validation Error: "' + workerName + '" is not in the allowed list for cell ' + cell.getA1Notation() + '.' 
          };
        }
      }
    }
    
    cell.setValue(workerName);
    SpreadsheetApp.flush();
    var t4 = new Date().getTime();
    
    var d = { r: row, c: col, s: sheetName };
    CacheService.getUserCache().put('active_cell', JSON.stringify(d), CONFIG.CACHE_POLL_TIME);
    var t5 = new Date().getTime();
    var qv = bumpQuarterVersion(sheetName);
    
    return { 
      success: true, 
      prevVersion: qv.prev, version: qv.next,
      times: { total: t5 - t0, ss: t1 - t0, sh: t2 - t1, rng: t3 - t2, set: t4 - t3, ca: t5 - t4 }
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  } finally {
    if (locked) lock.releaseLock();
  }
}

function getTeamsAndWorkers(forceRefresh) {
  var cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    var cached = cacheGetLargeJson(cache, 'teams_and_workers');
    if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  }
  var masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TEAMS_SHEET);
  var isRosters = false;
  if (!masterSheet) {
    masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rosters');
    if (masterSheet) isRosters = true;
  }
  if (!masterSheet) {
    try {
      if (typeof getMasterHWSSpreadsheet === 'function') {
        var masterSs = getMasterHWSSpreadsheet();
        if (masterSs) masterSheet = masterSs.getSheetByName(CONFIG.TEAMS_SHEET);
      }
    } catch(e) {}
  }
  if (!masterSheet) return { error: 'Teams or Rosters sheet not found', teams: {} };
  
  var maxR = Math.min(masterSheet.getMaxRows(), CONFIG.MAX_TEAM_ROWS);
  var maxC = Math.min(masterSheet.getMaxColumns(), CONFIG.MAX_TEAM_COLS);
  var data = masterSheet.getRange(1, 1, maxR, maxC).getValues();
  var teams = {};
  var headers = data[0].map(function(h){ return String(h).trim().toLowerCase(); });
  
  if (isRosters) {
    for (var col = 0; col < maxC; col++) {
      var tn = '', sRow = 1;
      var r0Val = String(data[0][col] !== undefined && data[0][col] !== null ? data[0][col] : '').trim();
      var r1Val = (data.length > 1) ? String(data[1][col] !== undefined && data[1][col] !== null ? data[1][col] : '').trim() : '';
      if (r1Val && (r0Val === '' || !isNaN(Number(r0Val)))) {
        tn = r1Val;
        sRow = 2;
      } else {
        tn = r0Val;
        sRow = 1;
      }
      if (!tn) continue;
      teams[tn] = [];
      for (var row = sRow; row < maxR; row++) {
        var w = String(data[row][col] || '').trim();
        if (w && teams[tn].indexOf(w) === -1) teams[tn].push(w);
      }
    }
  } else {
    var tCands = [
      ['team name','team_name','teamname','group name','group_name','groupname','role name','role_name','rolename','team name/role','team/role','role/team'],
      ['team','group','role','service'],
      ['team id','team_id','teamid','team no','team no.','team number','id']
    ];
    var nCands = [['worker name','worker_name','workername','member name','member_name','membername','full name','fullname','name','worker','member','person','schedule name']];
    
    var tIdx = -1, nIdx = -1;
    for (var p = 0; p < tCands.length && tIdx === -1; p++)
      for (var i = 0; i < headers.length; i++)
        if (tCands[p].indexOf(headers[i]) !== -1) { tIdx = i; break; }
        
    for (var p = 0; p < nCands.length && nIdx === -1; p++)
      for (var i = 0; i < headers.length; i++)
        if (nCands[p].indexOf(headers[i]) !== -1) { nIdx = i; break; }
        
    if (nIdx !== -1 && tIdx !== -1) {
      for (var r = 1; r < data.length; r++) {
        var n = String(data[r][nIdx]).trim(), tn = String(data[r][tIdx]).trim();
        if (n && tn) {
          if (!teams[tn]) teams[tn] = [];
          if (teams[tn].indexOf(n) === -1) teams[tn].push(n);
        }
      }
    } else {
      for (var col = 0; col < maxC; col++) {
        var tn = '', sRow = 1;
        if (String(data[1][col]).trim()) { tn = String(data[1][col]).trim(); sRow = 2; } 
        else { tn = String(data[0][col]).trim(); sRow = 1; }
        if (!tn) continue;
        teams[tn] = [];
        for (var row = sRow; row < maxR; row++) {
          var w = String(data[row][col]).trim();
          if (w) teams[tn].push(w);
        }
      }
    }
  }
  var result = { success: true, teams: teams };
  cachePutLarge(cache, 'teams_and_workers', result, CONFIG.CACHE_DATA_TIME);
  return result;
}

function getUnavailabilityData(forceRefresh) {
  var cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    var cached = cacheGetLargeJson(cache, 'unavailability_data');
    if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  }
  
  var tz = cache.get('sheet_tz');
  if (!tz) {
     tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
     try { cache.put('sheet_tz', tz, 21600); } catch(e) {}
  }
  
  var activeSs = SpreadsheetApp.getActiveSpreadsheet();
  var uvSheet = activeSs ? activeSs.getSheetByName(CONFIG.UNAVAILABILITY_SHEET) : null;
  if (!uvSheet && typeof getMasterHWSSpreadsheet === 'function') {
    try {
      var masterSs = getMasterHWSSpreadsheet();
      if (masterSs) uvSheet = masterSs.getSheetByName(CONFIG.UNAVAILABILITY_SHEET);
    } catch(e) {
      console.warn('Could not load Unavailability from master HWS sheet: ' + e.message);
    }
  }
  var map = {};
  if (uvSheet) {
    var maxUvR = Math.min(uvSheet.getMaxRows(), CONFIG.MAX_UNAVAIL_ROWS);
    var maxUvC = Math.min(uvSheet.getMaxColumns(), 3);
    if (maxUvR > 0 && maxUvC > 0) {
      var rows = uvSheet.getRange(1, 1, maxUvR, maxUvC).getValues();
      for (var r = 1; r < maxUvR; r++) {
        if (!rows[r]) continue;
        var dv = rows[r][0], nm = String(rows[r][1] || '').trim();
        if (!dv && !nm) continue;
        var rs = rows[r][2] ? String(rows[r][2]).trim() : 'Marked Unavailable';
        if (dv && nm) {
          var dateObj = dv instanceof Date ? dv : new Date(dv);
          if (!isNaN(dateObj.getTime())) {
            var ds = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
            if (!map[ds]) map[ds] = [];
            map[ds].push({ name: nm, reason: rs });
          }
        }
      }
    }
  }
  cachePutLarge(cache, 'unavailability_data', map, CONFIG.CACHE_DATA_TIME);
  return map;
}

/**
 * Detects whether headers are on Row 1 (data[0]) or Row 2 (data[1]) of Workers sheet.
 * Returns { headerRowIdx: 0 or 1, dataStartIdx: 1 or 2, headers: [...] }
 * @param {Array<Array<any>>} data
 * @returns {Object}
 */
function detectWorkersHeaderRow_(data) {
  if (!data || data.length === 0) {
    return { headerRowIdx: 0, dataStartIdx: 1, headers: [] };
  }
  var keywords = ['first name', 'schedule name', 'email', 'active', 'worker id', 'teams', 'status', 'chinese name', 'phone', 'available monday', 'resident church'];
  
  function scoreRow(row) {
    if (!row || !Array.isArray(row)) return 0;
    var count = 0;
    for (var i = 0; i < row.length; i++) {
      var str = String(row[i] || '').trim().toLowerCase();
      if (!str) continue;
      for (var k = 0; k < keywords.length; k++) {
        if (str.indexOf(keywords[k]) !== -1) {
          count++;
          break;
        }
      }
    }
    return count;
  }
  
  var score0 = scoreRow(data[0]);
  var score1 = (data.length > 1) ? scoreRow(data[1]) : 0;
  
  var headerRowIdx = (score1 > score0) ? 1 : 0;
  var dataStartIdx = headerRowIdx + 1;
  var headers = (data[headerRowIdx] || []).map(function(h) { return String(h || '').trim(); });
  
  return {
    headerRowIdx: headerRowIdx,
    dataStartIdx: dataStartIdx,
    headers: headers
  };
}

function getWorkersAvailability(forceRefresh) {
  var cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    var cached = cacheGetLargeJson(cache, 'workers_day_avail');
    if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  }
  var workerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.WORKERS_SHEET);
  var result = {};
  if (workerSheet) {
    var maxWR = Math.min(workerSheet.getMaxRows(), CONFIG.MAX_WORKER_ROWS);
    var maxWC = Math.min(workerSheet.getMaxColumns(), CONFIG.MAX_WORKER_COLS);
    var data = (maxWR > 0 && maxWC > 0) ? workerSheet.getRange(1, 1, maxWR, maxWC).getValues() : [];
    if (data && data.length > 0) {
      var detected = detectWorkersHeaderRow_(data);
      var headers = detected.headers.map(function(h) { return String(h).trim().toLowerCase(); });
      var nIdx = -1, mIdx = -1, fIdx = -1;
      for (var i = 0; i < headers.length; i++) {
        if (nIdx === -1 && headers[i].indexOf('schedule name') !== -1) nIdx = i;
        if (headers[i].indexOf('monday') !== -1) mIdx = i;
        if (headers[i].indexOf('friday') !== -1) fIdx = i;
      }
      if (nIdx === -1) nIdx = 4;
      if (mIdx === -1) mIdx = 2;
      if (fIdx === -1) fIdx = 3;
      
      for (var r = detected.dataStartIdx; r < Math.min(data.length, CONFIG.MAX_WORKER_ROWS); r++) {
        if (!data[r] || typeof data[r][nIdx] === 'undefined') continue;
        var n = String(data[r][nIdx]).trim();
        if (n) {
           var availM = true, availF = true;
           var vM = data[r][mIdx], vF = data[r][fIdx];
           if (vM === false || String(vM).trim().toLowerCase() === 'false') availM = false;
           if (vF === false || String(vF).trim().toLowerCase() === 'false') availF = false;
           result[n] = { monday: availM, friday: availF };
        }
      }
    }
  }
  cachePutLarge(cache, 'workers_day_avail', result, CONFIG.CACHE_DATA_TIME);
  return result;
}

function resolveTeamMapping(rawName, teamNames) {
  if (!rawName) return '';
  var clean = rawName.trim().toLowerCase();
  for (var i = 0; i < teamNames.length; i++) {
    if (teamNames[i].toLowerCase() === clean) return teamNames[i];
  }
  for (var i = 0; i < teamNames.length; i++) {
    var t = teamNames[i].toLowerCase();
    if (t.indexOf(clean) !== -1 || clean.indexOf(t) !== -1) return teamNames[i];
  }
  return '';
}

function getSheetChunkData(sheet, sName, forceRefresh) {
  try {
    var tr = getTeamsAndWorkers(!!forceRefresh);
    if (tr.error) return { success: false, error: tr.error };
    var teamsMap = tr.teams;
    var teamNames = Object.keys(teamsMap).sort();

    var cache = CacheService.getScriptCache();
    var tz = cache.get('sheet_tz');
    if (!tz) {
       tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
       try { cache.put('sheet_tz', tz, 21600); } catch(e) {}
    }
    
    var isESSC = (sName === 'Spring_ESSC_Grid' || sName === 'Fall_ESSC_Grid' || sName === 'ESSC_Data');
    if (isESSC) {
      var numRows = Math.max(sheet.getLastRow(), 45);
      var allData = sheet.getRange(1, 1, numRows, 16).getValues();
      var teamMap = {}, teamRawMap = {}, roleMap = {};
      var row1 = allData[0];
      
      for (var c = 5; c <= 16; c++) {
        var idx = c - 1;
        var role = String(row1[idx] || '').trim();
        if (role) roleMap[c] = cleanLabel(role) || role;
        var resolved = resolveTeamMapping(role, teamNames);
        teamMap[c] = resolved || role;
        teamRawMap[c] = role;
      }
      
      var rowDates = {}, rowAssignments = {};
      for (var r = 1; r < allData.length; r++) {
        var actualRow = r + 1;
        var rowArr = allData[r];
        var dVal = rowArr[0];
        var dayVal = String(rowArr[1] || '').trim();
        var timeVal = String(rowArr[2] || '').trim();
        var sessionVal = String(rowArr[3] || '').trim();
        
        var dateStr = '', tds = '';
        if (dVal) {
          var dateObj = dVal instanceof Date ? dVal : new Date(dVal);
          if (!isNaN(dateObj.getTime())) {
            tds = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
            dateStr = tds + ' (' + dayVal + ' ' + timeVal + ' - ' + sessionVal + ')';
          } else {
            tds = String(dVal).substring(0, 10);
            dateStr = String(dVal) + ' (' + dayVal + ' ' + timeVal + ' - ' + sessionVal + ')';
          }
        } else {
          dateStr = dayVal + ' ' + timeVal + ' - ' + sessionVal;
        }
        rowDates[actualRow] = { dateStr: dateStr, tds: tds };
        
        var awm = {};
        for (var c = 4; c < 16; c++) {
          var val = String(rowArr[c] || '').trim();
          if (val && val !== 'X') {
            var roleName = roleMap[c + 1] || ('Column ' + colNumToLetter(c + 1));
            if (!awm[val]) awm[val] = [];
            awm[val].push({ col: c + 1, role: roleName });
          }
        }
        rowAssignments[actualRow] = awm;
      }
      
      // --- C. Choir Presentation & Buffer Exclusions ---
      // 1. Resiliently find choir teams from teamsMap
      var adultChoirKey = '';
      var youthChoirKey = '';
      for (var tk in teamsMap) {
        var tkLower = tk.toLowerCase();
        if (!adultChoirKey && tkLower.indexOf('adult choir') !== -1) adultChoirKey = tk;
        if (!youthChoirKey && tkLower.indexOf('youth choir') !== -1) youthChoirKey = tk;
      }
      var adultChoirMembers = adultChoirKey ? (teamsMap[adultChoirKey] || []) : [];
      var youthChoirMembers = youthChoirKey ? (teamsMap[youthChoirKey] || []) : [];

      // 2. Scan rows for Choir assignments in Column P (Col index 15)
      for (var r = 1; r < allData.length; r++) {
        var actualRow = r + 1;
        var choirVal = String(allData[r][15] || '').trim();
        if (!choirVal || choirVal === 'X') continue;
        
        // Resolve placeholder tokens like {SATURDAY_CHOIR} if present
        if (choirVal.indexOf('{') !== -1) {
          var seasonName = (sName.toLowerCase().indexOf('spring') !== -1) ? 'Spring' : 'Fall';
          if (typeof resolveSpeakerPlaceholders === 'function') {
            choirVal = resolveSpeakerPlaceholders(sheet.getParent(), seasonName, choirVal);
          }
        }
        
        var choirValLower = choirVal.toLowerCase();
        // Ignore Children's Choir (do not flag adult worker conflicts)
        if (choirValLower.indexOf('children') !== -1 || choirVal.indexOf('兒童') !== -1) {
          continue;
        }
        
        var isAdultChoir = (choirValLower.indexOf('adult') !== -1 || choirVal.indexOf('成人') !== -1);
        var isYouthChoir = (choirValLower.indexOf('youth') !== -1 || choirVal.indexOf('青年') !== -1);
        
        var presentingTeams = [];
        if (isAdultChoir && adultChoirMembers.length > 0) {
          presentingTeams.push({ name: 'Adult Choir', members: adultChoirMembers });
        }
        if (isYouthChoir && youthChoirMembers.length > 0) {
          presentingTeams.push({ name: 'Youth Choir', members: youthChoirMembers });
        }
        
        if (presentingTeams.length === 0) continue;
        
        var dayValLower = cleanLabel(String(allData[r][1] || '')).toLowerCase();
        var sessionValLower = cleanLabel(String(allData[r][3] || '')).toLowerCase();
        
        var isSatEvangelical = (dayValLower.indexOf('sat') !== -1 || dayValLower.indexOf('安息') !== -1) &&
                               (sessionValLower.indexOf('evangelical') !== -1 || sessionValLower.indexOf('佈道') !== -1);
        var isSunCommunionOrChoir = (dayValLower.indexOf('sun') !== -1 || dayValLower.indexOf('日') !== -1) &&
                                    (sessionValLower.indexOf('communion') !== -1 || sessionValLower.indexOf('聖餐') !== -1 ||
                                     sessionValLower.indexOf('choir') !== -1 || sessionValLower.indexOf('詩班') !== -1);
        
        for (var pt = 0; pt < presentingTeams.length; pt++) {
          var pTeam = presentingTeams[pt];
          var teamName = pTeam.name;
          var members = pTeam.members;
          
          if (!rowAssignments[actualRow]) rowAssignments[actualRow] = {};
          var presRole = 'Choir Presentation (' + teamName + ')';
          for (var m = 0; m < members.length; m++) {
            var member = members[m];
            if (!rowAssignments[actualRow][member]) rowAssignments[actualRow][member] = [];
            var alreadyHasPres = rowAssignments[actualRow][member].some(function(item) { return item.role === presRole; });
            if (!alreadyHasPres) {
              rowAssignments[actualRow][member].push({ col: 16, role: presRole });
            }
          }
          
          // Pre-Slot Buffer Exclusion: Saturday AM Evangelical Service or Sunday Holy Communion / Choir
          if ((isSatEvangelical || isSunCommunionOrChoir) && actualRow > 2) {
            var prevRow = actualRow - 1;
            if (!rowAssignments[prevRow]) rowAssignments[prevRow] = {};
            var bufferRole = 'Choir Warm-up / Staging Buffer (' + teamName + ')';
            for (var m = 0; m < members.length; m++) {
              var member = members[m];
              if (!rowAssignments[prevRow][member]) rowAssignments[prevRow][member] = [];
              var alreadyHasBuffer = rowAssignments[prevRow][member].some(function(item) { return item.role === bufferRole; });
              if (!alreadyHasBuffer) {
                rowAssignments[prevRow][member].push({ col: 16, role: bufferRole });
              }
            }
          }
        }
      }
      
      // --- D. Cooking Team Meal Prep/Clean Windows ---
      for (var r = 1; r < allData.length; r++) {
        var rawSession = String(allData[r][3] || '').trim();
        var sessionLower = rawSession.toLowerCase();
        var isLunch = (sessionLower.indexOf('lunch') !== -1 || sessionLower.indexOf('午餐') !== -1);
        var isDinner = (sessionLower.indexOf('dinner') !== -1 || sessionLower.indexOf('晚餐') !== -1);
        if (!isLunch && !isDinner) continue;
        
        var mealRow = r + 1;
        var mealDay = cleanLabel(String(allData[r][1] || '')).toLowerCase();
        var mealTitle = cleanLabel(rawSession) || rawSession;
        
        // Read assigned cooking team from Col M (col index 12), or scan other rows on the same day if empty
        var rawCookingTeam = String(allData[r][12] || '').trim();
        if (!rawCookingTeam) {
          for (var scanR = 1; scanR < allData.length; scanR++) {
            var sDay = cleanLabel(String(allData[scanR][1] || '')).toLowerCase();
            if (sDay === mealDay && allData[scanR][12]) {
              rawCookingTeam = String(allData[scanR][12]).trim();
              break;
            }
          }
        }
        if (!rawCookingTeam) continue;
        
        var cookingMembers = resolveCookingTeamWorkers(rawCookingTeam, teamsMap);
        if (!cookingMembers || cookingMembers.length === 0) continue;
        
        // Window rows on the same day:
        // Lunch: ~10:30 AM (630m) to 1:30 PM (810m)
        // Dinner: ~4:15 PM (975m) to 7:00 PM (1140m)
        for (var w = 1; w < allData.length; w++) {
          var wRow = w + 1;
          var wDay = cleanLabel(String(allData[w][1] || '')).toLowerCase();
          if (wDay !== mealDay) continue;
          
          var inWindow = false;
          if (wRow === mealRow) {
            inWindow = true;
          } else {
            var tRange = parseTimeRangeMinutes(String(allData[w][2] || ''));
            if (tRange.start !== null && tRange.end !== null) {
              if (isLunch && tRange.start < 810 && tRange.end > 630) inWindow = true;
              else if (isDinner && tRange.start < 1140 && tRange.end > 975) inWindow = true;
            }
          }
          
          if (inWindow) {
            if (!rowAssignments[wRow]) rowAssignments[wRow] = {};
            var cookingRole = 'Cooking Duty (Prep/Clean: ' + mealTitle + ')';
            for (var m = 0; m < cookingMembers.length; m++) {
              var member = cookingMembers[m];
              if (!rowAssignments[wRow][member]) rowAssignments[wRow][member] = [];
              var alreadyHasDuty = rowAssignments[wRow][member].some(function(item) { return item.role === cookingRole; });
              if (!alreadyHasDuty) {
                rowAssignments[wRow][member].push({ col: 13, role: cookingRole });
              }
            }
          }
        }
      }
      
      // Build date-level and session-level unavailability
      var um = getUnavailabilityData(!!forceRefresh);
      var unavailByDate = {};
      var unavailByRow = {};
      
      for (var rk in rowDates) {
        var rkInt = parseInt(rk, 10);
        var tdsKey = rowDates[rk].tds;
        if (!tdsKey || !um[tdsKey]) continue;
        
        var uod = um[tdsKey];
        if (!unavailByDate[tdsKey]) {
          var dm = {};
          for (var u = 0; u < uod.length; u++) dm[uod[u].name] = uod[u].reason;
          unavailByDate[tdsKey] = dm;
        }
        
        // Find which ESSC block applies to rkInt
        var rIdx = rkInt - 1;
        var rRow = (rIdx < allData.length) ? allData[rIdx] : null;
        var rSession = rRow ? cleanLabel(String(rRow[3] || '')).toLowerCase() : '';
        var rDay = rRow ? cleanLabel(String(rRow[1] || '')).toLowerCase() : '';
        
        var applicableBlock = null;
        for (var b = 0; b < ESSC_BLOCK_DEFINITIONS.length; b++) {
          var bDef = ESSC_BLOCK_DEFINITIONS[b];
          var dayMatch = false;
          if (bDef.dayOffset === 0 && rDay.indexOf('wed') !== -1) dayMatch = true;
          else if (bDef.dayOffset === 1 && rDay.indexOf('thu') !== -1) dayMatch = true;
          else if (bDef.dayOffset === 2 && rDay.indexOf('fri') !== -1) dayMatch = true;
          else if (bDef.dayOffset === 3 && rDay.indexOf('sat') !== -1) dayMatch = true;
          else if (bDef.dayOffset === 4 && rDay.indexOf('sun') !== -1) dayMatch = true;
          
          if (dayMatch) {
            for (var ms = 0; ms < bDef.matchSessions.length; ms++) {
              if (rSession.indexOf(bDef.matchSessions[ms]) !== -1) {
                applicableBlock = bDef;
                break;
              }
            }
          }
          if (applicableBlock) break;
        }
        
        var rowDm = {};
        for (var u = 0; u < uod.length; u++) {
          var uReason = uod[u].reason || '';
          if (uReason.indexOf('ESSC:') !== -1) {
            if (applicableBlock && uReason.indexOf(applicableBlock.label) !== -1) {
              rowDm[uod[u].name] = uReason;
            }
          } else {
            rowDm[uod[u].name] = uReason;
          }
        }
        unavailByRow[rkInt] = rowDm;
      }
      
      var speakerCols = [];
      for (var c in roleMap) {
        if (roleMap[c].toLowerCase().indexOf('speaker') !== -1) {
          speakerCols.push(parseInt(c, 10));
        }
      }
      
      return {
        success: true, sheetName: sName, teamNames: teamNames, teamsMap: teamsMap,
        columnTeamMap: teamMap, columnTeamRawMap: teamRawMap, columnRoleMap: roleMap,
        rowDates: rowDates, rowAssignments: rowAssignments,
        unavailByDate: unavailByDate, unavailByRow: unavailByRow,
        weeklyUnavail: { monday: {}, friday: {} },
        speakerLanguages: getSpeakerLanguages(),
        speakerCols: speakerCols,
        isESSC: true,
        neighbors: []
      };
    }
    
    var maxR = Math.min(sheet.getMaxRows(), CONFIG.MAX_SCHEDULE_ROWS);
    var maxC = Math.min(sheet.getMaxColumns(), CONFIG.MAX_COLUMNS);
    var allData = sheet.getRange(1, 1, maxR, maxC).getValues();
    
    var teamMap = {}, teamRawMap = {}, roleMap = {};
    var row2 = allData.length > 1 ? allData[1] : null;
    var row4 = allData.length > 3 ? allData[3] : null;
    var row5 = allData.length > 4 ? allData[4] : null;
    var row6 = allData.length > 5 ? allData[5] : null;
    
    for (var c = 2; c <= maxC; c++) {
      var idx = c - 1;
      var role = (row4 && String(row4[idx] || '').trim()) ? String(row4[idx] || '').trim() : (row5 ? String(row5[idx] || '').trim() : '');
      if (role) roleMap[c] = role;
      
      var raw = '';
      if (row2) {
        for (var sc = idx; sc >= 0; sc--) { var v = String(row2[sc] || '').trim(); if (v) { raw = v; break; } }
      }
      if (!raw && row6) {
        for (var sc = idx; sc >= 0; sc--) { var v = String(row6[sc] || '').trim(); if (v) { raw = v; break; } }
      }
      
      var resolvedFromRole = resolveTeamMapping(role, teamNames);
      var resolvedFromRaw  = resolveTeamMapping(raw, teamNames);
      
      if (resolvedFromRole) {
          teamMap[c] = resolvedFromRole;
          teamRawMap[c] = role;
      } else if (resolvedFromRaw) {
          teamMap[c] = resolvedFromRaw;
          teamRawMap[c] = raw;
      } else {
          teamMap[c] = raw || role;
          teamRawMap[c] = raw || role;
      }
    }
    
    var rowDates = {}, rowAssignments = {};
    
    for (var r = 5; r < maxR; r++) { 
      var actualRow = r + 1;
      var rowArr = allData[r];
      if (!rowArr) continue;
      var dv = rowArr[0];
      
      var isEmpty = true;
      for (var i = 0; i < maxC; i++) {
         if (String(rowArr[i] || '').trim() !== '') { isEmpty = false; break; }
      }
      if (isEmpty) continue; 
      
      var dateStr = '', tds = '';
      if (dv) {
        var dateObj = dv instanceof Date ? dv : new Date(dv);
        if (!isNaN(dateObj.getTime())) {
          tds = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
          dateStr = Utilities.formatDate(dateObj, tz, 'EEEE, MMM d, yyyy');
        } else { dateStr = String(dv); }
      } else { dateStr = '(No date in Column A)'; }
      
      rowDates[actualRow] = { dateStr: dateStr, tds: tds };
      
      var awm = {};
      
      for (var c = 0; c < maxC; c++) {
        var val = String(rowArr[c] || '').trim();
        
        if (c >= 1 && val) {
          var roleName = roleMap[c + 1] || 'Column ' + colNumToLetter(c + 1);
          if (!awm[val]) awm[val] = [];
          awm[val].push({ col: c + 1, role: roleName });
        }
      }
      
      rowAssignments[actualRow] = awm;
    }
    
    var um = getUnavailabilityData(!!forceRefresh);
    var wda = getWorkersAvailability(!!forceRefresh);
    
    // Date-specific unavailability, only for dates present on this quarter sheet
    var unavailByDate = {};
    for (var rk in rowDates) {
      var tdsKey = rowDates[rk].tds;
      if (!tdsKey || unavailByDate[tdsKey] || !um[tdsKey]) continue;
      var dm = {};
      var uod = um[tdsKey];
      for (var u = 0; u < uod.length; u++) dm[uod[u].name] = uod[u].reason;
      unavailByDate[tdsKey] = dm;
    }
    
    // Weekly (Monday / Friday) unavailability: sent once instead of on every row
    var weeklyUnavail = { monday: {}, friday: {} };
    var workerNames = Object.keys(wda);
    for (var w = 0; w < workerNames.length; w++) {
      var wName = workerNames[w];
      if (wda[wName].monday === false) weeklyUnavail.monday[wName] = 'Not available on Mondays';
      if (wda[wName].friday === false) weeklyUnavail.friday[wName] = 'Not available on Fridays';
    }
    var speakerCols = [];
    for (var c in roleMap) {
      if (roleMap[c].toLowerCase().indexOf('speaker') !== -1) {
        speakerCols.push(parseInt(c, 10));
      }
    }
    
    return {
      success: true, sheetName: sName, teamNames: teamNames, teamsMap: teamsMap,
      columnTeamMap: teamMap, columnTeamRawMap: teamRawMap, columnRoleMap: roleMap,
      rowDates: rowDates, rowAssignments: rowAssignments,
      unavailByDate: unavailByDate, weeklyUnavail: weeklyUnavail, speakerLanguages: getSpeakerLanguages(),
      speakerCols: speakerCols,
      neighbors: getNeighborQuarters(sName)
    };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * Generates a static Annual Summary report across Q1-Q4.
 * Leverages the robust getSheetChunkData parser and strictly whitelists against the Workers sheet.
 */
function generateAnnualSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheetName = 'Annual Summary';
  
  var stats = {};    
  var allRoles = {}; 

  // Safe UI alert helper for both UI and trigger execution contexts
  function showAlert(title, msg) {
    try {
      var ui = SpreadsheetApp.getUi();
      if (ui) {
        ui.alert(title, msg, ui.ButtonSet.OK);
      }
    } catch (e) {
      console.log(title + ': ' + msg);
    }
  }
  
  // 1. Build a strict case-insensitive whitelist from the Workers sheet (Fix ID #24)
  var workerSheet = ss.getSheetByName(CONFIG.WORKERS_SHEET);
  var whitelist = {};
  if (workerSheet) {
    try {
      var wSchema = getWorkerSchema(workerSheet);
      var wMaxRows = Math.min(workerSheet.getMaxRows(), CONFIG.MAX_WORKER_ROWS);
      var wMaxCols = Math.min(workerSheet.getMaxColumns(), CONFIG.MAX_WORKER_COLS);
      var wData = (wMaxRows > 0 && wMaxCols > 0) ? workerSheet.getRange(1, 1, wMaxRows, wMaxCols).getValues() : [];
      var nameCols = (wSchema && wSchema.nameCols && wSchema.nameCols.length > 0) ? wSchema.nameCols : [4];
      var startIdx = (wSchema && wSchema.dataStartIdx !== undefined) ? wSchema.dataStartIdx : 2;
      for (var r = startIdx; r < wData.length; r++) {
        for (var nc = 0; nc < nameCols.length; nc++) {
          var val = String(wData[r][nameCols[nc]] || '').trim();
          if (val) {
            whitelist[val.toUpperCase()] = val;
          }
        }
      }
    } catch (e) {
      console.warn('Could not build worker whitelist from schema: ' + e.message);
    }
  }
  if (Object.keys(whitelist).length === 0) {
    var rawWorkers = getWorkersAvailability(false);
    for (var w in rawWorkers) {
       whitelist[w.toUpperCase()] = w; 
    }
  }
  
  CONFIG.QUARTERS.forEach(function(q) {
     var sheet = ss.getSheetByName(q);
     if (!sheet) return;
     
     // 2. Leverage our robust sidebar parser
     var chunk = getSheetChunkData(sheet, q, false);
     if (!chunk || !chunk.success) return;
     
     var assignments = chunk.rowAssignments;
     var teamMap = chunk.columnTeamRawMap;

     // Collect role headers from column definitions so role headers exist even if no assignments are made yet
     if (teamMap) {
        for (var col in teamMap) {
           var rName = teamMap[col];
           if (!rName || rName.trim() === '') continue;
           if (rName.indexOf('Column ') !== -1) continue;
           var engOnly = rName.replace(/[^\x00-\x7F]/g, "").replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
           if (engOnly.length > 1) {
              rName = engOnly;
           } else {
              rName = rName.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
           }
           if (rName.toUpperCase() === 'SPEAKER FULL NAME') continue;
           allRoles[rName] = true;
        }
     }
     
     for (var r in assignments) {
        var rowData = assignments[r];
        for (var worker in rowData) {
           
           // STRICT VALIDATION
           var matchedWorker = null;
           var lines = String(worker).split('\n');
           
           for (var l = 0; l < lines.length; l++) {
               var lineName = lines[l].trim().toUpperCase();
               if (whitelist.hasOwnProperty(lineName)) {
                   matchedWorker = whitelist[lineName];
                   break;
               }
           }
           
           var fullUpper = String(worker).trim().toUpperCase();
           if (!matchedWorker && whitelist.hasOwnProperty(fullUpper)) {
               matchedWorker = whitelist[fullUpper];
           }
           
           if (!matchedWorker) continue;
           
           // EXPLICIT FILTER: Remove cooking teams from the report
           if (matchedWorker.toUpperCase().indexOf('COOKING TEAM') !== -1) continue;
           
           var cleanWorker = matchedWorker;
           var jobs = rowData[worker];
           
           // 3. Count up their assignments
           for (var j = 0; j < jobs.length; j++) {
               var roleName = teamMap[jobs[j].col];
               if (!roleName || roleName.trim() === '') roleName = jobs[j].role;
               if (!roleName || roleName.indexOf('Column ') !== -1) continue; 
               
               // SANITIZE: Remove non-English (Chinese) characters and fix spacing/newlines
               var englishOnly = roleName.replace(/[^\x00-\x7F]/g, "").replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
               if (englishOnly.length > 1) { 
                   roleName = englishOnly;
               } else {
                   // Fallback if removing Chinese made it totally empty
                   roleName = roleName.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
               }
               
               // EXPLICIT FILTER: Remove the "Speaker Full Name" column
               if (roleName.toUpperCase() === 'SPEAKER FULL NAME') continue;
               
               if (!stats[cleanWorker]) stats[cleanWorker] = {};
               if (!stats[cleanWorker][roleName]) stats[cleanWorker][roleName] = 0;
               
               stats[cleanWorker][roleName]++;
               allRoles[roleName] = true;
           }
        }
     }
  });
  
  // 4. Sorting headers alphabetically 
  var roleHeaders = Object.keys(allRoles).sort();
  var output = [];
  
  // 5. Build Header Row
  var headerRow = ['Worker Name', 'Total Assignments'];
  roleHeaders.forEach(function(r) { headerRow.push(r); });
  output.push(headerRow);
  
  // 6. Build Worker Rows (include all whitelisted workers; fill 0 if unassigned)
  var workers = Object.keys(whitelist).map(function(k) { return whitelist[k]; })
    .filter(function(w) { return w.toUpperCase().indexOf('COOKING TEAM') === -1; })
    .sort();
  
  if (workers.length === 0) {
    workers = Object.keys(stats).sort();
  }

  workers.forEach(function(w) {
     var row = [w];
     var total = 0;
     var roleCounts = [];
     
     roleHeaders.forEach(function(r) {
        var count = (stats[w] && stats[w][r]) ? stats[w][r] : 0;
        roleCounts.push(count);
        total += count;
     });
     
     row.push(total); 
     row = row.concat(roleCounts); 
     output.push(row);
  });
  
  // 7. Write to Sheet
  var summarySheet = ss.getSheetByName(summarySheetName);
  if (!summarySheet) {
     summarySheet = ss.insertSheet(summarySheetName, 0); 
  } else {
     summarySheet.clear(); 
  }
  
  summarySheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  
  var headerRange = summarySheet.getRange(1, 1, 1, output[0].length);
  headerRange.setFontWeight('bold').setBackground('#e0e7ff');
  summarySheet.setFrozenRows(1);
  summarySheet.setFrozenColumns(1);
  summarySheet.autoResizeColumns(1, output[0].length);
  
  showAlert('Success!', 'The Annual Summary has been generated.');
}

// -------------------------------------------------------------------------
// Year-End Reset Functions
// -------------------------------------------------------------------------

function promptResetForNewYear() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.prompt(
    'Reset for New Year',
    'Please enter the NEW YEAR you want to reset the spreadsheet for (e.g., 2027):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK) {
    var newYearStr = response.getResponseText().trim();
    var newYear = parseInt(newYearStr, 10);
    
    if (isNaN(newYear) || newYear < 2000 || newYear > 2100) {
      ui.alert('Invalid Input', 'Please enter a valid 4-digit year.', ui.ButtonSet.OK);
      return;
    }
    
    var confirm = ui.alert(
      'Confirm Reset',
      'Are you absolutely sure you want to reset all Q1-Q4 sheets for the year ' + newYear + '?\n\n' +
      'This will erase all assigned names but will preserve your formulas, X boxes, and formatting.',
      ui.ButtonSet.YES_NO
    );
    
    if (confirm == ui.Button.YES) {
      try {
        resetForNewYear(newYear);
        ui.alert('Success', 'The spreadsheet has been reset for ' + newYear + '.', ui.ButtonSet.OK);
      } catch(e) {
        ui.alert('Error', 'An error occurred during reset: ' + e.toString(), ui.ButtonSet.OK);
      }
    }
  }
}

function resetForNewYear(year, targetSs) {
  var ss = targetSs || SpreadsheetApp.getActiveSpreadsheet();
  var quarters = CONFIG.QUARTERS;
  var startRow = CONFIG.START_ROW; // Row 5 is header, data starts at Row 6
  
  for (var i = 0; i < quarters.length; i++) {
    var sheet = ss.getSheetByName(quarters[i]);
    if (!sheet) continue;
    
    // 1. Update the Anchor Date in cell A2
    var anchorMonth = i * 3; 
    var newAnchorDate = new Date(year, anchorMonth, 1);
    sheet.getRange("A2").setValue(newAnchorDate);
    
    // 2. Clear assigned names while preserving formulas, 'X's, and blank cells
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    
    if (lastRow < startRow + 1 || lastCol < 2) continue; 
    
    var dataRange = sheet.getRange(startRow + 1, 2, lastRow - startRow, lastCol - 1);
    var values = dataRange.getValues();
    var formulas = dataRange.getFormulas();
    
    for (var r = 0; r < values.length; r++) {
      for (var c = 0; c < values[r].length; c++) {
        var cellFormula = String(formulas[r][c] || '').trim();
        var cellValue = String(values[r][c] || '').trim().toUpperCase();
        
        if (cellFormula !== '') {
          values[r][c] = formulas[r][c]; // preserve exact formula
        } else if (cellValue !== 'X' && cellValue !== '') {
          values[r][c] = ''; // clear assigned name
        }
      }
    }
    dataRange.setValues(values);
  }
  
  // 3. Clear the Email Log sheet if present
  var logSheet = ss.getSheetByName("Email Log");
  if (logSheet) {
    var lastLogRow = logSheet.getLastRow();
    var lastLogCol = logSheet.getLastColumn();
    if (lastLogRow > 1 && lastLogCol > 0) {
      logSheet.getRange(2, 1, lastLogRow - 1, lastLogCol).clearContent();
    }
  }

  // 4. Clear the SWR Sync Log sheet if present
  var swrLogSheet = ss.getSheetByName(CONFIG.SYNC_LOG_SHEET);
  if (swrLogSheet) {
    var lastSwrRow = swrLogSheet.getLastRow();
    var lastSwrCol = swrLogSheet.getLastColumn();
    if (lastSwrRow > 1 && lastSwrCol > 0) {
      swrLogSheet.getRange(2, 1, lastSwrRow - 1, lastSwrCol).clearContent();
    }
  }

  // 5. Automatically delete prior-year generated monthly tabs (e.g., January 2026 – December 2026)
  var monthPattern = /^(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})$/i;
  var allSheets = ss.getSheets();
  for (var s = 0; s < allSheets.length; s++) {
    var sh = allSheets[s];
    var match = sh.getName().match(monthPattern);
    if (match) {
      var tabYear = parseInt(match[2], 10);
      if (tabYear < year) {
        try {
          ss.deleteSheet(sh);
        } catch (delErr) {
          console.warn("Could not delete prior year monthly tab " + sh.getName() + ": " + delErr);
        }
      }
    }
  }
}

/**
 * UI Prompt to restructure Q1-Q4 sheets into the Universal 14-Row Architecture.
 */
function promptRestructureQuarterSheetsToUniversal14() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    'Restructure Q1-Q4 Sheets to Universal 14-Row Format',
    'This automated tool will restructure sheets Q1, Q2, Q3, and Q4 into the Universal 14-Row Architecture:\n\n' +
    '1. Standardizes all 6 day sections (Mondays, Tuesdays, Wednesdays, Fridays, Saturday AM, Saturday PM) to 14 rows across all 4 quarters (Rows 1–94).\n' +
    '2. Injects smart, quarter-bounded date formulas in Column A that automatically handle 12, 13, or 14 weeks without cross-quarter spill.\n' +
    '3. Automatically dims inactive 13th or 14th rows via conditional formatting.\n' +
    '4. Sets Speaker XLOOKUP formulas in Column AY and cleans legacy residual rows.\n' +
    '5. Ensures all future annual rollovers (2028+) automatically deploy this future-proof layout with zero manual edits.\n\n' +
    'Do you want to proceed?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  try {
    var res = restructureQuarterSheetsToUniversal14();
    ui.alert(
      'Restructuring Complete! 🎉',
      'Successfully upgraded ' + res.updatedQuarters.join(', ') + ' to the Universal 14-Row Architecture.\n\n' +
      '• All sections now have 14 fixed rows (Rows 1–94).\n' +
      '• Column A dates now dynamically calculate based on cell A2 and auto-terminate at quarter boundaries.\n' +
      '• Inactive weeks are cleanly grayed out.\n' +
      '• All future annual rollovers will automatically deploy this future-proof layout.',
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('Restructure Failed', 'Error: ' + err.message, ui.ButtonSet.OK);
  }
}

/**
 * Interactive UI Prompt allowing admins to specify column(s) to remove across Q1-Q4.
 */
function promptRemoveCustomColumns() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt(
    '🧹 Remove Column(s) (Interactive)',
    'Enter column letter(s) or numbers to remove across Q1-Q4 (and MonthlyTemplate).\n\n' +
    'Examples:\n' +
    '  • Column letters: AS, AT, AU  or  M, N\n' +
    '  • Column numbers: 45, 46, 47  or  45-47\n\n' +
    'Target column(s):',
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() !== ui.Button.OK) return;
  var inputStr = res.getResponseText();
  if (!inputStr || !inputStr.trim()) {
    ui.alert('Notice', 'No columns were specified.', ui.ButtonSet.OK);
    return;
  }

  var colIndices = parseColumnInputs_(inputStr);
  if (!colIndices || colIndices.length === 0) {
    ui.alert('Error', 'Invalid column input: "' + inputStr + '". Please enter valid column letters (e.g. AS) or numbers (e.g. 45).', ui.ButtonSet.OK);
    return;
  }

  // Inspect Q1 to get role details for confirmation
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quarters = (typeof CONFIG !== 'undefined' && CONFIG.QUARTERS) ? CONFIG.QUARTERS : ['Q1', 'Q2', 'Q3', 'Q4'];
  var q1Sheet = ss.getSheetByName('Q1') || ss.getSheetByName(quarters[0]);
  var previewLines = [];

  colIndices.forEach(function(cIdx) {
    var cLetter = colIdxToLetter_(cIdx);
    var teamName = '';
    var roleTitle = '';
    if (q1Sheet && cIdx <= q1Sheet.getLastColumn()) {
      teamName = String(q1Sheet.getRange(2, cIdx).getValue() || '').trim();
      roleTitle = String(q1Sheet.getRange(4, cIdx).getValue() || '').trim();
    }
    var infoStr = '• Column ' + cLetter + ' (Col ' + cIdx + ')';
    if (teamName || roleTitle) {
      infoStr += ': ' + [teamName, roleTitle].filter(Boolean).join(' - ');
    }
    previewLines.push(infoStr);
  });

  var confirmMsg = 'You are about to remove ' + colIndices.length + ' column(s) across Q1, Q2, Q3, Q4, and MonthlyTemplate:\n\n' +
    previewLines.join('\n') + '\n\n' +
    'This will:\n' +
    '1. Permanently delete these column(s) from master schedule grids.\n' +
    '2. Shift subsequent columns left and update Coordinators q_columns mapping.\n' +
    '3. Re-apply validation rules and standard layout formatting.\n\n' +
    'Do you want to proceed?';

  var confirm = ui.alert('⚠️ Confirm Column Deletion', confirmMsg, ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  ss.toast('Removing selected column(s)... Please wait', 'Processing', 20);

  try {
    var outcome = removeCustomColumns(colIndices, ss);
    ui.alert(
      'Column Deletion Complete! 🎉',
      'Successfully removed ' + outcome.removedCount + ' column(s) across ' + outcome.updatedQuarters.join(', ') + '.\n\n' +
      '• Deleted: ' + previewLines.map(function(l) { return l.replace(/^•\s*/, ''); }).join(', ') + '\n' +
      '• Standardized layout and validation rules have been re-applied.\n' +
      '• Coordinator column mappings updated.',
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('Deletion Failed', 'Error: ' + err.message, ui.ButtonSet.OK);
  }
}

/**
 * Converts column letter to 1-based index (e.g. A=1, B=2, AS=45).
 */
function colLetterToIdx_(letter) {
  var col = 0, str = String(letter || '').trim().toUpperCase();
  for (var i = 0; i < str.length; i++) {
    col = col * 26 + (str.charCodeAt(i) - 64);
  }
  return col;
}

/**
 * Converts 1-based index to column letter (e.g. 1=A, 2=B, 45=AS).
 */
function colIdxToLetter_(colIdx) {
  var temp = '', letter = '';
  while (colIdx > 0) {
    temp = (colIdx - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    colIdx = (colIdx - temp - 1) / 26;
  }
  return letter;
}

/**
 * Parses user input string into a sorted, unique array of 1-based column indices.
 */
function parseColumnInputs_(inputStr) {
  var tokens = inputStr.trim().split(/[\s,;]+/);
  var colMap = {};

  tokens.forEach(function(token) {
    token = token.trim();
    if (!token) return;

    var rangeMatch = token.match(/^([A-Za-z]+|\d+)-([A-Za-z]+|\d+)$/);
    if (rangeMatch) {
      var startIdx = isNaN(rangeMatch[1]) ? colLetterToIdx_(rangeMatch[1]) : parseInt(rangeMatch[1], 10);
      var endIdx = isNaN(rangeMatch[2]) ? colLetterToIdx_(rangeMatch[2]) : parseInt(rangeMatch[2], 10);
      if (startIdx > 0 && endIdx >= startIdx) {
        for (var i = startIdx; i <= endIdx; i++) colMap[i] = true;
      }
    } else if (token.match(/^\d+$/)) {
      var idx = parseInt(token, 10);
      if (idx > 1) colMap[idx] = true;
    } else if (token.match(/^[A-Za-z]+$/)) {
      var idx = colLetterToIdx_(token);
      if (idx > 1) colMap[idx] = true;
    }
  });

  return Object.keys(colMap).map(Number).sort(function(a, b) { return a - b; });
}

/**
 * Core engine that programmatically deletes specified columns across Q1-Q4 and MonthlyTemplate,
 * shifting Coordinators tab q_columns and MappingConfig, and restoring grid layout.
 *
 * @param {number[]|number} targetCols Column index or array of column indices (1-based)
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSs] Optional target spreadsheet
 */
function removeCustomColumns(targetCols, targetSs) {
  var ss = targetSs || SpreadsheetApp.getActiveSpreadsheet();
  var colIndices = Array.isArray(targetCols) ? targetCols.slice() : [targetCols];
  
  colIndices.sort(function(a, b) { return b - a; });

  var quarters = (typeof CONFIG !== 'undefined' && CONFIG.QUARTERS) ? CONFIG.QUARTERS : ['Q1', 'Q2', 'Q3', 'Q4'];
  var sheetsToProcess = quarters.slice();
  var updatedQuarters = [];

  sheetsToProcess.forEach(function(sName) {
    var sheet = ss.getSheetByName(sName);
    if (!sheet) return;

    var modified = false;
    colIndices.forEach(function(colIdx) {
      if (sheet.getMaxColumns() >= colIdx) {
        sheet.deleteColumn(colIdx);
        modified = true;
      }
    });
    if (modified) updatedQuarters.push(sName);
  });

  // Shift Coordinators tab q_columns mapping
  try {
    var coordSheet = ss.getSheetByName('Coordinators');
    if (coordSheet && coordSheet.getLastRow() > 1) {
      var dataRange = coordSheet.getRange(2, 1, coordSheet.getLastRow() - 1, coordSheet.getLastColumn());
      var values = dataRange.getValues();
      var headers = coordSheet.getRange(1, 1, 1, coordSheet.getLastColumn()).getValues()[0];
      var qColIdx = -1;

      for (var h = 0; h < headers.length; h++) {
        if (String(headers[h] || '').trim().toLowerCase() === 'q_columns') {
          qColIdx = h;
          break;
        }
      }

      if (qColIdx !== -1) {
        var coordModified = false;
        values.forEach(function(row) {
          var qColsStr = String(row[qColIdx] || '').trim();
          if (!qColsStr) return;

          var tokens = qColsStr.split(/[\s,;]+/);
          var newTokens = [];

          tokens.forEach(function(tok) {
            var cleanTok = tok.trim().toUpperCase();
            if (!cleanTok.match(/^[A-Z]+$/)) return;

            var cIdx = colLetterToIdx_(cleanTok);

            var wasDeleted = colIndices.indexOf(cIdx) !== -1;
            if (wasDeleted) {
              coordModified = true;
              return;
            }

            var shiftAmount = 0;
            colIndices.forEach(function(dIdx) {
              if (dIdx < cIdx) shiftAmount++;
            });

            var newIdx = cIdx - shiftAmount;
            newTokens.push(colIdxToLetter_(newIdx));
            if (shiftAmount > 0) coordModified = true;
          });

          row[qColIdx] = newTokens.join(', ');
        });

        if (coordModified) {
          coordSheet.getRange(2, qColIdx + 1, values.length, 1).setValues(values.map(function(r) { return [r[qColIdx]]; }));
        }
      }
    }
  } catch (e) {
    console.warn('Could not shift Coordinators tab q_columns: ' + e.message);
  }

  // Re-apply validation & layout
  try {
    if (typeof ensureValidationConfigSheet === 'function') {
      ensureValidationConfigSheet(ss);
    }
  } catch (e) {}

  if (typeof restructureQuarterSheetsToUniversal14 === 'function') {
    restructureQuarterSheetsToUniversal14(ss);
  }

  try {
    if (typeof bumpAllQuarterVersions === 'function') {
      bumpAllQuarterVersions();
    }
  } catch (e) {}

  return {
    success: true,
    removedCount: colIndices.length,
    updatedQuarters: updatedQuarters
  };
}

/**
 * Backward-compatible wrapper for promptDeleteLegacyCleaningColumns
 */
function promptDeleteLegacyCleaningColumns() {
  promptRemoveCustomColumns();
}

/**
 * Backward-compatible wrapper for deleteLegacyCleaningColumns
 */
function deleteLegacyCleaningColumns(targetSs) {
  return removeCustomColumns([45, 46, 47], targetSs);
}

/**
 * Dynamically computes activeCols for each Universal 14 section based on RoleConfig tab if available.
 * @param {Spreadsheet} ss
 * @param {Array<Object>} defaultSections
 * @returns {Array<Object>}
 */
function getActiveColsFromRoleConfig_(ss, defaultSections) {
  if (!ss) return defaultSections;
  var roleSheet = ss.getSheetByName('RoleConfig');
  if (!roleSheet || roleSheet.getLastRow() < 2) return defaultSections;
  
  var data = roleSheet.getDataRange().getValues();
  if (data.length < 2) return defaultSections;
  
  var headers = data[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
  var dayTypeIdx = headers.indexOf('day_type');
  var secIdx = headers.indexOf('section');
  var qColIdx = headers.indexOf('q_column');
  var enabledIdx = headers.indexOf('enabled');
  
  if (dayTypeIdx === -1 || qColIdx === -1) return defaultSections;
  
  var sectionColSets = {
    "MONDAYS": [],
    "TUESDAYS": [],
    "WEDNESDAYS": [],
    "FRIDAYS": [],
    "SATURDAYS AM": [],
    "SATURDAYS PM": []
  };
  
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (enabledIdx !== -1 && String(row[enabledIdx] || '').trim().toUpperCase() === 'N') continue;
    
    var dayType = String(row[dayTypeIdx] || '').trim().toUpperCase();
    var sec = (secIdx !== -1) ? String(row[secIdx] || '').trim().toUpperCase() : '';
    var qColStr = String(row[qColIdx] || '').trim().toUpperCase();
    if (!qColStr) continue;
    
    var secKey = '';
    if (dayType === 'MON') secKey = 'MONDAYS';
    else if (dayType === 'TUE') secKey = 'TUESDAYS';
    else if (dayType === 'WED') secKey = 'WEDNESDAYS';
    else if (dayType === 'FRI') secKey = 'FRIDAYS';
    else if (dayType === 'SAT') {
      if (sec === 'AM' || sec.indexOf('AM') !== -1) secKey = 'SATURDAYS AM';
      else if (sec === 'PM' || sec.indexOf('PM') !== -1) secKey = 'SATURDAYS PM';
      else if (sec === 'CLEANING') secKey = 'SATURDAYS PM';
      else secKey = 'SATURDAYS AM';
    }
    
    if (!sectionColSets[secKey]) continue;
    
    var colLetters = qColStr.split(/[\s,;]+/);
    for (var cl = 0; cl < colLetters.length; cl++) {
      var let = colLetters[cl].trim().toUpperCase();
      if (!let) continue;
      var cNum = colLetterToIdx_(let); // 1-indexed column number
      if (cNum >= 2 && cNum <= 51) {
        if (sectionColSets[secKey].indexOf(cNum) === -1) sectionColSets[secKey].push(cNum);
        // AY represents Speaker in RoleConfig; Col B (Col 2) is the actual input cell on Universal 14
        if (let === 'AY' && sectionColSets[secKey].indexOf(2) === -1) {
          sectionColSets[secKey].push(2);
        }
      }
    }
  }
  
  return defaultSections.map(function(sec) {
    var foundCols = sectionColSets[sec.name];
    if (foundCols && foundCols.length > 0) {
      // Preserve RE columns if Saturday AM/PM had them in default
      if (sec.name === 'SATURDAYS AM') {
        [12, 13, 14, 15, 16, 17].forEach(function(reCol) {
          if (foundCols.indexOf(reCol) === -1) foundCols.push(reCol);
        });
      } else if (sec.name === 'SATURDAYS PM') {
        [12, 13, 14, 15, 16, 17, 18, 20, 22].forEach(function(reCol) {
          if (foundCols.indexOf(reCol) === -1) foundCols.push(reCol);
        });
      }
      foundCols.sort(function(a, b) { return a - b; });
      var clone = Object.assign({}, sec);
      clone.activeCols = foundCols;
      return clone;
    }
    return sec;
  });
}

/**
 * Programmatically restructures Q1-Q4 sheets to the Universal 14-Row Architecture.
 * Can be called standalone or during annual provisioning.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSs] Optional target spreadsheet.
 * @return {{ success: boolean, updatedQuarters: string[] }}
 */
function restructureQuarterSheetsToUniversal14(targetSs) {
  var ss = targetSs || SpreadsheetApp.getActiveSpreadsheet();
  var quarters = CONFIG.QUARTERS; // ['Q1', 'Q2', 'Q3', 'Q4']
  var updatedQuarters = [];

  var BASE_SECTIONS_CONFIG = [
    {
      name: "MONDAYS",
      headerRow: 5,
      startRow: 6,
      endRow: 19,
      mask: "0111111",
      activeCols: [2, 3, 4, 5, 6, 7, 8, 9] // Speaker, Interpreter, Praise, Accompanist, Ushers, AV
    },
    {
      name: "TUESDAYS",
      headerRow: 20,
      startRow: 21,
      endRow: 34,
      mask: "1011111",
      activeCols: [45] // Joy Fellowship in AS
    },
    {
      name: "WEDNESDAYS",
      headerRow: 35,
      startRow: 36,
      endRow: 49,
      mask: "1101111",
      activeCols: [46, 47, 48, 49, 50] // Agape venues & Wed Study (Columns AT to AX)
    },
    {
      name: "FRIDAYS",
      headerRow: 50,
      startRow: 51,
      endRow: 64,
      mask: "1111011",
      activeCols: [2, 3, 4, 5, 6, 7, 8, 9, 30, 31, 32] // Speaker, Praise, Accompanist, Ushers, AV, Cleaning Chapel 1-3
    },
    {
      name: "SATURDAYS AM",
      headerRow: 65,
      startRow: 66,
      endRow: 79,
      mask: "1111101",
      activeCols: [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 23, 33, 34] // Main service, Grace, Morning Prayer, RE K-E2, Cooking, Trash & Bros Down
    },
    {
      name: "SATURDAYS PM",
      headerRow: 80,
      startRow: 81,
      endRow: 94,
      mask: null, // Mirrors Saturday AM
      activeCols: [2, 3, 4, 5, 6, 7, 8, 9, 12, 13, 14, 15, 16, 17, 18, 20, 22, 23, 24, 25, 26, 27, 28, 29, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44]
    }
  ];

  var SECTIONS_CONFIG = getActiveColsFromRoleConfig_(ss, BASE_SECTIONS_CONFIG);

  var NUM_GRID_ROWS = 90; // Rows 5 to 94 inclusive
  var NUM_GRID_COLS = 51; // Columns 1 to 51 (Col A to Col AY)

  for (var qIdx = 0; qIdx < quarters.length; qIdx++) {
    var qName = quarters[qIdx];
    var sheet = ss.getSheetByName(qName);
    if (!sheet) continue;

    // 1. Ensure Anchor Date in Cell A2
    var a2Val = sheet.getRange("A2").getValue();
    var currentYear = new Date().getFullYear();
    if (a2Val && a2Val instanceof Date) {
      currentYear = a2Val.getFullYear();
    } else if (typeof a2Val === 'string' && a2Val.match(/\b(20\d\d)\b/)) {
      currentYear = parseInt(a2Val.match(/\b(20\d\d)\b/)[1], 10);
    }
    var anchorMonth = qIdx * 3;
    var anchorDate = new Date(currentYear, anchorMonth, 1);
    sheet.getRange("A2").setValue(anchorDate).setNumberFormat("yyyy-MM-dd");

    // Ensure sheet has at least 94 rows and 51 columns
    if (sheet.getMaxRows() < 94) {
      sheet.insertRowsAfter(sheet.getMaxRows(), 94 - sheet.getMaxRows());
    }
    if (sheet.getMaxColumns() < 51) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), 51 - sheet.getMaxColumns());
    }

    // Initialize formatting batch arrays (90 rows x 51 cols)
    var backgrounds = [];
    var fontWeights = [];
    var fontColors = [];
    var horizAligns = [];
    var vertAligns = [];
    for (var r = 0; r < NUM_GRID_ROWS; r++) {
      backgrounds.push(new Array(NUM_GRID_COLS));
      fontWeights.push(new Array(NUM_GRID_COLS));
      fontColors.push(new Array(NUM_GRID_COLS));
      horizAligns.push(new Array(NUM_GRID_COLS));
      vertAligns.push(new Array(NUM_GRID_COLS));
    }

    // 2. Setup each section: Header, Date Formulas, Role X Masks, and Speaker Column BB
    for (var s = 0; s < SECTIONS_CONFIG.length; s++) {
      var sec = SECTIONS_CONFIG[s];
      var hRowIdx = sec.headerRow - 5; // offset in 90-row grid (0, 15, 30, 45, 60, 75)

      // A. Header Row Formatting & Values across Cols 1 to 51
      for (var c = 0; c < NUM_GRID_COLS; c++) {
        backgrounds[hRowIdx][c] = "#D9D9D9";
        fontWeights[hRowIdx][c] = "bold";
        fontColors[hRowIdx][c] = "#000000";
        horizAligns[hRowIdx][c] = "center";
        vertAligns[hRowIdx][c] = "middle";
      }
      var headerValues = [new Array(NUM_GRID_COLS)];
      headerValues[0][0] = sec.name;
      for (var c = 1; c < NUM_GRID_COLS; c++) {
        headerValues[0][c] = "";
      }
      sheet.getRange(sec.headerRow, 1, 1, NUM_GRID_COLS).setValues(headerValues);

      // B. Date Formulas in Column A
      var numRows = sec.endRow - sec.startRow + 1;
      var colAFormulas = [];
      if (sec.mask) {
        // Week 1
        var w1Formula = '=WORKDAY.INTL(DATE(YEAR(A$2), INT((MONTH(A$2)-1)/3)*3 + 1, 1) - 1, 1, "' + sec.mask + '")';
        colAFormulas.push([w1Formula]);

        // Weeks 2 to 14
        for (var r = sec.startRow + 1; r <= sec.endRow; r++) {
          var prevRow = r - 1;
          var weekFormula = '=IF(A' + prevRow + '="", "", IF(A' + prevRow + '+7 > EOMONTH(DATE(YEAR(A$2), INT((MONTH(A$2)-1)/3)*3 + 1, 1), 2), "", A' + prevRow + '+7))';
          colAFormulas.push([weekFormula]);
        }
      } else {
        // Saturday PM mirrors Saturday AM (Rows 66 to 79)
        var amStartRow = 66;
        for (var r = sec.startRow; r <= sec.endRow; r++) {
          var mirrorAmRow = amStartRow + (r - sec.startRow);
          colAFormulas.push(['=IF(A' + mirrorAmRow + '="", "", A' + mirrorAmRow + ')']);
        }
      }

      sheet.getRange(sec.startRow, 1, numRows, 1)
        .setFormulas(colAFormulas)
        .setNumberFormat("M/d/yyyy");

      // C. Role 'X' Masks and Active Columns across Cols B (col 2) to AX (col 50)
      var activeSet = {};
      for (var a = 0; a < sec.activeCols.length; a++) activeSet[sec.activeCols[a]] = true;

      var gridRange = sheet.getRange(sec.startRow, 2, numRows, 49); // Cols 2 to 50
      var gridValues = gridRange.getValues();
      var gridFormulas = gridRange.getFormulas();

      for (var rIdx = 0; rIdx < numRows; rIdx++) {
        var absRowIdx = (sec.startRow + rIdx) - 5; // row index in 90-row grid

        // Column A formatting (c = 0)
        backgrounds[absRowIdx][0] = "#FFFFFF";
        fontWeights[absRowIdx][0] = "normal";
        fontColors[absRowIdx][0] = "#000000";
        horizAligns[absRowIdx][0] = "center";
        vertAligns[absRowIdx][0] = "middle";

        // Columns B to AX (c = 1 to 49, colNum = 2 to 50)
        for (var colIdx = 0; colIdx < 49; colIdx++) {
          var colNum = colIdx + 2;
          var c = colIdx + 1;
          var cellVal = String(gridValues[rIdx][colIdx]).trim();
          var cellForm = String(gridFormulas[rIdx][colIdx]).trim();

          if (!activeSet[colNum]) {
            // Inactive column -> set 'X' and dark gray background
            gridValues[rIdx][colIdx] = 'X';
            backgrounds[absRowIdx][c] = "#D9D9D9";
          } else {
            // Active column -> white background; clear legacy 'X' if present
            if (cellVal === 'X' || cellForm === 'X') {
              gridValues[rIdx][colIdx] = '';
            }
            backgrounds[absRowIdx][c] = "#FFFFFF";
          }
          fontWeights[absRowIdx][c] = "normal";
          fontColors[absRowIdx][c] = "#000000";
          horizAligns[absRowIdx][c] = "center";
          vertAligns[absRowIdx][c] = "middle";
        }

        // Column AY formatting (c = 50, colNum = 51)
        if (sec.name === "TUESDAYS") {
          backgrounds[absRowIdx][50] = "#D9D9D9";
        } else {
          backgrounds[absRowIdx][50] = "#FFFFFF";
        }
        fontWeights[absRowIdx][50] = "normal";
        fontColors[absRowIdx][50] = "#000000";
        horizAligns[absRowIdx][50] = "center";
        vertAligns[absRowIdx][50] = "middle";
      }
      gridRange.setValues(gridValues);

      // D. Column AY (Speaker Full Name lookup)
      var ayRange = sheet.getRange(sec.startRow, 51, numRows, 1);
      if (sec.name === "TUESDAYS") {
        var ayVals = [];
        for (var k = 0; k < numRows; k++) ayVals.push(['X']);
        ayRange.setValues(ayVals);
      } else {
        var wSheet = ss.getSheetByName(CONFIG.WORKERS_SHEET);
        var wSchema = null;
        try { if (wSheet) wSchema = getWorkerSchema(wSheet); } catch (e) {}
        var nameColLet = (wSchema && wSchema.nameCols && wSchema.nameCols.length > 0) ? colIdxToLetter_(wSchema.nameCols[0] + 1) : 'E';
        var chnColLet = (wSchema && wSchema.chineseNameCol !== -1) ? colIdxToLetter_(wSchema.chineseNameCol + 1) : 'H';
        var wStartRow = (wSchema && wSchema.dataStartIdx !== undefined) ? (wSchema.dataStartIdx + 1) : 3;

        var ayFormulas = [];
        for (var r = sec.startRow; r <= sec.endRow; r++) {
          ayFormulas.push(['=IF(ISBLANK(B' + r + '), "", IFERROR(XLOOKUP(B' + r + ', Workers!$' + nameColLet + '$' + wStartRow + ':$' + nameColLet + ', Workers!$' + chnColLet + '$' + wStartRow + ':$' + chnColLet + ', B' + r + '), B' + r + '))']);
        }
        ayRange.setFormulas(ayFormulas);
      }
    }

    // 3. Batch apply all formatting across Rows 5 to 94, Columns 1 to 51 (A5:AY94)
    var masterRange = sheet.getRange(5, 1, NUM_GRID_ROWS, NUM_GRID_COLS);
    masterRange.setBackgrounds(backgrounds);
    masterRange.setFontWeights(fontWeights);
    masterRange.setFontColors(fontColors);
    masterRange.setHorizontalAlignments(horizAligns);
    masterRange.setVerticalAlignments(vertAligns);
    masterRange.setBorder(true, true, true, true, true, true, "#B7B7B7", SpreadsheetApp.BorderStyle.SOLID);

    // 4. Clear residual legacy rows beyond Row 94
    var maxRows = sheet.getMaxRows();
    if (maxRows > 94) {
      sheet.getRange(95, 1, maxRows - 94, sheet.getMaxColumns()).clear();
    }

    // 5. Inactive Row Visual Dimming (Conditional Formatting)
    var rules = sheet.getConditionalFormatRules();
    var newRules = [];
    var hasDimRule = false;
    for (var i = 0; i < rules.length; i++) {
      var rule = rules[i];
      var cond = rule.getBooleanCondition();
      if (cond && cond.getCriteriaValues()) {
        var cVal = cond.getCriteriaValues()[0];
        if (cVal === '=$A6=""') {
          hasDimRule = true;
        }
        // Normalize any old "X" conditional formatting rule to cover the full A6:AY94 range
        if (cVal === 'X' || cVal === '"X"') {
          try {
            rule = rule.copy().setRanges([sheet.getRange("A6:AY94")]).build();
          } catch (e) {}
        }
      }
      newRules.push(rule);
    }
    if (!hasDimRule) {
      var dimRange = sheet.getRange(6, 1, 89, sheet.getMaxColumns());
      var dimRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$A6=""')
        .setBackground("#F3F3F3")
        .setFontColor("#AAAAAA")
        .setRanges([dimRange])
        .build();
      newRules.push(dimRule);
    }
    sheet.setConditionalFormatRules(newRules);

    updatedQuarters.push(qName);
  }

  return { success: true, updatedQuarters: updatedQuarters };
}

/**
 * UI Prompt to reset a Sermon Topics & Hymns spreadsheet for a target year.
 */
function promptResetSermonTopicsSheet() {
  var ui = SpreadsheetApp.getUi();
  var currentYear = new Date().getFullYear();
  var defaultNextYear = currentYear + 1;
  
  var respYear = ui.prompt(
    'Reset Sermon Topics & Hymns Sheet',
    'Please enter the 4-digit NEW YEAR to reset for (e.g., ' + defaultNextYear + '):',
    ui.ButtonSet.OK_CANCEL
  );
  if (respYear.getSelectedButton() !== ui.Button.OK) return;
  var yearStr = respYear.getResponseText().trim();
  var year = parseInt(yearStr, 10);
  if (isNaN(year) || year < 2020 || year > 2100) {
    ui.alert('Invalid Year', 'Please enter a valid 4-digit year.', ui.ButtonSet.OK);
    return;
  }
  
  var sermonId = getLinkedSermonTopicsId();
  var respSheet = ui.prompt(
    'Target Sermon Topics Sheet',
    'Enter the Google Spreadsheet URL or File ID of the Sermon Topics & Hymns sheet to reset:\n\n(Detected ID: ' + (sermonId || 'None') + ')',
    ui.ButtonSet.OK_CANCEL
  );
  if (respSheet.getSelectedButton() !== ui.Button.OK) return;
  var inputId = respSheet.getResponseText().trim();
  if (inputId) {
    var match = inputId.match(/\/d\/([a-zA-Z0-9-_]+)/);
    sermonId = match ? match[1] : inputId;
  }
  if (!sermonId) {
    ui.alert('Error', 'No valid Sermon Topics spreadsheet ID provided.', ui.ButtonSet.OK);
    return;
  }
  
  var hwsUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var confirm = ui.alert(
    'Confirm Sermon Topics Reset for ' + year,
    'This will:\n' +
    '1. Set Q1!A1 anchor date to 1/1/' + year + ' (automatically updating all quarter dates).\n' +
    '2. Update IMPORTRANGE formulas in Q1-Q4 to point to this Holy Work Schedule.\n' +
    '3. Clear previous Sermon Titles, Hymns, and Usher Attendance across all quarters.\n' +
    '4. Rename ESSC tabs to ' + year + ' and clear past speakers (Column B), topics, and hymns (leaving cells with "X" and gray backgrounds).\n\n' +
    'Are you sure you want to proceed?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Resetting Sermon Topics sheet for ' + year + '...', 'Please wait', 10);
    var sermonSs = SpreadsheetApp.openById(sermonId);
    resetSermonTopicsSheet(year, sermonSs, hwsUrl);
    ui.alert('Success! 🎉', 'Sermon Topics & Hymns sheet "' + sermonSs.getName() + '" has been reset for ' + year + '!', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('Reset Failed', 'An error occurred during reset:\n\n' + err.message, ui.ButtonSet.OK);
  }
}

/**
 * Helper to determine if a hex color string represents a neutral gray background (e.g. #D9D9D9, #B7B7B7, #CCCCCC).
 */
function isGrayBackground(bg) {
  if (!bg || typeof bg !== 'string') return false;
  var hex = bg.toLowerCase().replace('#', '').trim();
  if (hex.length === 3) hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
  if (hex.length !== 6) return false;
  var r = parseInt(hex.substring(0, 2), 16);
  var g = parseInt(hex.substring(2, 4), 16);
  var b = parseInt(hex.substring(4, 6), 16);
  if (isNaN(r) || isNaN(g) || isNaN(b)) return false;
  return Math.abs(r - g) <= 18 && Math.abs(g - b) <= 18 && Math.abs(r - b) <= 18 && r < 245 && r > 50;
}

/**
 * Core engine for configuring, restructuring, and resetting the Sermon Topics & Hymns spreadsheet:
 * 1. Renames tabs if they contain old year numbers (e.g. "Spring ESSC 2026" -> "Spring ESSC 2027").
 * 2. Restructures Q1-Q4 into the Universal 14-Row Architecture (Rows 1–65):
 *    - Mondays: Rows 4–17 (14 data rows)
 *    - Fridays: Rows 20–33 (14 data rows)
 *    - Saturday AM: Rows 36–49 (14 data rows)
 *    - Saturday PM: Rows 52–65 (14 data rows)
 *    - Blank separator rows at Rows 18, 34, 50
 * 3. Sets anchor dates and 14-row quarter-bounded self-terminating date formulas in Column A.
 * 4. Rewrites IMPORTRANGE formulas in Column B to pull speaker from Column AY of Holy Work Schedule:
 *    - Mondays: Q!AY6:AY19
 *    - Fridays: Q!AY51:AY64
 *    - Saturday AM: Q!AY66:AY79
 *    - Saturday PM: Q!AY81:AY94
 * 5. Clears all spill cells in Column B (B5:B17, B21:B33, B37:B49, B53:B65) preventing #REF! spill collision.
 * 6. Clears user content (Topic/Title, Hymns, Usher attendance) in columns C, D, E and F-I.
 * 7. Populates Column J with attendance totals formulas (=SUM(F{r}:I{r})).
 * 8. Clears residual rows beyond Row 65.
 * 9. Clears previous speakers (Column B), broken formulas (#REF!), sermon topics, and hymns in ESSC tabs
 *    (Spring/Fall ESSC) across all service rows (including Evangelical Services and Spiritual Convocation)
 *    while strictly preserving headers and cells with 'X' and gray backgrounds.
 */
function resetSermonTopicsSheet(year, targetSs, hwsUrl) {
  var ss = targetSs || SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  
  // 1. Rename active seasonal event tabs if they contain an old 4-digit year (e.g., Spring ESSC 2026 -> Spring ESSC 2027)
  allSheets.forEach(function(s) {
    var sName = s.getName();
    if (/histor|archive|log|template|config/i.test(sName)) return;
    if (/(?:Spring|Fall|Winter|Summer|ESSC|Convocation|Evangelical|Spiritual)/i.test(sName)) {
      var updatedName = sName.replace(/\b20\d\d\b/g, String(year));
      if (updatedName !== sName) {
        try { s.setName(updatedName); } catch (e) {
          console.warn("Could not rename tab " + sName + ": " + e.message);
        }
      }
    }
  });
  
  // Resolve clean Holy Work Schedule URL
  if (!hwsUrl) {
    try {
      hwsUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    } catch (e) {}
  }
  var hwsCleanUrl = "";
  if (hwsUrl) {
    hwsCleanUrl = hwsUrl.split('#')[0].split('?')[0];
    if (!hwsCleanUrl.endsWith('/edit')) hwsCleanUrl += '/edit';
  }

  // 2. Process Quarter Sheets (Q1, Q2, Q3, Q4) with Universal 14-Row Architecture
  var quarters = [
    { name: "Q1", month: 1 },
    { name: "Q2", month: 4 },
    { name: "Q3", month: 7 },
    { name: "Q4", month: 10 }
  ];

  var sections = [
    { name: "MONDAYS", hRow: 3, sRow: 4, eRow: 17, mask: "0111111", hwsRange: "AY6:AY19" },
    { name: "FRIDAYS", hRow: 19, sRow: 20, eRow: 33, mask: "1111011", hwsRange: "AY51:AY64" },
    { name: "SATURDAY AM'S", hRow: 35, sRow: 36, eRow: 49, mask: "1111101", hwsRange: "AY66:AY79" },
    { name: "SATURDAY PM'S", hRow: 51, sRow: 52, eRow: 65, mask: null, hwsRange: "AY81:AY94" }
  ];

  quarters.forEach(function(q) {
    var s = ss.getSheetByName(q.name);
    if (!s) return;
    
    // Ensure sheet has at least 65 rows and 10 columns
    if (s.getMaxRows() < 65) {
      s.insertRowsAfter(s.getMaxRows(), 65 - s.getMaxRows());
    }
    if (s.getMaxColumns() < 10) {
      s.insertColumnsAfter(s.getMaxColumns(), 10 - s.getMaxColumns());
    }
    
    // Set Anchor Date in Cell A1
    if (q.name === "Q1") {
      s.getRange("A1").setValue(new Date(year, 0, 1)).setNumberFormat("yyyy-MM-dd");
    } else {
      s.getRange("A1").setFormula("=DATE(YEAR('Q1'!$A$1), " + q.month + ", 1)").setNumberFormat("yyyy-MM-dd");
    }
    
    // Configure each 14-row section
    sections.forEach(function(sec) {
      var numRows = sec.eRow - sec.sRow + 1; // 14
      
      // A. Header Row (Rows 3, 19, 35, 51)
      var hRange = s.getRange(sec.hRow, 1, 1, 10);
      hRange.setBackground("#D9D9D9").setFontWeight("bold");
      s.getRange(sec.hRow, 1).setValue(sec.name).setHorizontalAlignment("center");
      s.getRange(sec.hRow, 2, 1, 9).clearContent();
      
      // B. Date Formulas in Column A
      var colAFormulas = [];
      if (sec.mask) {
        var w1Formula = '=WORKDAY.INTL(DATE(YEAR(A$1), MONTH(A$1), 1) - 1, 1, "' + sec.mask + '")';
        colAFormulas.push([w1Formula]);
        for (var r = sec.sRow + 1; r <= sec.eRow; r++) {
          var prevRow = r - 1;
          var weekFormula = '=IF(A' + prevRow + '="", "", IF(A' + prevRow + '+7 > EOMONTH(A$1, 2), "", A' + prevRow + '+7))';
          colAFormulas.push([weekFormula]);
        }
      } else {
        // Saturday PM mirrors Saturday AM (Rows 36 to 49)
        colAFormulas.push(['=IF(A36="","",A36)']);
        for (var r = sec.sRow + 1; r <= sec.eRow; r++) {
          var mirrorAmRow = 36 + (r - sec.sRow);
          colAFormulas.push(['=IF(A' + mirrorAmRow + '="", "", A' + mirrorAmRow + ')']);
        }
      }
      s.getRange(sec.sRow, 1, numRows, 1)
        .setFormulas(colAFormulas)
        .setNumberFormat("M/d/yyyy")
        .setHorizontalAlignment("center");
        
      // C. Column B: IMPORTRANGE formula at start row, spill cells cleared below
      if (hwsCleanUrl) {
        var formulaCell = s.getRange(sec.sRow, 2);
        var newFormula = '=IMPORTRANGE("' + hwsCleanUrl + '", "' + q.name + '!' + sec.hwsRange + '")';
        formulaCell.setFormula(newFormula);
      }
      s.getRange(sec.sRow + 1, 2, numRows - 1, 1).clearContent();
      
      // D. Columns C to I (Topics, Hymns, Attendance): Completely cleared
      s.getRange(sec.sRow, 3, numRows, 7).clearContent();
      
      // E. Column J: Attendance totals formula (=SUM(F:I))
      var colJFormulas = [];
      for (var r = sec.sRow; r <= sec.eRow; r++) {
        colJFormulas.push(['=SUM(F' + r + ':I' + r + ')']);
      }
      s.getRange(sec.sRow, 10, numRows, 1)
        .setFormulas(colJFormulas)
        .setNumberFormat("0")
        .setHorizontalAlignment("center");
        
      // F. Formatting across section data cells
      s.getRange(sec.sRow, 1, numRows, 10)
        .setBackground("#FFFFFF")
        .setBorder(true, true, true, true, true, true, "#B7B7B7", SpreadsheetApp.BorderStyle.SOLID);
    });
    
    // Clear separator rows: 18, 34, 50
    [18, 34, 50].forEach(function(sepRow) {
      s.getRange(sepRow, 1, 1, 10).clearContent().clearFormat();
    });
    
    // Clear residual legacy rows beyond Row 65
    var maxRows = s.getMaxRows();
    if (maxRows > 65) {
      s.getRange(66, 1, maxRows - 65, s.getMaxColumns()).clear();
    }
  });

  // 3. Process ESSC Sheets (e.g. Spring ESSC, Fall ESSC)
  allSheets.forEach(function(s) {
    var sName = s.getName();
    if (!/(?:essc|convocation)/i.test(sName)) return;
    
    var lastRow = s.getLastRow();
    var lastCol = Math.max(s.getLastColumn(), 6);
    if (lastRow < 4) return;
    
    // Process from Row 3 (EVANGELICAL SERVICES header) through lastRow
    var startRow = 3;
    var numRows = lastRow - startRow + 1;
    var range = s.getRange(startRow, 1, numRows, lastCol);
    var values = range.getValues();
    var backgrounds = range.getBackgrounds();
    
    var inEvangelical = false;
    var inConvocation = false;
    
    for (var r = 0; r < values.length; r++) {
      var colA = String(values[r][0] || "").trim().toUpperCase();
      
      // Detect Section Headers (Row 3, Row 8)
      if (colA.indexOf("EVANGELICAL") !== -1) {
        inEvangelical = true;
        inConvocation = false;
        continue;
      }
      if (colA.indexOf("CONVOCATION") !== -1) {
        inEvangelical = false;
        inConvocation = true;
        continue;
      }
      
      // Detect Day Separator Headers in Convocation (Row 9 "FRIDAY", Row 20 "SATURDAY", Row 32 "SUNDAY")
      if (inConvocation && /^(FRIDAY|SATURDAY|SUNDAY|WEDNESDAY|THURSDAY)$/i.test(colA)) {
        continue;
      }
      
      // If row is outside active service sections or is an empty spacer, skip
      if (!inEvangelical && !inConvocation) continue;
      if (!colA) continue;
      
      // True service/data row:
      // Column B is index 1 (Speaker)
      // Columns C-F are indices 2-5 (Topics, Sermon Hymns, Hymnleader Hymns, Choir Presentation)
      for (var c = 1; c < lastCol; c++) {
        var cellVal = String(values[r][c] || "").trim();
        var bg = String(backgrounds[r][c] || "").toLowerCase();
        
        var isX = (cellVal.toUpperCase() === "X");
        var isGray = isGrayBackground(bg);
        
        if (isX) {
          // Preserve 'X' and ensure standard gray background
          values[r][c] = "X";
          backgrounds[r][c] = "#D9D9D9";
        } else if (isGray && cellVal === "") {
          // Preserve intentional empty gray background cells
          backgrounds[r][c] = "#D9D9D9";
        } else {
          // Clear speaker info, broken formulas (#REF!), sermon topics, hymns, and choir presentations
          values[r][c] = "";
          backgrounds[r][c] = "#FFFFFF";
        }
      }
    }
    
    // Batch write cleaned values and backgrounds
    range.setValues(values);
    range.setBackgrounds(backgrounds);
  });
}

// -------------------------------------------------------------------------
// Annual Provisioner & Google Drive Folder Automation
// -------------------------------------------------------------------------

/**
 * Resolves or creates the target year's folder inside the parent "Holy Work Schedules" folder in Google Drive.
 */
function getOrCreateTargetYearFolder(targetYear) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentFile = DriveApp.getFileById(ss.getId());
  var currentParents = currentFile.getParents();
  var currentYearFolder = currentParents.hasNext() ? currentParents.next() : null;
  
  var parentSchedulesFolder = null;
  if (currentYearFolder && currentYearFolder.getParents().hasNext()) {
    var candidateParent = currentYearFolder.getParents().next();
    if (candidateParent.getName().toLowerCase().indexOf("holy work") !== -1) {
      parentSchedulesFolder = candidateParent;
    }
  }
  
  // Fallback to explicit ID from Drive screenshot if needed
  if (!parentSchedulesFolder) {
    try {
      parentSchedulesFolder = DriveApp.getFolderById(CONFIG.HOLY_WORK_SCHEDULES_FOLDER_ID);
    } catch (e) {
      console.warn("Could not access parent folder via ID: " + e.message);
      if (currentYearFolder) parentSchedulesFolder = currentYearFolder;
    }
  }
  
  if (!parentSchedulesFolder) {
    throw new Error("Could not locate the 'Holy Work Schedules' folder in Google Drive.");
  }
  
  // Check if target year folder (e.g., "2027") already exists
  var yearFolderName = String(targetYear).trim();
  var subFolders = parentSchedulesFolder.getFoldersByName(yearFolderName);
  if (subFolders.hasNext()) {
    return subFolders.next();
  } else {
    return parentSchedulesFolder.createFolder(yearFolderName);
  }
}

/**
 * Retrieves the linked Unavailability Spreadsheet ID with multi-tiered fallback.
 */
function getLinkedUnavailabilityId() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSsId = ss ? ss.getId() : "";
  
  function cleanAndValidateId(val) {
    if (!val || typeof val !== 'string') return null;
    val = val.trim();
    var m = val.match(/\/d\/([a-zA-Z0-9-_]+)/);
    var candidate = m ? m[1] : val;
    if (currentSsId && candidate === currentSsId) return null;
    if (!/^[a-zA-Z0-9-_]{20,60}$/.test(candidate)) return null;
    try {
      DriveApp.getFileById(candidate);
      return candidate;
    } catch (e) {
      return null;
    }
  }

  // 1. Check Script Properties
  var id = props.getProperty(CONFIG.UNAVAILABILITY_SHEET_PROPERTY);
  var validId = cleanAndValidateId(id);
  if (validId) return validId;
  if (id) props.deleteProperty(CONFIG.UNAVAILABILITY_SHEET_PROPERTY);
  
  // 2. Check cell E1 in internal Unavailability tab
  if (ss) {
    var internalTab = ss.getSheetByName(CONFIG.UNAVAILABILITY_SHEET);
    if (internalTab) {
      var cellVal = String(internalTab.getRange("E1").getValue() || "").trim();
      validId = cleanAndValidateId(cellVal);
      if (validId) {
        props.setProperty(CONFIG.UNAVAILABILITY_SHEET_PROPERTY, validId);
        return validId;
      }
    }
    
    // 3. Search within the current Google Drive folder
    try {
      var currentFile = DriveApp.getFileById(currentSsId);
      var currentParents = currentFile.getParents();
      if (currentParents.hasNext()) {
        var folder = currentParents.next();
        var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
        while (files.hasNext()) {
          var f = files.next();
          if (f.getName().toLowerCase().indexOf('unavailability') !== -1 && f.getId() !== currentSsId) {
            props.setProperty(CONFIG.UNAVAILABILITY_SHEET_PROPERTY, f.getId());
            return f.getId();
          }
        }
      }
    } catch (e) {
      console.warn("Drive auto-detection error for Unavailability: " + e.message);
    }
  }
  
  // 4. Fallback to CONFIG.DEFAULT_UNAVAILABILITY_ID
  if (CONFIG.DEFAULT_UNAVAILABILITY_ID) {
    validId = cleanAndValidateId(CONFIG.DEFAULT_UNAVAILABILITY_ID);
    if (validId) return validId;
  }
  
  return null;
}

/**
 * UI prompt to link or update the external Unavailability Spreadsheet ID/URL.
 */
function promptLinkUnavailabilitySheet() {
  var ui = SpreadsheetApp.getUi();
  var currentId = getLinkedUnavailabilityId() || '';
  
  var resp = ui.prompt(
    'Link External Unavailability Sheet',
    'Enter the Google Spreadsheet URL or File ID of the Annual Unavailability Sheet:\n\n(Current ID: ' + (currentId || 'None') + ')',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  
  var input = resp.getResponseText().trim();
  if (!input) {
    ui.alert('Cancelled', 'No URL or ID was provided.', ui.ButtonSet.OK);
    return;
  }
  
  var id = input;
  var match = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match) id = match[1];
  
  var localSs = SpreadsheetApp.getActiveSpreadsheet();
  if (id === localSs.getId()) {
    ui.alert('Invalid Sheet', 'You entered the URL or ID of the Holy Work Schedule itself. Please enter the Unavailability spreadsheet URL.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    var testSs = SpreadsheetApp.openById(id);
    PropertiesService.getScriptProperties().setProperty(CONFIG.UNAVAILABILITY_SHEET_PROPERTY, id);
    
    // Also record in internal Unavailability tab for transparency
    var internalTab = localSs.getSheetByName(CONFIG.UNAVAILABILITY_SHEET);
    if (internalTab) {
      internalTab.getRange("D1").setValue("External Spreadsheet ID:");
      internalTab.getRange("E1").setValue(id);
    }
    
    ui.alert('Linked Successfully! ✅', 'Successfully connected to external sheet:\n"' + testSs.getName() + '"', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Connection Failed', 'Could not open spreadsheet with ID: ' + id + '\n\nError: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Retrieves the linked Sermon Topics & Hymns Spreadsheet ID with multi-tiered fallback.
 */
function getLinkedSermonTopicsId() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSsId = ss ? ss.getId() : "";
  
  function cleanAndValidateId(val) {
    if (!val || typeof val !== 'string') return null;
    val = val.trim();
    var m = val.match(/\/d\/([a-zA-Z0-9-_]+)/);
    var candidate = m ? m[1] : val;
    if (currentSsId && candidate === currentSsId) return null;
    if (!/^[a-zA-Z0-9-_]{20,60}$/.test(candidate)) return null;
    try {
      DriveApp.getFileById(candidate);
      return candidate;
    } catch (e) {
      console.warn("Candidate Sermon Topics ID " + candidate + " is not accessible in DriveApp: " + e.message);
      return null;
    }
  }

  // 1. Check Script Properties
  var id = props.getProperty(CONFIG.SERMON_TOPICS_SHEET_PROPERTY);
  var validId = cleanAndValidateId(id);
  if (validId) return validId;
  if (id) props.deleteProperty(CONFIG.SERMON_TOPICS_SHEET_PROPERTY);

  var url = props.getProperty(CONFIG.SERMON_TOPICS_SHEET_URL_PROPERTY);
  validId = cleanAndValidateId(url);
  if (validId) {
    props.setProperty(CONFIG.SERMON_TOPICS_SHEET_PROPERTY, validId);
    return validId;
  }
  if (url) props.deleteProperty(CONFIG.SERMON_TOPICS_SHEET_URL_PROPERTY);

  // 2. Check metadata cells in internal Unavailability tab (cell E3 or E2)
  if (ss) {
    var internalTab = ss.getSheetByName(CONFIG.UNAVAILABILITY_SHEET);
    if (internalTab) {
      var cellId = String(internalTab.getRange("E3").getValue() || "").trim();
      validId = cleanAndValidateId(cellId);
      if (validId) {
        props.setProperty(CONFIG.SERMON_TOPICS_SHEET_PROPERTY, validId);
        return validId;
      }
      
      var cellUrl = String(internalTab.getRange("E2").getValue() || "").trim();
      validId = cleanAndValidateId(cellUrl);
      if (validId) {
        props.setProperty(CONFIG.SERMON_TOPICS_SHEET_PROPERTY, validId);
        props.setProperty(CONFIG.SERMON_TOPICS_SHEET_URL_PROPERTY, cellUrl);
        return validId;
      }
    }
    
    // 3. Search within current Google Drive folder for a file with "sermon topic"
    try {
      var currentFile = DriveApp.getFileById(currentSsId);
      var currentParents = currentFile.getParents();
      if (currentParents.hasNext()) {
        var folder = currentParents.next();
        var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
        while (files.hasNext()) {
          var f = files.next();
          if (f.getName().toLowerCase().indexOf('sermon topic') !== -1 && f.getId() !== currentSsId) {
            props.setProperty(CONFIG.SERMON_TOPICS_SHEET_PROPERTY, f.getId());
            props.setProperty(CONFIG.SERMON_TOPICS_SHEET_URL_PROPERTY, f.getUrl());
            return f.getId();
          }
        }
      }
    } catch (e) {
      console.warn("Drive auto-detection error for Sermon Topics: " + e.message);
    }
  }

  // 4. Default baseline template ID (2026 Sermon Topic & Hymns workbook)
  var defaultId = CONFIG.DEFAULT_SERMON_TOPICS_ID || '1zrMaPubQ7uaUvgBjYZ3lWIBRUclg0GsGliHMqpvOark';
  validId = cleanAndValidateId(defaultId);
  if (validId) {
    props.setProperty(CONFIG.SERMON_TOPICS_SHEET_PROPERTY, validId);
    return validId;
  }

  return defaultId;
}

/**
 * Retrieves the linked Sermon Topics & Hymns Spreadsheet URL with multi-tiered fallback.
 */
function getLinkedSermonTopicsUrl() {
  var props = PropertiesService.getScriptProperties();
  var url = props.getProperty(CONFIG.SERMON_TOPICS_SHEET_URL_PROPERTY);
  if (url && url.indexOf("docs.google.com/spreadsheets/d/") !== -1) return url.trim();
  
  var id = getLinkedSermonTopicsId();
  if (id && /^[a-zA-Z0-9-_]{20,60}$/.test(id)) {
    var constructedUrl = "https://docs.google.com/spreadsheets/d/" + id + "/edit?usp=sharing";
    props.setProperty(CONFIG.SERMON_TOPICS_SHEET_URL_PROPERTY, constructedUrl);
    return constructedUrl;
  }
  return "https://docs.google.com/spreadsheets/d/" + (CONFIG.DEFAULT_SERMON_TOPICS_ID || '1zrMaPubQ7uaUvgBjYZ3lWIBRUclg0GsGliHMqpvOark') + "/edit?usp=sharing";
}

/**
 * UI prompt to link or update the Sermon Topics & Hymns Spreadsheet ID/URL.
 */
function promptLinkSermonTopicsSheet() {
  var ui = SpreadsheetApp.getUi();
  var currentUrl = getLinkedSermonTopicsUrl() || '';
  
  var resp = ui.prompt(
    'Link Sermon Topics & Hymns Sheet',
    'Enter the Google Spreadsheet URL or File ID of the Annual Sermon Topics & Hymns Sheet:\n\n(Current URL: ' + (currentUrl || 'None') + ')',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  
  var input = resp.getResponseText().trim();
  if (!input) {
    ui.alert('Cancelled', 'No URL or ID was provided.', ui.ButtonSet.OK);
    return;
  }
  
  var id = input;
  var match = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match) id = match[1];
  
  var localSs = SpreadsheetApp.getActiveSpreadsheet();
  if (id === localSs.getId()) {
    ui.alert('Invalid Sheet', 'You entered the URL or ID of the Holy Work Schedule itself. Please enter the Sermon Topics & Hymns spreadsheet URL.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    var testSs = SpreadsheetApp.openById(id);
    var fullUrl = "https://docs.google.com/spreadsheets/d/" + id + "/edit?usp=sharing";
    PropertiesService.getScriptProperties().setProperty(CONFIG.SERMON_TOPICS_SHEET_PROPERTY, id);
    PropertiesService.getScriptProperties().setProperty(CONFIG.SERMON_TOPICS_SHEET_URL_PROPERTY, fullUrl);
    
    // Also record in internal Unavailability tab for transparency
    var internalTab = localSs.getSheetByName(CONFIG.UNAVAILABILITY_SHEET);
    if (internalTab) {
      internalTab.getRange("D2").setValue("Sermon Topics URL:");
      internalTab.getRange("E2").setValue(fullUrl);
      internalTab.getRange("D3").setValue("Sermon Topics ID:");
      internalTab.getRange("E3").setValue(id);
    }
    
    ui.alert('Linked Successfully! ✅', 'Successfully connected to Sermon Topics sheet:\n"' + testSs.getName() + '"', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Connection Failed', 'Could not open spreadsheet with ID: ' + id + '\n\nError: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Syncs the external Unavailability Matrix into the local internal Unavailability tab.
 */
function syncUnavailabilityFromMatrix() {
  return executeWithLockAndRetry_('syncUnavailabilityFromMatrix', function() {
    var localSs = SpreadsheetApp.getActiveSpreadsheet();
  var tz = localSs.getSpreadsheetTimeZone();
  var unavailId = getLinkedUnavailabilityId();
  
  if (!unavailId) {
    if (localSs) {
      localSs.toast("No linked Unavailability sheet. Use 'Link Unavailability Sheet' first.", "Sync Aborted", 8);
    }
    return;
  }
  
  var unavailSs;
  try {
    unavailSs = SpreadsheetApp.openById(unavailId);
  } catch (e) {
    if (localSs) {
      localSs.toast("Error opening external Unavailability sheet: " + e.message, "Sync Failed", 10);
    }
    return;
  }
  
  var matrixSheet = unavailSs.getSheetByName("Unavailability") || unavailSs.getSheets()[0];
  if (!matrixSheet) return;
  
  var lastRow = matrixSheet.getLastRow();
  var lastCol = matrixSheet.getLastColumn();
  if (lastRow < 4 || lastCol < 2) {
    if (localSs) localSs.toast("Unavailability sheet has no data to sync.", "Sync Complete", 5);
    return;
  }
  
  // Auto-heal: If B2 contains the year date (e.g. 1/1/2027), move it to A2 and restore B2 to "Q1"
  var b2Val = matrixSheet.getRange("B2").getValue();
  if (b2Val && (b2Val instanceof Date || String(b2Val).indexOf("20") !== -1)) {
    matrixSheet.getRange("A2").setValue(b2Val);
    matrixSheet.getRange("B2").setValue("Q1");
  }
  
  // Scan up to column 57 (Col A + 56 date cols B:BE)
  var totalColsToScan = Math.min(lastCol, 57);
  var fullData = matrixSheet.getRange(1, 1, lastRow, totalColsToScan).getValues();
  
  var unavailableEntries = [];
  var currentSectionDates = null;
  var currentSection = "";
  
  for (var r = 0; r < fullData.length; r++) {
    var rawName = String(fullData[r][0] || "").trim();
    if (!rawName) continue;
    
    // Detect section headers in Column A (e.g., "Monday", "Friday", "Saturday AM", "Saturday PM", etc.)
    var upperName = rawName.toUpperCase();
    if (upperName === "MONDAY" || upperName === "FRIDAY" || upperName.indexOf("SATURDAY") !== -1 || upperName.indexOf("SABBATH") !== -1) {
      currentSection = rawName;
      currentSectionDates = fullData[r]; // Row r contains this section's specific date row!
      continue;
    }
    
    // Skip general header/metadata rows
    if (upperName === "Q1" || upperName === "DATE" || r < 3) continue;
    
    // Process worker row using active section dates
    var activeDates = currentSectionDates || (fullData.length > 2 ? fullData[2] : null);
    if (!activeDates) continue;
    
    for (var c = 1; c < totalColsToScan; c++) {
      var cellVal = fullData[r][c];
      var isUnavailable = (cellVal === true || String(cellVal).trim().toUpperCase() === "TRUE");
      
      if (isUnavailable) {
        var rawDate = activeDates[c];
        if (rawDate && rawDate !== "") {
          var dateObj = rawDate instanceof Date ? rawDate : new Date(rawDate);
          if (!isNaN(dateObj.getTime())) {
            var reason = currentSection ? "Unavailable (" + currentSection + ")" : "Marked Unavailable";
            unavailableEntries.push([dateObj, rawName, reason]);
          }
        }
      }
    }
  }
  
  // --- Dedicated ESSC Unavailability Tabs Sync (Spring ESSC & Fall ESSC) ---
  var esscTabs = [
    { name: 'Spring ESSC', season: 'Spring', configKey: 'SPRING_WEDNESDAY_START_DATE' },
    { name: 'Fall ESSC',   season: 'Fall',   configKey: 'FALL_WEDNESDAY_START_DATE' }
  ];
  
  for (var t = 0; t < esscTabs.length; t++) {
    var tabInfo = esscTabs[t];
    var tabSheet = unavailSs.getSheetByName(tabInfo.name);
    if (!tabSheet) continue;
    
    var tabLastRow = tabSheet.getLastRow();
    var tabLastCol = tabSheet.getLastColumn();
    if (tabLastRow < 4 || tabLastCol < 2) continue;
    
    // Cell B2: Event start date (Wednesday anchor date), fallback to MasterConfig
    var rawB2 = tabSheet.getRange("B2").getValue();
    var anchorDate = parseAnchorDate(rawB2);
    if (!anchorDate) {
      var fallbackDate = getMasterConfigStartDate(localSs, tabInfo.configKey);
      anchorDate = parseAnchorDate(fallbackDate);
    }
    if (!anchorDate) {
      console.warn("Could not determine anchor date for " + tabInfo.name + ", skipping.");
      continue;
    }
    
    // Row 3 Headers: Col A: Worker Name, Cols B-L: 11 shift block keys
    var headerRowData = tabSheet.getRange(3, 1, 1, Math.min(tabLastCol, 12)).getValues()[0];
    var colBlockMap = {};
    for (var c = 1; c < headerRowData.length; c++) {
      var hKey = String(headerRowData[c] || '').trim().toUpperCase();
      var matchedBlock = null;
      for (var b = 0; b < ESSC_BLOCK_DEFINITIONS.length; b++) {
        if (ESSC_BLOCK_DEFINITIONS[b].key.toUpperCase() === hKey ||
            ESSC_BLOCK_DEFINITIONS[b].label.toUpperCase() === hKey) {
          matchedBlock = ESSC_BLOCK_DEFINITIONS[b];
          break;
        }
      }
      if (!matchedBlock && (c - 1) < ESSC_BLOCK_DEFINITIONS.length) {
        matchedBlock = ESSC_BLOCK_DEFINITIONS[c - 1];
      }
      if (matchedBlock) {
        colBlockMap[c] = matchedBlock;
      }
    }
    
    // Row 2 Dates: Read dates across columns B:L if available
    var row2Dates = tabSheet.getRange(2, 1, 1, Math.min(tabLastCol, 12)).getValues()[0];

    // Rows 4+: Active rostered workers in Col A with checkboxes in Cols B:L
    var numDataRows = tabLastRow - 3;
    var tabData = tabSheet.getRange(4, 1, numDataRows, Math.min(tabLastCol, 12)).getValues();
    
    for (var r = 0; r < tabData.length; r++) {
      var workerName = String(tabData[r][0] || '').trim();
      if (!workerName) continue;
      
      for (var c = 1; c < tabData[r].length; c++) {
        var isChecked = (tabData[r][c] === true || String(tabData[r][c]).trim().toUpperCase() === "TRUE");
        if (isChecked && colBlockMap[c]) {
          var blockDef = colBlockMap[c];
          var cellDate = parseAnchorDate(row2Dates[c]);
          var sessionDate = cellDate || new Date(anchorDate.getFullYear(), anchorDate.getMonth(), anchorDate.getDate() + blockDef.dayOffset, 12, 0, 0);
          var reason = 'Unavailable (' + tabInfo.season + ' ESSC: ' + blockDef.label + ')';
          unavailableEntries.push([sessionDate, workerName, reason]);
        }
      }
    }
  }
  
  // Write to internal Unavailability tab safely without wiping metadata
  var internalSheet = localSs.getSheetByName(CONFIG.UNAVAILABILITY_SHEET);
  if (!internalSheet) {
    internalSheet = localSs.insertSheet(CONFIG.UNAVAILABILITY_SHEET);
  }
  
  var maxR = internalSheet.getMaxRows();
  if (maxR > 1) {
    internalSheet.getRange(2, 1, maxR - 1, 3).clearContent();
  }
  
  // Set headers & metadata (preserves D2:E5 external ID/URL storage)
  internalSheet.getRange("A1:C1").setValues([["Date", "Worker Name", "Reason"]])
    .setFontWeight("bold")
    .setBackground("#E0E7FF");
    
  internalSheet.getRange("D1:E1").setValues([["External Spreadsheet ID:", unavailId]])
    .setFontColor("#777777");
  internalSheet.setFrozenRows(1);
  
  if (unavailableEntries.length > 0) {
    var neededRows = unavailableEntries.length + 1; // 1 header row + data rows
    var currentRows = internalSheet.getMaxRows();
    if (neededRows > currentRows) {
      internalSheet.insertRowsAfter(currentRows, neededRows - currentRows);
    }
    var writeRange = internalSheet.getRange(2, 1, unavailableEntries.length, 3);
    writeRange.setValues(unavailableEntries);
    internalSheet.getRange(2, 1, unavailableEntries.length, 1).setNumberFormat("yyyy-MM-dd");
    internalSheet.autoResizeColumns(1, 3);
  }
  
  // Clear cache so sidebar updates immediately
  cacheRemoveLarge(CacheService.getScriptCache(), 'unavailability_data');
  bumpAllQuarterVersions();
  
  var statusMsg = "Successfully synced " + unavailableEntries.length + " unavailable assignments from external matrix (including ESSC tabs).";
  localSs.toast(statusMsg, "Unavailability Synced", 6);
  return { success: true, count: unavailableEntries.length };
  });
}

/**
 * Automated Provisioning / Sync Helper Function for ESSC Unavailability Tabs.
 * Ensures 'Spring ESSC' and/or 'Fall ESSC' tabs exist in the external Unavailability workbook,
 * sets up headers and anchor dates, populates deduplicated active rostered workers in Column A,
 * and sets up native checkbox validations in Columns B:L.
 *
 * @param {string} [targetSeason] 'Spring', 'Fall', or undefined/null for both.
 */
function syncESSCUnavailabilityRosters(targetSeason) {
  var localSs = SpreadsheetApp.getActiveSpreadsheet();
  var unavailId = getLinkedUnavailabilityId();
  if (!unavailId) {
    if (localSs) localSs.toast("No linked Unavailability spreadsheet found. Use 'Link Unavailability Sheet' first.", "Sync Aborted", 8);
    throw new Error("No linked Unavailability spreadsheet found.");
  }
  
  var unavailSs;
  try {
    unavailSs = SpreadsheetApp.openById(unavailId);
  } catch (e) {
    if (localSs) localSs.toast("Error opening external Unavailability sheet: " + e.message, "Sync Failed", 10);
    throw e;
  }
  
  // 1. Determine target seasons
  var seasonsToSync = [];
  var sTarget = String(targetSeason || '').trim().toLowerCase();
  if (sTarget.indexOf('spring') !== -1) {
    seasonsToSync = ['Spring'];
  } else if (sTarget.indexOf('fall') !== -1) {
    seasonsToSync = ['Fall'];
  } else {
    seasonsToSync = ['Spring', 'Fall'];
  }
  
  // 2. Extract active rostered workers from master Workers sheet
  var workerSheet = localSs.getSheetByName(CONFIG.WORKERS_SHEET);
  if (!workerSheet) {
    try {
      if (typeof getMasterHWSSpreadsheet === 'function') {
        var masterSs = getMasterHWSSpreadsheet();
        if (masterSs) workerSheet = masterSs.getSheetByName(CONFIG.WORKERS_SHEET);
      }
    } catch(e) {}
  }
  if (!workerSheet) throw new Error("Master 'Workers' sheet not found.");
  
  var workerData = workerSheet.getDataRange().getValues();
  if (workerData.length < 2) throw new Error("Workers sheet contains no data.");
  
  var detected = detectWorkersHeaderRow_(workerData);
  var headers = detected.headers.map(function(h) { return String(h || "").trim().toLowerCase(); });
  var nameCol = -1, activeCol = -1, rosterCol = -1, churchCol = -1;
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (nameCol === -1 && (h.indexOf("schedule name") !== -1 || h === "name")) nameCol = i;
    if (activeCol === -1 && h === "active") activeCol = i;
    if (rosterCol === -1 && h.indexOf("roster status") !== -1) rosterCol = i;
    if (churchCol === -1 && (h.indexOf("resident church") !== -1 || h === "church")) churchCol = i;
  }
  if (nameCol === -1) nameCol = 4;
  if (activeCol === -1) activeCol = 11;
  if (rosterCol === -1) rosterCol = 12;
  if (churchCol === -1) churchCol = 9; // Default Column J (Resident Church)
  
  var rosteredSet = {};
  var rosteredWorkers = [];
  
  for (var r = detected.dataStartIdx; r < workerData.length; r++) {
    var row = workerData[r];
    var name = String(row[nameCol] || "").trim();
    if (!name) {
      var first = String(row[0] || "").trim();
      var last = String(row[1] || "").trim();
      name = (first + " " + last).trim();
    }
    if (!name) continue;
    
    // Rule 1: Active members only
    var isActive = String(row[activeCol] || "").trim().toLowerCase() === "active";
    if (!isActive) continue;
    
    // Rule 2: Only active Garden Grove members (resident church in Column J must be blank)
    var residentChurch = String(row[churchCol] || "").trim();
    if (residentChurch !== "") continue;
    
    // Rule 3: Rostered status check (or role assignments)
    var statusStr = String(row[rosterCol] || "").trim().toLowerCase();
    var isRostered = (statusStr.indexOf("rostered") !== -1 && statusStr.indexOf("not rostered") === -1);
    
    if (!isRostered && statusStr === "") {
      var roleCnt = 0;
      for (var c = 14; c < Math.min(row.length, 53); c++) {
        if (row[c] === true || String(row[c]).toLowerCase() === "true") roleCnt++;
      }
      if (roleCnt > 0) isRostered = true;
    }
    
    if (isRostered && !rosteredSet[name]) {
      rosteredSet[name] = true;
      rosteredWorkers.push(name);
    }
  }
  
  rosteredWorkers.sort(function(a, b) { return a.localeCompare(b); });
  if (rosteredWorkers.length === 0) throw new Error("No active, rostered Garden Grove members found in Workers sheet.");
  
  var checkboxValidation = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  var syncedCount = 0;
  
  for (var sIdx = 0; sIdx < seasonsToSync.length; sIdx++) {
    var season = seasonsToSync[sIdx];
    var tabName = season + ' ESSC';
    var tabSheet = unavailSs.getSheetByName(tabName);
    if (!tabSheet) {
      tabSheet = unavailSs.insertSheet(tabName);
    }
    
    // Read existing checked states from Row 4+ if present to preserve user inputs
    var existingChecks = {};
    if (tabSheet.getLastRow() >= 4 && tabSheet.getLastColumn() >= 2) {
      var oldLastRow = tabSheet.getLastRow();
      var oldLastCol = Math.min(tabSheet.getLastColumn(), 12);
      var oldHeaders = tabSheet.getRange(3, 1, 1, oldLastCol).getValues()[0];
      var oldData = tabSheet.getRange(4, 1, oldLastRow - 3, oldLastCol).getValues();
      
      for (var or = 0; or < oldData.length; or++) {
        var oName = String(oldData[or][0] || '').trim();
        if (!oName) continue;
        if (!existingChecks[oName]) existingChecks[oName] = {};
        for (var oc = 1; oc < oldData[or].length; oc++) {
          var blockKey = String(oldHeaders[oc] || '').trim().toUpperCase();
          var isChecked = (oldData[or][oc] === true || String(oldData[or][oc]).toUpperCase() === 'TRUE');
          if (blockKey && isChecked) {
            existingChecks[oName][blockKey] = true;
          }
        }
      }
    }
    
    // Row 1: Title
    tabSheet.getRange("A1").setValue(season + " ESSC Worker Unavailability")
      .setFontWeight("bold")
      .setFontSize(12)
      .setBackground("#1F4E79")
      .setFontColor("#FFFFFF");
    tabSheet.getRange("B1:L1").setBackground("#1F4E79");
    
    // Row 2: Date formula/anchor across all 11 shift blocks (Cols B:L)
    tabSheet.getRange("A2").setValue("Date")
      .setFontWeight("bold")
      .setFontColor("#1F4E79")
      .setHorizontalAlignment("left")
      .setBackground("#F8FAFC");
    var currentB2 = tabSheet.getRange("B2").getValue();
    var parsedAnchor = parseAnchorDate(currentB2);
    if (!parsedAnchor) {
      var mcKey = (season === 'Spring') ? 'SPRING_WEDNESDAY_START_DATE' : 'FALL_WEDNESDAY_START_DATE';
      var fallbackDate = getMasterConfigStartDate(localSs, mcKey);
      parsedAnchor = parseAnchorDate(fallbackDate);
      if (parsedAnchor) {
        tabSheet.getRange("B2").setValue(parsedAnchor);
      } else if (fallbackDate) {
        tabSheet.getRange("B2").setValue(fallbackDate);
      }
    }
    
    // Dynamic date formulas for each shift block in Cols C through L (dayOffsets 1 to 4)
    var dateFormulas = [];
    for (var b = 1; b < ESSC_BLOCK_DEFINITIONS.length; b++) {
      var offset = ESSC_BLOCK_DEFINITIONS[b].dayOffset;
      dateFormulas.push('=IF(ISBLANK($B$2), "", IF(ISNUMBER($B$2), $B$2 + ' + offset + ', DATEVALUE(TO_TEXT($B$2)) + ' + offset + '))');
    }
    tabSheet.getRange(2, 3, 1, dateFormulas.length).setFormulas([dateFormulas]);
    
    // Format all of B2:L2 with yyyy-MM-dd date format, bold, center aligned
    tabSheet.getRange(2, 2, 1, 11)
      .setNumberFormat("yyyy-MM-dd")
      .setFontWeight("bold")
      .setFontColor("#1F4E79")
      .setHorizontalAlignment("center")
      .setBackground("#F8FAFC");
    
    // Row 3: Headers (Col A: Worker Name, Cols B-L: 11 shift block keys)
    var headerRow = ['Worker Name'].concat(ESSC_BLOCK_DEFINITIONS.map(function(b) { return b.key; }));
    tabSheet.getRange(3, 1, 1, 12).setValues([headerRow])
      .setFontWeight("bold")
      .setBackground("#E0E7FF")
      .setFontColor("#1F4E79")
      .setHorizontalAlignment("center");
    tabSheet.getRange("A3").setHorizontalAlignment("left");
    
    // Build Rows 4+
    var newRows = [];
    for (var w = 0; w < rosteredWorkers.length; w++) {
      var workerName = rosteredWorkers[w];
      var rowVals = [workerName];
      for (var b = 0; b < ESSC_BLOCK_DEFINITIONS.length; b++) {
        var bKey = ESSC_BLOCK_DEFINITIONS[b].key.toUpperCase();
        var isChecked = !!(existingChecks[workerName] && existingChecks[workerName][bKey]);
        rowVals.push(isChecked);
      }
      newRows.push(rowVals);
    }
    
    // Clear old data rows if previous roster had more rows
    var prevMaxRows = tabSheet.getLastRow();
    if (prevMaxRows > 3) {
      tabSheet.getRange(4, 1, prevMaxRows - 3, Math.max(tabSheet.getLastColumn(), 12)).clearContent().clearDataValidations();
    }
    
    // Write new roster and values
    var writeRange = tabSheet.getRange(4, 1, newRows.length, 12);
    writeRange.setValues(newRows);
    
    // Inject native checkbox data validations across Columns B:L (cols 2-12)
    tabSheet.getRange(4, 2, newRows.length, 11).setDataValidation(checkboxValidation);
    tabSheet.getRange(4, 1, newRows.length, 1).setFontWeight("bold").setHorizontalAlignment("left");
    tabSheet.getRange(4, 2, newRows.length, 11).setHorizontalAlignment("center");
    
    tabSheet.setFrozenRows(3);
    tabSheet.setFrozenColumns(1);
    tabSheet.autoResizeColumns(1, 12);
    
    syncedCount++;
  }
  
  if (localSs) {
    localSs.toast("Successfully provisioned/synced " + syncedCount + " ESSC tab(s) with " + rosteredWorkers.length + " active rostered workers.", "ESSC Sync Complete", 6);
  }
  return { success: true, seasons: seasonsToSync, workerCount: rosteredWorkers.length };
}

/**
 * UI prompt to trigger ESSC Unavailability Roster provisioning/syncing.
 */
function promptSyncESSCUnavailabilityRosters() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    "Sync ESSC Unavailability Rosters",
    "Do you want to provision and sync rosters for both Spring & Fall ESSC tabs in the external Unavailability workbook?\n\n(Click YES for Both, NO to choose a specific season, or CANCEL to abort)",
    ui.ButtonSet.YES_NO_CANCEL
  );
  if (resp === ui.Button.CANCEL) return;
  if (resp === ui.Button.YES) {
    syncESSCUnavailabilityRosters();
  } else if (resp === ui.Button.NO) {
    var seasonResp = ui.prompt(
      "Choose ESSC Season",
      "Enter 'Spring' or 'Fall':",
      ui.ButtonSet.OK_CANCEL
    );
    if (seasonResp.getSelectedButton() === ui.Button.OK) {
      var s = seasonResp.getResponseText().trim();
      syncESSCUnavailabilityRosters(s);
    }
  }
}

/**
 * Activates all background automation triggers for this Holy Work Schedule.
 */
function activateAllTriggers() {
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch(e) {}

  var triggers = ScriptApp.getProjectTriggers();
  var managedFunctions = ['sendAssignmentReminders', 'syncSWRSpeakers', 'syncUnavailabilityFromMatrix', 'autoSyncRESchedules', 'autoSyncScheduledMonths', 'generateAnnualSummary', 'sendMonthlyUnavailabilityReminder'];
  
  // Delete existing triggers for these functions to avoid duplicates
  for (var i = 0; i < triggers.length; i++) {
    if (managedFunctions.indexOf(triggers[i].getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // 1. Daily Reminder Emails: 8:00 AM - 9:00 AM (5 days prior to Mon/Fri/Sat services)
  ScriptApp.newTrigger('sendAssignmentReminders')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
    
  // 2. SWR Regional Speaker Sync: 1:00 AM - 2:00 AM
  ScriptApp.newTrigger('syncSWRSpeakers')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();
    
  // 3. Unavailability Matrix Sync: 2:00 AM - 3:00 AM
  ScriptApp.newTrigger('syncUnavailabilityFromMatrix')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
    
  // 4. Religious Education (RE) Schedule Sync: 3:00 AM - 4:00 AM
  ScriptApp.newTrigger('autoSyncRESchedules')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();

  // 5. Automated Monthly Schedule Sync (Active & Next Month): 4:00 AM - 5:00 AM
  ScriptApp.newTrigger('autoSyncScheduledMonths')
    .timeBased()
    .everyDays(1)
    .atHour(4)
    .create();

  // 6. Monthly Annual Summary Update: 1st of every month at 5:00 AM - 6:00 AM
  ScriptApp.newTrigger('generateAnnualSummary')
    .timeBased()
    .onMonthDay(1)
    .atHour(5)
    .create();

  // 7. Monthly Unavailability Email Reminder: 1st of every month at 10:00 AM - 11:00 AM (avoids collision with 8:00 AM assignment reminders)
  ScriptApp.newTrigger('sendMonthlyUnavailabilityReminder')
    .timeBased()
    .onMonthDay(1)
    .atHour(10)
    .create();

  // 8. Sync Coordinator Names from Coordinators config tab to Row 3 of Q1-Q4
  try {
    if (typeof syncCoordinatorsToQuarterSheets === 'function') {
      syncCoordinatorsToQuarterSheets(SpreadsheetApp.getActiveSpreadsheet());
    }
  } catch (e) {
    console.warn("Could not sync coordinators in activateAllTriggers: " + e.message);
  }

  // 9. Cache linked Sermon Topics Sheet URL & ID into Script Properties
  try {
    var sermonUrl = getLinkedSermonTopicsUrl();
    if (sermonUrl) {
      PropertiesService.getScriptProperties().setProperty(CONFIG.SERMON_TOPICS_SHEET_URL_PROPERTY, sermonUrl);
    }
    var sermonId = getLinkedSermonTopicsId();
    if (sermonId) {
      PropertiesService.getScriptProperties().setProperty(CONFIG.SERMON_TOPICS_SHEET_PROPERTY, sermonId);
    }
  } catch (e) {
    console.warn("Could not register Sermon Topics properties in activateAllTriggers: " + e.message);
  }
    
  if (ui) {
    ui.alert(
      'Automation Triggers Activated! ✅',
      'The following automated background jobs have been scheduled for this spreadsheet:\n\n' +
      '1. 📧 Assignment Email Reminders: Daily at 8:00 AM (5 days prior)\n' +
      '2. 🔄 SWR Regional Speaker Sync: Daily at 1:00 AM\n' +
      '3. 📋 Unavailability Matrix Sync: Daily at 2:00 AM\n' +
      '4. 📘 RE Schedule Sync: Daily at 3:00 AM\n' +
      '5. 📅 Monthly Schedule Auto-Sync: Daily at 4:00 AM (Active & Next Month)\n' +
      '6. 📊 Annual Summary Report Update: Monthly on the 1st at 5:00 AM\n' +
      '7. 📋 Monthly Unavailability Email Reminder: Monthly on the 1st at 8:00 AM\n\n' +
      'Your Holy Work Schedule is fully automated!',
      ui.ButtonSet.OK
    );
  }
}

/**
 * UI Prompt for the One-Click Annual Provisioner.
 */
function promptProvisionNewYear() {
  var ui = SpreadsheetApp.getUi();
  var currentYear = new Date().getFullYear();
  var defaultNextYear = currentYear + 1;
  
  var resp = ui.prompt(
    '🚀 Provision New Year Suite (Schedule, Unavailability & Sermon Topics)',
    'Please enter the 4-digit NEW YEAR you want to provision (e.g., ' + defaultNextYear + '):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  
  var newYearStr = resp.getResponseText().trim();
  var newYear = parseInt(newYearStr, 10);
  if (isNaN(newYear) || newYear < 2020 || newYear > 2100) {
    ui.alert('Invalid Input', 'Please enter a valid 4-digit year (e.g., ' + defaultNextYear + ').', ui.ButtonSet.OK);
    return;
  }
  
  var unavailId = getLinkedUnavailabilityId();
  if (!unavailId) {
    var unavailResp = ui.prompt(
      'Unavailability Sheet Required',
      'Could not automatically detect the linked Unavailability sheet.\nPlease paste the URL or ID of the current Unavailability Google Sheet:',
      ui.ButtonSet.OK_CANCEL
    );
    if (unavailResp.getSelectedButton() !== ui.Button.OK) return;
    var unavailInput = unavailResp.getResponseText().trim();
    var match = unavailInput.match(/\/d\/([a-zA-Z0-9-_]+)/);
    unavailId = match ? match[1] : unavailInput;
  }
  
  var sermonTopicsId = getLinkedSermonTopicsId();
  if (!sermonTopicsId) {
    var sermonResp = ui.prompt(
      'Sermon Topics & Hymns Sheet Required',
      'Could not automatically detect the Sermon Topics & Hymns sheet.\nPlease paste the URL or ID of the current Sermon Topics & Hymns Google Sheet:',
      ui.ButtonSet.OK_CANCEL
    );
    if (sermonResp.getSelectedButton() !== ui.Button.OK) return;
    var sermonInput = sermonResp.getResponseText().trim();
    var matchSermon = sermonInput.match(/\/d\/([a-zA-Z0-9-_]+)/);
    sermonTopicsId = matchSermon ? matchSermon[1] : sermonInput;
  }
  
  var confirm = ui.alert(
    'Confirm Provisioning for ' + newYear,
    'This will automatically:\n' +
    '1. Locate or create the "' + newYear + '" folder in Holy Work Schedules.\n' +
    '2. Duplicate Holy Work Schedule into "' + newYear + '" folder & reset Q1-Q4.\n' +
    '3. Duplicate Unavailability into "' + newYear + '" folder, set permissions so ANYONE WITH LINK CAN EDIT, set A2 to 1/1/' + newYear + ', set B2 to "Q1", clear checkboxes, and sync workers.\n' +
    '4. Duplicate Sermon Topics & Hymns into "' + newYear + '" folder, set permissions so ANYONE WITH LINK CAN EDIT, update IMPORTRANGE formulas, and clear past topics/hymns.\n' +
    '5. Link all three files together and configure dynamic email reminder URLs.\n\n' +
    'Do you want to proceed?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Provisioning ' + newYear + ' files in Google Drive...', 'Please wait', 15);
    var result = provisionNewYear(newYear, unavailId, sermonTopicsId);
    
    var modalShown = false;
    try {
      var html = HtmlService.createHtmlOutput(
        '<div style="font-family: -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, Helvetica, Arial, sans-serif; padding: 18px; line-height: 1.6; color: #333;">' +
        '<h2 style="color: ' + (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.THEME ? APP_CONFIG.THEME.PRIMARY_COLOR : '#1F4E79') + '; margin-top: 0; font-size: 20px;">🎉 ' + newYear + ' Suite Successfully Provisioned!</h2>' +
        '<p>All three files have been created, reset, and linked inside the <b>' + newYear + '</b> Google Drive folder:</p>' +
        '<div style="background: #F0F4F8; border-left: 4px solid #2E75B6; padding: 12px; margin: 15px 0; border-radius: 4px;">' +
        '<p style="margin: 4px 0;">📁 <b>Year Folder:</b> <a href="' + result.folderUrl + '" target="_blank" style="color: #2E75B6; font-weight: bold;">Open ' + newYear + ' Folder</a></p>' +
        '<p style="margin: 4px 0;">📅 <b>Holy Work Schedule:</b> <a href="' + result.scheduleUrl + '" target="_blank" style="color: #2E75B6; font-weight: bold;">Open ' + newYear + ' Schedule</a></p>' +
        '<p style="margin: 4px 0;">📋 <b>Unavailability Sheet:</b> <a href="' + result.unavailUrl + '" target="_blank" style="color: #2E75B6; font-weight: bold;">Open ' + newYear + ' Unavailability</a> <span style="font-size: 11px; background: #E2F0D9; color: #276A3C; padding: 2px 6px; border-radius: 3px; font-weight: bold;">Anyone with link can edit ✅</span></p>' +
        (result.sermonTopicsUrl ? '<p style="margin: 4px 0;">📖 <b>Sermon Topics & Hymns:</b> <a href="' + result.sermonTopicsUrl + '" target="_blank" style="color: #2E75B6; font-weight: bold;">Open ' + newYear + ' Sermon Topics</a> <span style="font-size: 11px; background: #E2F0D9; color: #276A3C; padding: 2px 6px; border-radius: 3px; font-weight: bold;">Anyone with link can edit ✅</span></p>' : '<p style="margin: 4px 0;">📖 <b>Sermon Topics & Hymns:</b> <span style="color: #C00000; font-weight: bold;">Not provisioned</span>' + (result.sermonTopicsError ? ' <span style="color: #C00000; font-size: 12px;">(' + result.sermonTopicsError + ')</span>' : ' <span style="color: #666; font-size: 12px;">(No source sheet linked; link via Admin Tools)</span>') + '</p>') +
        '</div>' +
        '<p style="font-size: 13px; color: #555;"><b>Next Steps for ' + newYear + ':</b></p>' +
        '<ol style="font-size: 13px; color: #555; padding-left: 20px;">' +
        '<li>Open the new <b>' + newYear + ' Holy Work Schedule</b> and click <code>⚙️ Admin Tools > ⏰ Automation > Activate All Automation Triggers</code>.</li>' +
        '<li>Open the new <b>' + newYear + ' Sermon Topic & Hymns</b> sheet, click cell <b>B4</b>, and click <b>"Allow access"</b> if prompted by Google Sheets to authorize the IMPORTRANGE link.</li>' +
        '<li>When the regional committee publishes the new <b>' + newYear + ' SWR Speaker Google Sheet</b>, click <code>⚙️ Admin Tools > ⏰ Automation > Setup Daily SWR Sync URL</code>.</li>' +
        '</ol>' +
        '<button onclick="google.script.host.close()" style="background: ' + (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.THEME ? APP_CONFIG.THEME.PRIMARY_COLOR : '#1F4E79') + '; color: white; border: none; padding: 8px 18px; border-radius: 4px; font-weight: 500; cursor: pointer; margin-top: 10px;">Close Window</button>' +
        '</div>'
      ).setWidth(580).setHeight(470);
      
      ui.showModalDialog(html, 'Provisioning Complete');
      modalShown = true;
    } catch (dialogErr) {
      console.warn("Could not display HTML modal dialog, falling back to native alert: " + dialogErr.message);
    }
    
    if (!modalShown) {
      ui.alert(
        '🎉 ' + newYear + ' Suite Successfully Provisioned!',
        'All three files have been provisioned and linked inside the "' + newYear + '" Google Drive folder:\n\n' +
        '1. 📁 Year Folder:\n' + result.folderUrl + '\n\n' +
        '2. 📅 Holy Work Schedule:\n' + result.scheduleUrl + '\n\n' +
        '3. 📋 Unavailability Sheet: (Anyone with link can edit ✅)\n' + result.unavailUrl + '\n\n' +
        '4. 📖 Sermon Topics & Hymns: (Anyone with link can edit ✅)\n' + (result.sermonTopicsUrl || ('Not provisioned ' + (result.sermonTopicsError ? '(' + result.sermonTopicsError + ')' : ''))) + '\n\n' +
        'Next Steps for ' + newYear + ':\n' +
        '• Open the new ' + newYear + ' Holy Work Schedule and click Admin Tools > Automation > Activate All Automation Triggers.\n' +
        '• Open the new ' + newYear + ' Sermon Topics sheet, click cell B4, and click "Allow access" if prompted.\n' +
        '• When regional SWR URL is published, run Admin Tools > Automation > Setup Daily SWR Sync URL.',
        ui.ButtonSet.OK
      );
    }
  } catch (err) {
    ui.alert('Provisioning Error', 'An error occurred during provisioning:\n\n' + err.message, ui.ButtonSet.OK);
    console.error(err);
  }
}

/**
 * Core engine for provisioning a new year: clones schedule, unavailability & sermon topics,
 * resets dates, clears past data, updates IMPORTRANGE formulas, and sets public edit sharing.
 */
function provisionNewYear(year, sourceUnavailId, sourceSermonTopicsId) {
  return executeWithLockAndRetry_('provisionNewYear', function() {
    var localSs = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!sourceSermonTopicsId) {
    sourceSermonTopicsId = getLinkedSermonTopicsId();
  }
  
  // 1. Resolve or create target year folder (e.g., 2027)
  var targetFolder = getOrCreateTargetYearFolder(year);
  
  // 2. Clone Holy Work Schedule into target folder
  var currentFile = DriveApp.getFileById(localSs.getId());
  var newScheduleName = year + " Holy Work Schedule";
  var newScheduleFile = currentFile.makeCopy(newScheduleName, targetFolder);
  var newScheduleSs = SpreadsheetApp.openById(newScheduleFile.getId());
  
  // 3. Reset the new Holy Work Schedule (Q1-Q4 anchor dates, wipe names, clear logs)
  resetForNewYear(year, newScheduleSs);
  try {
    restructureQuarterSheetsToUniversal14(newScheduleSs);
  } catch (restructErr) {
    console.warn("Could not enforce 14-row restructuring on new schedule: " + restructErr.message);
  }
  
  // 4. Clone Unavailability Spreadsheet into target folder
  var sourceUnavailFile = DriveApp.getFileById(sourceUnavailId);
  var newUnavailName = year + " Holy Work Unavailability";
  var newUnavailFile = sourceUnavailFile.makeCopy(newUnavailName, targetFolder);
  var newUnavailSs = SpreadsheetApp.openById(newUnavailFile.getId());
  
  // Set permissions: Anyone with the link can edit
  try {
    newUnavailFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  } catch (permErr) {
    console.warn("Could not set public edit sharing on Unavailability file: " + permErr.message);
  }
  
  // 5. Configure the new Unavailability Sheet (Universal 14-Column layout + worker roster sync)
  try {
    restructureAndSyncUnavailabilityMatrix(newUnavailSs, localSs, year, true);
  } catch (unavailErr) {
    console.warn("Could not restructure and sync Unavailability matrix: " + unavailErr.message);
  }
  
  // 6. Clone Sermon Topics & Hymns Sheet into target folder
  var newSermonFile = null;
  var newSermonSs = null;
  var sermonProvisionError = "";
  
  // Build candidate ID list: try provided ID, then fallback to default template
  var candidateSermonIds = [];
  if (sourceSermonTopicsId) candidateSermonIds.push(sourceSermonTopicsId);
  var fallbackTemplateId = CONFIG.DEFAULT_SERMON_TOPICS_ID || '1zrMaPubQ7uaUvgBjYZ3lWIBRUclg0GsGliHMqpvOark';
  if (fallbackTemplateId && candidateSermonIds.indexOf(fallbackTemplateId) === -1) {
    candidateSermonIds.push(fallbackTemplateId);
  }
  
  for (var sIdx = 0; sIdx < candidateSermonIds.length; sIdx++) {
    var trySermonId = candidateSermonIds[sIdx];
    try {
      var sourceSermonFile = DriveApp.getFileById(trySermonId);
      var newSermonName = year + " Sermon Topic & Hymns";
      newSermonFile = sourceSermonFile.makeCopy(newSermonName, targetFolder);
      
      // Update properties to point to the working template
      PropertiesService.getScriptProperties().setProperty(CONFIG.SERMON_TOPICS_SHEET_PROPERTY, trySermonId);
      PropertiesService.getScriptProperties().setProperty(CONFIG.SERMON_TOPICS_SHEET_URL_PROPERTY, sourceSermonFile.getUrl());
      
      // Set permissions: Anyone with the link can edit
      try {
        newSermonFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
      } catch (permErr) {
        console.warn("Could not set public edit sharing on Sermon Topics: " + permErr.message);
      }
      
      try {
        newSermonSs = SpreadsheetApp.openById(newSermonFile.getId());
        // Configure new Sermon Topics Sheet (tab names, IMPORTRANGE rewrite, clear data, dates)
        resetSermonTopicsSheet(year, newSermonSs, newScheduleSs.getUrl());
      } catch (resetErr) {
        console.warn("Could not complete in-sheet resets on Sermon Topics: " + resetErr.message);
      }
      
      sermonProvisionError = "";
      break; // Cloned successfully!
    } catch (sermonErr) {
      console.error("Could not clone Sermon Topics file (" + trySermonId + "): " + sermonErr.message);
      sermonProvisionError = sermonErr.message;
      newSermonFile = null;
    }
  }
  
  // 7. Link the new Unavailability & Sermon Topics IDs into the new Holy Work Schedule
  var newInternalUnavailSheet = newScheduleSs.getSheetByName(CONFIG.UNAVAILABILITY_SHEET);
  if (!newInternalUnavailSheet) {
    newInternalUnavailSheet = newScheduleSs.insertSheet(CONFIG.UNAVAILABILITY_SHEET);
  }
  newInternalUnavailSheet.clear();
  newInternalUnavailSheet.getRange("A1:C1").setValues([["Date", "Worker Name", "Reason"]])
    .setFontWeight("bold")
    .setBackground("#E0E7FF");
  newInternalUnavailSheet.getRange("D1:E1").setValues([["External Spreadsheet ID:", newUnavailFile.getId()]])
    .setFontColor("#777777");
    
  if (newSermonFile) {
    newInternalUnavailSheet.getRange("D2:E2").setValues([["Sermon Topics URL:", newSermonFile.getUrl()]])
      .setFontColor("#777777");
    newInternalUnavailSheet.getRange("D3:E3").setValues([["Sermon Topics ID:", newSermonFile.getId()]])
      .setFontColor("#777777");
  }
  newInternalUnavailSheet.setFrozenRows(1);
  
  // Return links for modal display
  return {
    year: year,
    folderUrl: targetFolder.getUrl(),
    scheduleUrl: newScheduleSs.getUrl(),
    scheduleId: newScheduleFile.getId(),
    unavailUrl: newUnavailSs.getUrl(),
    unavailId: newUnavailFile.getId(),
    sermonTopicsUrl: newSermonFile ? newSermonFile.getUrl() : "",
    sermonTopicsId: newSermonFile ? newSermonFile.getId() : "",
    sermonTopicsError: sermonProvisionError
  };
  }, { timeoutMs: 60000 });
}

/**
 * Restructures and synchronizes the Unavailability Matrix to the Universal 14-Column Architecture
 * across Q1-Q4 (Columns B through BE, 56 date columns) and synchronizes qualified workers
 * into four distinct sections: Monday, Friday, Saturday AM, and Saturday PM.
 *
 * Preserves existing worker checkboxes by matching (WorkerName, Date).
 *
 * @param {Spreadsheet} unavailSs Target Unavailability Spreadsheet.
 * @param {Spreadsheet} [masterSs] Master Holy Work Schedule Spreadsheet (defaults to active).
 * @param {number} [targetYear] Target 4-digit year.
 * @param {boolean} [isFreshYearReset] If true, resets all checkboxes to false (used during annual rollover).
 */
function restructureAndSyncUnavailabilityMatrix(unavailSs, masterSs, targetYear, isFreshYearReset) {
  if (!unavailSs) throw new Error("Target Unavailability Spreadsheet is missing or cannot be opened.");
  if (!masterSs) masterSs = SpreadsheetApp.getActiveSpreadsheet();
  
  var unavailSheet = unavailSs.getSheetByName("Unavailability") || unavailSs.getSheets()[0];
  if (!unavailSheet) throw new Error("Could not find 'Unavailability' tab in target spreadsheet.");
  
  // 1. Determine Target Year
  if (!targetYear) {
    var a2Val = unavailSheet.getRange("A2").getValue();
    if (a2Val instanceof Date) {
      targetYear = a2Val.getFullYear();
    } else if (a2Val && !isNaN(parseInt(a2Val, 10))) {
      var parsedY = parseInt(a2Val, 10);
      if (parsedY >= 2020 && parsedY <= 2100) targetYear = parsedY;
    }
  }
  if (!targetYear) {
    var q1A2 = masterSs.getSheetByName("Q1") ? masterSs.getSheetByName("Q1").getRange("A2").getValue() : null;
    if (q1A2 instanceof Date) {
      targetYear = q1A2.getFullYear();
    } else {
      targetYear = new Date().getFullYear() + 1;
    }
  }
  
  // 2. Read master Workers tab
  var workerSheet = masterSs.getSheetByName(CONFIG.WORKERS_SHEET);
  if (!workerSheet) throw new Error("Could not find master 'Workers' tab in Holy Work Schedule.");
  
  var workerData = workerSheet.getDataRange().getValues();
  if (workerData.length < 2) throw new Error("Master 'Workers' tab contains no data rows.");
  
  var detected = detectWorkersHeaderRow_(workerData);
  var headers = detected.headers.map(function(h) { return String(h || "").trim().toLowerCase(); });
  var nameCol = -1, activeCol = -1, monCol = -1, friCol = -1, rosterCol = -1, churchCol = -1;
  
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (nameCol === -1 && (h.indexOf("schedule name") !== -1 || h === "name")) nameCol = i;
    if (activeCol === -1 && h === "active") activeCol = i;
    if (monCol === -1 && (h.indexOf("available monday") !== -1 || h === "monday" || h === "mondays")) monCol = i;
    if (friCol === -1 && (h.indexOf("available friday") !== -1 || h === "friday" || h === "fridays")) friCol = i;
    if (rosterCol === -1 && h.indexOf("roster status") !== -1) rosterCol = i;
    if (churchCol === -1 && (h.indexOf("resident church") !== -1 || h === "church")) churchCol = i;
  }
  if (nameCol === -1) nameCol = 4;   // Default Column E
  if (activeCol === -1) activeCol = 11; // Default Column L
  if (monCol === -1) monCol = 2;     // Default Column C
  if (friCol === -1) friCol = 3;     // Default Column D
  if (rosterCol === -1) rosterCol = 12; // Default Column M
  if (churchCol === -1) churchCol = 9;  // Default Column J (Resident Church)
  
  var mondayWorkers = [];
  var fridayWorkers = [];
  var saturdayWorkers = [];
  
  for (var r = detected.dataStartIdx; r < workerData.length; r++) {
    var name = String(workerData[r][nameCol] || "").trim();
    if (!name) {
      // Fallback: construct name from Col A + Col B if Schedule Name formula wasn't evaluated
      var first = String(workerData[r][0] || "").trim();
      var last = String(workerData[r][1] || "").trim();
      name = (first + " " + last).trim();
    }
    if (!name) continue;
    
    var isActive = String(workerData[r][activeCol] || "").trim().toLowerCase() === "active";
    if (!isActive) continue;
    
    // Resident Church (Column J) must be blank for local Garden Grove members
    var residentChurch = String(workerData[r][churchCol] || "").trim();
    if (residentChurch !== "") continue;
    
    var statusStr = String(workerData[r][rosterCol] || "").trim().toLowerCase();
    var isRostered = (statusStr.indexOf("rostered") !== -1 && statusStr.indexOf("not rostered") === -1);
    
    // Fallback: check role checkboxes across columns O to BA (cols 14 to 52)
    if (!isRostered && statusStr === "") {
      var roleCnt = 0;
      for (var c = 14; c < Math.min(workerData[r].length, 53); c++) {
        if (workerData[r][c] === true || String(workerData[r][c]).toLowerCase() === "true") roleCnt++;
      }
      if (roleCnt >= 1) isRostered = true;
    }
    
    // Qualification rule: Active worker must be Rostered (Roster Status >= 1, not "Not Rostered")
    if (!isRostered) continue;
    
    var availMon = (workerData[r][monCol] === true || String(workerData[r][monCol]).toLowerCase() === "true");
    var availFri = (workerData[r][friCol] === true || String(workerData[r][friCol]).toLowerCase() === "true");
    
    if (availMon && mondayWorkers.indexOf(name) === -1) mondayWorkers.push(name);
    if (availFri && fridayWorkers.indexOf(name) === -1) fridayWorkers.push(name);
    if (saturdayWorkers.indexOf(name) === -1) saturdayWorkers.push(name);
  }
  
  mondayWorkers.sort();
  fridayWorkers.sort();
  saturdayWorkers.sort();
  
  // 3. Read Existing Unavailability Checkboxes to Preserve Data (by Worker & Date)
  var existingCheckedMap = {
    monday: {},
    friday: {},
    saturday_am: {},
    saturday_pm: {}
  };
  
  var tz = masterSs.getSpreadsheetTimeZone();
  
  if (!isFreshYearReset) {
    var curLastRow = unavailSheet.getLastRow();
    var curLastCol = unavailSheet.getLastColumn();
    if (curLastRow >= 3 && curLastCol >= 2) {
      var curData = unavailSheet.getRange(1, 1, curLastRow, Math.min(curLastCol, 60)).getValues();
      var curSec = "";
      var curSecDates = null;
      
      for (var r = 0; r < curData.length; r++) {
        var c1 = String(curData[r][0] || "").trim().toLowerCase();
        if (!c1) continue;
        
        if (c1 === "monday") { curSec = "monday"; curSecDates = curData[r]; continue; }
        if (c1 === "friday") { curSec = "friday"; curSecDates = curData[r]; continue; }
        if (c1.indexOf("saturday am") !== -1 || c1 === "sabbath am") { curSec = "saturday_am"; curSecDates = curData[r]; continue; }
        if (c1.indexOf("saturday pm") !== -1 || c1 === "sabbath pm") { curSec = "saturday_pm"; curSecDates = curData[r]; continue; }
        if (c1 === "saturday" || c1 === "sabbath") {
          curSec = (curSec !== "saturday_am" && curSec !== "saturday_pm") ? "saturday_am" : "saturday_pm";
          curSecDates = curData[r];
          continue;
        }
        
        if (curSec && curSecDates && r >= 3) {
          var wName = String(curData[r][0] || "").trim();
          if (!wName || wName === "Q1" || wName.toLowerCase() === "date") continue;
          
          if (!existingCheckedMap[curSec][wName]) existingCheckedMap[curSec][wName] = {};
          
          for (var c = 1; c < curData[r].length; c++) {
            var val = curData[r][c];
            if (val === true || String(val).toUpperCase() === "TRUE") {
              var dVal = curSecDates[c];
              if (dVal) {
                var dObj = dVal instanceof Date ? dVal : new Date(dVal);
                if (!isNaN(dObj.getTime())) {
                  var dStr = Utilities.formatDate(dObj, tz, "yyyy-MM-dd");
                  existingCheckedMap[curSec][wName][dStr] = true;
                }
              }
            }
          }
        }
      }
    }
  }
  
  // 4. Calculate exact dates for target year (to match preserved checkboxes)
  function get56DatesForDayOfWeek(y, weekdayNum) { // 1=Mon, 5=Fri, 6=Sat
    var dates = [];
    for (var q = 0; q < 4; q++) {
      var startMonth = q * 3;
      var qStart = new Date(y, startMonth, 1);
      var qEnd = new Date(y, startMonth + 3, 0); // last day of quarter
      
      var cur = new Date(qStart.getTime());
      while (cur.getDay() !== weekdayNum) {
        cur.setDate(cur.getDate() + 1);
      }
      
      var count = 0;
      while (cur <= qEnd && count < 14) {
        dates.push(new Date(cur.getTime()));
        cur.setDate(cur.getDate() + 7);
        count++;
      }
      while (count < 14) {
        dates.push(null); // inactive week
        count++;
      }
    }
    return dates;
  }
  
  var mondayDates = get56DatesForDayOfWeek(targetYear, 1);
  var fridayDates = get56DatesForDayOfWeek(targetYear, 5);
  var saturdayDates = get56DatesForDayOfWeek(targetYear, 6);
  
  // 5. Structure & Layout Dimensions
  // Row 1: Title
  // Row 2: Anchor Date & Quarters
  // Row 3: Monday Header (dates)
  // Rows 4 to (3 + monCount): Monday Data
  // Next Row: Friday Header (friHRow)
  // Next friCount rows: Friday Data
  // Next Row: Saturday AM Header (satAmHRow)
  // Next satCount rows: Saturday AM Data
  // Next Row: Saturday PM Header (satPmHRow)
  // Next satCount rows: Saturday PM Data
  var friHRow = 4 + mondayWorkers.length;
  var satAmHRow = friHRow + 1 + fridayWorkers.length;
  var satPmHRow = satAmHRow + 1 + saturdayWorkers.length;
  var totalRowsNeeded = satPmHRow + saturdayWorkers.length;
  
  // Ensure sufficient columns (at least 57: Col A + 56 date cols B:BE)
  if (unavailSheet.getMaxColumns() < 57) {
    unavailSheet.insertColumnsAfter(unavailSheet.getMaxColumns(), 57 - unavailSheet.getMaxColumns());
  }
  // Ensure sufficient rows
  if (unavailSheet.getMaxRows() < totalRowsNeeded) {
    unavailSheet.insertRowsAfter(unavailSheet.getMaxRows(), totalRowsNeeded - unavailSheet.getMaxRows());
  }
  
  // A. Clear all existing data validations across the sheet to prevent date cell violations
  unavailSheet.getDataRange().clearDataValidations();
  
  // B. Clear all existing row groups to eliminate offset/stale groupings
  try {
    for (var d = 0; d < 8; d++) {
      unavailSheet.getDataRange().shiftRowGroupDepth(-1);
    }
  } catch (e) {
    // shiftRowGroupDepth terminates cleanly when all depths removed
  }
  
  // C. Reset any existing basic filter so rows are never hidden on load
  var currentFilter = unavailSheet.getFilter();
  if (currentFilter) {
    try { currentFilter.remove(); } catch(e) {}
  }
  
  // Helper to get column letter (1 -> A, 2 -> B, ..., 57 -> BE)
  function toColLetter(colIdx) {
    var res = "";
    while (colIdx > 0) {
      var rem = (colIdx - 1) % 26;
      res = String.fromCharCode(65 + rem) + res;
      colIdx = Math.floor((colIdx - 1) / 26);
    }
    return res;
  }
  
  // Helper to generate 56 date formulas for a section header row
  function buildSectionDateFormulas(hRow, mask) {
    var formulas = [];
    // Q1: cols 2..15 (B..O)
    formulas.push('=WORKDAY.INTL(DATE(YEAR($A$2), 1, 1) - 1, 1, "' + mask + '")');
    for (var c = 3; c <= 15; c++) {
      var prev = toColLetter(c - 1);
      formulas.push('=IF(' + prev + hRow + '="", "", IF(' + prev + hRow + '+7 > EOMONTH(DATE(YEAR($A$2), 1, 1), 2), "", ' + prev + hRow + '+7))');
    }
    // Q2: cols 16..29 (P..AC)
    formulas.push('=WORKDAY.INTL(DATE(YEAR($A$2), 4, 1) - 1, 1, "' + mask + '")');
    for (var c = 17; c <= 29; c++) {
      var prev = toColLetter(c - 1);
      formulas.push('=IF(' + prev + hRow + '="", "", IF(' + prev + hRow + '+7 > EOMONTH(DATE(YEAR($A$2), 4, 1), 2), "", ' + prev + hRow + '+7))');
    }
    // Q3: cols 30..43 (AD..AQ)
    formulas.push('=WORKDAY.INTL(DATE(YEAR($A$2), 7, 1) - 1, 1, "' + mask + '")');
    for (var c = 31; c <= 43; c++) {
      var prev = toColLetter(c - 1);
      formulas.push('=IF(' + prev + hRow + '="", "", IF(' + prev + hRow + '+7 > EOMONTH(DATE(YEAR($A$2), 7, 1), 2), "", ' + prev + hRow + '+7))');
    }
    // Q4: cols 44..57 (AR..BE)
    formulas.push('=WORKDAY.INTL(DATE(YEAR($A$2), 10, 1) - 1, 1, "' + mask + '")');
    for (var c = 45; c <= 57; c++) {
      var prev = toColLetter(c - 1);
      formulas.push('=IF(' + prev + hRow + '="", "", IF(' + prev + hRow + '+7 > EOMONTH(DATE(YEAR($A$2), 10, 1), 2), "", ' + prev + hRow + '+7))');
    }
    return formulas;
  }
  
  // 6. Write Row 1 (Search Box & Title) & Row 2 (Anchor & Quarters)
  // Cell A1: Search dropdown populated with active workers, blank by default so all rows show
  var allUniqueWorkers = Array.from(new Set(mondayWorkers.concat(fridayWorkers).concat(saturdayWorkers))).sort();
  if (allUniqueWorkers.length > 0) {
    var nameDropdownRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(allUniqueWorkers, true)
      .setAllowInvalid(true)
      .build();
    unavailSheet.getRange("A1")
      .setDataValidation(nameDropdownRule)
      .setValue("") // Blank initially so all rows show!
      .setFontWeight("bold")
      .setHorizontalAlignment("left")
      .setNote("Select or type a worker name to filter unavailability, or leave blank to show all.");
  } else {
    unavailSheet.getRange("A1").setValue("").clearNote();
  }
  
  unavailSheet.getRange("B1").setValue(targetYear + " Holy Work Unavailability Matrix").setFontWeight("bold").setFontSize(11);
  unavailSheet.getRange("A2").setValue(new Date(targetYear, 0, 1)).setNumberFormat("yyyy-MM-dd");
  
  var qLabels = [];
  for (var c = 2; c <= 57; c++) {
    if (c <= 15) qLabels.push("Q1");
    else if (c <= 29) qLabels.push("Q2");
    else if (c <= 43) qLabels.push("Q3");
    else qLabels.push("Q4");
  }
  unavailSheet.getRange(2, 2, 1, 56)
    .setValues([qLabels])
    .setBackground("#D9D9D9")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
    
  // 7. Write Section Headers and Date Formulas (No Checkbox Validations on Headers!)
  // Monday Header
  unavailSheet.getRange(3, 1).setValue("Monday");
  unavailSheet.getRange(3, 2, 1, 56).setFormulas([buildSectionDateFormulas(3, "0111111")]).setNumberFormat("M/d");
  unavailSheet.getRange(3, 1, 1, 57).setBackground("#6FA8DC").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  
  // Friday Header
  unavailSheet.getRange(friHRow, 1).setValue("Friday");
  unavailSheet.getRange(friHRow, 2, 1, 56).setFormulas([buildSectionDateFormulas(friHRow, "1111011")]).setNumberFormat("M/d");
  unavailSheet.getRange(friHRow, 1, 1, 57).setBackground("#6FA8DC").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  
  // Saturday AM Header
  unavailSheet.getRange(satAmHRow, 1).setValue("Saturday AM");
  unavailSheet.getRange(satAmHRow, 2, 1, 56).setFormulas([buildSectionDateFormulas(satAmHRow, "1111101")]).setNumberFormat("M/d");
  unavailSheet.getRange(satAmHRow, 1, 1, 57).setBackground("#6FA8DC").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  
  // Saturday PM Header (mirrors Saturday AM)
  unavailSheet.getRange(satPmHRow, 1).setValue("Saturday PM");
  var satPmFormulas = [];
  for (var c = 2; c <= 57; c++) {
    var cLet = toColLetter(c);
    satPmFormulas.push('=IF(' + cLet + satAmHRow + '="", "", ' + cLet + satAmHRow + ')');
  }
  unavailSheet.getRange(satPmHRow, 2, 1, 56).setFormulas([satPmFormulas]).setNumberFormat("M/d");
  unavailSheet.getRange(satPmHRow, 1, 1, 57).setBackground("#6FA8DC").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  
  // 8. Helper to build and write section data rows
  function populateSectionData(startRow, workerList, sectionKey, targetDates) {
    if (workerList.length === 0) return;
    var rows = [];
    for (var w = 0; w < workerList.length; w++) {
      var wName = workerList[w];
      var row = [wName];
      var savedDates = existingCheckedMap[sectionKey] ? existingCheckedMap[sectionKey][wName] : null;
      
      for (var d = 0; d < 56; d++) {
        var isChecked = false;
        var tDate = targetDates[d];
        if (savedDates && tDate) {
          var tStr = Utilities.formatDate(tDate, tz, "yyyy-MM-dd");
          if (savedDates[tStr]) isChecked = true;
        }
        row.push(isChecked);
      }
      rows.push(row);
    }
    
    var range = unavailSheet.getRange(startRow, 1, rows.length, 57);
    range.setValues(rows);
    range.setBackground("#FFFFFF");
    range.setBorder(true, true, true, true, true, true, "#B7B7B7", SpreadsheetApp.BorderStyle.SOLID);
    
    // Checkbox data validation on cols 2..57 ONLY for data rows
    var checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    unavailSheet.getRange(startRow, 2, rows.length, 56).setDataValidation(checkRule).setHorizontalAlignment("center");
    unavailSheet.getRange(startRow, 1, rows.length, 1).setHorizontalAlignment("left").setFontWeight("normal");
  }
  
  // Write data rows for all 4 sections
  populateSectionData(4, mondayWorkers, "monday", mondayDates);
  populateSectionData(friHRow + 1, fridayWorkers, "friday", fridayDates);
  populateSectionData(satAmHRow + 1, saturdayWorkers, "saturday_am", saturdayDates);
  populateSectionData(satPmHRow + 1, saturdayWorkers, "saturday_pm", saturdayDates);
  
  // 9. Apply dynamic row groupings for each section
  try {
    if (mondayWorkers.length > 0) {
      unavailSheet.getRange(4, 1, mondayWorkers.length, 1).shiftRowGroupDepth(1);
    }
    if (fridayWorkers.length > 0) {
      unavailSheet.getRange(friHRow + 1, 1, fridayWorkers.length, 1).shiftRowGroupDepth(1);
    }
    if (saturdayWorkers.length > 0) {
      unavailSheet.getRange(satAmHRow + 1, 1, saturdayWorkers.length, 1).shiftRowGroupDepth(1);
    }
    if (saturdayWorkers.length > 0) {
      unavailSheet.getRange(satPmHRow + 1, 1, saturdayWorkers.length, 1).shiftRowGroupDepth(1);
    }
    // Expand all row groups by default so all rows are visible initially
    unavailSheet.expandAllRowGroups();
  } catch (grpApplyErr) {
    console.warn("Could not apply row groupings: " + grpApplyErr.message);
  }
  
  // 10. Re-create clean basic filter over A3:BE[totalRowsNeeded]
  try {
    unavailSheet.getRange(3, 1, totalRowsNeeded - 2, 57).createFilter();
  } catch (filterErr) {}
  
  // 11. Clean up residual rows beyond Saturday PM
  var maxRows = unavailSheet.getMaxRows();
  if (maxRows > totalRowsNeeded) {
    unavailSheet.getRange(totalRowsNeeded + 1, 1, maxRows - totalRowsNeeded, unavailSheet.getMaxColumns()).clear();
  }
  
  // 12. Sheet Layout & Freezes
  unavailSheet.setFrozenColumns(1);
  unavailSheet.setFrozenRows(3);
  unavailSheet.setColumnWidth(1, 160);
  unavailSheet.setColumnWidths(2, 56, 48);
  
  return {
    mondayCount: mondayWorkers.length,
    fridayCount: fridayWorkers.length,
    saturdayCount: saturdayWorkers.length,
    totalRows: totalRowsNeeded
  };
}

/**
 * UI prompt to manually synchronize workers and restructure the linked Unavailability Matrix.
 */
function promptSyncWorkersToUnavailability() {
  var ui = SpreadsheetApp.getUi();
  var localSs = SpreadsheetApp.getActiveSpreadsheet();
  var unavailId = getLinkedUnavailabilityId();
  
  if (!unavailId) {
    ui.alert('Unavailability Sheet Not Linked', 'No external Unavailability spreadsheet is linked. Please use "Link Unavailability Sheet" first.', ui.ButtonSet.OK);
    return;
  }
  
  var unavailSs;
  try {
    unavailSs = SpreadsheetApp.openById(unavailId);
    try {
      var unavailFile = DriveApp.getFileById(unavailId);
      unavailFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    } catch (permErr) {
      console.warn("Could not set public edit sharing on Unavailability file: " + permErr.message);
    }
  } catch (e) {
    ui.alert('Error Opening Sheet', 'Could not open external Unavailability spreadsheet with ID: ' + unavailId + '\n\nError: ' + e.message, ui.ButtonSet.OK);
    return;
  }
  
  localSs.toast('Restructuring 14-column layout and syncing worker roster...', 'Please Wait', 20);
  try {
    var stats = restructureAndSyncUnavailabilityMatrix(unavailSs, localSs);
    ui.alert(
      'Sync Complete',
      'Successfully restructured and synchronized the Unavailability Matrix:\n\n' +
      '• Mondays: ' + stats.mondayCount + ' workers\n' +
      '• Fridays: ' + stats.fridayCount + ' workers\n' +
      '• Saturday AM: ' + stats.saturdayCount + ' workers\n' +
      '• Saturday PM: ' + stats.saturdayCount + ' workers\n' +
      '• Architecture: Universal 14-Column Grid (Cols B-BE, 56 weeks)\n' +
      '• Total Matrix Rows: ' + stats.totalRows,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Sync Failed', 'Failed to restructure and sync Unavailability Matrix:\n\n' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Backward-compatible adapter for syncWorkersToUnavailabilityColumnA.
 */
function syncWorkersToUnavailabilityColumnA(masterSs, unavailSheet) {
  var ss = unavailSheet.getParent();
  return restructureAndSyncUnavailabilityMatrix(ss, masterSs);
}

/**
 * Standalone provisioner for Sermon Topics & Hymns sheet into an existing or new year folder.
 */
function promptProvisionSermonTopicsOnly() {
  var ui = SpreadsheetApp.getUi();
  var localSs = SpreadsheetApp.getActiveSpreadsheet();
  var currentYear = new Date().getFullYear();
  var defaultNextYear = currentYear + 1;
  
  var respYear = ui.prompt(
    'Provision Sermon Topics & Hymns Only',
    'Enter the 4-digit target year to provision (e.g., ' + defaultNextYear + '):',
    ui.ButtonSet.OK_CANCEL
  );
  if (respYear.getSelectedButton() !== ui.Button.OK) return;
  var targetYear = parseInt(respYear.getResponseText().trim(), 10);
  if (isNaN(targetYear) || targetYear < 2020 || targetYear > 2100) {
    ui.alert('Invalid Year', 'Please enter a valid 4-digit year.', ui.ButtonSet.OK);
    return;
  }
  
  var sourceId = getLinkedSermonTopicsId();
  if (!sourceId) {
    var respSource = ui.prompt(
      'Source Sermon Topics Sheet',
      'Enter the Google Spreadsheet URL or File ID of the baseline template:\n\n(Default: ' + CONFIG.DEFAULT_SERMON_TOPICS_ID + ')',
      ui.ButtonSet.OK_CANCEL
    );
    if (respSource.getSelectedButton() !== ui.Button.OK) return;
    var inputSource = respSource.getResponseText().trim();
    var match = inputSource.match(/\/d\/([a-zA-Z0-9-_]+)/);
    sourceId = match ? match[1] : inputSource;
  }
  if (!sourceId) sourceId = CONFIG.DEFAULT_SERMON_TOPICS_ID;
  
  try {
    localSs.toast('Locating target year folder and cloning Sermon Topics...', 'Please wait', 15);
    var targetFolder = getOrCreateTargetYearFolder(targetYear);
    var sourceFile = DriveApp.getFileById(sourceId);
    var newName = targetYear + " Sermon Topic & Hymns";
    var newFile = sourceFile.makeCopy(newName, targetFolder);
    
    try {
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    } catch (e) {
      console.warn("Could not set public edit permissions: " + e.message);
    }
    
    var newSs = SpreadsheetApp.openById(newFile.getId());
    resetSermonTopicsSheet(targetYear, newSs, localSs.getUrl());
    
    // Save to script properties and internal tab
    PropertiesService.getScriptProperties().setProperty(CONFIG.SERMON_TOPICS_SHEET_PROPERTY, newFile.getId());
    PropertiesService.getScriptProperties().setProperty(CONFIG.SERMON_TOPICS_SHEET_URL_PROPERTY, newFile.getUrl());
    
    var internalTab = localSs.getSheetByName(CONFIG.UNAVAILABILITY_SHEET);
    if (internalTab) {
      internalTab.getRange("D2").setValue("Sermon Topics URL:");
      internalTab.getRange("E2").setValue(newFile.getUrl());
      internalTab.getRange("D3").setValue("Sermon Topics ID:");
      internalTab.getRange("E3").setValue(newFile.getId());
    }
    
    ui.alert(
      '🎉 Sermon Topics Provisioned for ' + targetYear + '!',
      'File: ' + newName + '\n' +
      'Folder: ' + targetFolder.getName() + '\n' +
      'URL: ' + newFile.getUrl() + '\n\n' +
      'Permissions set to: Anyone with the link can edit ✅\n' +
      'IMPORTRANGE updated to point to this Holy Work Schedule.',
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('Provisioning Error', 'Could not provision Sermon Topics sheet:\n\n' + err.message, ui.ButtonSet.OK);
    console.error(err);
  }
}

/**
 * Directly updates all IMPORTRANGE formulas in the target Sermon Topics & Hymns sheet
 * to pull Speaker Name from Column AY across Q1-Q4 of the Holy Work Schedule.
 */
function promptUpdateSermonTopicsSpeakerFormulas() {
  var ui = SpreadsheetApp.getUi();
  var localSs = SpreadsheetApp.getActiveSpreadsheet();
  var sermonId = getLinkedSermonTopicsId();
  
  var resp = ui.prompt(
    'Update Sermon Topics Speaker IMPORTRANGEs',
    'Enter the Google Spreadsheet URL or File ID of the Sermon Topics & Hymns sheet to update:\n\n(Detected: ' + (sermonId || 'None') + ')',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  
  var input = resp.getResponseText().trim();
  if (input) {
    var match = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
    sermonId = match ? match[1] : input;
  }
  if (!sermonId) {
    ui.alert('Error', 'No valid Sermon Topics spreadsheet ID provided.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    localSs.toast('Updating IMPORTRANGE formulas to Column AY...', 'Please wait', 10);
    var targetSs = SpreadsheetApp.openById(sermonId);
    var hwsUrl = localSs.getUrl();
    var hwsCleanUrl = hwsUrl.split('#')[0].split('?')[0];
    if (!hwsCleanUrl.endsWith('/edit')) hwsCleanUrl += '/edit';
    
    var quarters = CONFIG.QUARTERS;
    var sections = [
      { start: 4, end: 17, formulaRow: 4, hwsRange: "AY6:AY19", label: "Mondays" },
      { start: 20, end: 33, formulaRow: 20, hwsRange: "AY51:AY64", label: "Fridays" },
      { start: 36, end: 49, formulaRow: 36, hwsRange: "AY66:AY79", label: "Saturday AM" },
      { start: 52, end: 65, formulaRow: 52, hwsRange: "AY81:AY94", label: "Saturday PM" }
    ];
    
    var updatedCount = 0;
    quarters.forEach(function(qName) {
      var s = targetSs.getSheetByName(qName);
      if (!s) return;
      
      sections.forEach(function(sec) {
        var formulaCell = s.getRange(sec.formulaRow, 2); // Column B
        var newFormula = '=IMPORTRANGE("' + hwsCleanUrl + '", "' + qName + '!' + sec.hwsRange + '")';
        formulaCell.setFormula(newFormula);
        
        // Clear spill cells below to prevent #REF! collision
        if (sec.end > sec.formulaRow) {
          s.getRange(sec.formulaRow + 1, 2, sec.end - sec.formulaRow, 1).clearContent();
        }
        updatedCount++;
      });
    });
    
    // Save to script properties and internal tab
    PropertiesService.getScriptProperties().setProperty(CONFIG.SERMON_TOPICS_SHEET_PROPERTY, targetSs.getId());
    PropertiesService.getScriptProperties().setProperty(CONFIG.SERMON_TOPICS_SHEET_URL_PROPERTY, targetSs.getUrl());
    
    var internalTab = localSs.getSheetByName(CONFIG.UNAVAILABILITY_SHEET);
    if (internalTab) {
      internalTab.getRange("D2").setValue("Sermon Topics URL:");
      internalTab.getRange("E2").setValue(targetSs.getUrl());
      internalTab.getRange("D3").setValue("Sermon Topics ID:");
      internalTab.getRange("E3").setValue(targetSs.getId());
    }
    
    ui.alert(
      'Formulas Updated! ✅',
      'Successfully updated ' + updatedCount + ' IMPORTRANGE formulas across Q1-Q4 in "' + targetSs.getName() + '".\n\n' +
      'All speaker names are now linked to Column AY of this Holy Work Schedule:\n' +
      '• Mondays: [Quarter]!AY6:AY19\n' +
      '• Fridays: [Quarter]!AY51:AY64\n' +
      '• Saturday AM: [Quarter]!AY66:AY79\n' +
      '• Saturday PM: [Quarter]!AY81:AY94',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Update Failed', 'Could not update formulas:\n\n' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * One-click comprehensive repair utility for 2027 Live Spreadsheets.
 * - Repairs Sermon Topics IMPORTRANGE formulas across Q1-Q4 (Column AY mapping & spill clears).
 * - Corrects Team Emails & Usher Team Emails lookup formulas (points to Workers!$F$3:$F instead of $F$4:$F).
 * - Verifies/sets timezone of linked ESSC and Sermon schedules to America/Los_Angeles.
 */
function repair2027LiveSpreadsheets() {
  return executeWithLockAndRetry_('repair2027LiveSpreadsheets', function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = null;
    try { ui = SpreadsheetApp.getUi(); } catch(e) {}
    var log = [];

    // 1. Repair Team Emails & Usher Team Emails formulas (Workers row 3 inclusion & syntax restoration)
    var emailTabs = ['Team Emails', 'Usher Team Emails'];
    var emailFormulasRepaired = 0;
    emailTabs.forEach(function(tabName) {
      var tab = ss.getSheetByName(tabName);
      if (!tab) return;
      var lastCol = tab.getLastColumn();
      var numRowsToScan = Math.min(2, tab.getLastRow());
      if (lastCol < 1 || numRowsToScan < 1) return;
      
      var formulasMatrix = tab.getRange(1, 1, numRowsToScan, lastCol).getFormulas();
      var modified = false;
      for (var r = 0; r < formulasMatrix.length; r++) {
        for (var c = 0; c < formulasMatrix[r].length; c++) {
          var f = formulasMatrix[r][c];
          if (!f) continue;
          var newF = f
            .replace(/Workers!\$(?:1\$3|[EF]\$[34]):\$([EF])/g, function(m, col) { return 'Workers!$' + col + '$3:$' + col; })
            .replace(/Workers!([A-Z]+)4:([A-Z]+)/g, function(m, c1, c2) { return 'Workers!' + c1 + '3:' + c2; });
          if (newF !== f) {
            formulasMatrix[r][c] = newF;
            modified = true;
            emailFormulasRepaired++;
          }
        }
      }
      if (modified) {
        tab.getRange(1, 1, numRowsToScan, lastCol).setFormulas(formulasMatrix);
        log.push('Repaired ' + tabName + ' formulas to include Workers row 3 (Aaron Young) and fix syntax.');
      }
    });

    // 1b. Populate blank ValidationConfig Column C cleaning entries (AL-AU)
    var valSheet = ss.getSheetByName('ValidationConfig');
    if (valSheet && valSheet.getLastRow() >= 2) {
      var valData = valSheet.getRange(2, 1, valSheet.getLastRow() - 1, 3).getValues();
      var valModified = 0;
      for (var vr = 0; vr < valData.length; vr++) {
        var colLet = String(valData[vr][0] || '').trim().toUpperCase();
        var roleDesc = String(valData[vr][1] || '').trim();
        var teamsVal = String(valData[vr][2] || '').trim();
        if (colLet && !teamsVal && typeof inferCleaningFallbackTeams_ === 'function') {
          var cIdx0 = colLetterToIdx_(colLet) - 1;
          var inferred = inferCleaningFallbackTeams_(cIdx0, roleDesc);
          if (inferred.length > 0) {
            valData[vr][2] = inferred.join(', ');
            valModified++;
          }
        }
      }
      if (valModified > 0) {
        valSheet.getRange(2, 1, valData.length, 3).setValues(valData);
        log.push('Populated ' + valModified + ' blank cleaning team mapping(s) in ValidationConfig Column C.');
      }
    }

    // 1c. Deduplicate disabled Coordinator rows where an identical active (email + q_columns) row exists
    var coordSheet = ss.getSheetByName('Coordinators');
    if (coordSheet && coordSheet.getLastRow() >= 2) {
      var cData = coordSheet.getDataRange().getValues();
      var cHeaders = cData[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
      var emailIdx = cHeaders.indexOf('email');
      var qColsIdx = cHeaders.indexOf('q_columns');
      var incIdx = cHeaders.indexOf('include_in_email');
      if (emailIdx !== -1 && qColsIdx !== -1 && incIdx !== -1) {
        var activeKeys = {};
        for (var cr = 1; cr < cData.length; cr++) {
          var incVal = String(cData[cr][incIdx] || '').trim().toUpperCase();
          if (incVal === 'Y' || incVal === 'YES' || incVal === 'TRUE') {
            var k = String(cData[cr][emailIdx] || '').trim().toLowerCase() + '|' + String(cData[cr][qColsIdx] || '').replace(/\s+/g, '').toUpperCase();
            activeKeys[k] = true;
          }
        }
        var removedCoords = 0;
        for (var dr = cData.length - 1; dr >= 1; dr--) {
          var dInc = String(cData[dr][incIdx] || '').trim().toUpperCase();
          var dKey = String(cData[dr][emailIdx] || '').trim().toLowerCase() + '|' + String(cData[dr][qColsIdx] || '').replace(/\s+/g, '').toUpperCase();
          if ((dInc === 'N' || dInc === 'NO' || dInc === 'FALSE') && activeKeys[dKey]) {
            coordSheet.deleteRow(dr + 1);
            removedCoords++;
          }
        }
        if (removedCoords > 0) {
          log.push('Removed ' + removedCoords + ' redundant disabled Coordinator row(s).');
        }
      }
    }

    // 2. Repair Sermon Topics IMPORTRANGE formulas
    var sermonId = PropertiesService.getScriptProperties().getProperty(CONFIG.SERMON_TOPICS_SHEET_PROPERTY) ||
                   (typeof APP_CONFIG !== 'undefined' ? APP_CONFIG.DEFAULT_SERMON_TOPICS_ID : '') ||
                   '17qRYYuZTV3PWf92CnMmqwARaC2bk4BUPxz8DQ_9f3hU';
    if (sermonId) {
      try {
        var sermonSs = SpreadsheetApp.openById(sermonId);
        var hwsUrl = ss.getUrl().split('#')[0].split('?')[0];
        if (!hwsUrl.endsWith('/edit')) hwsUrl += '/edit';

        var quarters = CONFIG.QUARTERS || ['Q1', 'Q2', 'Q3', 'Q4'];
        var sections = [
          { formulaRow: 4,  end: 17, hwsRange: "AY6:AY19" },
          { formulaRow: 20, end: 33, hwsRange: "AY51:AY64" },
          { formulaRow: 36, end: 49, hwsRange: "AY66:AY79" },
          { formulaRow: 52, end: 65, hwsRange: "AY81:AY94" }
        ];

        var sCount = 0;
        quarters.forEach(function(qName) {
          var qSheet = sermonSs.getSheetByName(qName);
          if (!qSheet) return;
          sections.forEach(function(sec) {
            var cell = qSheet.getRange(sec.formulaRow, 2);
            var f = cell.getFormula();
            if (!f || f.indexOf(sec.hwsRange) === -1) {
              cell.setFormula('=IMPORTRANGE("' + hwsUrl + '", "' + qName + '!' + sec.hwsRange + '")');
              if (sec.end > sec.formulaRow) {
                qSheet.getRange(sec.formulaRow + 1, 2, sec.end - sec.formulaRow, 1).clearContent();
              }
              sCount++;
            }
          });
        });
        log.push('Verified/Repaired ' + sCount + ' Sermon Topics IMPORTRANGE formulas in "' + sermonSs.getName() + '".');
      } catch (sErr) {
        log.push('Warning: Could not repair Sermon Topics sheet: ' + sErr.message);
      }
    }

    // 3. Verify/Repair ESSC Schedule Timezone
    var esscId = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.ESSC_SCHEDULE_ID) || '1YupXj1wjmjTefQCPArSAJKNy1aF95CIxMxUV6JeFi3w';
    if (esscId) {
      try {
        var esscSs = SpreadsheetApp.openById(esscId);
        if (esscSs.getSpreadsheetTimeZone() !== 'America/Los_Angeles') {
          esscSs.setSpreadsheetTimeZone('America/Los_Angeles');
          log.push('Updated ESSC Schedule (' + esscId + ') timezone to America/Los_Angeles.');
        } else {
          log.push('ESSC Schedule timezone already America/Los_Angeles.');
        }
      } catch (eErr) {
        log.push('Warning: Could not check/set ESSC Schedule timezone: ' + eErr.message);
      }
    }

    var resultMsg = log.join('\n');
    Logger.log(resultMsg);
    if (ui) {
      ui.alert('2027 Live Spreadsheets Repair Complete', resultMsg, ui.ButtonSet.OK);
    } else {
      ss.toast('Repairs complete: ' + log.length + ' operations executed.', 'Live Repair Complete', 8);
    }
    return { success: true, log: log };
  });
}

function getSpeakerLanguages() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var workerSheet = ss.getSheetByName(CONFIG.WORKERS_SHEET);
  if (!workerSheet) return {};
  
  var data = workerSheet.getDataRange().getValues();
  var result = {}; 
  
  if (data && data.length > 0) {
    var detected = detectWorkersHeaderRow_(data);
    var headers = detected.headers.map(function(h) { return String(h).trim().toLowerCase(); });
    var nameIdx = -1, chineseIdx = -1, englishIdx = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var h = headers[i];
      if (nameIdx === -1 && (h.indexOf('schedule name') !== -1 || h === 'name')) nameIdx = i;
      if (chineseIdx === -1 && (h.indexOf('chinese') !== -1 || h.indexOf('中文') !== -1)) chineseIdx = i;
      if (englishIdx === -1 && (h.indexOf('english') !== -1 || h.indexOf('英文') !== -1)) englishIdx = i;
    }
    
    if (chineseIdx === -1) chineseIdx = 15; // Column P
    if (englishIdx === -1) englishIdx = 16; // Column Q
    if (nameIdx === -1) nameIdx = 4; // Column E
    
    for (var r = detected.dataStartIdx; r < data.length; r++) {
      if (!data[r]) continue;
      var name = String(data[r][nameIdx] || '').trim();
      if (!name) continue;
      
      var isChinese = data[r][chineseIdx] === true || String(data[r][chineseIdx]).toUpperCase() === 'TRUE';
      var isEnglish = data[r][englishIdx] === true || String(data[r][englishIdx]).toUpperCase() === 'TRUE';
      
      if (isChinese) {
        result[name.toLowerCase()] = "CHINESE";
      } else if (isEnglish) {
        result[name.toLowerCase()] = "ENGLISH";
      }
    }
  }
  return result;
}

// -----------------------------------------------------------------------------
// WORKER DIRECTORY SCHEMA MANAGEMENT
// -----------------------------------------------------------------------------

/**
 * Resolves column positions based on headers to eliminate hardcoded dependencies.
 * Used by all scripts that need to read worker data.
 * @param {Sheet} sheet The Workers sheet object
 * @returns {Object} schema The resolved column indices (0-indexed)
 */
function getWorkerSchema(sheet) {
  if (!sheet) return null;
  var maxCols = Math.min(sheet.getMaxColumns(), 60);
  var maxRows = Math.min(sheet.getMaxRows(), 5);
  var sampleData = sheet.getRange(1, 1, maxRows, maxCols).getValues();
  var detected = detectWorkersHeaderRow_(sampleData);
  var headers = detected.headers;
  var schema = {
    emailCol: -1,
    nameCols: [],
    statusCol: -1,
    idCol: -1,
    qualCol: -1,
    teamsCol: -1,
    chineseNameCol: -1,
    phoneCol: -1,
    headerRowIdx: detected.headerRowIdx,
    dataStartIdx: detected.dataStartIdx
  };
  
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c]).trim().toLowerCase();
    if (!h) continue;
    if (h.indexOf('email') !== -1 && schema.emailCol === -1) schema.emailCol = c;
    else if (h.indexOf('chinese') !== -1 || h.indexOf('中文') !== -1) schema.chineseNameCol = c;
    else if (h.indexOf('phone') !== -1 || h.indexOf('contact') !== -1 || h.indexOf('tel') !== -1) schema.phoneCol = c;
    else if (h.indexOf('name') !== -1 || h.indexOf('worker') !== -1 || h.indexOf('member') !== -1 || h.indexOf('schedule name') !== -1) {
      if (schema.nameCols.indexOf(c) === -1) schema.nameCols.push(c);
    }
    else if (h === 'status') schema.statusCol = c;
    else if (h === 'worker id') schema.idCol = c;
    else if (h === 'qualifications') schema.qualCol = c;
    else if (h === 'teams') schema.teamsCol = c;
  }
  
  if (schema.emailCol === -1 || schema.nameCols.length === 0) {
    throw new Error("Could not find required 'Name' or 'Email' column on the Workers tab. Please ensure headers are correct.");
  }
  return schema;
}

/**
 * Appends missing canonical columns to the Workers tab without disturbing existing data.
 * Useful for bootstrapping legacy sheets for multi-church deployment.
 * @param {Spreadsheet} ss The active spreadsheet
 */
function ensureWorkersSchema(ss) {
  var sheet = ss.getSheetByName(CONFIG.WORKERS_SHEET);
  if (!sheet) return;
  
  var maxCols = Math.min(sheet.getMaxColumns(), 60);
  var sampleData = sheet.getRange(1, 1, Math.min(sheet.getMaxRows(), 5), maxCols).getValues();
  var detected = detectWorkersHeaderRow_(sampleData);
  var headerRow = detected.headerRowIdx + 1; // 1-indexed row
  var headersRange = sheet.getRange(headerRow, 1, 1, sheet.getMaxColumns());
  var headers = headersRange.getValues()[0];
  var requiredHeaders = ["Worker ID", "Name", "Chinese Name", "Email", "Phone", "Status", "Teams", "Qualifications", "Weekday Availability", "Added On", "Notes"];
  
  var existingHeaders = headers.map(function(h) { return String(h).trim().toLowerCase(); });
  var colsToAdd = [];
  
  for (var i = 0; i < requiredHeaders.length; i++) {
    var target = requiredHeaders[i].toLowerCase();
    var found = false;
    // Allow slight variations to pass
    for (var j = 0; j < existingHeaders.length; j++) {
      if (existingHeaders[j] === target || existingHeaders[j].indexOf(target) !== -1 || target.indexOf(existingHeaders[j]) !== -1) {
        if (existingHeaders[j].length > 0) {
           found = true;
           break;
        }
      }
    }
    // Strict check for exactly required new columns
    if (!found || (target === 'qualifications' && existingHeaders.indexOf('qualifications') === -1) || (target === 'worker id' && existingHeaders.indexOf('worker id') === -1) || (target === 'status' && existingHeaders.indexOf('status') === -1)) {
       if (existingHeaders.indexOf(target) === -1) {
         colsToAdd.push(requiredHeaders[i]);
       }
    }
  }
  
  if (colsToAdd.length > 0) {
    var lastCol = 0;
    for (var c = headers.length - 1; c >= 0; c--) {
      if (String(headers[c]).trim() !== "") {
        lastCol = c + 1;
        break;
      }
    }
    if (lastCol === 0) lastCol = 1;
    
    // Auto-expand sheet if we need more columns
    if (lastCol + colsToAdd.length > sheet.getMaxColumns()) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), colsToAdd.length);
    }
    
    sheet.getRange(headerRow, lastCol + 1, 1, colsToAdd.length)
         .setValues([colsToAdd])
         .setFontWeight("bold");
         
    SpreadsheetApp.getActiveSpreadsheet().toast("Added new scalable columns to Workers tab: " + colsToAdd.join(", "), "Migration Complete");
  }
}

/**
 * Rebuilds the Teams sheet in long format from Workers.Teams only if Teams does not use
 * formulas or a wide multi-column layout. Always safely purges chunked script caches.
 */
function syncRosters() {
  try {
    return executeWithLockAndRetry_('syncRosters', function() {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var workersSheet = ss.getSheetByName(CONFIG.WORKERS_SHEET);
      var teamsSheet = ss.getSheetByName(CONFIG.TEAMS_SHEET);
      if (!workersSheet || !teamsSheet) return { success: false, error: 'Missing Workers or Teams sheet' };
      
      var sc = CacheService.getScriptCache();
      
      // Inspect teamsSheet: check if any cell in rows 1-3 contains a formula or if it uses wide multi-column layout
      var teamsLastCol = teamsSheet.getLastColumn();
      var teamsLastRow = Math.min(teamsSheet.getLastRow(), 5);
      var hasFormula = false;
      if (teamsLastRow >= 1 && teamsLastCol >= 1) {
        var formulas = teamsSheet.getRange(1, 1, teamsLastRow, teamsLastCol).getFormulas();
        for (var fr = 0; fr < formulas.length; fr++) {
          for (var fc = 0; fc < formulas[fr].length; fc++) {
            if (formulas[fr][fc] && String(formulas[fr][fc]).trim().indexOf('=') === 0) {
              hasFormula = true;
              break;
            }
          }
          if (hasFormula) break;
        }
      }
      
      var isWide = teamsLastCol > 2;
      var schema = getWorkerSchema(workersSheet);
      
      // If teamsSheet uses live formulas or wide multi-column layout, or schema has no teamsCol, preserve it!
      if (hasFormula || isWide || !schema || schema.teamsCol === -1) {
        cacheRemoveLarge(sc, 'teams_and_workers');
        cacheRemoveLarge(sc, 'workers_day_avail');
        sc.remove('workers_data');
        if (typeof bumpAllQuarterVersions === 'function') bumpAllQuarterVersions();
        return { success: true, mode: 'formula_preserved' };
      }
      
      var data = workersSheet.getDataRange().getValues();
      var longFormat = [["Team", "Name"]]; // Row 1 headers for the long format
      var dataStart = (schema.dataStartIdx !== undefined) ? schema.dataStartIdx : 2;
      var hasWorkerTeamText = false;
      
      for (var r = dataStart; r < data.length; r++) {
        var status = "";
        if (schema.statusCol !== -1) {
           status = String(data[r][schema.statusCol]).trim().toLowerCase();
        }
        if (status === 'inactive') continue; // Skip inactive workers
        
        var name = String(data[r][schema.nameCols[0]]).trim();
        var teamsStr = String(data[r][schema.teamsCol]).trim();
        
        if (!name || !teamsStr) continue;
        hasWorkerTeamText = true;
        
        var workerTeams = teamsStr.split(/[,;]/).map(function(t) { return t.trim(); }).filter(function(t) { return t.length > 0; });
        for (var t = 0; t < workerTeams.length; t++) {
          longFormat.push([workerTeams[t], name]);
        }
      }
      
      if (!hasWorkerTeamText) {
        cacheRemoveLarge(sc, 'teams_and_workers');
        cacheRemoveLarge(sc, 'workers_day_avail');
        sc.remove('workers_data');
        if (typeof bumpAllQuarterVersions === 'function') bumpAllQuarterVersions();
        return { success: true, mode: 'formula_preserved' };
      }
      
      // Sort by team name, then worker name
      var headersRow = longFormat.shift();
      longFormat.sort(function(a, b) {
        var teamCmp = a[0].localeCompare(b[0]);
        if (teamCmp !== 0) return teamCmp;
        return a[1].localeCompare(b[1]);
      });
      longFormat.unshift(headersRow);
      
      teamsSheet.clearContent();
      if (longFormat.length > 0) {
         teamsSheet.getRange(1, 1, longFormat.length, 2).setValues(longFormat);
      }
      
      cacheRemoveLarge(sc, 'teams_and_workers');
      cacheRemoveLarge(sc, 'workers_day_avail');
      sc.remove('workers_data');
      if (typeof bumpAllQuarterVersions === 'function') bumpAllQuarterVersions();
      return { success: true, mode: 'rebuilt' };
    });
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function showSetupWizard() {
  var ui = SpreadsheetApp.getUi();
  var htmlOutput = HtmlService.createHtmlOutputFromFile('SetupWizard')
      .setTitle('Church Setup Wizard — Visual Worker Directory Import')
      .setWidth(980)
      .setHeight(720);
  ui.showModalDialog(htmlOutput, 'Church Setup Wizard');
}

/**
 * Returns available sheet tab names and the standard Garden Grove teams catalog.
 */
function getWizardInitialData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets().map(function(s) { return s.getName(); });
    
    // Complete Garden Grove Teams Catalog
    var ggTeams = [
      // Worship & Service
      "Board On Duty", "Sermon Speaker (Chinese)", "Sermon Speaker (English)",
      "Interpreter (Chinese to English)", "Interpreter (English to Chinese)",
      "Hymn Leader", "Pianist", "Usher", "AV", "AV Support",
      "Sat. AM Prayer", "Lunch Grace", "Sat Lunch Prayer",
      // Fellowships & Bible Classes
      "English Bible Fellowship", "Sisters' Fellowship", "Vessel Fellowship",
      "Truth Seeker Class", "Joy Fellowship", "Agape",
      // Religious Education
      "K", "E1", "E2", "J1", "J2", "Sr",
      // Cooking
      "Cooking Team 1", "Cooking Team 2", "Cooking Team 3", "Cooking Team 4",
      "Cooking Team 5", "Cooking Team 6", "Cooking Team 7", "Cooking Team 8",
      // Cleaning
      "Friday Chapel Cleaning", "Saturday Chapel Cleaning",
      "Saturday Side Chapel Cleaning", "Saturday Trash / Love Hall",
      "Bro Bathroom (Shed)", "Sis Bathroom (Shed)"
    ];
    
    return {
      success: true,
      sheets: sheets,
      defaultTeams: ggTeams
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Fetches headers and up to 5 preview rows from a designated tab in the workbook.
 */
function fetchTabHeadersAndPreview(sheetName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sheet tab not found: ' + sheetName };
    
    var data = sheet.getDataRange().getValues();
    if (data.length === 0) return { success: false, error: 'Selected sheet is empty.' };
    
    var headerRowIdx = 0;
    if (data.length > 1) {
      var row0EmptyOrNums = 0;
      for (var c = 0; c < data[0].length; c++) {
        var v = data[0][c];
        if (v === '' || !isNaN(Number(v))) row0EmptyOrNums++;
      }
      if (row0EmptyOrNums > data[0].length * 0.7) {
        headerRowIdx = 1; // Row 2 has text headers
      }
    }
    
    var headers = data[headerRowIdx].map(function(h) { return String(h !== null && h !== undefined ? h : '').trim(); });
    var previewRows = [];
    for (var r = headerRowIdx + 1; r < Math.min(data.length, headerRowIdx + 6); r++) {
      previewRows.push(data[r].map(function(cell) { return String(cell !== null && cell !== undefined ? cell : '').trim(); }));
    }
    
    return {
      success: true,
      headerRow: headerRowIdx + 1,
      headers: headers,
      totalRows: data.length - (headerRowIdx + 1),
      previewRows: previewRows
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Parses raw text pasted from clipboard (Excel or Google Sheets).
 */
function parsePastedTable(rawText) {
  try {
    var data = parseCSV_(rawText);
    if (data.length === 0) return { success: false, error: 'Pasted text is empty.' };
    
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var previewRows = [];
    for (var r = 1; r < Math.min(data.length, 6); r++) {
      previewRows.push(data[r]);
    }
    
    return {
      success: true,
      headers: headers,
      totalRows: data.length - 1,
      previewRows: previewRows
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Helper to parse CSV/TSV supporting quoted fields with commas, tabs, and newlines.
 */
function parseCSV_(text) {
  var isTabDelimited = (text.indexOf('\t') !== -1 && text.indexOf('\t') < text.indexOf('\n'));
  var delimiter = isTabDelimited ? '\t' : ',';
  
  var p = '', row = [''], result = [row], i = 0, r = 0, s = true;
  for (var k = 0; k < text.length; k++) {
    var c = text[k];
    if (c === '"') {
      if (s && p === '"') { row[i] += '"'; }
      s = !s;
    } else if (c === delimiter && s) {
      row[++i] = '';
    } else if ((c === '\n' || c === '\r') && s) {
      if (p !== '\r' || c !== '\n') {
        row = [''];
        result[++r] = row;
        i = 0;
      }
    } else {
      row[i] += c;
    }
    p = c;
  }
  var clean = [];
  for (var a = 0; a < result.length; a++) {
    var hasVal = false;
    for (var b = 0; b < result[a].length; b++) {
      result[a][b] = String(result[a][b] || '').trim();
      if (result[a][b]) hasVal = true;
    }
    if (hasVal) clean.push(result[a]);
  }
  return clean;
}

/**
 * Visual Column Mapper Import Processor: evaluates mapping, matches teams, and runs pre-flight diff.
 */
function processVisualWorkerImport(config) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceData = [];
    
    if (config.sourceType === 'sheet') {
      var sheet = ss.getSheetByName(config.sheetName);
      if (!sheet) return { success: false, error: 'Source sheet not found: ' + config.sheetName };
      var allData = sheet.getDataRange().getValues();
      var startRow = (config.headerRow || 1);
      for (var r = startRow; r < allData.length; r++) {
        sourceData.push(allData[r]);
      }
    } else {
      var parsed = parseCSV_(config.rawText);
      if (parsed.length < 2) return { success: false, error: 'Pasted data must contain a header row and at least one data row.' };
      sourceData = parsed.slice(1);
    }
    
    var mapping = config.mapping;
    var workersSheet = ss.getSheetByName(CONFIG.WORKERS_SHEET);
    if (!workersSheet) return { success: false, error: 'Workers sheet not found.' };
    ensureWorkersSchema(ss);
    var existingSchema = getWorkerSchema(workersSheet);
    var existingData = workersSheet.getDataRange().getValues();
    
    var existingMap = {};
    for (var r = 2; r < existingData.length; r++) {
      var n = String(existingData[r][existingSchema.nameCols[0]]).trim().toLowerCase();
      if (n) existingMap[n] = { row: r + 1, data: existingData[r] };
    }
    
    var records = [];
    var stats = { new: 0, updated: 0, unchanged: 0, conflicts: 0 };
    
    for (var i = 0; i < sourceData.length; i++) {
      var row = sourceData[i];
      var name = '';
      if (mapping.nameMode === 'split') {
        var fn = mapping.firstNameCol !== -1 && row[mapping.firstNameCol] !== undefined ? String(row[mapping.firstNameCol]).trim() : '';
        var ln = mapping.lastNameCol !== -1 && row[mapping.lastNameCol] !== undefined ? String(row[mapping.lastNameCol]).trim() : '';
        name = (fn + ' ' + ln).trim();
      } else {
        name = mapping.nameCol !== -1 && row[mapping.nameCol] !== undefined ? String(row[mapping.nameCol]).trim() : '';
      }
      
      if (!name) continue;
      
      var chName = mapping.chineseNameCol !== -1 && row[mapping.chineseNameCol] !== undefined ? String(row[mapping.chineseNameCol]).trim() : '';
      var email = mapping.emailCol !== -1 && row[mapping.emailCol] !== undefined ? String(row[mapping.emailCol]).trim() : '';
      var phone = mapping.phoneCol !== -1 && row[mapping.phoneCol] !== undefined ? String(row[mapping.phoneCol]).trim() : '';
      var qual = mapping.qualCol !== -1 && row[mapping.qualCol] !== undefined ? String(row[mapping.qualCol]).trim() : '';
      
      var status = mapping.defaultStatus || 'Active';
      if (mapping.statusCol !== -1 && row[mapping.statusCol] !== undefined) {
        var sVal = String(row[mapping.statusCol]).trim();
        if (sVal) status = sVal;
      }
      
      // Determine teams
      var teamsList = [];
      if (mapping.teamMode === 'matrix' && mapping.teamMatrix && mapping.teamMatrix.length > 0) {
        for (var m = 0; m < mapping.teamMatrix.length; m++) {
          var entry = mapping.teamMatrix[m];
          var cellVal = row[entry.colIdx];
          if (cellVal !== undefined && cellVal !== null) {
            var cStr = String(cellVal).trim().toLowerCase();
            if (cellVal === true || cStr === 'true' || cStr === 'yes' || cStr === 'y' || cStr === 'x' || cStr === '1' || cStr === '✓' || cStr === 'checked') {
              if (teamsList.indexOf(entry.teamName) === -1) {
                teamsList.push(entry.teamName);
              }
            }
          }
        }
      } else if (mapping.teamMode === 'single' && mapping.singleTeamCol !== -1) {
        var tStr = row[mapping.singleTeamCol] !== undefined ? String(row[mapping.singleTeamCol]).trim() : '';
        if (tStr) {
          var parsedTeams = tStr.split(/[,;/]/).map(function(t) { return t.trim(); }).filter(function(t) { return t.length > 0; });
          for (var pt = 0; pt < parsedTeams.length; pt++) {
            if (teamsList.indexOf(parsedTeams[pt]) === -1) teamsList.push(parsedTeams[pt]);
          }
        }
      }
      
      var teamsStr = teamsList.join(', ');
      var rec = {
        Name: name,
        ChineseName: chName,
        Email: email,
        Phone: phone,
        Teams: teamsStr,
        Qualifications: qual,
        Status: status
      };
      
      var action = 'new';
      var nKey = name.toLowerCase();
      if (existingMap[nKey]) {
        var eEmail = existingSchema.emailCol !== -1 ? String(existingMap[nKey].data[existingSchema.emailCol]).trim() : '';
        var eTeams = existingSchema.teamsCol !== -1 ? String(existingMap[nKey].data[existingSchema.teamsCol]).trim() : '';
        var eQual = existingSchema.qualCol !== -1 ? String(existingMap[nKey].data[existingSchema.qualCol]).trim() : '';
        var eStatus = existingSchema.statusCol !== -1 ? String(existingMap[nKey].data[existingSchema.statusCol]).trim() : '';
        
        if (eEmail === email && eTeams === teamsStr && eQual === qual && eStatus === status) {
          action = 'unchanged';
        } else {
          action = 'update';
        }
      }
      
      records.push({ action: action, record: rec });
      stats[action]++;
    }
    
    var stagingId = 'import_' + new Date().getTime();
    CacheService.getUserCache().put(stagingId, JSON.stringify(records), 600); // 10 minutes
    
    return {
      success: true,
      stagingId: stagingId,
      records: records,
      stats: stats
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Legacy wrapper for CSV text import flow.
 */
function processWorkerImportFlow(csvText) {
  return processVisualWorkerImport({
    sourceType: 'raw',
    rawText: csvText,
    mapping: {
      nameMode: 'single',
      nameCol: 0,
      emailCol: 1,
      teamsCol: 2,
      qualCol: 3,
      statusCol: 4,
      chineseNameCol: -1,
      phoneCol: -1,
      teamMode: 'single',
      singleTeamCol: 2,
      defaultStatus: 'Active'
    }
  });
}

function commitWorkerImport(stagingId, options) {
  try {
    return executeWithLockAndRetry_('commitWorkerImport', function() {
      var cached = CacheService.getUserCache().get(stagingId);
      if (!cached) return { success: false, error: 'Import session expired. Please try again.' };
      
      var records = JSON.parse(cached);
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(CONFIG.WORKERS_SHEET);
      var schema = getWorkerSchema(sheet);
      var existingData = sheet.getDataRange().getValues();
      
      var existingMap = {};
      for (var r = 2; r < existingData.length; r++) {
         var n = String(existingData[r][schema.nameCols[0]]).trim().toLowerCase();
         if (n) existingMap[n] = { row: r + 1, data: existingData[r] };
      }
      
      // Build map of team names to column index on Workers tab (for checkbox support)
      var headersRow1 = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues()[0];
      var headersRow2 = sheet.getRange(2, 1, 1, sheet.getMaxColumns()).getValues()[0];
      var teamColMap = {};
      for (var c = 0; c < sheet.getMaxColumns(); c++) {
        var h1 = String(headersRow1[c] || '').trim().toLowerCase();
        var h2 = String(headersRow2[c] || '').trim().toLowerCase();
        if (h1 && teamColMap[h1] === undefined) teamColMap[h1] = c + 1;
        if (h2 && teamColMap[h2] === undefined) teamColMap[h2] = c + 1;
      }
      
      var newRows = [];
      var updates = [];
      
      for (var i = 0; i < records.length; i++) {
         var r = records[i];
         if (r.action === 'unchanged' || r.action === 'conflict') continue;
         
         var rec = r.record;
         var nKey = String(rec.Name).trim().toLowerCase();
         if (r.action === 'update' && existingMap[nKey]) {
            updates.push({ row: existingMap[nKey].row, rec: rec });
         } else if (r.action === 'new') {
            var newRow = new Array(sheet.getMaxColumns());
            for (var c = 0; c < newRow.length; c++) newRow[c] = '';
            if (schema.nameCols.length > 0) newRow[schema.nameCols[0]] = rec.Name;
            if (schema.chineseNameCol !== -1 && rec.ChineseName) newRow[schema.chineseNameCol] = rec.ChineseName;
            if (schema.emailCol !== -1) newRow[schema.emailCol] = rec.Email;
            if (schema.phoneCol !== -1 && rec.Phone) newRow[schema.phoneCol] = rec.Phone;
            if (schema.teamsCol !== -1) newRow[schema.teamsCol] = rec.Teams;
            if (schema.qualCol !== -1) newRow[schema.qualCol] = rec.Qualifications;
            if (schema.statusCol !== -1) newRow[schema.statusCol] = rec.Status;
            
            // Checkboxes on new row if team columns exist
            if (rec.Teams) {
              var teamsArr = rec.Teams.split(',').map(function(t){ return t.trim().toLowerCase(); });
              for (var t = 0; t < teamsArr.length; t++) {
                if (teamColMap[teamsArr[t]] !== undefined) {
                  newRow[teamColMap[teamsArr[t]] - 1] = true;
                }
              }
            }
            newRows.push(newRow);
         }
      }
      
      // Process updates (batched in-memory to prevent quota/rate limits - Fix ID #14)
      if (updates.length > 0) {
        var maxCols = sheet.getMaxColumns();
        var minUpdateRow = updates[0].row;
        var maxUpdateRow = updates[0].row;
        for (var i = 1; i < updates.length; i++) {
          if (updates[i].row < minUpdateRow) minUpdateRow = updates[i].row;
          if (updates[i].row > maxUpdateRow) maxUpdateRow = updates[i].row;
        }
        var updateNumRows = maxUpdateRow - minUpdateRow + 1;
        var updateRange = sheet.getRange(minUpdateRow, 1, updateNumRows, maxCols);
        var updateValues = updateRange.getValues();

        for (var i = 0; i < updates.length; i++) {
          var u = updates[i];
          var rowOffset = u.row - minUpdateRow;
          var rowVals = updateValues[rowOffset];
          if (schema.emailCol !== -1 && u.rec.Email) rowVals[schema.emailCol] = u.rec.Email;
          if (schema.chineseNameCol !== -1 && u.rec.ChineseName) rowVals[schema.chineseNameCol] = u.rec.ChineseName;
          if (schema.phoneCol !== -1 && u.rec.Phone) rowVals[schema.phoneCol] = u.rec.Phone;
          if (schema.teamsCol !== -1 && u.rec.Teams) rowVals[schema.teamsCol] = u.rec.Teams;
          if (schema.qualCol !== -1 && u.rec.Qualifications) rowVals[schema.qualCol] = u.rec.Qualifications;
          if (schema.statusCol !== -1 && u.rec.Status) rowVals[schema.statusCol] = u.rec.Status;
          
          if (u.rec.Teams) {
            var teamsArr = u.rec.Teams.split(',').map(function(t){ return t.trim().toLowerCase(); });
            for (var t = 0; t < teamsArr.length; t++) {
              if (teamColMap[teamsArr[t]] !== undefined) {
                rowVals[teamColMap[teamsArr[t]] - 1] = true;
              }
            }
          }
        }
        updateRange.setValues(updateValues);
      }
      
      // Process new rows (guarded with row expansion - Fix ID #7)
      if (newRows.length > 0) {
         var lastRow = sheet.getLastRow();
         var startNewRow = lastRow + 1;
         var totalCols = sheet.getMaxColumns();
         var neededRows = startNewRow + newRows.length - 1;
         var currentMaxRows = sheet.getMaxRows();
         if (neededRows > currentMaxRows) {
           sheet.insertRowsAfter(currentMaxRows, neededRows - currentMaxRows);
         }
         var lastRowFormulas = lastRow >= 2 ? sheet.getRange(lastRow, 1, 1, totalCols).getFormulas()[0] : [];
         
         sheet.getRange(startNewRow, 1, newRows.length, totalCols).setValues(newRows);
         
         // Copy down formulas for any column that had a formula on lastRow and where newRows was blank
         if (lastRowFormulas && lastRowFormulas.length > 0) {
           for (var colIdx = 0; colIdx < totalCols; colIdx++) {
             var f = lastRowFormulas[colIdx];
             if (f && f !== '') {
               var hasValues = false;
               for (var nr = 0; nr < newRows.length; nr++) {
                 if (newRows[nr][colIdx] !== '' && newRows[nr][colIdx] !== null && typeof newRows[nr][colIdx] !== 'undefined') {
                   hasValues = true;
                   break;
                 }
               }
               if (!hasValues) {
                 var sourceCell = sheet.getRange(lastRow, colIdx + 1, 1, 1);
                 var targetRange = sheet.getRange(startNewRow, colIdx + 1, newRows.length, 1);
                 sourceCell.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
               }
             }
           }
         }
      }
      
      // Clean up
      CacheService.getUserCache().remove(stagingId);
      
      // Rebuild Teams & clear cache
      syncRosters();
      
      return { success: true };
    });
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Update a worker's status (e.g. Active to Inactive).
 */
function setWorkerStatus(workerName, newStatus) {
  try {
    return executeWithLockAndRetry_('setWorkerStatus', function() {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(CONFIG.WORKERS_SHEET);
      var schema = getWorkerSchema(sheet);
      if (!schema || schema.nameCols.length === 0 || schema.statusCol === -1) {
         return { success: false, error: 'Workers sheet missing Name or Status columns' };
      }
      
      var data = sheet.getDataRange().getValues();
      var targetLower = String(workerName).trim().toLowerCase();
      var found = false;
      
      for (var r = 2; r < data.length; r++) {
         var n = String(data[r][schema.nameCols[0]]).trim().toLowerCase();
         if (n === targetLower) {
            sheet.getRange(r + 1, schema.statusCol + 1).setValue(newStatus);
            found = true;
            break;
         }
      }
      
      if (found) {
         syncRosters();
         return { success: true };
      } else {
         return { success: false, error: 'Worker not found' };
      }
    });
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Rename a worker across the Workers sheet and Q1-Q4 sheets.
 */
function renameWorker(oldName, newName, options) {
  try {
    return executeWithLockAndRetry_('renameWorker', function() {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var oldLower = String(oldName).trim().toLowerCase();
      var newTrim = String(newName).trim();
      if (!oldLower || !newTrim) return { success: false, error: 'Invalid names' };
      
      var stats = { workersSheet: 0, qSheets: 0 };
      var isDryRun = options && options.dryRun;
      
      // 1. Rename in Workers tab
      var wSheet = ss.getSheetByName(CONFIG.WORKERS_SHEET);
      if (wSheet) {
         var wSchema = null;
         try { wSchema = getWorkerSchema(wSheet); } catch (e) {}
         if (wSchema && wSchema.nameCols.length > 0) {
            var wData = wSheet.getDataRange().getValues();
            var wFormulas = wSheet.getDataRange().getFormulas();
            var wModified = false;
            var startIdx = (wSchema.dataStartIdx !== undefined) ? wSchema.dataStartIdx : 2;
            for (var r = startIdx; r < wData.length; r++) {
               for (var i = 0; i < wSchema.nameCols.length; i++) {
                  var col = wSchema.nameCols[i];
                  if (wFormulas[r] && wFormulas[r][col] !== '') continue;
                  var n = String(wData[r][col] || '').trim().toLowerCase();
                  if (n === oldLower) {
                     stats.workersSheet++;
                     if (!isDryRun) {
                        wData[r][col] = newTrim;
                        wModified = true;
                     }
                  }
               }
            }
            if (wModified) {
               for (var r = 0; r < wData.length; r++) {
                  for (var c = 0; c < wData[r].length; c++) {
                     if (wFormulas[r] && wFormulas[r][c] !== '') {
                        wData[r][c] = wFormulas[r][c];
                     }
                  }
               }
               wSheet.getRange(1, 1, wData.length, wData[0].length).setValues(wData);
            }
         }
      }
      
      // 2. Rename in Q1-Q4 sheets
      var qSheets = CONFIG.QUARTERS;
      for (var s = 0; s < qSheets.length; s++) {
         var sheet = ss.getSheetByName(qSheets[s]);
         if (!sheet) continue;
         var data = sheet.getDataRange().getValues();
         var formulas = sheet.getDataRange().getFormulas();
         var qModified = false;
         for (var r = 0; r < data.length; r++) {
            for (var c = 0; c < data[r].length; c++) {
               if (formulas[r] && formulas[r][c] !== '') continue;
               if (String(data[r][c] || '').trim().toLowerCase() === oldLower) {
                  stats.qSheets++;
                  if (!isDryRun) {
                     data[r][c] = newTrim;
                     qModified = true;
                  }
               }
            }
         }
         if (qModified) {
            for (var r = 0; r < data.length; r++) {
               for (var c = 0; c < data[r].length; c++) {
                  if (formulas[r] && formulas[r][c] !== '') {
                     data[r][c] = formulas[r][c];
                  }
               }
            }
            sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
         }
      }
      
      if (!isDryRun && stats.workersSheet > 0) {
         syncRosters();
      }
      
      return { success: true, stats: stats };
    });
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Run preflight checks on assignments against active rosters.
 */
function validateWorkers() {
  // To be implemented fully in future as it requires mapping back all assignments.
  return { success: true, message: 'Validation stub.' };
}

// -----------------------------------------------------------------------------
// CHURCH CONFIGURATION ENGINE (Phase B)
// -----------------------------------------------------------------------------

function getChurchConfig(forceRefresh) {
  var cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    var cached = cache.get('church_config');
    if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = {
    profile: getChurchProfileFallback_(),
    serviceLayout: getServiceLayoutFallback_()
  };
  
  // 1. Load ChurchProfile if exists
  var profileSheet = ss.getSheetByName('ChurchProfile');
  if (profileSheet) {
    var data = profileSheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var key = String(data[r][0]).trim().toLowerCase();
      var val = String(data[r][1]).trim();
      if (key) {
        config.profile[key] = val;
      }
    }
  }
  
  // 2. Load ServiceLayoutConfig if exists
  var layoutSheet = ss.getSheetByName('ServiceLayoutConfig');
  if (layoutSheet) {
    var data = layoutSheet.getDataRange().getValues();
    var layout = [];
    if (data.length > 1) {
      var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
      var colMap = {};
      for (var i = 0; i < headers.length; i++) colMap[headers[i]] = i;
      
      for (var r = 1; r < data.length; r++) {
        var row = data[r];
        if (!String(row[colMap['day_type']]).trim()) continue;
        layout.push({
          day_type: String(row[colMap['day_type']]).trim().toUpperCase(),
          section_name: String(row[colMap['section_name']]).trim(),
          header_row: parseInt(row[colMap['header_row']], 10),
          start_row: parseInt(row[colMap['start_row']], 10),
          end_row: parseInt(row[colMap['end_row']], 10),
          workday_mask: String(row[colMap['workday_mask']]).trim().toUpperCase(),
          enabled: String(row[colMap['enabled']]).trim().toUpperCase()
        });
      }
      if (layout.length > 0) {
        config.serviceLayout = layout;
      }
    }
  }
  
  cache.put('church_config', JSON.stringify(config), 21600);
  return config;
}

function getChurchProfileFallback_() {
  var isAppConfig = typeof APP_CONFIG !== 'undefined';
  var editors = [];
  if (isAppConfig && APP_CONFIG.AUTHORIZED_EDITORS && Array.isArray(APP_CONFIG.AUTHORIZED_EDITORS)) {
    editors = APP_CONFIG.AUTHORIZED_EDITORS;
  } else if (typeof AUTHORIZED_EDITORS !== 'undefined' && Array.isArray(AUTHORIZED_EDITORS)) {
    editors = AUTHORIZED_EDITORS;
  }
  
  var biblePlan = '';
  if (isAppConfig && APP_CONFIG.BIBLE_PLAN_URL) {
    biblePlan = APP_CONFIG.BIBLE_PLAN_URL;
  } else if (typeof BIBLE_PLAN_URL !== 'undefined') {
    biblePlan = BIBLE_PLAN_URL;
  }

  return {
    church_code: (isAppConfig && APP_CONFIG.CHURCH_CODE) ? APP_CONFIG.CHURCH_CODE : 'GG',
    church_name: (isAppConfig && APP_CONFIG.CHURCH_NAME) ? APP_CONFIG.CHURCH_NAME : 'Garden Grove',
    church_name_zh: (isAppConfig && APP_CONFIG.CHURCH_NAME_ZH) ? APP_CONFIG.CHURCH_NAME_ZH : '園林教會',
    contact_email: (isAppConfig && APP_CONFIG.CONTACT_EMAIL) ? APP_CONFIG.CONTACT_EMAIL : 'gg.info@tjc.org',
    timezone: (isAppConfig && APP_CONFIG.TIMEZONE) ? APP_CONFIG.TIMEZONE : 'America/Los_Angeles',
    authorized_editors: editors.join(','),
    bible_plan_url: biblePlan,
    year: new Date().getFullYear(),
    week_anchor_day: 'SAT',
    cleaning_enabled: 'Y',
    cleaning_header_row: 58,
    cleaning_date_row: 59,
    cleaning_start_col: 3
  };
}

function getServiceLayoutFallback_() {
  // Hardcoded GG Universal 14-Row blocks (Rows 1–94)
  return [
    { day_type: 'MON', section_name: 'Evening', header_row: 5, start_row: 6, end_row: 19, workday_mask: 'MON', enabled: 'Y' },
    { day_type: 'TUE', section_name: 'Evening', header_row: 20, start_row: 21, end_row: 34, workday_mask: 'TUE', enabled: 'Y' },
    { day_type: 'WED', section_name: 'Evening', header_row: 35, start_row: 36, end_row: 49, workday_mask: 'WED', enabled: 'Y' },
    { day_type: 'FRI', section_name: 'Evening', header_row: 50, start_row: 51, end_row: 64, workday_mask: 'FRI', enabled: 'Y' },
    { day_type: 'SAT', section_name: 'AM', header_row: 65, start_row: 66, end_row: 79, workday_mask: 'SAT', enabled: 'Y' },
    { day_type: 'SAT', section_name: 'PM', header_row: 80, start_row: 81, end_row: 94, workday_mask: 'SAT', enabled: 'Y' }
  ];
}

function promptCompileRoleConfig() {
  var ui = SpreadsheetApp.getUi();
  var result = compileRoleConfig();
  if (result.success) {
    ui.alert("Success", result.message, ui.ButtonSet.OK);
  } else {
    ui.alert("Error", result.error, ui.ButtonSet.OK);
  }
}

/**
 * Reads RoleConfig and compiles it into ValidationConfig and MappingConfig.
 * This guarantees backwards compatibility with existing legacy scripts.
 */
function compileRoleConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var roleSheet = ss.getSheetByName('RoleConfig');
  if (!roleSheet) {
    return { success: false, error: 'RoleConfig tab not found.' };
  }
  
  var data = roleSheet.getDataRange().getValues();
  if (data.length < 2) {
    return { success: false, error: 'RoleConfig tab is empty.' };
  }
  
  var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var colMap = {};
  for (var i = 0; i < headers.length; i++) colMap[headers[i]] = i;
  
  var validationRows = [];
  var mappingRows = [];
  var seenQCols = {}; // For ValidationConfig deduping
  
  // Collision tracking
  var qColDayTypeMap = {}; // (q_column + day_type) -> role
  var templateRowMap = {}; // (day_type + section + template_row) -> role
  var errors = [];
  
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var roleKey = String(row[colMap['role_key']] || '').trim();
    if (!roleKey) continue; // Skip empty rows
    
    var enabled = String(row[colMap['enabled']] || '').trim().toUpperCase();
    if (enabled !== 'Y') continue;
    
    var qCol = String(row[colMap['q_column']] || '').trim().toUpperCase();
    var dayType = String(row[colMap['day_type']] || '').trim().toUpperCase();
    var section = String(row[colMap['section']] || '').trim().toUpperCase();
    var templateRow = parseInt(row[colMap['template_row']], 10);
    var labelEn = String(row[colMap['label_en']] || '').trim();
    var teamSources = String(row[colMap['team_sources']] || '').trim();
    
    if (!qCol) continue;
    
    // Collision Checks (per day and section)
    var qKey = qCol + "_" + dayType + "_" + section;
    if (qColDayTypeMap[qKey]) {
      errors.push("Collision: Multiple roles assigned to Master Column " + qCol + " for " + dayType + " " + section);
    }
    qColDayTypeMap[qKey] = true;
    
    if (!isNaN(templateRow) && templateRow > 0) {
      var tKey = dayType + "_" + section + "_" + templateRow + "_" + qCol;
      if (templateRowMap[tKey]) {
        errors.push("Collision: Master Column " + qCol + " already mapped to template row " + templateRow + " for " + dayType + " " + section);
      }
      templateRowMap[tKey] = true;
    }
    
    // Check against reserved RE columns (L through V)
    var reserved = ['L','M','N','O','P','Q','R','S','T','U','V'];
    if (reserved.indexOf(qCol) !== -1 && roleKey.indexOf('RE_') !== 0) {
      // Reserved for RE sync
    }
    
    // Build ValidationConfig row (Deduplicate by Q Column)
    if (!seenQCols[qCol] && teamSources) {
      validationRows.push([qCol, labelEn, teamSources]);
      seenQCols[qCol] = true;
    }
    
    // Build MappingConfig row (Only if template_row is set)
    if (!isNaN(templateRow) && templateRow > 0) {
      mappingRows.push([dayType, labelEn, section, qCol, templateRow]);
    }
  }
  
  if (errors.length > 0) {
    return { success: false, error: errors.join('\n') };
  }
  
  // Invalidate configuration cache
  try { CacheService.getScriptCache().remove('church_config'); } catch(e) {}
  
  // Write ValidationConfig
  var valSheet = ss.getSheetByName('ValidationConfig');
  if (!valSheet) {
    valSheet = ss.insertSheet('ValidationConfig');
  }
  valSheet.clear();
  var valHeaders = ["Master Column (Letter)", "Role / Description", "Teams Sheet Name(s)"];
  valSheet.getRange(1, 1, 1, valHeaders.length).setValues([valHeaders]).setFontWeight("bold");
  if (validationRows.length > 0) {
    valSheet.getRange(2, 1, validationRows.length, validationRows[0].length).setValues(validationRows);
  }
  valSheet.autoResizeColumns(1, valHeaders.length);
  valSheet.hideSheet();
  
  // Write MappingConfig
  var mapSheet = ss.getSheetByName('MappingConfig');
  if (!mapSheet) {
    mapSheet = ss.insertSheet('MappingConfig');
  }
  mapSheet.clear();
  var mapHeaders = ["Day Type (SAT/FRI etc)", "Role Name", "Section (AM/PM/CLEANING)", "Master Column (Letter)", "Template Row (Number)"];
  mapSheet.getRange(1, 1, 1, mapHeaders.length).setValues([mapHeaders]).setFontWeight("bold");
  if (mappingRows.length > 0) {
    mapSheet.getRange(2, 1, mappingRows.length, mappingRows[0].length).setValues(mappingRows);
  }
  mapSheet.autoResizeColumns(1, mapHeaders.length);
  mapSheet.hideSheet();
  
  return { success: true, message: 'Successfully compiled ' + validationRows.length + ' validation rules and ' + mappingRows.length + ' mapping rules.' };
}

function seedChurchProfileSheet(ss) {
  var sheet = ss.getSheetByName('ChurchProfile');
  if (!sheet) {
    sheet = ss.insertSheet('ChurchProfile');
  } else {
    sheet.clear();
  }
  var headers = ['key', 'value', 'note'];
  var fallback = getChurchProfileFallback_();
  var rows = [
    ['church_code', fallback.church_code || 'GG', 'Unique church identifier code (e.g. GG, CR, CP)'],
    ['church_name', fallback.church_name || 'Garden Grove', 'Full English Church Name'],
    ['church_name_zh', fallback.church_name_zh || '園林教會', 'Full Chinese Church Name'],
    ['contact_email', fallback.contact_email || 'gg.info@tjc.org', 'Admin / Info Dept reply-to email address'],
    ['timezone', fallback.timezone || 'America/Los_Angeles', 'Local IANA timezone identifier'],
    ['swr_location_tag', 'Garden Grove', 'Location tag for filtering SWR speaker synchronization'],
    ['authorized_editors', fallback.authorized_editors || '', 'Comma-separated authorized admin emails for monthly schedules'],
    ['bible_plan_url', fallback.bible_plan_url || '', 'URL to fetch monthly Bible reading plan (blank to skip)'],
    ['year', String(fallback.year || new Date().getFullYear()), 'Active calendar year for Holy Work Schedule'],
    ['week_anchor_day', fallback.week_anchor_day || 'SAT', 'Day of the week defining the end of a weekly cycle (SAT/SUN)'],
    ['cleaning_enabled', fallback.cleaning_enabled || 'Y', 'Whether cleaning duty groups are populated in monthly sheets (Y/N)'],
    ['cleaning_header_row', String(fallback.cleaning_header_row || 58), 'Row index where cleaning headers start in MonthlyTemplate'],
    ['cleaning_date_row', String(fallback.cleaning_date_row || 59), 'Row index where cleaning dates start in MonthlyTemplate'],
    ['cleaning_start_col', String(fallback.cleaning_start_col || 3), 'Start column index for cleaning duties (3 = Col C)']
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#d9ead3');
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  sheet.autoResizeColumns(1, headers.length);
  sheet.hideSheet();
  try { CacheService.getScriptCache().remove('church_config'); } catch(e) {}
  return sheet;
}

function seedServiceLayoutConfigSheet(ss) {
  var sheet = ss.getSheetByName('ServiceLayoutConfig');
  if (!sheet) {
    sheet = ss.insertSheet('ServiceLayoutConfig');
  } else {
    sheet.clear();
  }
  var headers = ['day_type', 'section_name', 'header_row', 'start_row', 'end_row', 'workday_mask', 'enabled'];
  var fallback = getServiceLayoutFallback_();
  var rows = fallback.map(function(item) {
    return [item.day_type, item.section_name, item.header_row, item.start_row, item.end_row, item.workday_mask, item.enabled];
  });
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#d9ead3');
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  sheet.autoResizeColumns(1, headers.length);
  sheet.hideSheet();
  try { CacheService.getScriptCache().remove('church_config'); } catch(e) {}
  return sheet;
}

function seedRoleConfigSheet(ss) {
  var sheet = ss.getSheetByName('RoleConfig');
  if (!sheet) {
    sheet = ss.insertSheet('RoleConfig');
  } else {
    sheet.clear();
  }
  var headers = ['role_key', 'label_en', 'label_zh', 'day_type', 'section', 'q_column', 'template_row', 'team_sources', 'slots', 'required', 'requires_qual', 'enabled'];
  var rows = [
    // Saturday AM
    ['SAT_AM_PRAYER', 'Prayer', '禱告', 'SAT', 'AM', 'K', 8, 'Prayer, 禱告', 1, 'Y', '', 'Y'],
    ['SAT_AM_PRAISE', 'Praise', '詩頌', 'SAT', 'AM', 'D', 9, 'Hymn Leader, 詩頌', 1, 'Y', '', 'Y'],
    ['SAT_AM_PIANIST', 'Accompanist', '司琴', 'SAT', 'AM', 'E', 10, 'Pianist, 司琴', 1, 'Y', 'PIANO', 'Y'],
    ['SAT_AM_SERMON', 'Sermon', '講道', 'SAT', 'AM', 'AY', 11, 'Speaker', 1, 'Y', 'SPEAK', 'Y'],
    ['SAT_AM_INTERPRETER', 'Interpreter', '翻譯', 'SAT', 'AM', 'C', 13, 'Interpreter (E-to-C), Interpreter (C-to-E), Interpreter, 翻譯', 1, 'Y', '', 'Y'],
    ['SAT_AM_AV1', 'A/V 1', '視聽 1', 'SAT', 'AM', 'H', 14, 'AV, 視聽', 1, 'Y', '', 'Y'],
    ['SAT_AM_AV2', 'A/V 2', '視聽 2', 'SAT', 'AM', 'I', 14, 'AV, 視聽', 1, 'Y', '', 'Y'],
    ['SAT_AM_USHER1', 'Usher 1', '接待 1', 'SAT', 'AM', 'F', 15, 'Usher, 接待', 1, 'Y', '', 'Y'],
    ['SAT_AM_USHER2', 'Usher 2', '接待 2', 'SAT', 'AM', 'G', 15, 'Usher, 接待', 1, 'Y', '', 'Y'],
    ['SAT_AM_GRACE', 'Grace', '謝飯', 'SAT', 'AM', 'J', 16, 'Grace, 謝飯', 1, 'Y', '', 'Y'],
    // Saturday PM
    ['SAT_PM_PRAYER', 'Prayer', '禱告', 'SAT', 'PM', 'X', 19, 'Prayer, 禱告', 1, 'Y', '', 'Y'],
    ['SAT_PM_PRAISE', 'Praise', '詩頌', 'SAT', 'PM', 'D', 20, 'Hymn Leader, 詩頌', 1, 'Y', '', 'Y'],
    ['SAT_PM_PIANIST', 'Accompanist', '司琴', 'SAT', 'PM', 'E', 21, 'Pianist, 司琴', 1, 'Y', 'PIANO', 'Y'],
    ['SAT_PM_SERMON', 'Sermon', '講道', 'SAT', 'PM', 'AY', 22, 'Speaker', 1, 'Y', 'SPEAK', 'Y'],
    ['SAT_PM_INTERPRETER', 'Interpreter', '翻譯', 'SAT', 'PM', 'C', 23, 'Interpreter (E-to-C), Interpreter (C-to-E), Interpreter, 翻譯', 1, 'Y', '', 'Y'],
    ['SAT_PM_AV1', 'A/V 1', '視聽 1', 'SAT', 'PM', 'H', 24, 'AV, 視聽', 1, 'Y', '', 'Y'],
    ['SAT_PM_AV2', 'A/V 2', '視聽 2', 'SAT', 'PM', 'I', 24, 'AV, 視聽', 1, 'Y', '', 'Y'],
    ['SAT_PM_USHER1', 'Usher 1', '接待 1', 'SAT', 'PM', 'F', 25, 'Usher, 接待', 1, 'Y', '', 'Y'],
    ['SAT_PM_USHER2', 'Usher 2', '接待 2', 'SAT', 'PM', 'G', 25, 'Usher, 接待', 1, 'Y', '', 'Y'],
    ['SAT_PM_EBF', 'English Bible', '英文查經', 'SAT', 'PM', 'AA', 26, 'English Bible Fellowship', 1, 'Y', '', 'Y'],
    ['SAT_PM_SISTERS', 'Sisters', '姐妹團契', 'SAT', 'PM', 'AB', 27, "Sisters' Fellowship", 1, 'Y', '', 'Y'],
    ['SAT_PM_VESSEL', 'Vessel', '器皿團契', 'SAT', 'PM', 'AC', 28, 'Vessel Fellowship', 1, 'Y', '', 'Y'],
    // Monday
    ['MON_PRAISE', 'Praise', '詩頌', 'MON', 'AM', 'D', 31, 'Hymn Leader, 詩頌', 1, 'Y', '', 'Y'],
    ['MON_PIANIST', 'Accompanist', '司琴', 'MON', 'AM', 'E', 32, 'Pianist, 司琴', 1, 'Y', 'PIANO', 'Y'],
    ['MON_SERMON', 'Sermon', '講道', 'MON', 'AM', 'AY', 33, 'Speaker', 1, 'Y', 'SPEAK', 'Y'],
    ['MON_INTERPRETER', 'Interpreter', '翻譯', 'MON', 'AM', 'C', 34, 'Interpreter (E-to-C), Interpreter (C-to-E), Interpreter, 翻譯', 1, 'Y', '', 'Y'],
    ['MON_AV1', 'A/V 1', '視聽 1', 'MON', 'AM', 'H', 35, 'AV, 視聽', 1, 'Y', '', 'Y'],
    ['MON_AV2', 'A/V 2', '視聽 2', 'MON', 'AM', 'I', 35, 'AV, 視聽', 1, 'Y', '', 'Y'],
    ['MON_USHER1', 'Usher 1', '接待 1', 'MON', 'AM', 'F', 36, 'Usher, 接待', 1, 'Y', '', 'Y'],
    ['MON_USHER2', 'Usher 2', '接待 2', 'MON', 'AM', 'G', 36, 'Usher, 接待', 1, 'Y', '', 'Y'],
    // Tuesday
    ['TUE_JOY', 'Joy Fellowship', '喜樂團契', 'TUE', 'AM', 'AS', 39, 'Joy Fellowship', 1, 'Y', '', 'Y'],
    // Wednesday
    ['WED_AGAPE', 'Agape Venue', '愛加倍', 'WED', 'AM', 'AX', 42, 'Wednesday Location', 1, 'Y', '', 'Y'],
    ['WED_STUDY', 'Wednesday Study', '週三查經', 'WED', 'AM', 'AT', '', 'Wednesday Study', 1, 'Y', '', 'Y'],
    ['WED_SOUTH', 'Wednesday South', '週三南區', 'WED', 'AM', 'AU', '', 'Wednesday Location', 1, 'Y', '', 'Y'],
    ['WED_WEST', 'Wednesday West', '週三西區', 'WED', 'AM', 'AV', '', 'Wednesday Location', 1, 'Y', '', 'Y'],
    ['WED_NORTH', 'Wednesday North', '週三北區', 'WED', 'AM', 'AW', '', 'Wednesday Location', 1, 'Y', '', 'Y'],
    // Friday
    ['FRI_PRAISE', 'Praise', '詩頌', 'FRI', 'AM', 'D', 49, 'Hymn Leader, 詩頌', 1, 'Y', '', 'Y'],
    ['FRI_PIANIST', 'Accompanist', '司琴', 'FRI', 'AM', 'E', 50, 'Pianist, 司琴', 1, 'Y', 'PIANO', 'Y'],
    ['FRI_SERMON', 'Sermon', '講道', 'FRI', 'AM', 'AY', 51, 'Speaker', 1, 'Y', 'SPEAK', 'Y'],
    ['FRI_INTERPRETER', 'Interpreter', '翻譯', 'FRI', 'AM', 'C', 52, 'Interpreter (E-to-C), Interpreter (C-to-E), Interpreter, 翻譯', 1, 'Y', '', 'Y'],
    ['FRI_AV1', 'A/V 1', '視聽 1', 'FRI', 'AM', 'H', 53, 'AV, 視聽', 1, 'Y', '', 'Y'],
    ['FRI_AV2', 'A/V 2', '視聽 2', 'FRI', 'AM', 'I', 53, 'AV, 視聽', 1, 'Y', '', 'Y'],
    ['FRI_USHER1', 'Usher 1', '接待 1', 'FRI', 'AM', 'F', 54, 'Usher, 接待', 1, 'Y', '', 'Y'],
    ['FRI_USHER2', 'Usher 2', '接待 2', 'FRI', 'AM', 'G', 54, 'Usher, 接待', 1, 'Y', '', 'Y'],
    // Friday Cleaning
    ['FRI_CLEAN_CHAPEL1', 'Friday Chapel Cleaning 1', '星期五會堂打掃 1', 'FRI', 'CLEANING', 'AD', 60, 'Friday Chapel Cleaning', 1, 'Y', '', 'Y'],
    ['FRI_CLEAN_CHAPEL2', 'Friday Chapel Cleaning 2', '星期五會堂打掃 2', 'FRI', 'CLEANING', 'AE', 61, 'Friday Chapel Cleaning', 1, 'Y', '', 'Y'],
    ['FRI_CLEAN_CHAPEL3', 'Friday Chapel Cleaning 3', '星期五會堂打掃 3', 'FRI', 'CLEANING', 'AF', 62, 'Friday Chapel Cleaning', 1, 'Y', '', 'Y'],
    // Saturday Cleaning
    ['SAT_CLEAN_CHAPEL1', 'Saturday Chapel 1', '星期六會堂打掃 1', 'SAT', 'CLEANING', 'AI', 60, 'Saturday Cleaning', 1, 'Y', '', 'Y'],
    ['SAT_CLEAN_CHAPEL2', 'Saturday Chapel 2', '星期六會堂打掃 2', 'SAT', 'CLEANING', 'AJ', 61, 'Saturday Cleaning', 1, 'Y', '', 'Y'],
    ['SAT_CLEAN_CHAPEL3', 'Saturday Chapel 3', '星期六會堂打掃 3', 'SAT', 'CLEANING', 'AK', 62, 'Saturday Cleaning', 1, 'Y', '', 'Y'],
    ['SAT_CLEAN_TRASH', 'Saturday Trash', '星期六倒垃圾', 'SAT', 'CLEANING', 'AL', 63, 'Saturday Cleaning', 1, 'Y', '', 'Y'],
    ['SAT_CLEAN_BROS_DN', 'Saturday Bros Down', '星期六弟兄樓下', 'SAT', 'CLEANING', 'AM', 64, 'Saturday Cleaning', 1, 'Y', '', 'Y'],
    ['SAT_CLEAN_BROS_DN2', 'Saturday Bros Down 2', '星期六弟兄樓下 2', 'SAT', 'CLEANING', 'AN', 64, 'Saturday Cleaning', 1, 'Y', '', 'Y'],
    ['SAT_CLEAN_SIS_DN', 'Saturday Sis Down', '星期六姐妹樓下', 'SAT', 'CLEANING', 'AO', 68, 'Saturday Cleaning', 1, 'Y', '', 'Y'],
    ['SAT_CLEAN_SIS_DN2', 'Saturday Sis Down 2', '星期六姐妹樓下 2', 'SAT', 'CLEANING', 'AP', 68, 'Saturday Cleaning', 1, 'Y', '', 'Y'],
    ['SAT_CLEAN_BROS_UP', 'Saturday Bros Up', '星期六弟兄樓上', 'SAT', 'CLEANING', 'AQ', 66, 'Saturday Cleaning', 1, 'Y', '', 'Y'],
    ['SAT_CLEAN_SIS_UP', 'Saturday Sis Up', '星期六姐妹樓上', 'SAT', 'CLEANING', 'AR', 70, 'Saturday Cleaning', 1, 'Y', '', 'Y']
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#d9ead3');
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  sheet.autoResizeColumns(1, headers.length);
  sheet.hideSheet();
  try { CacheService.getScriptCache().remove('church_config'); } catch(e) {}
  return sheet;
}

function promptSeedNewChurch() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var existing = ss.getSheetByName('ChurchProfile') || ss.getSheetByName('RoleConfig');
  if (existing) {
     var res = ui.alert('Confirmation', 'Configuration sheets already exist. Proceeding may overwrite settings or add duplicates. Are you sure you want to run Seed New Church?', ui.ButtonSet.YES_NO);
     if (res !== ui.Button.YES) return;
  }
  
  try {
     ensureWorkersSchema(ss);
     seedChurchProfileSheet(ss);
     seedServiceLayoutConfigSheet(ss);
     seedRoleConfigSheet(ss);
     var compileRes = compileRoleConfig();
     try { CacheService.getScriptCache().remove('church_config'); } catch(e) {}
     
     if (compileRes.success) {
       ui.alert('Success', 'Seeded initial configurations (ChurchProfile, ServiceLayoutConfig, RoleConfig) and compiled ValidationConfig & MappingConfig successfully.', ui.ButtonSet.OK);
     } else {
       ui.alert('Seeded with Warning', 'Configuration tabs created, but role compilation had an error: ' + compileRes.error, ui.ButtonSet.OK);
     }
  } catch(e) {
     ui.alert('Error', e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * One-click Tab Hygiene Tool:
 * 1. Hides all Tier 2 (Config) and Tier 3 (Log/Cache) tabs, plus completed past-month tabs.
 * 2. Re-orders visible tabs: Q1–Q4 -> Active Monthly Schedules -> Workers -> Teams -> Announcements -> Annual Summary -> Unavailability (and others).
 * 3. Color-codes bottom tab borders:
 *    - Blue: Q1–Q4
 *    - Gold: Active Monthly Schedules
 *    - Green: Workers & Teams
 *    - Gray: Announcements, Annual Summary, Unavailability, etc.
 */
function organizeAndCleanUpTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var blueColor  = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.THEME ? APP_CONFIG.THEME.PRIMARY_COLOR : "#1F4E79");
  var goldColor  = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.THEME ? APP_CONFIG.THEME.ACCENT_WARNING : "#F4B400");
  var greenColor = (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.THEME ? APP_CONFIG.THEME.ACCENT_SUCCESS : "#0F9D58");
  var grayColor  = "#757575";

  var configTabs = [
    'ChurchProfile', 'ServiceLayoutConfig', 'RoleConfig', 'ValidationConfig', 'MappingConfig',
    'MonthlyTemplate', 'EmailSettings', 'EmailTemplate', 'EmailLinks', 'Coordinators', 'PublishedLinks',
    'MasterConfig', 'ESSC_MappingConfig', 'Rosters'
  ];
  var logTabs = [
    'SWR Sync Log', 'RE Sync Log', 'MonthlyGeneratorLog', 'MonthlyOverridesLog', 'Email Log',
    'WorkerImport_Staging', 'WorkerImportLog'
  ];

  // 1. Ensure Q1 (or first available quarter/sheet) is visible and active so we never attempt to hide the active sheet
  var anchorSheet = ss.getSheetByName('Q1') || (CONFIG.QUARTERS ? ss.getSheetByName(CONFIG.QUARTERS[0]) : null) || ss.getSheets()[0];
  if (anchorSheet) {
    if (anchorSheet.isSheetHidden()) {
      anchorSheet.showSheet();
    }
    ss.setActiveSheet(anchorSheet);
  }

  // 2. Hide Config and Log tabs
  var hiddenConfigCount = 0;
  for (var i = 0; i < configTabs.length; i++) {
    var cSheet = ss.getSheetByName(configTabs[i]);
    if (cSheet && !cSheet.isSheetHidden()) {
      cSheet.hideSheet();
      hiddenConfigCount++;
    }
  }
  for (var j = 0; j < logTabs.length; j++) {
    var lSheet = ss.getSheetByName(logTabs[j]);
    if (lSheet && !lSheet.isSheetHidden()) {
      lSheet.hideSheet();
      hiddenConfigCount++;
    }
  }

  // Hide backup tabs (_bak_...)
  var allSheets = ss.getSheets();
  for (var b = 0; b < allSheets.length; b++) {
    var bName = allSheets[b].getName();
    if ((bName.indexOf('_bak_') === 0 || bName.indexOf('_backup_') === 0) && !allSheets[b].isSheetHidden()) {
      allSheets[b].hideSheet();
      hiddenConfigCount++;
    }
  }

  // 3. Classify and hide completed past-month tabs
  var monthPattern = /^(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})$/i;
  var monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var now = new Date();
  var startOfCurrentMonth = new Date(now.getFullYear(), now.getMonth(), 1);

  var activeMonthlySheets = [];
  var hiddenPastMonthsCount = 0;

  for (var s = 0; s < allSheets.length; s++) {
    var sh = allSheets[s];
    var match = sh.getName().match(monthPattern);
    if (match) {
      var mIndex = -1;
      for (var mi = 0; mi < monthNames.length; mi++) {
        if (monthNames[mi].toLowerCase() === match[1].toLowerCase()) {
          mIndex = mi;
          break;
        }
      }
      var y = parseInt(match[2], 10);
      var endOfMonth = new Date(y, mIndex + 1, 0, 23, 59, 59);

      if (endOfMonth < startOfCurrentMonth) {
        if (!sh.isSheetHidden()) {
          sh.hideSheet();
          hiddenPastMonthsCount++;
        }
      } else {
        if (sh.isSheetHidden()) {
          sh.showSheet();
        }
        activeMonthlySheets.push({ sheet: sh, year: y, monthIndex: mIndex, name: sh.getName() });
      }
    }
  }

  // Sort active monthly schedules chronologically
  activeMonthlySheets.sort(function(a, b) {
    if (a.year !== b.year) return a.year - b.year;
    return a.monthIndex - b.monthIndex;
  });

  // 4. Construct desired visible tabs order:
  // Q1–Q4 -> Active Monthly Schedules -> Workers -> Teams -> Announcements -> Annual Summary -> Unavailability
  var orderedVisibleSheets = [];
  var processedNames = {};

  // Q1-Q4
  var quarters = CONFIG.QUARTERS || ['Q1', 'Q2', 'Q3', 'Q4'];
  for (var q = 0; q < quarters.length; q++) {
    var qs = ss.getSheetByName(quarters[q]);
    if (qs) {
      if (qs.isSheetHidden()) qs.showSheet();
      orderedVisibleSheets.push({ sheet: qs, color: blueColor });
      processedNames[qs.getName()] = true;
    }
  }

  // Active Monthly Schedules
  for (var m = 0; m < activeMonthlySheets.length; m++) {
    var ms = activeMonthlySheets[m].sheet;
    orderedVisibleSheets.push({ sheet: ms, color: goldColor });
    processedNames[ms.getName()] = true;
  }

  // Workers & Teams
  var workersSheet = ss.getSheetByName(CONFIG.WORKERS_SHEET || 'Workers');
  if (workersSheet && !workersSheet.isSheetHidden()) {
    orderedVisibleSheets.push({ sheet: workersSheet, color: greenColor });
    processedNames[workersSheet.getName()] = true;
  }
  var teamsSheet = ss.getSheetByName(CONFIG.TEAMS_SHEET || 'Teams');
  if (teamsSheet && !teamsSheet.isSheetHidden()) {
    orderedVisibleSheets.push({ sheet: teamsSheet, color: greenColor });
    processedNames[teamsSheet.getName()] = true;
  }

  // Announcements
  var annSheet = ss.getSheetByName('Announcements');
  if (annSheet && !annSheet.isSheetHidden()) {
    orderedVisibleSheets.push({ sheet: annSheet, color: grayColor });
    processedNames[annSheet.getName()] = true;
  }

  // Annual Summary
  var sumSheet = ss.getSheetByName('Annual Summary');
  if (sumSheet && !sumSheet.isSheetHidden()) {
    orderedVisibleSheets.push({ sheet: sumSheet, color: grayColor });
    processedNames[sumSheet.getName()] = true;
  }

  // Unavailability
  var unavailSheet = ss.getSheetByName(CONFIG.UNAVAILABILITY_SHEET || 'Unavailability');
  if (unavailSheet && !unavailSheet.isSheetHidden()) {
    orderedVisibleSheets.push({ sheet: unavailSheet, color: grayColor });
    processedNames[unavailSheet.getName()] = true;
  }

  // Any remaining visible sheets
  for (var rem = 0; rem < allSheets.length; rem++) {
    var rSheet = allSheets[rem];
    var rName = rSheet.getName();
    if (!rSheet.isSheetHidden() && !processedNames[rName]) {
      orderedVisibleSheets.push({ sheet: rSheet, color: grayColor });
      processedNames[rName] = true;
    }
  }

  // 5. Execute positioning and color application
  var pos = 1;
  for (var d = 0; d < orderedVisibleSheets.length; d++) {
    var targetSheet = orderedVisibleSheets[d].sheet;
    try {
      ss.setActiveSheet(targetSheet);
      ss.moveActiveSheet(pos);
      targetSheet.setTabColor(orderedVisibleSheets[d].color);
      pos++;
    } catch(moveErr) {
      console.warn('Could not move sheet ' + targetSheet.getName() + ': ' + moveErr);
    }
  }

  // Return focus to Q1
  if (anchorSheet) {
    try { ss.setActiveSheet(anchorSheet); } catch(e) {}
  }

  ui.alert(
    'Tab Hygiene Complete! 🗂️',
    'Sheet tabs have been organized and color-coded:\n\n' +
    '• Q1–Q4: Blue\n' +
    '• Active Monthly Schedules: Gold (' + activeMonthlySheets.length + ' active)\n' +
    '• Workers & Teams: Green\n' +
    '• Admin & Announcements: Gray\n\n' +
    'Hidden Tabs:\n' +
    '• ' + hiddenConfigCount + ' Configuration & Log tab(s) hidden.\n' +
    '• ' + hiddenPastMonthsCount + ' Completed past-month schedule(s) hidden.\n\n' +
    'Tip: Click "👁️ Toggle Admin Config & Log Tabs" anytime you need to view or edit configuration sheets.',
    ui.ButtonSet.OK
  );
}

/**
 * Toggles visibility of all admin configuration, log, and backup tabs in one click.
 */
function toggleAdminConfigTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var configTabs = [
    'ChurchProfile', 'ServiceLayoutConfig', 'RoleConfig', 'ValidationConfig', 'MappingConfig',
    'MonthlyTemplate', 'EmailSettings', 'EmailTemplate', 'EmailLinks', 'Coordinators', 'PublishedLinks',
    'MasterConfig', 'ESSC_MappingConfig', 'Rosters'
  ];
  var logTabs = [
    'SWR Sync Log', 'RE Sync Log', 'MonthlyGeneratorLog', 'MonthlyOverridesLog', 'Email Log',
    'WorkerImport_Staging', 'WorkerImportLog'
  ];
  var allAdminTabs = configTabs.concat(logTabs);

  // Determine current visibility state: if ANY admin tab is hidden, our toggle will REVEAL ALL
  var anyHidden = false;
  for (var i = 0; i < allAdminTabs.length; i++) {
    var sh = ss.getSheetByName(allAdminTabs[i]);
    if (sh && sh.isSheetHidden()) {
      anyHidden = true;
      break;
    }
  }
  var allSheets = ss.getSheets();
  if (!anyHidden) {
    for (var b = 0; b < allSheets.length; b++) {
      if ((allSheets[b].getName().indexOf('_bak_') === 0 || allSheets[b].getName().indexOf('_backup_') === 0) && allSheets[b].isSheetHidden()) {
        anyHidden = true;
        break;
      }
    }
  }

  // Make sure an anchor sheet is active so we never hide an active sheet
  var anchorSheet = ss.getSheetByName('Q1') || (CONFIG.QUARTERS ? ss.getSheetByName(CONFIG.QUARTERS[0]) : null) || ss.getSheets()[0];
  if (anchorSheet && anchorSheet.isSheetHidden()) anchorSheet.showSheet();
  if (anchorSheet) ss.setActiveSheet(anchorSheet);

  var count = 0;
  if (anyHidden) {
    // Reveal all admin tabs
    for (var j = 0; j < allAdminTabs.length; j++) {
      var s = ss.getSheetByName(allAdminTabs[j]);
      if (s) {
        if (s.isSheetHidden()) s.showSheet();
        s.setTabColor('#757575');
        count++;
      }
    }
    for (var k = 0; k < allSheets.length; k++) {
      var bkName = allSheets[k].getName();
      if ((bkName.indexOf('_bak_') === 0 || bkName.indexOf('_backup_') === 0) && allSheets[k].isSheetHidden()) {
        allSheets[k].showSheet();
        allSheets[k].setTabColor('#757575');
        count++;
      }
    }
    ui.alert(
      'Admin Tabs Visible 👁️',
      'Unhid ' + count + ' configuration and log tab(s).\n\n' +
      'You can now view and edit configuration tables directly.\n\n' +
      'Tip: Click "👁️ Toggle Admin Config & Log Tabs" or "🗂️ Organize & Clean Up Tabs" when finished to hide them again.',
      ui.ButtonSet.OK
    );
  } else {
    // Hide all admin tabs
    for (var m = 0; m < allAdminTabs.length; m++) {
      var sHide = ss.getSheetByName(allAdminTabs[m]);
      if (sHide && !sHide.isSheetHidden()) {
        sHide.hideSheet();
        count++;
      }
    }
    for (var n = 0; n < allSheets.length; n++) {
      var bNameHide = allSheets[n].getName();
      if ((bNameHide.indexOf('_bak_') === 0 || bNameHide.indexOf('_backup_') === 0) && !allSheets[n].isSheetHidden()) {
        allSheets[n].hideSheet();
        count++;
      }
    }
    ui.alert(
      'Admin Tabs Hidden 🙈',
      'Successfully hid ' + count + ' configuration and log tab(s).\n\n' +
      'Your sheet tabs are now clean and focused on active quarterly and monthly schedules.',
      ui.ButtonSet.OK
    );
  }
}

/**
 * 1-Click Unified Worker Roster Sync:
 * 1. Calls syncRosters() (safe formula-preserving sync + cache invalidation)
 * 2. Calls applyAllValidation() (rebuilds Q1-Q4 dropdowns)
 * 3. Syncs active workers to external Unavailability Matrix (if linked)
 * 4. Calls bumpAllQuarterVersions()
 */
function applyWorkerRosterChanges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stepsCompleted = [];
  var errors = [];
  
  // Step 1: Sync Rosters & clear worker caches
  try {
    var rosterRes = syncRosters();
    if (rosterRes && rosterRes.success) {
      stepsCompleted.push("Rosters & Caches Updated (" + (rosterRes.mode || "synced") + ")");
    } else {
      errors.push("Roster Sync: " + (rosterRes ? rosterRes.error : "Unknown error"));
    }
  } catch(e) {
    errors.push("Roster Sync: " + e.message);
  }
  
  // Step 2: Apply Validation Dropdowns to Quarters
  try {
    if (typeof applyAllValidation === 'function') {
      applyAllValidation();
      stepsCompleted.push("Quarter Validation Dropdowns Refreshed");
    }
  } catch(e) {
    errors.push("Validation: " + e.message);
  }
  
  // Step 3: Sync to external Unavailability Matrix if linked
  try {
    var unavailId = (typeof getLinkedUnavailabilityId === 'function') ? getLinkedUnavailabilityId() : null;
    if (unavailId) {
      var unavailSs = SpreadsheetApp.openById(unavailId);
      if (unavailSs && typeof restructureAndSyncUnavailabilityMatrix === 'function') {
        restructureAndSyncUnavailabilityMatrix(unavailSs, ss);
        stepsCompleted.push("External Unavailability Matrix Rosters Synced");
      }
    }
  } catch(e) {
    console.warn("Unavailability matrix sync skipped or failed: " + e.message);
  }
  
  // Step 4: Bump Quarter Versions
  try {
    if (typeof bumpAllQuarterVersions === 'function') {
      bumpAllQuarterVersions();
      stepsCompleted.push("All Quarter Cache Versions Bumped");
    }
  } catch(e) {}
  
  var success = errors.length === 0;
  var msg = "1-Click Worker Roster Apply Complete!\n\n" +
            "Steps Completed:\n• " + stepsCompleted.join("\n• ") +
            (errors.length > 0 ? "\n\n⚠️ Notices / Warnings:\n• " + errors.join("\n• ") : "");
  
  try {
    var ui = SpreadsheetApp.getUi();
    if (ui) {
      ui.alert(success ? "Rosters Applied 🎉" : "Completed with Notices ⚠️", msg, ui.ButtonSet.OK);
    }
  } catch(uiErr) {}
  
  return {
    success: success,
    steps: stepsCompleted,
    errors: errors
  };
}

/**
 * Scans next month's schedule dates across Q1-Q4 for blank/unassigned slots in active role columns,
 * groups them by coordinator from the Coordinators tab, and presents a modal audit summary dialog.
 */
function auditUpcomingMonthMissingSlots() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  
  // 1. Determine next month target
  var now = new Date();
  var nextDate = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  var targetYear = nextDate.getFullYear();
  var targetMonth = nextDate.getMonth() + 1; // 1-12
  var monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var targetMonthName = monthNames[targetMonth - 1] + " " + targetYear;
  
  var qNum = Math.ceil(targetMonth / 3);
  var qSheetName = "Q" + qNum;
  var qSheet = ss.getSheetByName(qSheetName);
  if (!qSheet) {
    ui.alert("Target Sheet Missing", "Could not find quarterly sheet '" + qSheetName + "' for " + targetMonthName + ".", ui.ButtonSet.OK);
    return;
  }
  
  // 2. Load Coordinators mapping
  var colToCoord = {}; // colIdx (1-based) -> { name, email, role, qCols }
  var coordinatorsList = [];
  var coordSheet = ss.getSheetByName('Coordinators');
  if (coordSheet && coordSheet.getLastRow() >= 2) {
    var cData = coordSheet.getDataRange().getValues();
    var cHeaders = cData[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
    var nameIdx = cHeaders.indexOf('name');
    var emailIdx = cHeaders.indexOf('email');
    var roleIdx = cHeaders.indexOf('role_en') !== -1 ? cHeaders.indexOf('role_en') : cHeaders.indexOf('role');
    var qColsIdx = cHeaders.indexOf('q_columns');
    var enabledIdx = cHeaders.indexOf('include_in_email') !== -1 ? cHeaders.indexOf('include_in_email') : cHeaders.indexOf('enabled');
    
    for (var cr = 1; cr < cData.length; cr++) {
      var cRow = cData[cr];
      if (enabledIdx !== -1) {
        var incFlag = String(cRow[enabledIdx] || '').trim().toUpperCase();
        if (incFlag === 'N' || incFlag === 'NO' || incFlag === 'FALSE') continue;
      }
      var cName = (nameIdx !== -1) ? String(cRow[nameIdx] || '').trim() : '';
      var cEmail = (emailIdx !== -1) ? String(cRow[emailIdx] || '').trim() : '';
      var cRole = (roleIdx !== -1) ? String(cRow[roleIdx] || '').trim() : '';
      var qColsStr = (qColsIdx !== -1) ? String(cRow[qColsIdx] || '').trim() : '';
      if (!cName && !cEmail) continue;
      
      var coordObj = { name: cName, email: cEmail, role: cRole, qColsStr: qColsStr, missingSlots: [] };
      coordinatorsList.push(coordObj);
      
      if (qColsStr) {
        var tokens = qColsStr.split(/[\s,;]+/);
        tokens.forEach(function(tok) {
          var cleanTok = tok.trim().toUpperCase();
          if (cleanTok.match(/^[A-Z]+$/)) {
            var colIdx = colLetterToIdx_(cleanTok); // 1-indexed
            if (!colToCoord[colIdx]) colToCoord[colIdx] = coordObj;
          }
        });
      }
    }
  }
  
  var unassignedCoordObj = { name: "General / Unassigned Coordinator", email: "", role: "Various", qColsStr: "", missingSlots: [] };
  
  // 3. Scan Q sheet
  var numCols = Math.min(qSheet.getLastColumn(), 51);
  var numRows = Math.min(qSheet.getLastRow(), 95);
  var values = qSheet.getRange(1, 1, numRows, numCols).getValues();
  var backgrounds = qSheet.getRange(1, 1, numRows, numCols).getBackgrounds();
  
  var row4 = values[3] || [];
  var row5 = values[4] || [];
  
  var totalMissing = 0;
  var daysWithMissing = {};
  
  for (var r = 5; r < values.length; r++) { // rows 6 to 95
    var dateVal = values[r][0];
    if (!dateVal) continue;
    var dObj = dateVal instanceof Date ? dateVal : new Date(dateVal);
    if (isNaN(dObj.getTime())) continue;
    
    // Check if this row is in next month
    if (dObj.getFullYear() !== targetYear || (dObj.getMonth() + 1) !== targetMonth) continue;
    
    var dateFormatted = Utilities.formatDate(dObj, tz, "EEE, MMM d");
    
    for (var c = 1; c < numCols; c++) { // Col B to AY (1-indexed 2 to 51)
      var colNum = c + 1;
      var cellVal = String(values[r][c] || '').trim();
      var bg = String(backgrounds[r][c] || '').toLowerCase();
      
      // Skip inactive/grayed cells (#d9d9d9, #b7b7b7, #cccccc, #efefef)
      var isGray = (bg === '#d9d9d9' || bg === '#b7b7b7' || bg === '#cccccc' || bg === '#efefef');
      if (isGray) continue;
      
      // Skip cells marked with 'X'
      if (cellVal === 'X' || cellVal === 'x') continue;
      
      // Check if slot is empty
      if (!cellVal) {
        var roleName = String(row4[c] || '').trim() || String(row5[c] || '').trim() || ('Col ' + colIdxToLetter_(colNum));
        var colLetter = colIdxToLetter_(colNum);
        var slotInfo = {
          date: dateFormatted,
          role: roleName,
          col: colLetter,
          row: r + 1
        };
        
        var assignedCoordinator = colToCoord[colNum];
        if (assignedCoordinator) {
          assignedCoordinator.missingSlots.push(slotInfo);
        } else {
          unassignedCoordObj.missingSlots.push(slotInfo);
        }
        
        totalMissing++;
        daysWithMissing[dateFormatted] = (daysWithMissing[dateFormatted] || 0) + 1;
      }
    }
  }
  
  if (unassignedCoordObj.missingSlots.length > 0) {
    coordinatorsList.push(unassignedCoordObj);
  }
  
  // 4. Build report
  if (totalMissing === 0) {
    ui.alert(
      "All Slots Filled! 🎉",
      "Great news! There are no missing or unassigned slots found for " + targetMonthName + " in " + qSheetName + ".\n\nEverything is fully scheduled!",
      ui.ButtonSet.OK
    );
    return;
  }
  
  var coordsWithMissing = coordinatorsList.filter(function(coord) {
    return coord.missingSlots.length > 0;
  });
  
  var summaryText = "Schedule Audit for " + targetMonthName + " (" + qSheetName + ")\n";
  summaryText += "Total Unassigned Slots: " + totalMissing + " across " + Object.keys(daysWithMissing).length + " service dates.\n\n";
  summaryText += "Open Slots by Coordinator:\n";
  
  coordsWithMissing.forEach(function(coord) {
    summaryText += "• " + coord.name + (coord.role ? " (" + coord.role + ")" : "") + ": " + coord.missingSlots.length + " slot(s)\n";
    // Show sample of first 3 slots
    var sample = coord.missingSlots.slice(0, 3).map(function(s) {
      return "    - " + s.date + ": " + s.role + " (" + s.col + ")";
    }).join("\n");
    summaryText += sample;
    if (coord.missingSlots.length > 3) {
      summaryText += "\n    ... and " + (coord.missingSlots.length - 3) + " more";
    }
    summaryText += "\n\n";
  });
  
  summaryText += "Would you like to email a friendly reminder notification to each coordinator with their open slots?";
  
  var choice = ui.alert("Missing Slot Audit (" + totalMissing + " Open Slots)", summaryText, ui.ButtonSet.YES_NO);
  if (choice === ui.Button.YES) {
    var sentCount = 0;
    var failCount = 0;
    coordsWithMissing.forEach(function(coord) {
      if (!coord.email) return;
      try {
        var emailSubject = "True Jesus Church — Open Schedule Slots for " + targetMonthName;
        var emailBody = "Shalom " + coord.name + ",\n\n" +
          "Here is a summary of the remaining open schedule slot(s) for your ministry for " + targetMonthName + " in " + qSheetName + ":\n\n";
        
        coord.missingSlots.forEach(function(s) {
          emailBody += "• " + s.date + " — " + s.role + " (Column " + s.col + ")\n";
        });
        
        emailBody += "\nPlease log into the Holy Work Schedule at your earliest convenience to assign workers to these slots:\n" +
          ss.getUrl() + "\n\n" +
          "Thank you for your holy work and dedication!\n" +
          "TJC Info Department";
        
        MailApp.sendEmail(coord.email, emailSubject, emailBody);
        sentCount++;
      } catch (err) {
        console.error("Failed to email coordinator " + coord.name + " (" + coord.email + "): " + err.message);
        failCount++;
      }
    });
    
    ui.alert(
      "Reminders Sent 📧",
      "Successfully sent reminders to " + sentCount + " coordinator(s)." +
      (failCount > 0 ? "\n(" + failCount + " failed due to email limits or invalid addresses)." : ""),
      ui.ButtonSet.OK
    );
  }
}

