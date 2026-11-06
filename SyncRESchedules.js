/**
 * Script to Synchronize Religious Education Schedules to the Master Holy Work Schedule
 */

const RE_SYNC_LOG_SHEET = 'RE Sync Log';

/**
 * Configuration for the RE Classes
 * Note: The script now features dynamic column detection! 
 * It will use these values as fallbacks, but will automatically override 
 * them if it finds the "Morning", "Afternoon", and "Teacher" headers.
 */
const RE_CLASSES_CONFIG = {
  'K': {
    spreadsheetId: '1bem2hLeeCfFRkOrY1T93VvTNkKZJz-_4qOYuC0usi6c',
    dateCol: 2, amTeacherCol: 4, amAssistantCol: 5, pmTeacherCol: 6, pmAssistantCol: 7, startRow: 9, hwsTeacherCol: 12, hwsAssistantCol: 13
  },
  'E1': {
    spreadsheetId: '13Bb3yObUEhjUp4kfL9eIkIv6esRFPMI8vw5B1SSlvWQ',
    dateCol: 2, amTeacherCol: 4, amAssistantCol: null, pmTeacherCol: 6, pmAssistantCol: null, startRow: 8, hwsTeacherCol: 14, hwsAssistantCol: 15
  },
  'E2': {
    spreadsheetId: '1tfeeGfM8FlKvMl_-gr5AVkh2GOdS5_2TBXgYXq4rsig',
    dateCol: 1, amTeacherCol: 3, amAssistantCol: null, pmTeacherCol: 6, pmAssistantCol: null, startRow: 9, hwsTeacherCol: 16, hwsAssistantCol: 17
  },
  'J1': {
    spreadsheetId: '1QsSsdcY4av6Jg5oCoZOTk01pizsrfDSPVQIbpj4rHXI', 
    dateCol: 1, amTeacherCol: null, amAssistantCol: null, pmTeacherCol: 4, pmAssistantCol: 5, startRow: 9, hwsTeacherCol: 18, hwsAssistantCol: 19
  },
  'J2': {
    spreadsheetId: '1b6uKhjO8i60-9DyCw-aOqh7zAxpFMtxSeCbWFX7VDn8', 
    dateCol: 2, amTeacherCol: null, amAssistantCol: null, pmTeacherCol: 4, pmAssistantCol: null, startRow: 11, hwsTeacherCol: 20, hwsAssistantCol: 21
  },
  'Sr': {
    spreadsheetId: '1wBRnKZ5undUFsDhsdTi3BPXVDfPXJoIY7komJXiqRkA', 
    dateCol: 1, amTeacherCol: null, amAssistantCol: null, pmTeacherCol: 5, pmAssistantCol: null, startRow: 8, hwsTeacherCol: 22, hwsAssistantCol: null 
  }
};

function getREConfig() {
  let configOverrides = {};
  try {
    const props = PropertiesService.getScriptProperties().getProperty('RE_CLASSES_CONFIG');
    if (props) {
      configOverrides = JSON.parse(props);
    } else {
      console.warn("RE_CLASSES_CONFIG property not found. Falling back to hardcoded IDs.");
    }
  } catch (e) {
    console.warn("Failed to parse RE_CLASSES_CONFIG property. Falling back to hardcoded IDs.", e);
  }
  
  const mergedConfig = JSON.parse(JSON.stringify(RE_CLASSES_CONFIG));
  for (const [key, val] of Object.entries(configOverrides)) {
    if (mergedConfig[key]) {
      Object.assign(mergedConfig[key], val);
    } else {
      // Allow multi-church scaling: if it's a new class key, add it
      mergedConfig[key] = val;
    }
  }
  return mergedConfig;
}

/**
 * ---------------------------------------------------------
 * 1. MANUAL RUN FUNCTION (For buttons or custom menus)
 * ---------------------------------------------------------
 */
function syncRESchedules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    console.error("This function requires an active UI. If you are running this from a time-based trigger, please change your trigger to run autoSyncRESchedules() instead.");
    return;
  }
  
  // Basic validation that we are on a Quarterly sheet (e.g., "Q1", "Q2", "Q3", "Q4")
  if (!/^Q[1-4]$/.test(sheet.getName())) {
    ui.alert('Please run this script from one of the Quarter sheets (Q1, Q2, Q3, Q4).');
    return;
  }
  
  const response = ui.alert(
    'Sync RE Schedules', 
    `Are you sure you want to sync Religious Education schedules into the ${sheet.getName()} sheet? This will search all sheets in the RE schedule files and update matching dates for AM and PM classes.`, 
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }

  // Execute the core logic
  const results = performSync(ss, sheet);
  
  // Report Results
  if (results.failed.length > 0) {
    ui.alert('Sync Completed with Errors', `Successfully synced ${results.successful} out of ${results.total} classes.\n\nChanges Made: ${results.logEntries.length}\nErrors encountered for IDs:\n${results.failed.join('\n')}`, ui.ButtonSet.OK);
  } else {
    ui.alert('Sync Complete', `Religious Education schedules have been synchronized successfully (${results.successful}/${results.total}).\n\nChanges Made: ${results.logEntries.length}\nCheck the '${RE_SYNC_LOG_SHEET}' tab for details.`, ui.ButtonSet.OK);
  }
}

/**
 * ---------------------------------------------------------
 * 2. AUTOMATED RUN FUNCTION (For Time-Driven Triggers)
 * ---------------------------------------------------------
 */
function autoSyncRESchedules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quarterSheets = ss.getSheets().filter(s => /^Q[1-4]$/.test(s.getName()));
  
  for (const sheet of quarterSheets) {
    console.log(`Auto-syncing sheet: ${sheet.getName()}`);
    performSync(ss, sheet);
  }
}

/**
 * ---------------------------------------------------------
 * 3. CORE LOGIC (Shared by both manual and automated runs)
 * ---------------------------------------------------------
 */
function performSync(ss, sheet) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { total: 0, successful: 0, failed: [], logEntries: [] };
  }
  
  try {
    const validNames = getValidNamesFromTeams(ss);
    const hwsDates = getDatesFromHWS(sheet);
    
    if (Object.keys(hwsDates.amRows).length === 0 && Object.keys(hwsDates.pmRows).length === 0) {
      return { total: 0, successful: 0, failed: [], logEntries: [] };
    }
    
    const targetDateStrings = [...new Set([...Object.keys(hwsDates.amRows), ...Object.keys(hwsDates.pmRows)])];
    
    const activeConfig = getREConfig();
    let successCount = 0;
    let failedIds = [];
    let logEntries = [];

    for (const [className, config] of Object.entries(activeConfig)) {
      try {
        const classData = fetchREDataForDates(config.spreadsheetId, config, targetDateStrings);
        writeREDataToHWS(sheet, hwsDates, classData, config, className, logEntries, validNames);
        successCount++;
      } catch (e) {
        console.error(`Failed to sync RE schedule sheet ID: ${config.spreadsheetId} for class ${className}`, e);
        failedIds.push(config.spreadsheetId);
      }
    }
    
    if (logEntries.length > 0) {
      let logSheet = ss.getSheetByName(RE_SYNC_LOG_SHEET);
      if (!logSheet) {
         logSheet = ss.insertSheet(RE_SYNC_LOG_SHEET);
         logSheet.appendRow(["Timestamp", "Quarter", "Date", "Class", "Role", "Old Value", "New Value"]);
         logSheet.getRange("A1:G1").setFontWeight("bold");
         logSheet.setFrozenRows(1);
         logSheet.hideSheet();
      }
      logSheet.getRange(logSheet.getLastRow() + 1, 1, logEntries.length, logEntries[0].length).setValues(logEntries);
    }
    
    if (successCount > 0) bumpQuarterVersion(sheet.getName());
  return { total: Object.keys(activeConfig).length, successful: successCount, failed: failedIds, logEntries };
  } finally {
    lock.releaseLock();
  }
}

function getValidNamesFromTeams(ss) {
  const teamsSheet = ss.getSheetByName('Teams');
  if (!teamsSheet) return {};
  
  const data = teamsSheet.getDataRange().getValues();
  if (data.length < 2) return {}; 
  
  const headerRow = data[1]; 
  const classCols = {};
  const rosterByClass = { 'K': [], 'E1': [], 'E2': [], 'J1': [], 'J2': [], 'Sr': [] };
  
  for (let c = 0; c < headerRow.length; c++) {
    const header = String(headerRow[c]).trim();
    if (rosterByClass.hasOwnProperty(header)) {
      classCols[header] = c;
    }
  }
  
  for (let r = 2; r < data.length; r++) { 
    for (const [className, colIdx] of Object.entries(classCols)) {
      const name = String(data[r][colIdx]).trim();
      if (name) {
        rosterByClass[className].push(name);
      }
    }
  }
  
  return rosterByClass;
}

function resolveWorkerName(rawName, validNames) {
  if (!rawName) return '';
  const trimmedRaw = String(rawName).trim();
  const lowerRaw = trimmedRaw.toLowerCase();
  
  let match = validNames.find(n => n.toLowerCase() === lowerRaw);
  if (match) return match;
  
  match = validNames.find(n => n.toLowerCase().startsWith(lowerRaw));
  return match || trimmedRaw; 
}

function getDatesFromHWS(sheet) {
  const data = sheet.getRange("A1:A" + sheet.getLastRow()).getValues();
  
  let currentSection = null;
  const amRows = {};
  const pmRows = {};
  
  for (let i = 0; i < data.length; i++) {
    const val = data[i][0];
    
    if (typeof val === 'string' && val.toUpperCase().includes('SATURDAYS AM')) {
      currentSection = 'AM';
      continue;
    } else if (typeof val === 'string' && val.toUpperCase().includes('SATURDAYS PM')) {
      currentSection = 'PM';
      continue;
    }
    
    if (currentSection) {
      let dateObj = null;
      if (val instanceof Date) {
        dateObj = val;
      } else if (typeof val === 'string' && val.includes('/')) {
        const parsed = Date.parse(val);
        if (!isNaN(parsed)) dateObj = new Date(parsed);
      }

      if (dateObj) {
        const dateStr = formatDate(dateObj);
        if (currentSection === 'AM') {
          amRows[dateStr] = i + 1; 
        } else if (currentSection === 'PM') {
          pmRows[dateStr] = i + 1;
        }
      }
    }
  }
  
  return { amRows: amRows, pmRows: pmRows };
}

/**
 * Scans all sheets in the given spreadsheet to find schedule data for the target dates.
 * Features DYNAMIC column detection.
 */
const _reSpreadsheetCache = {};

function fetchREDataForDates(spreadsheetId, config, targetDateStrings) {
  const sourceSs = _reSpreadsheetCache[spreadsheetId] || (_reSpreadsheetCache[spreadsheetId] = SpreadsheetApp.openById(spreadsheetId));
  const sheets = sourceSs.getSheets();
  const scheduleData = {};
  
  for (const sheet of sheets) {
    const sheetNameLower = sheet.getName().toLowerCase();
    if (sheetNameLower.includes('summary') || 
        sheetNameLower.includes('workers') ||
        sheetNameLower.includes('teams') ||
        sheetNameLower.includes('log')) {
      continue;
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow < 2) continue;
    
    // --- DYNAMIC COLUMN DETECTION ---
    let dynamicConfig = { ...config };
    let headerRowIdx = -1;
    const searchRows = Math.min(15, lastRow);
    const topData = sheet.getRange(1, 1, searchRows, lastCol).getValues();
    
    // 1. Search the first 15 rows to find the row containing 'Teacher'
    for (let r = 0; r < searchRows; r++) {
      const rowStrings = topData[r].map(c => String(c).toLowerCase().trim());
      if (rowStrings.some(c => c.includes('teacher'))) {
        headerRowIdx = r;
        break;
      }
    }
    
    // 2. If we found the header row, dynamically map columns
    if (headerRowIdx !== -1) {
      const periodRowIdx = Math.max(0, headerRowIdx - 1);
      const periodRow = topData[periodRowIdx].map(c => String(c).toLowerCase().trim());
      const headerRow = topData[headerRowIdx].map(c => String(c).toLowerCase().trim());
      
      let currentPeriod = null;
      let dateCol = null;
      let amTeacherCol = null, amAssistantCol = null;
      let pmTeacherCol = null, pmAssistantCol = null;
      
      for (let c = 0; c < lastCol; c++) {
        const pVal = periodRow[c];
        const hVal = headerRow[c];
        
        // Update current period based on merged headers
        if (pVal.includes('morning') || pVal.includes('am')) {
          currentPeriod = 'AM';
        } else if (pVal.includes('afternoon') || pVal.includes('pm')) {
          currentPeriod = 'PM';
        }
        
        // Handle vertically merged 'Date' cell
        if (hVal === 'date' || hVal === 'class date' || pVal === 'date' || pVal === 'class date') {
          dateCol = c + 1; // 1-indexed
        } else if (hVal === 'teacher' || hVal === 'teachers' || hVal === 'lead teacher') {
          if (currentPeriod === 'AM') {
            amTeacherCol = c + 1;
          } else if (currentPeriod === 'PM') {
            pmTeacherCol = c + 1;
          } else {
            // Fallback: If no period found, rely on original config's expectations
            if (config.amTeacherCol && !amTeacherCol) {
              amTeacherCol = c + 1;
            } else if (config.pmTeacherCol && !pmTeacherCol) {
              pmTeacherCol = c + 1;
            }
          }
        } else if (hVal.includes('assistant') || hVal === 'ta' || hVal === 'teaching assistant') {
          if (currentPeriod === 'AM') {
            amAssistantCol = c + 1;
          } else if (currentPeriod === 'PM') {
            pmAssistantCol = c + 1;
          } else {
            if (config.amAssistantCol && !amAssistantCol) {
              amAssistantCol = c + 1;
            } else if (config.pmAssistantCol && !pmAssistantCol) {
              pmAssistantCol = c + 1;
            }
          }
        }
      }
      
      // Update dynamic config with found columns
      if (dateCol) dynamicConfig.dateCol = dateCol;
      if (amTeacherCol !== null) dynamicConfig.amTeacherCol = amTeacherCol;
      if (pmTeacherCol !== null) dynamicConfig.pmTeacherCol = pmTeacherCol;
      if (amAssistantCol !== null) dynamicConfig.amAssistantCol = amAssistantCol;
      if (pmAssistantCol !== null) dynamicConfig.pmAssistantCol = pmAssistantCol;
      
      // Auto-Adjust startRow to skip the headers entirely
      if (dynamicConfig.startRow <= headerRowIdx + 1) {
        dynamicConfig.startRow = headerRowIdx + 2; 
      }
    }
    // --- END DYNAMIC COLUMN DETECTION ---

    if (lastRow < dynamicConfig.startRow || lastCol < dynamicConfig.dateCol) continue;
    
    const data = sheet.getRange(dynamicConfig.startRow, 1, lastRow - dynamicConfig.startRow + 1, lastCol).getValues();
    
    for (const row of data) {
      const dateVal = row[dynamicConfig.dateCol - 1]; 
      let dateObj = null;
      
      if (dateVal instanceof Date) {
        dateObj = dateVal;
      } else if (typeof dateVal === 'string' && dateVal.includes('/')) {
        const parsed = Date.parse(dateVal);
        if (!isNaN(parsed)) {
          dateObj = new Date(parsed);
        }
      }
      
      if (dateObj) {
        const dateStr = formatDate(dateObj);
        
        if (targetDateStrings.includes(dateStr)) {
          scheduleData[dateStr] = {
            amTeacher: dynamicConfig.amTeacherCol ? row[dynamicConfig.amTeacherCol - 1] : '',
            amAssistant: dynamicConfig.amAssistantCol ? row[dynamicConfig.amAssistantCol - 1] : '',
            pmTeacher: dynamicConfig.pmTeacherCol ? row[dynamicConfig.pmTeacherCol - 1] : '',
            pmAssistant: dynamicConfig.pmAssistantCol ? row[dynamicConfig.pmAssistantCol - 1] : ''
          };
        }
      }
    }
  }
  
  return scheduleData;
}

function writeREDataToHWS(sheet, hwsDates, classData, config, className, logEntries, validNames) {
  const timestamp = new Date();
  const quarter = sheet.getName();
  const pendingUpdates = [];

  for (const [dateStr, data] of Object.entries(classData)) {
    function processRole(row, hwsCol, rawName, roleName) {
      if (!hwsCol || String(rawName).trim() === '') return;
      
      let rawStr = String(rawName).trim();
      let newVal = '';
      let logNewVal = '';
      
      if (rawStr.toLowerCase() === 'all') {
        newVal = ''; 
        logNewVal = '(Cleared)';
      } else {
        const classRoster = validNames[className] || [];
        newVal = resolveWorkerName(rawStr, classRoster);
        logNewVal = newVal;
      }
      
      pendingUpdates.push({
        row: row,
        col: hwsCol,
        newVal: newVal,
        logNewVal: logNewVal,
        dateStr: dateStr,
        roleName: roleName
      });
    }

    if (hwsDates.amRows[dateStr]) {
      const row = hwsDates.amRows[dateStr];
      processRole(row, config.hwsTeacherCol, data.amTeacher, 'AM Teacher');
      processRole(row, config.hwsAssistantCol, data.amAssistant, 'AM Assistant');
    }
    
    if (hwsDates.pmRows[dateStr]) {
      const row = hwsDates.pmRows[dateStr];
      processRole(row, config.hwsTeacherCol, data.pmTeacher, 'PM Teacher');
      processRole(row, config.hwsAssistantCol, data.pmAssistant, 'PM Assistant');
    }
  }

  if (pendingUpdates.length === 0) return;

  // Bounding box for batch read/write
  let minRow = pendingUpdates[0].row, maxRow = pendingUpdates[0].row;
  let minCol = pendingUpdates[0].col, maxCol = pendingUpdates[0].col;
  for (let i = 1; i < pendingUpdates.length; i++) {
    const u = pendingUpdates[i];
    if (u.row < minRow) minRow = u.row;
    if (u.row > maxRow) maxRow = u.row;
    if (u.col < minCol) minCol = u.col;
    if (u.col > maxCol) maxCol = u.col;
  }

  const numRows = maxRow - minRow + 1;
  const numCols = maxCol - minCol + 1;
  const range = sheet.getRange(minRow, minCol, numRows, numCols);
  const values = range.getValues();
  let changed = false;

  for (const u of pendingUpdates) {
    const rIdx = u.row - minRow;
    const cIdx = u.col - minCol;
    const oldVal = values[rIdx][cIdx];
    
    if (oldVal !== u.newVal) {
      values[rIdx][cIdx] = u.newVal;
      logEntries.push([timestamp, quarter, u.dateStr, className, u.roleName, oldVal, u.logNewVal]);
      changed = true;
    }
  }

  if (changed) {
    range.setValues(values);
  }
}

function formatDate(dateObj, tz) {
  if (!tz) {
    try {
      tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    } catch (e) {
      tz = 'America/Los_Angeles';
    }
  }
  return Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
}