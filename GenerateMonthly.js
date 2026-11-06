/**
 * Monthly Schedule Generator - Apps Script
 * File: GenerateMonthly.gs
 */

// =========================================================================
// CONFIGURATION
// =========================================================================

var TEMPLATE_SHEET_NAME = "MonthlyTemplate";
var CLEANING_START_COLUMN = 3; // 3 = Column C. Change to 2 for Col B, 4 for Col D, etc.

// Add the Google account emails of the admins allowed to edit the generated sheets.
var AUTHORIZED_EDITORS = ["david.liou@tjc.org", "anthony.tsai@tjc.org", "gg.info@tjc.org", "jessica.chen@tjc.org", "anthony.chang@tjc.org", "jessica.jeh.chen@gmail.com"]; 

// URL for the external Bible Reading Plan spreadsheet
var BIBLE_PLAN_URL = "https://docs.google.com/spreadsheets/d/1Omlqzklhiw7I1Wyp-5tgxuVZQdUlszZEufPh5O5zCpM/edit";

// =========================================================================
// MENU & UI
// =========================================================================

function promptMonthlyGenerator() {
  var ui = SpreadsheetApp.getUi();
  var now = new Date();
  var nextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  var exampleDate = ('0' + (nextMonth.getMonth() + 1)).slice(-2) + '/' + nextMonth.getFullYear();
  var response = ui.prompt(
    'Generate / Sync Monthly Schedule',
    'Enter the Month and Year you want to generate or sync (e.g., ' + exampleDate + '):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK) {
    var input = response.getResponseText().trim();
    if (!input.match(/^(0?[1-9]|1[0-2])\/\d{4}$/)) {
      ui.alert("Invalid format. Please use MM/YYYY (e.g., " + exampleDate + ").");
      return;
    }
    
    var parts = input.split("/");
    var month = parseInt(parts[0], 10);
    var year = parseInt(parts[1], 10);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    var targetName = monthNames[month - 1] + " " + year;
    var existingSheet = ss.getSheetByName(targetName);

    var genOptions = {};
    if (existingSheet) {
      var syncChoice = ui.alert(
        "Sync Existing Schedule: " + targetName,
        "The schedule for " + targetName + " already exists.\n\n" +
        "• Click YES to Sync In-Place (Updates data from Q sheets while preserving published links & manual overrides).\n" +
        "• Click NO for Full Rebuild from Template (Resets all manual overrides while keeping the same tab and link).\n" +
        "• Click CANCEL to abort.",
        ui.ButtonSet.YES_NO_CANCEL
      );
      
      if (syncChoice === ui.Button.YES) {
        genOptions = { preserveManualOverrides: true, inPlace: true };
      } else if (syncChoice === ui.Button.NO) {
        genOptions = { preserveManualOverrides: false, fullRebuildKeepGid: true };
      } else {
        return; 
      }
    }
    
    try {
      var result = generateMonthlySheet(year, month, genOptions);
      var alertTitle = "Success! 🎉";
      var alertMsg = "Successfully generated and synced the schedule for " + input + " (" + targetName + ").\n\nThe sheet has been locked to authorized editors.";
      
      if (result && result.biblePlanStatus && !result.biblePlanStatus.success) {
        alertTitle = "Generated with Warning ⚠️";
        alertMsg += "\n\n⚠️ Bible Reading Plan Notice:\n" + result.biblePlanStatus.message + "\n(Monthly schedule roles were populated, but Bible reading passages could not be synced).";
      }
      
      if (result && result.missingCount > 0) {
        alertMsg += "\n\n📌 Note: " + result.missingCount + " unassigned slot(s) were highlighted in light pink for coordinator attention.";
      }
      
      ui.alert(alertTitle, alertMsg, ui.ButtonSet.OK);
    } catch (e) {
      ui.alert("Error", e.message, ui.ButtonSet.OK);
    }
  }
}

// =========================================================================
// CORE LOGIC
// =========================================================================

function generateMonthlySheet(year, month, options) {
  options = options || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error("Could not acquire lock. Another schedule generation or update is in progress. Please wait and try again.");
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone();
  
  var runTime = Utilities.formatDate(new Date(), tz, "MM/dd/yyyy hh:mm:ss a");
  var logs = [["Run Time", "Action", "DayType", "Date", "Template Row", "Target Col", "Value", "Notes"]];
  function addLog(action, dayType, date, roleRow, col, val, notes) {
      logs.push([runTime, action, dayType, date, roleRow, col, val, notes]);
  }
  addLog("Start Run", "ALL", month + "/" + year, "-", "-", "-", "Starting generator script");

  var MAPPINGS = loadMappings(ss);

  var template = ss.getSheetByName(TEMPLATE_SHEET_NAME);
  if (!template) {
    throw new Error("Could not find a sheet named '" + TEMPLATE_SHEET_NAME + "'. Please create one.");
  }
  
  var monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var targetName = monthNames[month - 1] + " " + year;
  var existingSheet = ss.getSheetByName(targetName);
  
  var newSheet;
  var isInPlace = false;
  if (existingSheet) {
    if (options.fullRebuildKeepGid) {
      newSheet = existingSheet;
      newSheet.clear();
      template.getDataRange().copyTo(newSheet.getRange(1, 1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      template.getDataRange().copyTo(newSheet.getRange(1, 1), SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
    } else if (options.inPlace || options.preserveManualOverrides) {
      newSheet = existingSheet;
      isInPlace = true;
    } else {
      throw new Error("A sheet named '" + targetName + "' already exists.");
    }
  } else {
    newSheet = template.copyTo(ss).setName(targetName);
    newSheet.showSheet();
  }
  
  var qSheetName = "Q" + Math.ceil(month / 3);
  var qSheet = ss.getSheetByName(qSheetName);
  if (!qSheet) {
    throw new Error("Could not find the master database sheet named '" + qSheetName + "'.");
  }
  
  // 4. Calculate all dates for the month
  var daysInMonth = new Date(year, month, 0).getDate();
  var datesMap = { WEEK: [] };
  
  var dayMapNum = { 'SUN':0, 'MON':1, 'TUE':2, 'WED':3, 'THU':4, 'FRI':5, 'SAT':6 };
  var anchorDay = 6; // Default SAT
  var activeDays = {};
  if (typeof getChurchConfig === 'function') {
    var config = getChurchConfig(false);
    var layout = config.serviceLayout || [];
    for (var l = 0; l < layout.length; l++) {
      if (layout[l].enabled === 'Y') {
        activeDays[layout[l].day_type] = dayMapNum[layout[l].day_type];
        if (!datesMap[layout[l].day_type]) datesMap[layout[l].day_type] = [];
      }
    }
    if (config.profile.week_anchor_day) {
      anchorDay = dayMapNum[config.profile.week_anchor_day.toUpperCase()] !== undefined ? dayMapNum[config.profile.week_anchor_day.toUpperCase()] : 6;
    }
  }
  
  if (Object.keys(activeDays).length === 0) {
    activeDays = { 'SAT':6, 'MON':1, 'TUE':2, 'WED':3, 'FRI':5 };
    datesMap.SAT = []; datesMap.MON = []; datesMap.TUE = []; datesMap.WED = []; datesMap.FRI = [];
  }
  
  for (var i = 1; i <= daysInMonth; i++) {
    var d = new Date(Date.UTC(year, month - 1, i, 12, 0, 0));
    var dayOfWeek = d.getUTCDay(); 
    var formattedDate = (month) + "/" + i;
    
    for (var dayType in activeDays) {
      if (dayOfWeek === activeDays[dayType]) {
        datesMap[dayType].push({ label: formattedDate, dateObj: d });
      }
    }
  }

  // Calculate WEEK ranges based on anchorDay
  var currentStart = 1;
  var anchorCount = 0;
  for (var i = 1; i <= daysInMonth; i++) {
    var d = new Date(Date.UTC(year, month - 1, i, 12, 0, 0));
    if (d.getUTCDay() === anchorDay) {
      anchorCount++;
      if (anchorCount <= 4) { 
        if (currentStart === i) {
          datesMap.WEEK.push({ label: month + "/" + i, dateObj: d, startDay: currentStart, endDay: i });
        } else {
          datesMap.WEEK.push({ label: month + "/" + currentStart + "-" + month + "/" + i, dateObj: d, startDay: currentStart, endDay: i });
        }
        currentStart = i + 1;
      }
    }
  }
  if (currentStart <= daysInMonth) {
    var lastD = new Date(Date.UTC(year, month - 1, daysInMonth, 12, 0, 0));
    if (currentStart === daysInMonth) {
      datesMap.WEEK.push({ label: month + "/" + daysInMonth, dateObj: lastD, startDay: currentStart, endDay: daysInMonth });
    } else {
      datesMap.WEEK.push({ label: month + "/" + currentStart + "-" + month + "/" + daysInMonth, dateObj: lastD, startDay: currentStart, endDay: daysInMonth });
    }
  }
  
  var qData = qSheet.getDataRange().getValues();
  var qDataDisplay = qSheet.getDataRange().getDisplayValues(); 
  
  var formattedRange = newSheet.getDataRange();
  formattedRange.setWrap(true);
  formattedRange.setVerticalAlignment("middle");
  
  // 7. Find placeholders and populate dates
  // Always use template values to locate placeholders reliably even during in-place sync
  var templateRange = template.getDataRange();
  var templateValues = templateRange.getValues();
  var templateBackgrounds = templateRange.getBackgrounds();
  var placeholderLocations = {};
  
  for (var r = 0; r < templateValues.length; r++) {
    for (var c = 0; c < templateValues[r].length; c++) {
      var cellVal = String(templateValues[r][c]);
      if (cellVal.indexOf("{{") !== -1 && cellVal.indexOf("}}") !== -1) {
        var updatedCellVal = cellVal;

        // Month / Year replacement
        if (updatedCellVal.indexOf("{{MONTH_YEAR}}") !== -1) {
          var titleStr = getMonthYearTitle(year, month);
          updatedCellVal = updatedCellVal.replace("{{MONTH_YEAR}}", titleStr);
          newSheet.getRange(r + 1, c + 1).setValue(updatedCellVal);
        }

        // Council Member replacement
        if (updatedCellVal.indexOf("{{COUNCIL_MEMBER}}") !== -1) {
          var councilStr = getCouncilMemberString(ss, month);
          updatedCellVal = updatedCellVal.replace("{{COUNCIL_MEMBER}}", councilStr);
          newSheet.getRange(r + 1, c + 1).setValue(updatedCellVal);
        }

        var match = updatedCellVal.match(/\{\{([A-Z]+)_(\d+)\}\}/);
        if (match) {
          var placeholderKey = match[1] + "_" + match[2]; 
          var dayType = match[1];
          var index = parseInt(match[2], 10) - 1;
          
          if (datesMap[dayType] && index < datesMap[dayType].length) {
            var dateItem = datesMap[dayType][index];
            var dateStr = (dateItem.dateObj.getUTCMonth() + 1) + "/" + dateItem.dateObj.getUTCDate();
            
            if (dayType === "WEEK") { 
                dateStr = dateItem.label; 
                dateItem.targetCol = c + 1; // Save target column for Bible Plan injection
            }
            
            var originalCellVal = String(updatedCellVal).trim();
            newSheet.getRange(r + 1, c + 1).setValue(updatedCellVal.replace("{{" + placeholderKey + "}}", dateStr));
            
            if (!placeholderLocations[placeholderKey]) placeholderLocations[placeholderKey] = [];
            placeholderLocations[placeholderKey].push({ 
                r: r + 1, 
                c: c + 1, 
                exact: (originalCellVal === "{{" + placeholderKey + "}}") 
            });
            
          } else {
            newSheet.getRange(r + 1, c + 1).setValue(updatedCellVal.replace("{{" + placeholderKey + "}}", ""));
          }
        }
      }
    }
  }

  // Read manual overrides from MonthlyOverridesLog if preserving overrides
  var overrideMap = {};
  if (options && options.preserveManualOverrides) {
    var logSheet = ss.getSheetByName("MonthlyOverridesLog");
    if (logSheet && logSheet.getLastRow() > 1) {
      var logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 6).getValues();
      for (var lr = 0; lr < logData.length; lr++) {
        var logMonth = String(logData[lr][2] || "").trim();
        var cellA1 = String(logData[lr][3] || "").trim().toUpperCase();
        var newVal = logData[lr][5];
        if (logMonth === targetName && cellA1) {
          if (overrideMap[cellA1] === undefined) {
            overrideMap[cellA1] = (newVal === "(empty)") ? "" : newVal;
          }
        }
      }
    }
  }

  // --- FETCH AND INJECT BIBLE READING PLAN ---
  var biblePlanStatus = { success: false, message: "" };
  try {
      var planUrl = "";
      if (typeof getChurchConfig === 'function') {
        var churchConfig = getChurchConfig(false);
        planUrl = churchConfig.profile.bible_plan_url || "";
      }
      if (!planUrl && typeof BIBLE_PLAN_URL !== 'undefined') {
        planUrl = BIBLE_PLAN_URL;
      }
      if (!planUrl) {
        throw new Error("No Bible Plan URL configured.");
      }
      var extSs = SpreadsheetApp.openByUrl(planUrl);
      var extSheetName = year + " Bible Reading Plan";
      var extSheet = extSs.getSheetByName(extSheetName);

      if (extSheet) {
          var maxRows = Math.max(extSheet.getLastRow(), 100);
          var maxCols = Math.max(extSheet.getLastColumn(), 30);
          var extData = extSheet.getRange(1, 1, maxRows, maxCols).getValues(); 
          
          var searchMonth = monthNames[month - 1]; 
          var headerRowIdx = -1;
          var bookColIdx = -1;
          
          // Formatter to prevent Google Sheet objects from failing
          function formatVal(v) {
              if (v instanceof Date) return (v.getMonth() + 1) + "-" + v.getDate();
              return String(v || "").trim();
          }

          // Locate the exact coordinates of this month's header
          for (var r = 0; r < extData.length; r++) {
              var rowStr = extData[r].map(function(c) { return formatVal(c).toLowerCase(); }).join(" ");
              if (rowStr.indexOf(searchMonth.toLowerCase()) !== -1 && (rowStr.indexOf("day") !== -1 || rowStr.indexOf("日") !== -1)) {
                  for (var c = 0; c < extData[r].length; c++) {
                      if (formatVal(extData[r][c]).toLowerCase().indexOf(searchMonth.toLowerCase()) !== -1) {
                          headerRowIdx = r;
                          bookColIdx = c;
                          break;
                      }
                  }
                  break;
              }
          }

          if (headerRowIdx !== -1 && bookColIdx !== -1) {
              
              var startDataRow = headerRowIdx + 1; 
              for (var rOffset = 1; rOffset <= 3; rOffset++) {
                  if (headerRowIdx + rOffset >= extData.length) break;
                  if (formatVal(extData[headerRowIdx + rOffset][bookColIdx]) !== "") {
                      startDataRow = headerRowIdx + rOffset;
                      break;
                  }
              }

              var weekItems = datesMap.WEEK;
              for (var w = 0; w < weekItems.length; w++) {
                  var wItem = weekItems[w];
                  if (!wItem.targetCol) continue;

                  var readings = [];
                  var currentBook = null;
                  var startChap = null;
                  var endChap = null;

                  for (var d = wItem.startDay; d <= wItem.endDay; d++) {
                      var rowIdx = startDataRow + (d - 1);
                      if (rowIdx >= extData.length) break;

                      var book = formatVal(extData[rowIdx][bookColIdx]);

                      if (!book || book.toLowerCase().indexOf("sabbath") !== -1 || book.indexOf("安息") !== -1) {
                          continue; 
                      }

                      var chapStr = "";
                      var extraBookText = "";
                      
                      // Handle the 3-column split (English | Chinese | Chapter)
                      for (var offset = 1; offset <= 2; offset++) {
                          var testCol = bookColIdx + offset;
                          if (testCol >= extData[0].length) break;
                          var val = formatVal(extData[rowIdx][testCol]);
                          
                          if (val !== "") {
                              if (/\d/.test(val)) {
                                  // Found a digit, it's the chapter!
                                  chapStr = val;
                                  break;
                              } else if (/[\u4e00-\u9fa5]/.test(val)) {
                                  // Found Chinese characters, add to the book name
                                  extraBookText += " " + val;
                              }
                          }
                      }
                      
                      book = (book + extraBookText).trim();

                      if (book !== currentBook) {
                          if (currentBook) {
                              var entry = (startChap === endChap) ? (currentBook + " " + startChap) : (currentBook + " " + startChap + "-" + endChap);
                              readings.push(entry);
                          }
                          currentBook = book;
                          startChap = chapStr;
                          endChap = chapStr;
                      } else {
                          endChap = chapStr; 
                      }
                  }
                  
                  if (currentBook) {
                      var entry = (startChap === endChap) ? (currentBook + " " + startChap) : (currentBook + " " + startChap + "-" + endChap);
                      readings.push(entry);
                  }

                  var finalString = readings.join("\n").trim();
                  if (finalString !== "") {
                      var bibleCellA1 = columnIndexToLetter(wItem.targetCol) + "5";
                      if (overrideMap.hasOwnProperty(bibleCellA1)) {
                          newSheet.getRange(5, wItem.targetCol).setValue(overrideMap[bibleCellA1]);
                      } else {
                          newSheet.getRange(5, wItem.targetCol).setValue(finalString);
                      }
                  }
              }
              biblePlanStatus = { success: true, message: "Bible Reading Plan successfully injected." };
              addLog("Bible Plan Sync", "ALL", month + "/" + year, "-", "Row 5", "-", "Successfully dynamically located and injected Bible reading plan.");
          } else {
              biblePlanStatus = { success: false, message: "Could not find calendar column header for '" + searchMonth + "' in '" + extSheetName + "' tab." };
              addLog("Bible Plan Sync", "ALL", month + "/" + year, "-", "-", "-", biblePlanStatus.message);
          }
      } else {
          biblePlanStatus = { success: false, message: "Could not find tab named '" + extSheetName + "' in external Bible Reading Plan spreadsheet." };
          addLog("Bible Plan Sync", "ALL", month + "/" + year, "-", "-", "-", biblePlanStatus.message);
      }
  } catch (e) {
      biblePlanStatus = { success: false, message: "Error opening or reading Bible Reading Plan: " + e.message };
      addLog("Bible Plan Sync", "ALL", month + "/" + year, "-", "-", "-", biblePlanStatus.message);
  }

  // 7.1 Resolve True Target Columns
  var minTemplateRows = {};
  for (var m = 0; m < MAPPINGS.length; m++) {
      var map = MAPPINGS[m];
      if (!map.isCleaning) {
          if (!minTemplateRows[map.dayType] || map.templateRow < minTemplateRows[map.dayType]) {
              minTemplateRows[map.dayType] = map.templateRow;
          }
      }
  }

  Object.keys(datesMap).forEach(function(dayType) {
    if (dayType === "WEEK") return;
    var datesArray = datesMap[dayType];
    var minRow = minTemplateRows[dayType] || 9999;
    
    for (var i = 0; i < datesArray.length; i++) {
        var dateItem = datesArray[i];
        var placeholderKey = dayType + "_" + (i + 1);
        var locs = placeholderLocations[placeholderKey] || [];
        var bestCol = null;
        var bestDistance = 9999;
        
        for (var l = 0; l < locs.length; l++) {
            var loc = locs[l];
            if (loc.exact && loc.r <= minRow) {
                var dist = minRow - loc.r;
                if (dist < bestDistance) {
                    bestDistance = dist;
                    bestCol = loc.c;
                }
            }
        }
        
        if (!bestCol && locs.length > 0) {
            bestCol = locs[locs.length - 1].c; 
        }
        
        dateItem.targetCol = bestCol;
        if (bestCol) {
            var msg = "Selected header location.";
            if (bestDistance > 3 && bestDistance !== 9999) {
                msg = "WARNING: Hijack averted. Ignored rogue placeholder located " + bestDistance + " rows away.";
            }
            addLog("Resolved Target Col", dayType, dateItem.label, "Header Row " + (minRow - bestDistance), bestCol, placeholderKey, msg);
        } else {
            addLog("Missing Target Col", dayType, dateItem.label, "-", "-", placeholderKey, "Could not find this placeholder in the template.");
        }
    }
  });

  // 7.5 Fix Cleaning Schedule Headers
  var cleaningHeaderRow = 58;
  var cleaningDateRow = 59;
  var cleaningEnabled = 'Y';
  var cleaningStartCol = CLEANING_START_COLUMN;
  
  if (typeof getChurchConfig === 'function') {
    var churchConfig = getChurchConfig(false);
    if (churchConfig.profile.cleaning_header_row) cleaningHeaderRow = parseInt(churchConfig.profile.cleaning_header_row, 10);
    if (churchConfig.profile.cleaning_date_row) cleaningDateRow = parseInt(churchConfig.profile.cleaning_date_row, 10);
    if (churchConfig.profile.cleaning_enabled) cleaningEnabled = churchConfig.profile.cleaning_enabled;
    if (churchConfig.profile.cleaning_start_col) cleaningStartCol = parseInt(churchConfig.profile.cleaning_start_col, 10);
  }
  
  if (cleaningEnabled === 'Y') {
    var satLen = (datesMap.SAT && datesMap.SAT.length) ? datesMap.SAT.length : 0;
    var friLen = (datesMap.FRI && datesMap.FRI.length) ? datesMap.FRI.length : 0;
    var maxCleaningWeeks = Math.max(satLen, friLen);
  
    for (var w = 0; w < 5; w++) {
      var pairBaseCol = cleaningStartCol + (w * 2); 
      if (w < maxCleaningWeeks) {
        var satDateItem = datesMap.SAT ? datesMap.SAT[w] : null;
        var friDateItem = datesMap.FRI ? datesMap.FRI[w] : null;
        
        var satFirst = true;
        if (satDateItem && friDateItem) {
            if (friDateItem.dateObj.getTime() < satDateItem.dateObj.getTime()) { satFirst = false; }
        } else if (friDateItem && !satDateItem) { satFirst = false; }
        
        var col1Name = satFirst ? "Saturday" : "Friday";
        var col2Name = satFirst ? "Friday" : "Saturday";
        var col1DateStr = satFirst ? (satDateItem ? satDateItem.label : "") : (friDateItem ? friDateItem.label : "");
        var col2DateStr = satFirst ? (friDateItem ? friDateItem.label : "") : (satDateItem ? satDateItem.label : "");
        
        newSheet.getRange(cleaningHeaderRow, pairBaseCol).setValue(col1Name);
        newSheet.getRange(cleaningHeaderRow, pairBaseCol + 1).setValue(col2Name);
        newSheet.getRange(cleaningDateRow, pairBaseCol).setValue(col1DateStr);
        newSheet.getRange(cleaningDateRow, pairBaseCol + 1).setValue(col2DateStr);
      } else {
        // Clear week 5 headers and dates if month only has 4 weeks
        newSheet.getRange(cleaningHeaderRow, pairBaseCol, 2, 2).clearContent();
        // Clear week 5 cleaning assignment worker rows (rows 60 to 72)
        newSheet.getRange(cleaningDateRow + 1, pairBaseCol, 15, 2).clearContent();
        // Restore template background for week 5 cleaning cells
        var w5Bgs = [];
        for (var br = cleaningDateRow + 1; br <= cleaningDateRow + 15; br++) {
          var rowBgs = [];
          for (var bc = pairBaseCol; bc <= pairBaseCol + 1; bc++) {
            var tBg = (templateBackgrounds[br - 1] && templateBackgrounds[br - 1][bc - 1]) ? templateBackgrounds[br - 1][bc - 1] : "#ffffff";
            rowBgs.push(tBg || "#ffffff");
          }
          w5Bgs.push(rowBgs);
        }
        newSheet.getRange(cleaningDateRow + 1, pairBaseCol, 15, 2).setBackgrounds(w5Bgs);
      }
    }
  }
  
  // 8. Populate Assignments
  var cellWrites = {};
  var targetedCells = {};

  Object.keys(datesMap).forEach(function(dayType) {
    if (dayType === "WEEK") return; 
    var datesArray = datesMap[dayType];
    
    for (var i = 0; i < datesArray.length; i++) {
      var dateItem = datesArray[i];
      
      for (var m = 0; m < MAPPINGS.length; m++) {
        var mapping = MAPPINGS[m];
        if (mapping.dayType === dayType) {
          
          var targetCol = null;
          if (mapping.isCleaning) {
             var pairBaseCol = cleaningStartCol + (i * 2);
             var satDateItem = datesMap.SAT ? datesMap.SAT[i] : null;
             var friDateItem = datesMap.FRI ? datesMap.FRI[i] : null;
             var satFirst = true;
             
             if (satDateItem && friDateItem) {
                if (friDateItem.dateObj.getTime() < satDateItem.dateObj.getTime()) { satFirst = false; }
             } else if (friDateItem && !satDateItem) { satFirst = false; }
             
             if (dayType === "SAT") {
                targetCol = satFirst ? pairBaseCol : pairBaseCol + 1;
             } else if (dayType === "FRI") {
                targetCol = satFirst ? pairBaseCol + 1 : pairBaseCol;
             }
          } else {
             targetCol = dateItem.targetCol;
          }
          
          if (!targetCol) continue;

          var writeCol = targetCol;
          if (mapping.colOffset) writeCol += mapping.colOffset;
          var key = mapping.templateRow + "_" + writeCol;
          
          targetedCells[key] = { dayType: dayType, isCleaning: mapping.isCleaning };

          var masterRows = [];
          var targetMonth = dateItem.dateObj.getUTCMonth() + 1;
          var targetDay = dateItem.dateObj.getUTCDate();

          for (var qr = 5; qr < qData.length; qr++) { 
            var qDateStr = qDataDisplay[qr][0];
            var qDateObj = qData[qr][0];
            
            if (!qDateStr && !qDateObj) continue;

            var qMonth = -1, qDay = -1;
            
            var parts = String(qDateStr).trim().split("/");
            if (parts.length >= 2 && !isNaN(parseInt(parts[0], 10))) {
                qMonth = parseInt(parts[0], 10);
                qDay = parseInt(parts[1], 10);
            } 
            else if (qDateObj && typeof qDateObj.getMonth === 'function') {
                var dStr = Utilities.formatDate(qDateObj, tz, "M/d/yyyy");
                var dParts = dStr.split("/");
                qMonth = parseInt(dParts[0], 10);
                qDay = parseInt(dParts[1], 10);
            }
            else if (qDateStr) {
                var isoMatch = String(qDateStr).match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
                if (isoMatch) {
                    qMonth = parseInt(isoMatch[2], 10);
                    qDay = parseInt(isoMatch[3], 10);
                } else {
                    var parsedDate = new Date(qDateStr);
                    if (!isNaN(parsedDate.getTime())) {
                        var dStr = Utilities.formatDate(parsedDate, tz, "M/d/yyyy");
                        var dParts = dStr.split("/");
                        qMonth = parseInt(dParts[0], 10);
                        qDay = parseInt(dParts[1], 10);
                    }
                }
            }
            
            if (qMonth === targetMonth && qDay === targetDay) {
                masterRows.push(qr);
            }
          }
          
          if (masterRows.length > 0) {
              var mRowIndex = 0;
              if (dayType === "SAT") {
                  if ((mapping.isPM || mapping.isCleaning) && masterRows.length > 1) {
                      mRowIndex = 1; 
                  }
              }
              var masterRow = masterRows[mRowIndex];
              
              function isValidWorker(name) {
                if (!name) return false;
                var lower = String(name).trim().toLowerCase();
                if (lower === "" || lower === "-" || lower === "s/a" || lower === "x") return false;
                return true;
              }

              if (masterRow !== undefined) {
                 var cellData = qData[masterRow][mapping.masterCol];
                 var workerName = cellData === undefined ? "" : String(cellData).trim();
                 
                 // If the selected row does not have a valid worker, check backup rows for the same date
                 if (!isValidWorker(workerName) && masterRows.length > 1) {
                     for (var bIdx = 0; bIdx < masterRows.length; bIdx++) {
                         var backupRow = masterRows[bIdx];
                         var bData = qData[backupRow][mapping.masterCol];
                         var bName = bData === undefined ? "" : String(bData).trim();
                         if (isValidWorker(bName)) {
                             masterRow = backupRow;
                             workerName = bName;
                             break;
                         }
                     }
                 }

                 if (isValidWorker(workerName)) {
                    if (cellWrites[key]) {
                        if (cellWrites[key].indexOf(workerName) === -1) {
                            cellWrites[key] += "\n" + workerName; 
                        }
                    } else {
                        cellWrites[key] = workerName;
                    }
                 }
              }
          }
        }
      }
    }
  });
  
  // 8. Batch write cell assignments and warning formatting
  var rowGroups = {};
  var redCellA1s = [];
  var bgGroups = {};
  
  Object.keys(targetedCells).forEach(function(key) {
      var parts = key.split("_");
      var r = parseInt(parts[0], 10);
      var c = parseInt(parts[1], 10);
      var cellA1 = columnIndexToLetter(c) + r;
      var text = overrideMap.hasOwnProperty(cellA1) ? overrideMap[cellA1] : (cellWrites[key] || "");
      
      if (!rowGroups[r]) rowGroups[r] = [];
      rowGroups[r].push({ col: c, text: text });
      
      var isBlank = (!text || String(text).trim() === "");
      var tInfo = targetedCells[key];
      var isFriCleaningExempt = (tInfo && tInfo.dayType === "FRI" && tInfo.isCleaning && r >= 63 && r <= 72);
      
      if (isBlank && !isFriCleaningExempt) {
          redCellA1s.push(cellA1);
      } else {
          // Restore original template background color (clears previous #ffcccc if cell is now filled)
          var origBg = (templateBackgrounds[r - 1] && templateBackgrounds[r - 1][c - 1]) ? templateBackgrounds[r - 1][c - 1] : "#ffffff";
          if (!origBg) origBg = "#ffffff";
          if (!bgGroups[origBg]) bgGroups[origBg] = [];
          bgGroups[origBg].push(cellA1);
      }
  });

  // Batch write row by row
  Object.keys(rowGroups).forEach(function(rowStr) {
      var r = parseInt(rowStr, 10);
      var items = rowGroups[r];
      var minCol = items[0].col;
      var maxCol = items[0].col;
      for (var i = 1; i < items.length; i++) {
          if (items[i].col < minCol) minCol = items[i].col;
          if (items[i].col > maxCol) maxCol = items[i].col;
      }
      var width = maxCol - minCol + 1;
      var rowRange = newSheet.getRange(r, minCol, 1, width);
      var rowVals = rowRange.getValues();
      
      items.forEach(function(item) {
          rowVals[0][item.col - minCol] = item.text;
      });
      
      rowRange.setValues(rowVals);
  });

  // Batch restore original template backgrounds for filled/exempt cells (clears stale #ffcccc from previous runs)
  Object.keys(bgGroups).forEach(function(bg) {
      if (bgGroups[bg].length > 0) {
          newSheet.getRangeList(bgGroups[bg]).setBackground(bg);
      }
  });

  // Batch apply all missing assignment warning highlights in a single call
  if (redCellA1s.length > 0) {
      newSheet.getRangeList(redCellA1s).setBackground("#ffcccc");
  }
  
  // 9. Write Logs to Persistent Sheet
  var logSheet = ss.getSheetByName("MonthlyGeneratorLog");
  var isNewLogSheet = false;
  
  if (!logSheet) {
      logSheet = ss.insertSheet("MonthlyGeneratorLog");
      logSheet.hideSheet();
      isNewLogSheet = true;
  }
  
  var lastRow = logSheet.getLastRow();
  
  if (isNewLogSheet || lastRow === 0 || logSheet.getLastColumn() !== logs[0].length) {
      logSheet.clear();
      logSheet.getRange(1, 1, 1, logs[0].length).setValues([logs[0]]);
      logSheet.getRange(1, 1, 1, logs[0].length).setFontWeight("bold");
      logSheet.setFrozenRows(1);
      try { logSheet.getRange(1, 1, logSheet.getMaxRows(), logs[0].length).createFilter(); } catch(e) {}
      lastRow = 1;
  } else {
      logSheet.getRange(lastRow + 1, 1, 1, logs[0].length).setBackground("#e0e0e0");
      lastRow++;
  }
  
  var logData = logs.slice(1);
  if (logData.length > 0) {
      logSheet.getRange(lastRow + 1, 1, logData.length, logData[0].length).setValues(logData);
  }
  logSheet.autoResizeColumns(1, logs[0].length);
  

  // =======================================================================
  // 10. LOCK THE SHEET TO PREVENT MANUAL OVERRIDES
  // =======================================================================
  var existingProtections = newSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  var protection = (existingProtections && existingProtections.length > 0) ? existingProtections[0] : newSheet.protect().setDescription('Locked: Generated from Q Sheets');
  
  var currentEditors = protection.getEditors();
  protection.removeEditors(currentEditors);
  
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
  
  var editors = [];
  if (typeof getChurchConfig === 'function') {
    var churchConfig = getChurchConfig(false);
    var authString = churchConfig.profile.authorized_editors || '';
    if (authString) {
      editors = authString.split(',').map(function(e) { return e.trim(); }).filter(function(e) { return e; });
    }
  }
  if (editors.length === 0 && typeof AUTHORIZED_EDITORS !== 'undefined' && AUTHORIZED_EDITORS.length > 0) {
    editors = AUTHORIZED_EDITORS;
  }
  if (editors.length > 0) {
    protection.addEditors(editors);
  }
  
  ss.setActiveSheet(newSheet);
  return {
    success: true,
    sheetName: targetName,
    biblePlanStatus: biblePlanStatus,
    missingCount: redCellA1s.length
  };
  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// HELPERS & MAPPING TAB LOGIC
// =========================================================================

function getCouncilMemberString(ss, month) {
  var annSheet = ss.getSheetByName("Announcements");
  if (!annSheet) return "";
  
  var data = annSheet.getDataRange().getValues();
  var rowIndex = month; 
  
  if (rowIndex < data.length) {
    var primary = String(data[rowIndex][2]).trim(); // Col C
    var backup = String(data[rowIndex][3]).trim();  // Col D
    
    if (primary && backup) return primary + "/" + backup;
    if (primary) return primary;
    if (backup) return backup;
  }
  return "";
}

function loadMappings(ss) {
  var configSheet = ss.getSheetByName("MappingConfig");
  if (!configSheet) configSheet = createDefaultConfigSheet(ss);
  
  var data = configSheet.getDataRange().getValues();
  var mappings = [];
  
  for (var i = 1; i < data.length; i++) {
    var dayType = String(data[i][0]).trim().toUpperCase();
    var section = String(data[i][2]).trim().toUpperCase();
    var masterColStr = String(data[i][3]).trim().toUpperCase();
    var templateRow = parseInt(data[i][4], 10);
    
    if (!dayType || !masterColStr || isNaN(templateRow)) continue;
    
    var masterCol = letterToColumnIndex(masterColStr);
    
    mappings.push({
      dayType: dayType,
      isPM: (section === "PM"),
      isCleaning: (section === "CLEANING"),
      masterCol: masterCol,
      templateRow: templateRow,
      colOffset: 0
    });
  }
  return mappings;
}

function letterToColumnIndex(letter) {
  var col = 0;
  var len = letter.length;
  for (var i = 0; i < len; i++) {
    col += (letter.charCodeAt(i) - 64) * Math.pow(26, len - i - 1);
  }
  return col - 1; 
}

function columnIndexToLetter(n) {
  var r = '';
  while (n > 0) {
    var rem = (n - 1) % 26;
    r = String.fromCharCode(65 + rem) + r;
    n = Math.floor((n - 1) / 26);
  }
  return r;
}

function createDefaultConfigSheet(ss) {
  var configSheet = ss.insertSheet("MappingConfig");
  var headers = ["Day Type (SAT/FRI etc)", "Role Name", "Section (AM/PM/CLEANING)", "Master Column (Letter)", "Template Row (Number)"];
  
  var defaultData = [];
  var isGG = true;
  if (typeof getChurchConfig === 'function') {
    isGG = (getChurchConfig(false).profile.church_code === 'GG');
  }
  
  if (isGG) {
    defaultData = [
      ["SAT", "PRAYER", "AM", "J", 8],
      ["SAT", "PRAISE", "AM", "D", 9],
      ["SAT", "ACCOMPANIST", "AM", "E", 10],
      ["SAT", "SERMON", "AM", "AY", 11],
      ["SAT", "INTERPRETER", "AM", "C", 13],
      ["SAT", "A/V 1", "AM", "H", 14],
      ["SAT", "A/V 2", "AM", "I", 14],
      ["SAT", "USHER 1", "AM", "F", 15],
      ["SAT", "USHER 2", "AM", "G", 15],
      ["SAT", "GRACE", "AM", "K", 16],
      ["SAT", "PRAYER", "PM", "X", 19],
      ["SAT", "PRAISE", "PM", "D", 20],
      ["SAT", "ACCOMPANIST", "PM", "E", 21],
      ["SAT", "SERMON", "PM", "AY", 22],
      ["SAT", "INTERPRETER", "PM", "C", 23],
      ["SAT", "A/V 1", "PM", "H", 24],
      ["SAT", "A/V 2", "PM", "I", 24],
      ["SAT", "USHER 1", "PM", "F", 25],
      ["SAT", "USHER 2", "PM", "G", 25],
      ["SAT", "ENGLISH BIBLE", "PM", "AA", 26],
      ["SAT", "SISTERS", "PM", "AB", 27],
      ["SAT", "VESSEL", "PM", "AC", 28],
      ["MON", "PRAISE", "AM", "D", 31],
      ["MON", "ACCOMPANIST", "AM", "E", 32],
      ["MON", "SERMON", "AM", "AY", 33],
      ["MON", "INTERPRETER", "AM", "C", 34],
      ["MON", "A/V 1", "AM", "H", 35],
      ["MON", "A/V 2", "AM", "I", 35],
      ["MON", "USHER 1", "AM", "F", 36],
      ["MON", "USHER 2", "AM", "G", 36],
      ["TUE", "JOY FELLOWSHIP", "AM", "AS", 39],
      ["WED", "AGAPE VENUE", "AM", "AX", 42],
      ["FRI", "PRAISE", "AM", "D", 49],
      ["FRI", "ACCOMPANIST", "AM", "E", 50],
      ["FRI", "SERMON", "AM", "AY", 51],
      ["FRI", "INTERPRETER", "AM", "C", 52],
      ["FRI", "A/V 1", "AM", "H", 53],
      ["FRI", "A/V 2", "AM", "I", 53],
      ["FRI", "USHER 1", "AM", "F", 54],
      ["FRI", "USHER 2", "AM", "G", 54],
      ["SAT", "Chapel 1", "CLEANING", "AI", 60],
      ["SAT", "Chapel 2", "CLEANING", "AJ", 61],
      ["SAT", "Chapel 3", "CLEANING", "AK", 62],
      ["SAT", "Trash", "CLEANING", "AL", 63],
      ["SAT", "Bros Down", "CLEANING", "AM", 64],
      ["SAT", "Bros Down 2", "CLEANING", "AN", 64],
      ["SAT", "Sis Down", "CLEANING", "AO", 68],
      ["SAT", "Sis Down 2", "CLEANING", "AP", 68],
      ["SAT", "Bros Up", "CLEANING", "AQ", 66],
      ["SAT", "Sis Up", "CLEANING", "AR", 70],
      ["FRI", "Chapel 1", "CLEANING", "AD", 60],
      ["FRI", "Chapel 2", "CLEANING", "AE", 61],
      ["FRI", "Chapel 3", "CLEANING", "AF", 62]
    ];
  }
  
  var writeData = [headers].concat(defaultData);
  configSheet.getRange(1, 1, writeData.length, writeData[0].length).setValues(writeData);
  if (defaultData.length === 0) {
    configSheet.getRange(1, 6).setValue("NOTE: Run 'Compile Roles' from the Admin Menu to populate this tab.").setFontStyle("italic");
  }
  
  configSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground(typeof APP_CONFIG !== 'undefined' && APP_CONFIG.THEME ? APP_CONFIG.THEME.ACCENT_SUCCESS : "#d9ead3");
  configSheet.setFrozenRows(1);
  configSheet.autoResizeColumns(1, headers.length);
  configSheet.hideSheet();
  
  return configSheet;
}

function getMonthYearTitle(year, month) {
  var englishMonths = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var chineseMonths = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二"];
  
  var chineseDigits = {
    '0': '零', '1': '一', '2': '二', '3': '三', '4': '四',
    '5': '五', '6': '六', '7': '七', '8': '八', '9': '九'
  };
  
  var yearStr = String(year);
  var chineseYear = "";
  for (var i = 0; i < yearStr.length; i++) {
    chineseYear += chineseDigits[yearStr[i]];
  }
  
  var engPart = englishMonths[month - 1] + " " + year;
  var chiPart = chineseYear + "年" + chineseMonths[month - 1] + "月";
  
  return engPart + "/" + chiPart;
}

// =========================================================================
// AUTOMATED NIGHTLY AUTO-SYNC (STRATEGY 3)
// =========================================================================

/**
 * Automated Daily Month Sync (Nightly Trigger at 4:00 AM)
 * Automatically scans Q1-Q4 and re-synchronizes the current active month and next month's schedule tabs.
 * Any swaps or personnel schedule changes made in Q1-Q4 are cleanly propagated into the monthly generated schedules.
 */
function autoSyncScheduledMonths() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  
  var currentYear = now.getFullYear();
  var currentMonth = now.getMonth() + 1; // 1-indexed
  
  var nextDate = new Date(currentYear, currentMonth, 1);
  var nextYear = nextDate.getFullYear();
  var nextMonth = nextDate.getMonth() + 1; // 1-indexed
  
  var monthsToSync = [
    { year: currentYear, month: currentMonth },
    { year: nextYear, month: nextMonth }
  ];
  
  var monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var syncedResults = [];
  
  monthsToSync.forEach(function(target) {
    var targetName = monthNames[target.month - 1] + " " + target.year;
    var existingSheet = ss.getSheetByName(targetName);
    
    // Verify master Q sheet exists for this target month
    var qSheetName = "Q" + Math.ceil(target.month / 3);
    if (!ss.getSheetByName(qSheetName)) {
      console.warn("Skipping auto-sync for " + targetName + ": Master sheet " + qSheetName + " does not exist.");
      return;
    }
    
    try {
      var result = generateMonthlySheet(target.year, target.month, { preserveManualOverrides: true, inPlace: true });
      if (result && result.success) {
        syncedResults.push(targetName + " (Success)");
        console.log("Successfully auto-synced " + targetName);
      }
    } catch (e) {
      console.error("Auto-sync failed for " + targetName + ": " + e.message);
      syncedResults.push(targetName + " (Error: " + e.message + ")");
    }
  });
  
  return {
    success: true,
    syncedMonths: syncedResults
  };
}

/**
 * UI prompt to manually trigger auto-sync for the current active month and next month.
 */
function promptAutoSyncScheduledMonths() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    'Auto-Sync Active & Next Month Schedules?',
    'This will re-sync the current active month and next month\'s generated schedule tabs directly from Q1-Q4.\n\nAny personnel swaps or schedule changes made in Q1-Q4 will be reflected in the monthly schedules.\n\nDo you want to proceed?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Auto-syncing current active month and next month from Q sheets...', 'Please wait', 30);
  
  try {
    var res = autoSyncScheduledMonths();
    if (res.syncedMonths && res.syncedMonths.length > 0) {
      ui.alert(
        'Auto-Sync Complete! 🎉',
        'The following monthly schedules were successfully re-synced from Q1-Q4:\n\n• ' + res.syncedMonths.join('\n• '),
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('Auto-Sync Notice', 'No active or next month schedules were eligible to sync.', ui.ButtonSet.OK);
    }
  } catch(e) {
    ui.alert('Auto-Sync Failed', 'Error: ' + e.message, ui.ButtonSet.OK);
  }
}