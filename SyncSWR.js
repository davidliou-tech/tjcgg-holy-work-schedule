/**
 * The actual sync function that runs daily or when triggered manually.
 * It reads the URL from PropertiesService instead of hardcoded config.
 */
const SWR_CONFIG = { TARGET_LOCATION: 'Garden Grove', EXCLUDE_LANGUAGE_SUBSTRING: 'english' };

function syncSWRSpeakers() {
  var url = PropertiesService.getScriptProperties().getProperty('SWR_SPEAKER_SHEET_URL');
  
  if (!url) {
    try {
      // Try to alert the user if running manually
      SpreadsheetApp.getUi().alert("Configuration Missing", "Please run 'Setup Daily Automatic Sync (1 AM)' first to set the SWR Speaker Sheet URL.", SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      // If getUi() fails (e.g. running on a time-based trigger), log the error quietly
      console.error("SWR Sync Failed: No URL configured in Script Properties.");
    }
    return;
  }
  
  try {
    return executeWithLockAndRetry_('syncSWRSpeakers', function() {
      var sourceSs = SpreadsheetApp.openByUrl(url);
      var localSs = SpreadsheetApp.getActiveSpreadsheet();
      var tz = localSs.getSpreadsheetTimeZone() || APP_CONFIG.TIMEZONE || 'America/Los_Angeles';
      
      // Dictionary to hold all SWR assignments mapped by date: swrAssignments["yyyy-MM-dd"] = { AM: "Name", PM: "Name" }
      var swrAssignments = {};
      
      // 1. EXTRACT DATA FROM SWR SPREADSHEET
      CONFIG.SWR_SOURCE_QUARTERS.forEach(function(q) {
         var sourceSheet = sourceSs.getSheetByName(q);
         if (!sourceSheet) return;
         
         var data = sourceSheet.getDataRange().getValues();
         var currentDateRow = null;
         
         for (var r = 0; r < data.length; r++) {
             var row = data[r];
             
             // Detect Date Rows (Check if Col C or Col D is a Date object or Date string)
             var isDateRow = false;
             for (var colIdx = 2; colIdx <= 3; colIdx++) {
                 if (row[colIdx] instanceof Date) {
                     isDateRow = true; break;
                 } else if (typeof row[colIdx] === 'string' && row[colIdx].indexOf('/') !== -1) {
                     var d = null;
                     try {
                         d = Utilities.parseDate(row[colIdx].trim(), tz, "M/d/yyyy");
                     } catch (parseErr) {
                         var fallbackD = new Date(row[colIdx].trim());
                         if (!isNaN(fallbackD.getTime())) d = fallbackD;
                     }
                     if (d && !isNaN(d.getTime())) {
                         isDateRow = true; break;
                     }
                 }
             }
             
             if (isDateRow) {
                 currentDateRow = row;
                 continue;
             }
             
             if (!currentDateRow) continue; // Skip headers until we find the first date row
             
             // Check if this row is for target location (and explicitly ignore excluded language to prevent conflicts)
             var location = String(row[0]).trim();
             if (location.indexOf(SWR_CONFIG.TARGET_LOCATION) !== -1 && location.toLowerCase().indexOf(SWR_CONFIG.EXCLUDE_LANGUAGE_SUBSTRING) === -1) {
                 var roleName = String(row[1]).trim(); // e.g. "AM", "PM", "FPM"
                 if (!roleName.match(/am|pm/i)) {
                     Logger.log('Sync Warning: Skipped unmatched role cell "' + roleName + '" at row ' + (r + 1));
                     continue;
                 }
                 
                 // Scan across the row for assignments (starting at Column C / index 2)
                 for (var c = 2; c < row.length; c++) {
                     var speaker = String(row[c]).trim();
                     
                     // Failsafe: If the SWR sheet is messy and the cell accidentally contains the role prefix 
                     // (e.g. "PMAlan Wang" instead of "Alan Wang"), automatically strip it out!
                     if (speaker.toLowerCase().startsWith(roleName.toLowerCase()) && speaker.length > roleName.length) {
                         var charAfterPrefix = speaker.charAt(roleName.length);
                         if (!/[a-z]/i.test(charAfterPrefix)) {
                             speaker = speaker.substring(roleName.length).trim();
                         }
                     }
                     
                     var upperSpeaker = speaker.toUpperCase();
                     var isIgnored = (upperSpeaker === "S/A" || upperSpeaker === "SA" || upperSpeaker === "N/A" || upperSpeaker === "NA" || upperSpeaker === "-");
                     
                     if (speaker && !isIgnored && speaker.toLowerCase() !== roleName.toLowerCase()) {
                         var dateVal = currentDateRow[c];
                         if (dateVal) {
                             var d = null;
                             if (dateVal instanceof Date) {
                                 d = dateVal;
                             } else {
                                 var rawDateStr = String(dateVal).trim();
                                 try {
                                     d = Utilities.parseDate(rawDateStr, tz, "M/d/yyyy");
                                 } catch (parseErr) {
                                     var fallbackDate = new Date(rawDateStr);
                                     if (!isNaN(fallbackDate.getTime())) d = fallbackDate;
                                 }
                             }
                             if (d && !isNaN(d.getTime())) {
                                 var dateKey = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
                                 if (!swrAssignments[dateKey]) swrAssignments[dateKey] = {};
                                 
                                 swrAssignments[dateKey][roleName] = speaker;
                             }
                         }
                     }
                 }
             }
         }
      });
      
      // 2. INJECT DATA INTO LOCAL SPREADSHEET
      var updatedCount = 0;
      var logEntries = [];
      
      CONFIG.QUARTERS.forEach(function(q) {
         var localSheet = localSs.getSheetByName(q);
         if (!localSheet) return;
         
         var lastRow = localSheet.getLastRow();
         if (lastRow < CONFIG.START_ROW) return;
         
         // Read Column A (Date/Headers) and Column B (Speaker)
         var range = localSheet.getRange(CONFIG.START_ROW, 1, lastRow - CONFIG.START_ROW + 1, 2); 
         var values = range.getValues();
         var changed = false;
         var currentHeader = "";
         
         for (var r = 0; r < values.length; r++) {
             var dateVal = values[r][0];
             
             // Track section headers (Look at both Col A and Col B in case they are not merged)
             if (typeof dateVal === 'string' && dateVal.trim() !== "" && dateVal.indexOf('/') === -1) {
                 currentHeader = dateVal.trim().toLowerCase();
                 var colB = values[r][1];
                 if (typeof colB === 'string' && colB.trim() !== "") {
                     currentHeader += " " + colB.trim().toLowerCase();
                 }
             }
             
             // Robustly detect the Date (handles both Date objects and typed Strings)
             var localDateObj = null;
             if (dateVal instanceof Date) {
                 localDateObj = dateVal;
             } else if (typeof dateVal === 'string' && dateVal.indexOf('/') !== -1) {
                 try {
                     localDateObj = Utilities.parseDate(dateVal.trim(), tz, "M/d/yyyy");
                 } catch(pe) {
                     var parsed = new Date(dateVal.trim());
                     if (!isNaN(parsed.getTime())) localDateObj = parsed;
                 }
             } else if (typeof dateVal === 'string' && dateVal.indexOf('-') !== -1) {
                 try {
                     localDateObj = Utilities.parseDate(dateVal.trim(), tz, "yyyy-MM-dd");
                 } catch(pe) {
                     var parsed = new Date(dateVal.trim());
                     if (!isNaN(parsed.getTime())) localDateObj = parsed;
                 }
             }
             
             if (localDateObj) {
                 var dateKey = Utilities.formatDate(localDateObj, tz, 'yyyy-MM-dd');
                 
                 // If SWR assigned speakers for this date
                 if (swrAssignments[dateKey]) {
                     var assignmentsObj = swrAssignments[dateKey];
                     var speakerStr = "";
                     
                     // Determine the correct speaker based on the day of the week and section header
                     var dayStr = Utilities.formatDate(localDateObj, tz, 'EEE'); // e.g., 'Sat', 'Mon'
                     
                     if (dayStr === 'Sat') { 
                         // On Saturdays, strictly match AM or PM to the section header
                         if (currentHeader.indexOf("am") !== -1 && assignmentsObj["AM"]) {
                             speakerStr = assignmentsObj["AM"];
                         } else if (currentHeader.indexOf("pm") !== -1 && assignmentsObj["PM"]) {
                             speakerStr = assignmentsObj["PM"];
                         }
                     } else {
                         // On Weekdays, grab whatever role is assigned
                         for (var swrRole in assignmentsObj) {
                             speakerStr = assignmentsObj[swrRole];
                             break; // Usually only 1 shift on weekdays
                         }
                     }
                     
                     // Only overwrite if it is different
                     if (speakerStr && values[r][1] !== speakerStr) {
                         var oldSpeaker = values[r][1] || "(blank)";
                         logEntries.push([new Date(), q, dateKey, (currentHeader || "Weekday"), oldSpeaker, speakerStr]);
                         values[r][1] = speakerStr;
                         changed = true;
                         updatedCount++;
                     }
                 }
             }
         }
         
         if (changed) {
             try {
                 range.setValues(values);
             } finally {
                 bumpQuarterVersion(q);
             }
         }
      });
      
      // 3. WRITE TO SYNC LOG SHEET (Batch Operation)
      if (logEntries.length > 0) {
         var logSheet = localSs.getSheetByName(CONFIG.SYNC_LOG_SHEET);
         
         // Create the sheet if it doesn't exist yet
         if (!logSheet) {
             logSheet = localSs.insertSheet(CONFIG.SYNC_LOG_SHEET);
             logSheet.appendRow(["Timestamp", "Quarter", "Date", "Section", "Old Name", "New Name"]);
             logSheet.getRange("A1:F1").setFontWeight("bold");
             logSheet.setFrozenRows(1);
             logSheet.hideSheet();
         }
         
         // Write all changes in one single rapid API call
         logSheet.getRange(logSheet.getLastRow() + 1, 1, logEntries.length, logEntries[0].length).setValues(logEntries);
      }
      
      // Only show an alert if a user is actively clicking the manual sync button
      try {
        SpreadsheetApp.getUi().alert("SWR Sync Complete", "Successfully synced " + updatedCount + " assignments from the SWR Speaker sheet into your schedule.", SpreadsheetApp.getUi().ButtonSet.OK);
      } catch(e) {
        // Ignore UI exception when running in background trigger
      }
      return { success: true, updatedCount: updatedCount };
    });
  } catch(e) {
    try {
      SpreadsheetApp.getUi().alert("Sync Error", "Could not sync SWR Speakers. Error: " + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch(uiError) {
      // Ignore UI exception when running in background trigger
    }
    console.error("SWR Sync Failed: " + e.message);
  }
}