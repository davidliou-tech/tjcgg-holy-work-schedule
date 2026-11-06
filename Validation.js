/**
 * VALIDATION.GS
 * Programmatically applies all dropdown data validation rules
 * to Q1, Q2, Q3, and Q4 sheets based on the ValidationConfig tab.
 *
 * HOW TO USE:
 *   From the menu: Schedule Assistant > Data Validation > Apply Validation (All Quarters)
 *   Or run applyAllValidation() directly from the Apps Script editor.
 */

function ensureValidationConfigSheet(ss) {
  var sheet = ss.getSheetByName("ValidationConfig");
  if (sheet) return sheet;
  
  var isGG = true;
  if (typeof getChurchConfig === 'function') {
    var churchConfig = getChurchConfig(false);
    isGG = (churchConfig.profile.church_code === 'GG');
  }
  
  sheet = ss.insertSheet("ValidationConfig");
    var headers = ["Master Column (Letter)", "Role / Description", "Teams Sheet Name(s)"];
    var defaultData = [];
    
    if (isGG) {
      defaultData = [
        ["C", "Interpreter", "Interpreter (E-to-C), Interpreter (C-to-E), Interpreter, 翻譯"],
        ["D", "Praise / Hymn Leader", "Hymn Leader, 詩頌"],
        ["E", "Accompanist / Pianist", "Pianist, 司琴"],
        ["F", "Usher 1", "Usher, 接待"],
        ["G", "Usher 2", "Usher, 接待"],
        ["H", "AV 1", "AV, 視聽"],
        ["I", "AV 2", "AV, 視聽"],
        ["J", "Morning Prayer", "Prayer, 禱告"],
        ["K", "Grace / Lunch Prayer", "Grace, 謝飯"],
        ["AA", "EBF", "English Bible Fellowship"],
        ["AB", "Sisters' Fellowship", "Sisters' Fellowship"],
        ["AC", "Vessel Fellowship", "Vessel Fellowship"],
        ["AD", "Friday Chapel Cleaning 1", "Friday Chapel Cleaning"],
        ["AE", "Friday Chapel Cleaning 2", "Friday Chapel Cleaning"],
        ["AF", "Friday Chapel Cleaning 3", "Friday Chapel Cleaning"],
        ["AS", "Joy Fellowship", "Joy Fellowship"],
        ["AT", "Wednesday Evening Bible Study", "Wednesday Study"],
        ["AU", "Wednesday South Location", "Wednesday Location"],
        ["AV", "Wednesday West Location", "Wednesday Location"],
        ["AW", "Wednesday North Location", "Wednesday Location"],
        ["AX", "Wednesday Agape Location", "Wednesday Location"]
      ];
    }
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground(typeof APP_CONFIG !== 'undefined' && APP_CONFIG.THEME ? APP_CONFIG.THEME.ACCENT_SUCCESS : "#d9ead3");
    if (defaultData.length > 0) {
      sheet.getRange(2, 1, defaultData.length, defaultData[0].length).setValues(defaultData);
    } else {
      sheet.getRange(1, 4).setValue("NOTE: Run 'Compile Roles' from the Admin Menu to populate this tab.").setFontStyle("italic");
    }
    sheet.autoResizeColumns(1, headers.length);
    sheet.hideSheet();
  return sheet;
}

function colLetterToIdx(letter) {
  var str = String(letter || "").toUpperCase().trim();
  var column = 0, length = str.length;
  for (var i = 0; i < length; i++) {
    column += (str.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

function applyAllValidation() {
  var sheets = ["Q1","Q2","Q3","Q4"];
  var totalApplied = 0;
  var allErrors = [];
  
  for (var i = 0; i < sheets.length; i++) {
    try {
      var res = applyDataValidation(sheets[i]);
      totalApplied += res.applied;
      if (res.errors && res.errors.length > 0) {
        allErrors = allErrors.concat(res.errors);
      }
    } catch(e) { 
      allErrors.push(sheets[i] + " Sheet Error: " + e.message); 
    }
  }
  
  if (allErrors.length > 0) {
    var msg = "Applied " + totalApplied + " dropdown validations with warnings/errors:\n";
    msg += "\n\nFailed Rules (" + allErrors.length + "):\n" + allErrors.slice(0, 10).join("\n");
    if (allErrors.length > 10) {
      msg += "\n...and " + (allErrors.length - 10) + " more. See Apps Script Execution Log for full details.";
    }
    SpreadsheetApp.getUi().alert("Validation Completed with Warnings", msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert("Success! 🎉", "Successfully applied " + totalApplied + " dropdown validation rules across active service slots in Q1, Q2, Q3, and Q4.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function applyValidationForCurrentSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var name = sheet.getName();
  if (!["Q1","Q2","Q3","Q4"].includes(name)) {
    SpreadsheetApp.getUi().alert("Please select a quarter sheet (Q1, Q2, Q3, or Q4) first.");
    return;
  }
  var res = applyDataValidation(name);
  if (res.errors && res.errors.length > 0) {
    SpreadsheetApp.getUi().alert("Completed with Warnings", "Applied " + res.applied + " rules to " + name + ", but encountered " + res.errors.length + " errors.\n\nFirst errors:\n" + res.errors.slice(0, 5).join("\n"), SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert("Success! 🎉", "Data validation applied to " + name + " successfully (" + res.applied + " active dropdowns).", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function applyDataValidation(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet not found: " + sheetName);
  
  var appliedCount = 0;
  var ruleErrors = [];
  
  // 1. Fetch master teams object
  var teamsObj = getTeamsAndWorkers(true);
  if (teamsObj.error) throw new Error(teamsObj.error);
  var teams = teamsObj.teams; // format: { "Interpreter (E-to-C)": [...], ... }
  
  // 2. Read ValidationConfig sheet
  var valConfigSheet = ensureValidationConfigSheet(ss);
  var configData = valConfigSheet.getDataRange().getValues();
  
  var colValidationRules = {}; // colIndex -> DataValidation rule
  
  for (var i = 1; i < configData.length; i++) {
    var colLetter = String(configData[i][0]).trim().toUpperCase();
    var targetTeamsStr = String(configData[i][2]).trim();
    
    if (!colLetter || !targetTeamsStr) continue;
    
    var colIndex = colLetterToIdx(colLetter);
    if (colIndex < 2) continue;
    
    var teamNames = targetTeamsStr.split(",").map(function(t) { return t.trim(); });
    var combinedWorkers = [];
    var seenWorkers = {};
    
    for (var t = 0; t < teamNames.length; t++) {
      var tName = teamNames[t];
      if (!tName) continue;
      
      // Look up in Teams tab (exact match or fuzzy match)
      for (var actualTeam in teams) {
        var aClean = actualTeam.toLowerCase();
        var tClean = tName.toLowerCase();
        
        if (aClean === tClean || aClean.indexOf(tClean) !== -1 || tClean.indexOf(aClean) !== -1) {
          var wList = teams[actualTeam] || [];
          for (var w = 0; w < wList.length; w++) {
            var workerName = String(wList[w]).trim();
            if (workerName && !seenWorkers[workerName]) {
              seenWorkers[workerName] = true;
              combinedWorkers.push(workerName);
            }
          }
        }
      }
    }
    
    combinedWorkers.sort();
    if (combinedWorkers.length > 500) combinedWorkers = combinedWorkers.slice(0, 500);
    
    if (combinedWorkers.length > 0) {
      colValidationRules[colIndex] = SpreadsheetApp.newDataValidation()
        .requireValueInList(combinedWorkers, true)
        .setAllowInvalid(false)
        .build();
    }
  }
  
  // 3. Define Architecture blocks
  var blocks = [];
  if (typeof getChurchConfig === 'function') {
    var churchConfig = getChurchConfig(true);
    var layout = churchConfig.serviceLayout || [];
    for (var i = 0; i < layout.length; i++) {
      if (layout[i].enabled === 'Y') {
        blocks.push({ start: layout[i].start_row, end: layout[i].end_row });
      }
    }
  }
  
  if (blocks.length === 0) {
    blocks = [
      { start: 6, end: 19 },   // MONDAY
      { start: 21, end: 34 },  // TUESDAY
      { start: 36, end: 49 },  // WEDNESDAY
      { start: 51, end: 64 },  // FRIDAY
      { start: 66, end: 79 },  // SATURDAY AM
      { start: 81, end: 94 }   // SATURDAY PM
    ];
  }
  
  // 4. Batch apply per block
  var maxCols = sheet.getMaxColumns();
  var numCols = maxCols - 1; // From Col 2 (B) to maxCols
  
  for (var b = 0; b < blocks.length; b++) {
    var block = blocks[b];
    var numRows = block.end - block.start + 1;
    
    try {
      var range = sheet.getRange(block.start, 2, numRows, numCols);
      var values = range.getValues();
      var validations = [];
      
      for (var r = 0; r < numRows; r++) {
        var valRow = [];
        for (var c = 0; c < numCols; c++) {
          var colIndex = c + 2; 
          var cellVal = String(values[r][c]).trim().toUpperCase();
          var rule = colValidationRules[colIndex];
          
          if (!rule || cellVal === "X") {
            valRow.push(null);
          } else {
            valRow.push(rule);
            appliedCount++;
          }
        }
        validations.push(valRow);
      }
      
      range.setDataValidations(validations);
    } catch(e) {
      ruleErrors.push(sheetName + " Block " + block.start + "-" + block.end + ": " + e.message);
    }
  }

  return { sheet: sheetName, applied: appliedCount, errors: ruleErrors };
}

function applyValidationForESSC() {
  var activeSs = SpreadsheetApp.getActiveSpreadsheet();
  var targetSs = (typeof getTargetESSCSpreadsheet === 'function') ? getTargetESSCSpreadsheet() : activeSs;
  if (!targetSs) {
    SpreadsheetApp.getUi().alert("Error", "No connected ESSC workbook found. Please run 'Connect Standalone ESSC Workbook...' first.", SpreadsheetApp.getUi().ButtonSet.OK);
    return { applied: 0, errors: ["No target ESSC spreadsheet found"] };
  }
  var ss = targetSs;
  var grids = ["Spring_ESSC_Grid", "Fall_ESSC_Grid"];
  var totalApplied = 0;
  var errors = [];

  // 1. Fetch live worker rosters from active Master HWS or target ESSC
  var teams = {};
  var teamsSheet = activeSs.getSheetByName("Teams") || activeSs.getSheetByName("Rosters") || targetSs.getSheetByName("Rosters");
  if (teamsSheet) {
    var maxR = teamsSheet.getLastRow();
    var maxC = teamsSheet.getLastColumn();
    if (maxR > 1 && maxC > 0) {
      var rData = teamsSheet.getRange(1, 1, maxR, maxC).getValues();
      for (var col = 0; col < maxC; col++) {
        var tName = '', sRow = 1;
        var row0Val = String(rData[0][col] || "").trim();
        var row1Val = (maxR > 1) ? String(rData[1][col] || "").trim() : "";
        if (row1Val && (!row0Val || !isNaN(Number(row0Val)) || row1Val.match(/[a-zA-Z]/))) {
          tName = row1Val;
          sRow = 2;
        } else {
          tName = row0Val;
          sRow = 1;
        }
        if (!tName) continue;
        teams[tName] = [];
        for (var row = sRow; row < maxR; row++) {
          var wName = String(rData[row][col] || "").trim();
          if (wName && teams[tName].indexOf(wName) === -1) {
            teams[tName].push(wName);
          }
        }
      }
    }
  }

  // Sync Rosters sheet to targetSs if external
  if (activeSs.getId() !== targetSs.getId() && teamsSheet) {
    try {
      var targetRosters = targetSs.getSheetByName("Rosters");
      if (!targetRosters) targetRosters = targetSs.insertSheet("Rosters");
      var sData = teamsSheet.getDataRange().getValues();
      targetRosters.clear();
      targetRosters.getRange(1, 1, sData.length, sData[0].length).setValues(sData);
    } catch(rErr) {
      console.warn("Could not sync Rosters tab to target ESSC: " + rErr.message);
    }
  }

  if (Object.keys(teams).length === 0) {
    try {
      var teamsObj = getTeamsAndWorkers(true);
      if (teamsObj && teamsObj.teams) teams = teamsObj.teams;
    } catch(e) {}
  }

  if (Object.keys(teams).length === 0) {
    SpreadsheetApp.getUi().alert("Error", "No worker rosters found in Master HWS 'Teams' or 'Rosters'.", SpreadsheetApp.getUi().ButtonSet.OK);
    return { applied: 0, errors: ["No workers found"] };
  }

  // 3. Auto-verify and recalculate Spring & Fall dates from MasterConfig
  try {
    var springStart = (typeof getConfigValue === 'function') ? getConfigValue(ss, 'SPRING_WEDNESDAY_START_DATE') : '';
    if (springStart && typeof recalculateGridDates === 'function') {
      var spGrid = ss.getSheetByName("Spring_ESSC_Grid");
      if (spGrid && spGrid.getLastRow() >= 2) {
        var dVal = spGrid.getRange(2, 1).getValue();
        var dStr = (typeof formatDateSafe === 'function') ? formatDateSafe(dVal, ss) : String(dVal || "").trim();
        if (dStr !== springStart) {
          recalculateGridDates(ss, 'Spring', springStart);
        }
      }
    }
    var fallStart = (typeof getConfigValue === 'function') ? getConfigValue(ss, 'FALL_WEDNESDAY_START_DATE') : '';
    if (fallStart && typeof recalculateGridDates === 'function') {
      var flGrid = ss.getSheetByName("Fall_ESSC_Grid");
      if (flGrid && flGrid.getLastRow() >= 2) {
        var fdVal = flGrid.getRange(2, 1).getValue();
        var fdStr = (typeof formatDateSafe === 'function') ? formatDateSafe(fdVal, ss) : String(fdVal || "").trim();
        if (fdStr !== fallStart) {
          recalculateGridDates(ss, 'Fall', fallStart);
        }
      }
    }
  } catch(dErr) {
    console.warn("Date auto-check skipped: " + dErr.message);
  }

  // 4. Inject Dynamic Date & Speaker Formulas (XLOOKUP) so MasterConfig edits update real-time
  var formulasInjected = 0;
  try {
    if (typeof injectDynamicGridFormulas === 'function') {
      formulasInjected = injectDynamicGridFormulas(ss);
    }
  } catch(fErr) {
    console.warn("Formula injection skipped: " + fErr.message);
  }

  // 4b. Also resolve any raw Speaker placeholders ({VISITING_MINISTER}, etc.) in-place
  var placeholdersResolved = 0;
  try {
    if (typeof resolveGridSpeakerPlaceholders === 'function') {
      placeholdersResolved = resolveGridSpeakerPlaceholders(ss);
    }
  } catch(pErr) {
    console.warn("Placeholder resolution skipped: " + pErr.message);
  }

  // 5. Read MasterConfig minister names & tokens to whitelist for Column E (Speakers)
  var ministerNames = [];
  var ministerTokens = [
    "{VISITING_MINISTER}", "{RESIDENT_MINISTER}",
    "{SPRING_VISITING_MINISTER}", "{FALL_VISITING_MINISTER}",
    "{LAYMINISTER_1}", "{LAYMINISTER_2}", "{LAYMINISTER_3}", "{LAYMINISTER_4}",
    "{CHAIRPERSON}"
  ];
  var mcSheet = ss.getSheetByName("MasterConfig");
  if (mcSheet) {
    var mcData = mcSheet.getDataRange().getValues();
    for (var m = 1; m < mcData.length; m++) {
      var mKey = String(mcData[m][0] || "").trim();
      var mVal = String(mcData[m][1] || "").trim();
      if (!mVal) continue;
      if (/(?:VISITING_MINISTER|RESIDENT_MINISTER|LAYMINISTER|CHAIRPERSON)/i.test(mKey)) {
        if (ministerNames.indexOf(mVal) === -1) ministerNames.push(mVal);
        // Also add bilingual components without newline if applicable
        if (mVal.indexOf("\n") !== -1) {
          var lines = mVal.split("\n").map(function(s){ return s.trim(); }).filter(Boolean);
          lines.forEach(function(l){ if (ministerNames.indexOf(l) === -1) ministerNames.push(l); });
          var slashClean = lines.join(" / ");
          if (ministerNames.indexOf(slashClean) === -1) ministerNames.push(slashClean);
        }
      }
    }
  }

  // 6. Read ValidationConfig
  var valConfigSheet = ss.getSheetByName("ValidationConfig");
  var configData = valConfigSheet ? valConfigSheet.getDataRange().getValues() : [];
  
  var sheetRules = {};
  for (var g = 0; g < grids.length; g++) {
    sheetRules[grids[g]] = {};
  }
  
  for (var i = 1; i < configData.length; i++) {
    var colLetterStr = String(configData[i][0] || "").trim();
    var targetTeamsStr = String(configData[i][2] || "").trim();
    if (!colLetterStr || !targetTeamsStr) continue;
    
    var targetSheet = "";
    var colLetter = "";
    if (colLetterStr.indexOf(":") !== -1) {
      var parts = colLetterStr.split(":");
      targetSheet = parts[0].trim();
      colLetter = parts[1].trim().toUpperCase();
    } else {
      colLetter = colLetterStr.toUpperCase();
    }
    
    var colIndex = colLetterToIdx(colLetter);
    if (!colIndex) continue;
    
    var teamNames = targetTeamsStr.split(",").map(function(t) { return t.trim(); });
    var combinedWorkers = [];
    var seenWorkers = {};
    
    for (var t = 0; t < teamNames.length; t++) {
      var tName = teamNames[t];
      if (!tName) continue;
      for (var actualTeam in teams) {
        var aClean = actualTeam.toLowerCase();
        var tClean = tName.toLowerCase();
        if (aClean === tClean || aClean.indexOf(tClean) !== -1 || tClean.indexOf(aClean) !== -1) {
          var wList = teams[actualTeam] || [];
          for (var w = 0; w < wList.length; w++) {
            var workerName = String(wList[w]).trim();
            if (workerName && !seenWorkers[workerName]) {
              seenWorkers[workerName] = true;
              combinedWorkers.push(workerName);
            }
          }
        }
      }
    }
    
    // For Speaker column (Col 5), always whitelist visiting ministers and placeholder tokens
    if (colIndex === 5) {
      ministerNames.forEach(function(mn) {
        if (mn && !seenWorkers[mn]) { seenWorkers[mn] = true; combinedWorkers.push(mn); }
      });
      ministerTokens.forEach(function(mt) {
        if (mt && !seenWorkers[mt]) { seenWorkers[mt] = true; combinedWorkers.push(mt); }
      });
    }

    if (combinedWorkers.length > 0) {
      combinedWorkers.sort();
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(combinedWorkers, true)
        .setAllowInvalid(true)
        .build();
        
      if (targetSheet && sheetRules[targetSheet]) {
        sheetRules[targetSheet][colIndex] = rule;
      } else {
        for (var g = 0; g < grids.length; g++) {
          sheetRules[grids[g]][colIndex] = rule;
        }
      }
    }
  }

  // Fallback default role mapping if ValidationConfig is missing specific columns
  var defaultColTeams = {
    5: ["Speakers", "講員"],
    6: ["Interpreters", "Interpreter", "翻譯"],
    7: ["Hymn Leaders", "Hymn Leader", "詩頌"],
    8: ["Pianists", "Pianist", "司琴"],
    9: ["AV", "視聽"],
    10: ["AV", "視聽"],
    11: ["Ushers", "Usher", "接待"],
    12: ["Ushers", "Usher", "接待"],
    13: ["Cleaning", "Fellowships", "區域/契頁"],
    14: ["Grace", "謝飯"],
    15: ["Grace", "謝飯"],
    16: ["Choir", "詩班"]
  };

  for (var g = 0; g < grids.length; g++) {
    var gName = grids[g];
    var gSheet = ss.getSheetByName(gName);
    if (!gSheet) continue;
    
    var lastRow = Math.max(gSheet.getLastRow(), 42);
    var rulesForGrid = sheetRules[gName] || {};
    
    // Read session titles (Col 4) and speaker cells (Col 5) to clean up non-duty rows
    var sessionRange = gSheet.getRange(2, 4, lastRow - 1, 2);
    var sessionVals = sessionRange.getValues();

    for (var colIdx = 5; colIdx <= 16; colIdx++) {
      var rule = rulesForGrid[colIdx];
      if (!rule && defaultColTeams[colIdx]) {
        var cTeams = defaultColTeams[colIdx];
        var cWorkers = [];
        var cSeen = {};
        for (var ct = 0; ct < cTeams.length; ct++) {
          for (var act in teams) {
            if (act.toLowerCase().indexOf(cTeams[ct].toLowerCase()) !== -1) {
              (teams[act] || []).forEach(function(w) {
                if (w && !cSeen[w]) { cSeen[w] = true; cWorkers.push(w); }
              });
            }
          }
        }
        if (colIdx === 5) {
          ministerNames.forEach(function(mn) {
            if (mn && !cSeen[mn]) { cSeen[mn] = true; cWorkers.push(mn); }
          });
          ministerTokens.forEach(function(mt) {
            if (mt && !cSeen[mt]) { cSeen[mt] = true; cWorkers.push(mt); }
          });
        }
        if (cWorkers.length > 0) {
          cWorkers.sort();
          rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(cWorkers, true)
            .setAllowInvalid(true)
            .build();
        }
      }
      
      if (rule) {
        gSheet.getRange(2, colIdx, lastRow - 1, 1).setDataValidation(rule);
        totalApplied += (lastRow - 1);
      }
    }

    // Clean up empty Speaker cells on meal / hymn rows so no unnecessary dropdown chip appears
    for (var rIdx = 0; rIdx < sessionVals.length; rIdx++) {
      var title = String(sessionVals[rIdx][0] || "").toLowerCase();
      var speakerCellVal = String(sessionVals[rIdx][1] || "").trim();
      if (!speakerCellVal && (title.indexOf("dinner") !== -1 || title.indexOf("lunch") !== -1 || (title.indexOf("hymn") !== -1 && title.indexOf("sermon") === -1))) {
        gSheet.getRange(rIdx + 2, 5).clearDataValidations();
      }
    }
  }

  var msg = "Successfully pushed " + totalApplied + " dropdown validation rules to '" + targetSs.getName() + "'.";
  if (placeholdersResolved > 0) {
    msg += "\n\n✨ Resolved " + placeholdersResolved + " speaker placeholders to official minister names.";
  }
  SpreadsheetApp.getUi().alert("Success! 🎉", msg, SpreadsheetApp.getUi().ButtonSet.OK);
  return { applied: totalApplied, errors: errors, placeholdersResolved: placeholdersResolved };
}


