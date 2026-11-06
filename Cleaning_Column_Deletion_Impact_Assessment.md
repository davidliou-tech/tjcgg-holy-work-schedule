# 📊 Impact Assessment & Implementation Plan: Cleaning Column Deletion & Grid Re-alignment

**Project:** Holy Work Schedule (HWS) — 2027 System  
**Scope:** Deletion of 3 Side Chapel Restroom Cleaning Columns (`AS`, `AT`, `AU`), Grid Shrinkage (54 Columns $\rightarrow$ 51 Columns), and Script Alignment  
**Target Script ID:** `1fdNNZB8Uvm5yOFPasUzE6MgYZ4srSLYFbYIZElgjJXEp5wsfFhbz9lvC`  
**Date:** October 3, 2026  

---

## 🙋‍♂️ ANSWER TO KEY QUESTION

> **"Do I need to manually delete these columns across Q1-Q4, or can the script handle that?"**

**No, you do NOT need to manually delete these columns.**  
An interactive column removal utility function is included in `Admin_Tools.js` (`promptRemoveCustomColumns()` / `removeCustomColumns()`, with backward-compatible alias `deleteLegacyCleaningColumns()`) accessible via **⚙️ Admin Tools > 🧰 Repair & Maintenance (Advanced) > 🧹 Remove Column(s) (Interactive)**. 

When executed, the script will:
1. Prompt the admin for column letter(s) or numbers (e.g. `AS, AT, AU` or `45-47`).
2. Inspect `Q1` Row 2 (Team) and Row 4 (Role Title) and display a confirmation preview of the affected roles.
3. Automatically delete the specified columns across `Q1`, `Q2`, `Q3`, `Q4`, and `MonthlyTemplate` in descending order.
4. Shift downstream `Coordinators` tab `q_columns` mappings and `MappingConfig` targets left by the deleted width.
5. Re-apply batch validation rules (`ensureValidationConfigSheet`) and section layout formatting (`restructureQuarterSheetsToUniversal14`).

---

## 🗺️ BEFORE & AFTER COLUMN MAPPING MATRIX (Cols AD to AY)

Below is the exact column shift mapping resulting from deleting columns 45, 46, and 47:

| Col Index (1-based) | Old Col Letter | New Col Letter | Category / Team | Specific Role / Sub-location | Status |
| :---: | :---: | :---: | :--- | :--- | :--- |
| **30 – 32** | `AD – AF` | `AD – AF` | `CLEANING TEAM` | Friday Cleaning (Chapel 1, Chapel 2, Chapel 3) | **Unchanged** |
| **33 – 34** | `AG – AH` | `AG – AH` | `CLEANING TEAM` | Saturday Window Cleaning (2 slots) | **Unchanged** |
| **35 – 37** | `AI – AK` | `AI – AK` | `CLEANING TEAM` | Saturday Chapel, MPR & Mezzanine (3 slots) | **Unchanged** |
| **38** | `AL` | `AL` | `CLEANING TEAM` | Saturday Trash (Dining Hall) | **Unchanged** |
| **39 – 40** | `AM – AN` | `AM – AN` | `CLEANING TEAM` | Saturday Brothers Restroom (Chapel Downstairs x2) | **Unchanged** |
| **41 – 42** | `AO – AP` | `AO – AP` | `CLEANING TEAM` | Saturday Sisters Restroom (Chapel Downstairs x2) | **Unchanged** |
| **43** | `AQ` | `AQ` | `CLEANING TEAM` | Saturday Brothers Restroom (Chapel Upstairs Left) | **Unchanged** |
| **44** | `AR` | `AR` | `CLEANING TEAM` | Saturday Sisters Restroom (Chapel Upstairs Right) | **Unchanged** |
| ~~45~~ | ~~`AS`~~ | — | ~~`CLEANING TEAM`~~ | ~~Sisters Restroom (Side Chapel Inside)~~ | ❌ **DELETED** |
| ~~46~~ | ~~`AT`~~ | — | ~~`CLEANING TEAM`~~ | ~~Sisters Restroom (Side Chapel Outside)~~ | ❌ **DELETED** |
| ~~47~~ | ~~`AU`~~ | — | ~~`CLEANING TEAM`~~ | ~~Brothers Restroom (Side Chapel Outside)~~ | ❌ **DELETED** |
| **45** | `AV` | **`AS`** | `JOY FELLOWSHIP` | CHURCH | ⬅️ Shifted -3 |
| **46** | `AW` | **`AT`** | `WEDNESDAY STUDY` | Leader / Location / Bible Passage | ⬅️ Shifted -3 |
| **47** | `AX` | **`AU`** | `WEDNESDAY STUDY` | 南區 (地點) (South Location) | ⬅️ Shifted -3 |
| **48** | `AY` | **`AV`** | `WEDNESDAY STUDY` | 西區 (地點) (West Location) | ⬅️ Shifted -3 |
| **49** | `AZ` | **`AW`** | `WEDNESDAY STUDY` | 北區 (地點) (North Location) | ⬅️ Shifted -3 |
| **50** | `BA` | **`AX`** | `WEDNESDAY STUDY` | AGAPE (Location) | ⬅️ Shifted -3 |
| **51** | `BB` | **`AY`** | `SPEAKER FULL NAME` | Dynamic `XLOOKUP` Speaker Formula Bridge | ⬅️ Shifted -3 |

---

## 🔍 SCRIPT IMPACT ANALYSIS & REQUIRED CHANGES

| Script File | Function / Scope | Required Modification |
| :--- | :--- | :--- |
| **`Admin_Tools.js`** | Universal Grid Layout & Formatter | • Update `NUM_GRID_COLS` constant from `54` to `51`.<br>• Update grid scanning boundaries to Cols 2–50 (49 cols).<br>• Update `TUESDAYS` active column to `45` (`AS` for Joy Fellowship).<br>• Update `WEDNESDAYS` active columns to `[46, 47, 48, 49, 50]` (`AT`–`AX`).<br>• Update Speaker `XLOOKUP` target column to `51` (`AY`) and update variable names/comments from `BB` to `AY`.<br>• Update full grid range references from `A6:BB94` to `A6:AY94`.<br>• Update SWR IMPORTRANGE formulas to target Column `AY` (`AY6:AY19`, `AY51:AY64`, `AY66:AY79`, `AY81:AY94`).<br>• Add automated column deletion utility `deleteLegacyCleaningColumns()`. |
| **`GenerateMonthly.js`** | Monthly Schedule PDF Generator & Default Mapping | • Update `createDefaultConfigSheet()` default mapping table:<br>  - `TUE JOY FELLOWSHIP AM`: Column letter updated to `AS`.<br>  - `WED AGAPE VENUE AM`: Column letter updated to `AX`.<br>  - `SERMON` (SAT AM/PM, MON AM, FRI AM): Master Column letter updated to `AY`. |
| **`Validation.js`** | Batch Data Validation Dropdowns | • Update `ensureValidationConfigSheet()` default configuration:<br>  - `AS` $\rightarrow$ `Joy Fellowship`<br>  - `AT` $\rightarrow$ `Wednesday Evening Bible Study`<br>  - `AU`–`AX` $\rightarrow$ `Wednesday Locations` |
| **`Service_Reminders.js` & `service_reminders.gs`** | Automated Email Reminders | • Update fallback dictionary column indices (`COL.TUE_JOY` = 44, `COL.WED_BIBLE` = 45, `COL.WED_LOCATIONS` = `[46, 47, 48, 49]`, `COL.SPEAKER_FULL` = 50).<br>• Verify dynamic column loading (`loadColMappingsFromConfig()`) reads updated mappings from `MappingConfig`. |
| **`SyncSWR.js`** | Sermon Topic & Speaker Sync | • Update inline comments and documentation references from Column `BB` to Column `AY`. |
| **`Sidebar.html`** | Interactive Schedule Assistant | • Confirm Speaker Full Name column index check in `renderWorkers()` is set to `51` (`AY`). |
| **Documentation Files** | System Guides & Logs | • Update `Cleaning_Column_Deletion_Impact_Assessment.md`, `README.md`, `SYSTEM_FEATURES.md`, `RELEASE_NOTES.md`, and `Admin_Guide.md` to reflect the 51-column standard grid. |

---

## 🚀 IMPLEMENTATION PLAN

Upon your approval, I will execute the plan in 4 phases:

### Phase 1: Local Code Updates
* Modify `Admin_Tools.js`, `GenerateMonthly.js`, `Validation.js`, `Service_Reminders.js`, `service_reminders.gs`, `SyncSWR.js`, and `Sidebar.html` with all column adjustments.

### Phase 2: Create Automated Deletion Script
* Add `deleteLegacyCleaningColumns()` in `Admin_Tools.js` and add menu option under **Admin Tools > Maintenance > Remove 3 Legacy Cleaning Columns (AS, AT, AU)**.

### Phase 3: Deployment & Documentation
* Update documentation files (`README.md`, `SYSTEM_FEATURES.md`, `RELEASE_NOTES.md`, `Admin_Guide.md`).
* Deploy code changes to Google Apps Script (`clasp push`) targeting Script ID `1fdNNZB8Uvm5yOFPasUzE6MgYZ4srSLYFbYIZElgjJXEp5wsfFhbz9lvC`.

### Phase 4: Execution & Verification
* Run automated column deletion and layout standardization on the target Google Sheet.
* Perform verification checks across Monthly PDF generation, Data Validation dropdowns, and Sidebar Assistant conflict checking.
