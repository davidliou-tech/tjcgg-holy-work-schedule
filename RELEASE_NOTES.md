# Release Notes — Holy Work Schedule System

## [v2027.11.0] — 2026-11-07

### 🚀 Overview
Release `v2027.11.0` delivers the complete remediation and QA verification of all 32 defects cataloged in the Pass 1 Master Defect Inventory across 7 core JavaScript modules. This release hardens formula preservation against overwrite corruption, bounds dynamic range queries against sheet boundaries, prevents Apps Script execution timeouts and quota exhaustion, enforces concurrency locking (`LockService`), eliminates data leakage in email distributions, and harmonizes schema across the entire 2027 multi-workbook ecosystem.

### 🔴 Critical Runtime, Formula Preservation & Range Bounds (Defects #1–#13)
1. **Typo Fix in Roster Reset (#1)**: Fixed `teamsSheet.clearContents()` typo to `clearContent()` in [`Admin_Tools.js`](Admin_Tools.js).
2. **Formula Preservation in Worker Rename (#2)**: Preserved date formulas, dynamic `XLOOKUP`, and `IMPORTRANGE` formulas during worker renaming across all sheets in [`Admin_Tools.js`](Admin_Tools.js) by reading `getFormulas()` in parallel with `getValues()`.
3. **Bounded Sheet Chunk Queries (#3, #4, #5)**: Clamped unbounded row/column dimension queries in `getSheetChunkData()`, `getWorkerUnavailabilityFromSheet()`, and `getWorkersAvailability()` in [`Admin_Tools.js`](Admin_Tools.js) against `sheet.getMaxRows()` and `sheet.getMaxColumns()`.
4. **Sheet Row Capacity Guards (#6, #7)**: Added automated capacity checks and `insertRowsAfter()` calls before batch-writing unavailability records in `syncUnavailabilityMatrix()` and new worker profiles in `commitWorkerImport()` in [`Admin_Tools.js`](Admin_Tools.js).
5. **ESSC Presentation Formula Preservation (#8)**: Preserved all header dates, theme lookups, and presentation formulas during `targetRange.setValues()` in `generateESSCForSeason()` in [`GenerateESSC.js`](GenerateESSC.js).
6. **Decoupled ESSC Spreadsheet Routing (#9)**: Switched `assignESSCWorker()` in [`GenerateESSC.js`](GenerateESSC.js) from `SpreadsheetApp.getActiveSpreadsheet()` to `getTargetESSCSpreadsheet()` to properly target external convocation workbooks.
7. **Dynamic Worker Email Row Offset (#10)**: Bound `startDataRow` dynamically to `schema.dataStartIdx` in `getWorkerEmails()` in [`Service_Reminders.js`](Service_Reminders.js), preventing header corruption when metadata rows shift.
8. **Headless Execution Guard in Validation Alerts (#11)**: Implemented `safeValidationAlert_()` in [`Validation.js`](Validation.js) to catch and log non-interactive trigger environments where `SpreadsheetApp.getUi()` is unavailable.
9. **Automated Trigger Staggering (#12)**: Shifted `sendMonthlyUnavailabilityReminder` trigger from 8:00 AM to 10:00 AM in `activateAllTriggers()` ([`Admin_Tools.js`](Admin_Tools.js)), preventing quota collision with `sendAssignmentReminders`.
10. **Email Concurrency Locking (#13, #27)**: Wrapped `sendAssignmentReminders` and `sendMonthlyUnavailabilityReminder` in `executeWithLockAndRetry_()` in [`Service_Reminders.js`](Service_Reminders.js) with `try ... finally { flushEmailLogs(); }` to prevent duplicate blasts and log loss.

### ⚡ Performance, Quota Optimization & Concurrency (Defects #14–#19, #23)
1. **Batch Worker Import Updates (#14)**: Consolidated row-by-row cell reads and writes in `commitWorkerImport()` into a single in-memory bounding box read and write pass in [`Admin_Tools.js`](Admin_Tools.js).
2. **2D Batch Replacement in Monthly Generation (#15)**: Replaced per-cell `setValue()` loops in `generateMonthlySheet()` with in-memory 2D array token substitutions written via a single `setValues()` in [`GenerateMonthly.js`](GenerateMonthly.js).
3. **Batch Reset for New Year (#16)**: Eliminated nested `sheet.getRange(...).getA1Notation()` loops over thousands of cells in `resetForNewYear()` in [`Admin_Tools.js`](Admin_Tools.js), batching updates via 2D array manipulation and `dataRange.setValues()`.
4. **RE Schedule Tab Filtering (#17)**: Filtered out `'archive'`, `'template'`, and `'old'` sheets in `fetchREDataForDates()` in [`SyncRESchedules.js`](SyncRESchedules.js), avoiding redundant cross-workbook RPCs.
5. **Year Provisioning Concurrency Lock (#18)**: Guarded `provisionNewYear()` in [`Admin_Tools.js`](Admin_Tools.js) with `executeWithLockAndRetry_('provisionNewYear', ..., { timeoutMs: 60000 })`, preventing concurrent triggers or double-clicks from spawning duplicate Drive folders.
6. **Recipient Quota Batching & Throttling (#19)**: Chunked recipients in `sendMonthlyUnavailabilityReminder()` into batches of 50 with `MailApp.getRemainingDailyQuota()` pre-checks in [`Service_Reminders.js`](Service_Reminders.js).
7. **Log Appending Optimization (#23)**: Replaced expensive `logSheet.insertRowAfter(1)` in `onEdit()` ([`Admin_Tools.js`](Admin_Tools.js)) with `logSheet.appendRow()` and bounded top-trimming.

### 📋 Schema Integrity, Privacy, Validation & Cleanups (Defects #20–#22, #24–#32)
1. **Worker Email Privacy Protection (#20)**: Placed email distribution lists in `bcc` instead of public `to` in `sendScheduleEmail_()` in [`ExportSchedulePDF.js`](ExportSchedulePDF.js).
2. **Deterministic Timezone Date Parsing (#21)**: Replaced native `Date.parse()` in [`SyncRESchedules.js`](SyncRESchedules.js) with explicit local noon `(y, m - 1, d, 12, 0, 0)` instantiation, eliminating UTC midnight backward-day shifts in `America/Los_Angeles`.
3. **Dynamic Formula Generation for Column AY (#22)**: Resolved Name and Chinese name column coordinates dynamically via `getWorkerSchema()` and `colIdxToLetter_()` in [`Admin_Tools.js`](Admin_Tools.js).
4. **Whitelist-Driven Annual Summary Workers (#24)**: Replaced index guessing with strict whitelist derived from `getWorkerSchema().nameCols` in `generateAnnualSummary()` ([`Admin_Tools.js`](Admin_Tools.js)).
5. **Validation Mappings for Columns AS & AT–AX (#25)**: Added default dropdown configurations for Column `AS` (Joy Fellowship) and `AT`–`AX` (Wednesday Bible Study & Locations) in `ensureValidationConfigSheet()` ([`Validation.js`](Validation.js)).
6. **Saturday Window Cleaning Integration (#26)**: Added Saturday Window Cleaning (Columns `AG` and `AH`) to `COL`, `loadColMappingsFromConfig()`, and `buildCleaningTable()` in [`Service_Reminders.js`](Service_Reminders.js).
7. **Quarter Sheet Cache Invalidation (#28)**: Added `bumpQuarterVersion(qName)` to `syncCoordinatorsToQuarterSheets()` in [`ExportSchedulePDF.js`](ExportSchedulePDF.js).
8. **Editor Permission Safeguard in Monthly Protection (#29)**: Wrapped protection assignment in `protectGeneratedMonthlySheet_()` ([`GenerateMonthly.js`](GenerateMonthly.js)) in `try...catch` so non-owner editors can generate monthly sheets without crashing.
9. **Default Published Link Sanitization (#30)**: Cleared legacy hardcoded October 2026 links in `EMAIL_TAB_PUBLISHED` default definition in [`ExportSchedulePDF.js`](ExportSchedulePDF.js).
10. **Unauthenticated URL Fallback Removal (#31)**: Stripped unauthenticated `uc?export=view` Google Drive fallback in `syncLogoToPresentationSheet()` ([`GenerateESSC.js`](GenerateESSC.js)) and `getLogoBlob()` ([`Service_Reminders.js`](Service_Reminders.js)).
11. **Orphaned Trigger Deprecation (#32)**: Deprecated unused `onSelectionChangeESSC()` trigger in [`GenerateESSC.js`](GenerateESSC.js) and routed event handling to `onSelectionChange(e)`.

---

## [v2027.10.0] — 2026-11-06

### 🚀 Overview
Release `v2027.10.0` completes the comprehensive full-stack architectural audit and live spreadsheet remediation across all four 2027 workbooks (`2027 Holy Work Schedule`, `2027 Holy Work Unavailability`, `2027 Sermon Topic & Hymns`, and `2027 GG ESSC Schedule`). Full details are documented in [`ARCHITECTURAL_AUDIT_REPORT_2027.md`](ARCHITECTURAL_AUDIT_REPORT_2027.md).

### 🔴 Critical Runtime, Concurrency & Live Sheet Fixes
1. **Live `Team Emails` & `Usher Team Emails` Formula Restoration (`#ERROR!` Fix)**:
   - Fixed corrupted `Workers!$1$3:$F` / `Workers!L4:L` formulas across all 36 columns of `Team Emails` and both columns of `Usher Team Emails` in `2027 Holy Work Schedule` (restored to `Workers!$F$3:$F`, `Workers!L3:L="Active"`, `Workers!<Col>3:<Col>=TRUE`).
   - Hardened `repair2027LiveSpreadsheets()` in [`Admin_Tools.js`](Admin_Tools.js) using function replacers to prevent `$1` regex capture-group corruption.
2. **Shared Drive Null-Owner Protection (`isAdminUser`)**:
   - Fixed `TypeError: Cannot read properties of null (reading 'getEmail')` in [`Admin_Tools.js`](Admin_Tools.js) when `ss.getOwner()` returns `null` on Google Shared Drive files (`0ABVpJUswSZTNUk9PVA`), with fallback to `APP_CONFIG.AUTHORIZED_EDITORS`.
3. **ESSC `Rosters` Dynamic Header Detection (`getTeamsAndWorkers` & `extractRostersFromMasterWorkers`)**:
   - Fixed `getTeamsAndWorkers()` in [`Admin_Tools.js`](Admin_Tools.js) (`isRosters === true`) to detect `=COUNTA(...)` numeric counts in Row 1 and read team headers from Row 2.
   - Rewrote `extractRostersFromMasterWorkers()` in [`GenerateESSC.js`](GenerateESSC.js) to use `detectWorkersHeaderRow_(data)` and dynamic header name matching instead of hardcoded 1-based column indices.
4. **Eliminated Global Namespace Collision (`letterToColumnIndex`)**:
   - Renamed the 0-based `letterToColumnIndex` helper in [`Service_Reminders.js`](Service_Reminders.js) to `letterToColIdx0_()`, preventing it from colliding with the 1-based `letterToColumnIndex` in [`GenerateMonthly.js`](GenerateMonthly.js).
5. **Comprehensive `LockService` Concurrency & TOCTOU Hardening**:
   - Fixed TOCTOU in `assignWorkerExact()` ([`Admin_Tools.js`](Admin_Tools.js)) by acquiring `LockService.getScriptLock()` *before* reading `cell.getDataValidation()`, respecting `rule.getAllowInvalid()`, and calling `SpreadsheetApp.flush()` before `bumpQuarterVersion()`.
   - Added `LockService.getScriptLock()`, `SpreadsheetApp.flush()`, and `bumpQuarterVersion()` to `assignESSCWorker()` in [`GenerateESSC.js`](GenerateESSC.js).
   - Wrapped `syncSWRSpeakers`, `applyDataValidation`, `applyValidationForESSC`, `syncRosters`, `commitWorkerImport`, `setWorkerStatus`, and `renameWorker` in `executeWithLockAndRetry_()`.
   - Protected `onEdit()` audit logging to `MonthlyOverridesLog` with a 2-second lock.

### ⚡ Quota & Performance Optimizations
1. **Batch `MasterConfig` Updates & In-Memory A1 Coordinate Parsing (`GenerateESSC.js`)**:
   - Replaced per-key `updateOrAppendConfigRow()` loops in `promptSetESSCDates()` and `resolveGridSpeakerPlaceholders()` with single-pass `batchUpdateOrAppendConfigRows_()`.
   - Replaced `targetSheet.getRange(a1).getRow()/getColumn()` inside `generateESSCForSeason()` with pure in-memory `parseA1Coords_(a1)` in [`AppConfig.js`](AppConfig.js) (eliminating 80+ RPCs).
2. **Memoized Cross-Workbook RE Schedule Parsing (`SyncRESchedules.js`)**:
   - Added `_reParsedScheduleCache` and consolidated two `.getRange().getValues()` reads per tab into a single `sheet.getDataRange().getValues()` call in `fetchREDataForDates()`, reducing external workbook reads by **87.5%** across `Q1`–`Q4`.
3. **2D Batch Writes (`Validation.js`, `Service_Reminders.js`, `GenerateMonthly.js`)**:
   - Replaced per-row `clearDataValidations()` loops in `applyValidationForESSC()` with a single 2D `setDataValidations(spkRulesMatrix)` call.
   - Replaced per-row `setBackground()` loops in `flushEmailLogs()` with a single `setBackgrounds(bgMatrix)` call.
   - Batched weekly cleaning header and date writes in `generateMonthlySheet()`.

### 📋 Schema & Cross-File Integrity
1. **Harmonized Columns `J` (Lunch Grace) & `K` (Morning Prayer)**:
   - Fixed inverted `J` and `K` defaults in `ensureValidationConfigSheet()` ([`Validation.js`](Validation.js)), `createDefaultConfigSheet()` ([`GenerateMonthly.js`](GenerateMonthly.js)), and `seedRoleConfigSheet()` ([`Admin_Tools.js`](Admin_Tools.js)) to match live `Q1`–`Q4`, `ValidationConfig`, and `MappingConfig`.
2. **Cleaning Team Fallbacks (`AL`–`AU`) & ESSC Team Aliases**:
   - Populated blank `ValidationConfig!C23:C32` (`AL`–`AU`) on live `2027 Holy Work Schedule` and added `VALIDATION_TEAM_ALIASES_`, `inferCleaningFallbackTeams_()`, and `expandTeamSearchKeys_()` in [`Validation.js`](Validation.js).
3. **Updated 2027 Default Spreadsheet & Folder IDs**:
   - Updated [`AppConfig.js`](AppConfig.js), [`Admin_Tools.js`](Admin_Tools.js), and [`Service_Reminders.js`](Service_Reminders.js) to reference the 2027 Google Drive folder (`1sH-ELkxwRLB78ts2rJeTpPmYtDY8KfzE`) and 2027 spreadsheets.

---

## [v2027.9.0] — 2026-11-06

### 🚀 Overview
Release `v2027.9.0` represents a comprehensive architectural audit and stability overhaul across the Holy Work Schedule ecosystem. This release resolves critical off-by-one column indexing risks, removes silent column truncation across large rosters, hardens concurrency with `LockService` and exponential backoff retry patterns, fixes live formula mismatches, and introduces automated administrative repair tooling.

---

### 🔴 Critical Bug Fixes & Runtime Safeguards
1. **Off-by-One Column Indexing in Quarter Restructuring & Slot Audits**:
   - Fixed `colLetterToIdx_` usage in `Admin_Tools.js` (`getActiveColsFromRoleConfig_` and `auditUpcomingMonthMissingSlots`). Because `colLetterToIdx_` is already 1-based (`A` -> `1`), redundant `+ 1` additions shifted active role columns right by 1, corrupting quarter sheets with `'X'` / `#D9D9D9` and misattributing coordinator slot ownership.
   - Fixed `colIdxToLetter_(c)` -> `colIdxToLetter_(colNum)` in `auditUpcomingMonthMissingSlots` so missing slots are accurately reported.
2. **Eliminated Silent Team & Worker Column Truncation**:
   - Increased `CONFIG.MAX_TEAM_COLS` (`30` -> `65`) and `CONFIG.MAX_WORKER_COLS` (`30` -> `75`) in `Admin_Tools.js`. Prevents silent truncation of `Teams` columns 31–41 (Cooking Teams 3–8, Cleaning, Joy Fellowship, Wednesday Bible Study) and `Workers` columns 31–55.
3. **Protected `MonthlyTemplate` Against Column Deletion**:
   - Removed `'MonthlyTemplate'` from the column-deletion loop in `removeCustomColumns`. Preserves the vertical 13-column month printout template layout when deleting Quarter sheet columns.
4. **Resolved Simple Trigger Authorization Exception**:
   - Removed `showSidebar()` from `onOpen()` in `Admin_Tools.js` and `showSidebarESSC()` from `onOpenESSC()` in `GenerateESSC.js`. Simple triggers execute in `AuthMode.LIMITED` and cannot invoke UI dialogs/sidebars without user interaction.
5. **Harmonized `Rosters` Row-1 vs. Row-2 Header Detection in ESSC**:
   - Added dynamic header row detection in `getESSCInitialData` (`GenerateESSC.js`) and `applyValidationForESSC` (`Validation.js`). Automatically detects when Row 1 contains `=COUNTA(...)` numeric summaries and reads team names from Row 2, preventing numeric team titles and corrupted sidebar worker rosters.
6. **Defensive Error Handling & Resource Cleanups**:
   - **`SyncSWR.js`**: Added per-row `try/catch` around `Utilities.parseDate` to prevent unformatted or date-range strings from crashing regional speaker sync.
   - **`Service_Reminders.js`**: Added `null` checks on `getLogoBlob()` before building `inlineImages` in `sendAssignmentReminders`, and guaranteed execution of `flushEmailLogs()` via `finally` blocks in `sendAssignmentReminders`, `sendMonthlyUnavailabilityReminder`, and `testSendAssignmentReminders`.
   - **`ExportSchedulePDF.js`**: Switched `findPublishedLink_` and `savePublishedLink_` to use `getDisplayValues()` to prevent Sheets date auto-coercion on `"October 2026"` from breaking link lookups and creating duplicate records. Deduplicated `standardRoles` in `upgradeCoordinatorsTabStructure_`.
   - **Global Scope**: Removed duplicate declarations of `letterToColumnIndex`, `AUTHORIZED_EDITORS`, and `BIBLE_PLAN_URL` across script files.

---

### ⚡ Quota & Performance Bottlenecks
1. **Eliminated 1,000+ Sequential RPC Calls in Unavailability Restructuring**:
   - Removed the per-row `getRowGroupDepth` loop in `restructureAndSyncUnavailabilityMatrix` (`Admin_Tools.js`). Replaced 56 individual `setColumnWidth` calls with a single batch `setColumnWidths(2, 56, 48)`.
2. **Atomic In-Memory Batch Operations**:
   - Replaced per-cell `.setValue()` loops in `renameWorker` with native `TextFinder` batch replacements across `Q1`–`Q4`.
   - Converted `commitWorkerImport` from per-row `getValues`/`setValues` calls to in-memory array modification and a single batch `setValues` call.
   - Optimized `flushEmailLogs` to apply background colors in a single `setBackgrounds(bgMatrix)` call.
3. **External Spreadsheet Caching**:
   - In `GenerateESSC.js` (`resolveSpeakerPlaceholders`), cached `MasterConfig` into memory once instead of reading the sheet up to 48 times per active mapping rule.
   - In `SyncRESchedules.js` (`autoSyncRESchedules`), cached each of the 6 external RE workbooks across quarters to eliminate 24 redundant `openById` invocations.
4. **Standardized Concurrency & Retry Architecture**:
   - Introduced `executeWithLockAndRetry_(actionName, fn, opts)` utilizing `LockService.getScriptLock()`, exponential backoff for transient Google service timeouts, and structured JSON Cloud Logging.
   - Wrapped `syncUnavailabilityFromMatrix` in atomic script locks, overwriting in-place without blank-sheet windows, and invalidating `CacheService` after `SpreadsheetApp.flush()`.

---

### 📋 Schema & Cross-File Integrity
1. **Dynamic Bilingual Speaker Formula (`AY` vs `BB`)**:
   - Updated `resetSermonTopicsSheet` in `Admin_Tools.js` to dynamically detect the bilingual speaker `XLOOKUP` column on `Q1`–`Q4` (Column `AY` in the 51-column layout), permanently eliminating `#REF!` errors in `2027 Sermon Topic & Hymns`.
2. **Spreadsheet Timezone Normalization**:
   - Standardized timezone handling across all scripts to use `America/Los_Angeles` matching `appsscript.json`.
3. **Mailing List Row Offset Fix**:
   - Updated `FILTER` formula ranges in `Team Emails` and `Usher Team Emails` from Row 4 to Row 3 (`Workers!$F$3:$F1001`), ensuring Row 3 members are included in automated email dispatches.
4. **Live Administrative Repair Suite**:
   - Added `repair2027LiveSpreadsheets()` to `Admin_Tools.js` to execute authorized, atomic remediation across all 4 production workbooks.

---

## [v2027.8.0] — 2026-09-30
- Initial implementation of Universal 14-Row Architecture (Rows 1–94).
- Multi-Church Config Engine (`ChurchProfile`, `ServiceLayoutConfig`, `RoleConfig`).
- Automated Tab Hygiene and dynamic monthly generator prompts.
