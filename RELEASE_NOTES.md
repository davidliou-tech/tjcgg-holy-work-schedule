# Release Notes — Holy Work Schedule System

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
