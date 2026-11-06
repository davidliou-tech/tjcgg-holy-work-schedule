# Holy Work Schedule - Administrator & Technical Guide

This guide is the single authoritative reference for church administrators and technical leads who maintain and operate the **Holy Work Schedule** ecosystem. It covers system architecture, spreadsheet structure, automated synchronizations, monthly schedule generation, email reminders, annual rollover provisioning, and troubleshooting.

---

## Table of Contents
1. [System Architecture](#1-system-architecture)
2. [Required Spreadsheet Structure](#2-required-spreadsheet-structure)
3. [Configuration & Constants](#3-configuration--constants)
4. [Data Synchronizations](#4-data-synchronizations)
5. [Monthly Schedule Generator & Template Engine](#5-monthly-schedule-generator--template-engine)
   - [5.1 PDF Export & Email Distribution Engine](#51-pdf-export--email-distribution-engine)
   - [5.2 Coordinator Missing-Slot Audit & Alert Tool](#52-coordinator-missing-slot-audit--alert-tool)
6. [Automated Email Reminders](#6-automated-email-reminders)
7. [Administrator Maintenance & Annual Rollover Tools](#7-administrator-maintenance--annual-rollover-tools)
   - [7.0 Multi-Church Onboarding & Directory Tools](#70-multi-church-onboarding--directory-tools)
   - [7.1 Annual Rollover & Year Folder Provisioner](#71-annual-rollover--year-folder-provisioner)
   - [7.2 Automation Triggers (Mandatory Annual Activation)](#72-automation-triggers-mandatory-annual-activation)
   - [7.3 External Synchronizations](#73-external-synchronizations)
   - [7.4 Standalone ESSC Scheduling Engine & Generator (`GenerateESSC.js`)](#74-standalone-essc-scheduling-engine--generator-generateesscjs)
   - [7.5 Auditing & Performance Tools](#75-auditing--performance-tools)
   - [7.6 Data Validation Applicator (`Validation.js`)](#76-data-validation-applicator-validationjs)
   - [7.7 Automated Tab Hygiene & Workbook Organization Suite](#77-automated-tab-hygiene--workbook-organization-suite)
8. [Performance, Polling & Caching](#8-performance-polling--caching)
   - [8.7 Load-Balanced Sorting & Workload Badges](#87-load-balanced-sorting--quarterly-workload-badges-sidebarhtml)
   - [8.8 One-Click Slot Unassignment & Clearing](#88-one-click-slot-unassignment--clearing-sidebarhtml)
9. [Configuring System Settings](#9-configuring-system-settings)
10. [Administrator FAQ & Troubleshooting](#10-administrator-faq--troubleshooting)
11. [Future Enhancement Roadmap](#11-future-enhancement-roadmap)

---

## 1. System Architecture

The Holy Work Schedule is a Google Apps Script project bound to the master spreadsheet. It comprises four primary functional domains:

```
+-----------------------------------------------------------------------------------+
|                           HOLY WORK SCHEDULE ECOSYSTEM                            |
+-----------------------------------------------------------------------------------+
       |                                       |                          |
       v                                       v                          v
[1. Sidebar Assistant]              [2. Automated Syncs]        [3. Email Reminders]
- Sidebar.html                      - syncSWRSpeakers           - Service_Reminders.js
- Admin_Tools.js                    - syncRESchedules           - Daily 8:00 AM Trigger
- Real-time conflict detection      - syncUnavailability        - 5-Day Lead Reminders
- Proximity & fatigue alerts        - Daily Bible Plan          - HTML Bilingual Tables
- Direct cell assignment            - Council on Duty           - Sermon Topics Link
       |                                       |                          |
       +---------------------------------------+--------------------------+
                                       |
                                       v
                     [4. Monthly Generator & Printouts]
                     - GenerateMonthly.js
                     - MonthlyTemplate (Print view)
                     - MappingConfig (Column-to-row map)
                     - Sheet locking & audit trail
```

1. **Sidebar Assistant (`Sidebar.html` & `Admin_Tools.js`)**: An interactive dockable UI for ministry coordinators to assign volunteers with live double-booking detection, proximity fatigue warnings, and member unavailability checking.
2. **Automated Integrations**: Background scripts that pull external data from the regional SWR Speaker sheet, Religious Education (RE) class sheets, and the annual Unavailability matrix.
3. **Automated Email Reminders (`Service_Reminders.js`)**: Daily automated email system notifying scheduled workers 5 days prior to Monday, Friday, and Saturday worship services.
4. **Monthly Schedule Generator (`GenerateMonthly.js`)**: Production engine that maps horizontal quarterly master data into the formatted, printable vertical `MonthlyTemplate` layout.

---

## 2. Required Spreadsheet Structure

The system relies on standardized tab names, header configurations, and row layouts.

### Core Tabs
* **`Q1`, `Q2`, `Q3`, `Q4`**: The quarterly schedule sheets standardized on the **Universal 14-Row Architecture (Rows 1–94)**:
  * **Rows 1–4**: Fixed metadata headers (Row 1: Notes/Anchor, Row 2: Category/Team, Row 3: Coordinator, Row 4: Specific Role title).
  * **Rows 5–19**: **MONDAYS** (Row 5 Header `#D9D9D9` bold; Rows 6–19 data).
  * **Rows 20–34**: **TUESDAYS** (Row 20 Header `#D9D9D9` bold; Rows 21–34 data).
  * **Rows 35–49**: **WEDNESDAYS** (Row 35 Header `#D9D9D9` bold; Rows 36–49 data).
  * **Rows 50–64**: **FRIDAYS** (Row 50 Header `#D9D9D9` bold; Rows 51–64 data).
  * **Rows 65–79**: **SATURDAYS AM** (Row 65 Header `#D9D9D9` bold; Rows 66–79 data).
  * **Rows 80–94**: **SATURDAYS PM** (Row 80 Header `#D9D9D9` bold; Rows 81–94 data).
  * **Visual & Formatting Standard**:
    * **Headers & Inapplicable Cells**: Dark gray background (`#D9D9D9`), centered text, with `'X'` on non-applicable service slots.
    * **Active Scheduling Cells**: Pure white background (`#FFFFFF`), centered text, with thin solid borders (`#B7B7B7`).
    * **Column A Dates**: Dynamic, quarter-bounded formulas (`=WORKDAY.INTL(...)` and `=IF(prev+7 > EOMONTH(...), "", prev+7)`) that auto-terminate cleanly without spilling into adjacent quarters.
    * **Column AY (Speaker Lookup)**: Dynamic `XLOOKUP` formulas linking speaker full names to the `Workers` tab (and `'X'` for Tuesdays).
    * **Inactive Row Dimming**: Conditional formatting `=$A6=""` soft-mutes inactive 13th or 14th weeks to `#F3F3F3` with `#AAAAAA` text.
* **`Teams`**: Defines team categories and lists the authorized worker roster under each team column using live `=FILTER(...)` formulas referencing the `Workers` tab (or rebuilt via `syncRosters()`).
* **`Workers`**: The **Single Source of Truth** for master worker profiles, email addresses, general weekday availability flags, speaker language fluency, and the 39 ministry qualification checkboxes:
  * **Cols A & B (`First Name`, `Last Name`) & Col E (`Schedule Name`)**: Column E (`=A3&" "&(B3)`) is the primary key used across `Teams`, `Q1–Q4`, `Unavailability`, and email lookups.
  * **Cols C & D (`Available Mondays`, `Available Fridays`)**: Checkboxes (`TRUE`/`FALSE`) read by `getWorkersAvailability()` to flag general weekday restrictions in the Sidebar and determine which workers are placed into the **Monday** and **Friday** sections of the external Unavailability Matrix.
  * **Col F (`Email`)**: Dynamically resolved via `getWorkerSchema()` to send 5-day advance service reminders and monthly unavailability reminders.
  * **Cols G & H (`Chinese Name`, `Full Speaker Name`)**: Referenced by `Column AY` (`XLOOKUP`) in `Q1–Q4` to format bilingual speaker titles (`English Name 中文名 (Church)`) for the Monthly Schedule and *Sermon Topics & Hymns* sheet.
  * **Col J (`Resident Church`)**: Blank for local members; set to home church for visiting speakers (visiting speakers are automatically excluded from monthly unavailability reminder emails).
  * **Col L (`Active` / `Status`) & Col M (`Roster Status`)**: Setting `Col L` to `Active` vs. `Inactive` gates whether a worker appears in `Teams`, `Q1–Q4` dropdowns, the Sidebar, and the Unavailability Matrix. `Col M` (`=IF(COUNTIF(O3:BA3,TRUE),"Rostered "&COUNTIF(O3:BA3,TRUE),"Not Rostered")`) verifies that the member is assigned to at least one ministry.
  * **Col N (`Unavailability Reminder`)**: Auto-formula determining whether the worker receives the 1st-of-the-month unavailability reminder email (`"Email"` vs `"No Email"`).
  * **Cols O–BA (39 Ministry / Team Checkboxes)**: Individual checkbox columns (`TRUE`/`FALSE`) across Worship & Prayer (`O–AA`), Religious Education (`AB–AG`), Fellowships (`AH–AM`), Cooking (`AN–AU`), and Cleaning (`AV–BA`).
  * **Smart Speaker-to-Interpreter Language Qualifications (`Cols P–S`)**:
    * **Col P (`Sermon Speaker (Chinese)`)** & **Col Q (`Sermon Speaker (English)`)** define a speaker's preaching language.
    * **Col R (`Interpreter (Chinese to English)`)** & **Col S (`Interpreter (English to Chinese)`)** define interpreter direction.
    * When a coordinator clicks an **Interpreter** cell (`Column C`) in `Q1–Q4`, the Sidebar checks the language of the **Speaker** assigned on that row (`getSpeakerLanguages()`) and automatically filters the candidate list to `C-to-E` or `E-to-C` interpreters.
  * **Multi-Church Schema Columns (`ensureWorkersSchema`)**: Appends `Worker ID`, `Status`, `Teams`, `Qualifications`, `Weekday Availability`, `Added On`, and `Notes` when running **Seed New Church** or the **Setup Wizard**. When committing via the Setup Wizard (`commitWorkerImport`), the system populates these columns **and automatically checks (`TRUE`) the matching native checkbox columns (`Cols O–BA`)** to keep `Teams` `=FILTER()` formulas 100% synchronized.
* **`Unavailability`**: Flattened internal cache of specific dates when workers are unavailable, populated daily from the external Unavailability spreadsheet.
* **`MonthlyTemplate`**: The master visual template for monthly printouts with bilingual formatting and dynamic `{{PLACEHOLDER}}` tags.
* **`Announcements`**: Contains annual announcements, including the monthly Council on Duty schedule.

### Multi-Church & Email Configuration Tabs (Hidden by Default)
> **💡 Tip:** All configuration and log tabs are hidden by default via **Automated Tab Hygiene**. Click **`⚙️ Admin Tools > 🗂️ Tabs > Toggle Admin Config & Log Tabs`** anytime to reveal or re-hide them. See [Workers_Rosters_and_Configuration_Tabs_Guide.md](Workers_Rosters_and_Configuration_Tabs_Guide.md) for the complete reference.

* **`ChurchProfile`**: Key-value store for church identity, timezone, `swr_location_tag`, `authorized_editors`, `bible_plan_url`, `year`, and cleaning row/column anchors.
* **`ServiceLayoutConfig`**: Defines the row boundaries on `Q1–Q4` for the Universal 14-Row Architecture (`MON` 6–19, `TUE` 21–34, `WED` 36–49, `FRI` 51–64, `SAT AM` 66–79, `SAT PM` 81–94) and enables/disables blocks (`Y`/`N`).
* **`RoleConfig`**: Master catalog of every assignable role, linking `q_column` in `Q1–Q4`, `template_row` in `MonthlyTemplate`, `team_sources` on `Teams`, and `requires_qual`.
* **`ValidationConfig` & `MappingConfig`**: Auto-compiled translation tables generated whenever you run **`⚙️ Admin Tools > 🔄 Sync Data > Compile Role Configurations`**. `ValidationConfig` drives `Q1–Q4` in-cell dropdowns; `MappingConfig` drives `MonthlyTemplate` row injection.
* **`Coordinators`, `EmailSettings`, `EmailTemplate`, `EmailLinks`, `PublishedLinks`**: Configuration tabs for ministry coordinator mappings (Row 3 of `Q1–Q4`) and the bilingual monthly PDF email distribution engine.

### Auto-Generated Log Tabs
* **`SWR Sync Log`**: Logs regional guest speaker imports and updates.
* **`RE Sync Log`**: Logs RE teacher and assistant imports.
* **`Email Log`**: Records every automated reminder batch (Success, Warning, Error, Skipped).
* **`MonthlyOverridesLog`**: Audit trail capturing manual in-cell edits made to generated monthly tabs via `onEdit` (automatically bounded to the newest 5,000 entries with batch trimming).
* **`MonthlyGeneratorLog`**: Diagnostic log generated whenever a monthly schedule is assembled.

---

## 3. Configuration & Constants

System settings can be adjusted in the script constants:

### General & Sidebar Settings (`Admin_Tools.js`)
```javascript
const CONFIG = {
  PROXIMITY_WARNING_DAYS: 14, // Proximity fatigue window (days)
  START_ROW: 5,               // Row index for header detection
  MAX_SCHEDULE_ROWS: 250,     // Max rows scanned in Q sheets
  CACHE_POLL_TIME: 1200,      // Sidebar active cell poll time (seconds)
  CACHE_DATA_TIME: 1800,      // Sidebar data cache duration (seconds)
};
```
*(Note: Color thresholds for fatigue in `Sidebar.html` are `PROXIMITY_RED_DAYS = 1` and `PROXIMITY_ORANGE_DAYS = 7`).*

### Email Reminder Settings (`Service_Reminders.js`)
```javascript
var REMINDER_LEAD_DAYS  = 5; // Sends emails 5 days prior to service
var CC_RECIPIENT        = "garden_grove@tjc.org";
var SHEET_URL           = getActiveSheetUrl(); // Self-referencing dynamic URL
```

### RE Schedule Settings (`SyncRESchedules.js`)
`RE_CLASSES_CONFIG` maps each class (`K`, `E1`, `E2`, `J1`, `J2`, `Sr`) to its source Google Sheet ID and specific column indices for AM/PM Teachers and Assistants.

### Monthly Generator Settings (`GenerateMonthly.js`)
```javascript
var AUTHORIZED_EDITORS = ["admin1@tjc.org", "admin2@tjc.org"]; // Allowed editors on locked monthly sheets
var BIBLE_PLAN_URL     = "https://docs.google.com/spreadsheets/d/..."; // Bible reading plan sheet
```

---

## 4. Data Synchronizations

The spreadsheet operates as a central hub synchronizing data through external and internal pipelines:

### 4.1 External Synchronizations (Cross-Spreadsheet)
* **A. Regional SWR Speaker Sync (`syncSWRSpeakers`)**:
  * **Function:** Ingests regional guest speaker assignments to keep local schedules aligned.
  * **Configuration:** Configured via **⚙️ Admin Tools > ⏰ Automation > Setup Daily SWR Sync URL**. Saved to `PropertiesService`.
  * **Annual Rollover Requirement:** Because the regional church committee releases a new SWR spreadsheet each year, the administrator must paste the new year's URL once annually.
  * **Failsafes:** Automatically skips rows labeled "English" to prevent overwriting bilingual services.
* **B. Religious Education (RE) Schedule Sync (`syncRESchedules`)**:
  * **Function:** Pulls Saturday RE teachers and assistants into the master schedule.
  * **Source:** Individual class Google Sheets (K, E1, E2, J1, J2, Senior) referenced in `RE_CLASSES_CONFIG`.
* **C. Daily Bible Reading Plan Sync**:
  * **Function:** Pulls daily Scripture chapters and injects them into the generated monthly schedule.
  * **Source:** External Bible Reading Google Sheet (`BIBLE_PLAN_URL`).
* **D. Annual Unavailability Matrix Sync (`syncUnavailabilityFromMatrix`)**:
  * **Function:** Reads checked blackout boxes from the external annual unavailability spreadsheet and flattens them into the local `Unavailability` tab for instant, zero-latency Sidebar conflict lookups.
  * **ESSC Seasonal Tab Support:** In addition to standard quarterly columns (`Cols B–BE`), the sync scans for dedicated `'Spring ESSC'` and `'Fall ESSC'` tabs. It evaluates the B2 anchor Wednesday date, calculates exact calendar dates for all 11 canonical shift blocks (`WED_EVE` to `SUN_PM`), and flattens blackouts into `Unavailability` with descriptive labels (e.g. `Spring ESSC (FRI_PM1)`).
  * **Configuration:** Configured via **⚙️ Admin Tools > 🔗 Connections > Link Unavailability Sheet** (auto-configured during annual rollover).

### 4.2 Internal Synchronizations (Same-Spreadsheet)
* **A. Council on Duty Sync**:
  * Injects the monthly Council Member names from the `Announcements` tab (Column C Primary, Column D Backup) into the monthly schedule header.
* **B. Automated Monthly Schedule Sync (`autoSyncScheduledMonths`)**:
  * Transposes horizontal data from quarterly sheets into the vertical monthly layout. Can be scheduled as a daily trigger to keep current and upcoming monthly sheets auto-updated.

### 4.3 Sheet Protection & Locking
When a monthly tab is generated, the script automatically locks the tab to prevent unauthorized edits that would break sync consistency. Only users listed in `AUTHORIZED_EDITORS` can edit locked monthly tabs directly.

---

## 5. Monthly Schedule Generator & Template Engine

The **`GenerateMonthly.js`** engine bridges the quarterly master sheets (`Q1`–`Q4`) with the printable `MonthlyTemplate`.

### Template Placeholders
The `MonthlyTemplate` uses double curly braces `{{TAG}}` to inject dynamic dates and assignments:

* **Title Tag**:
  * `{{MONTH_YEAR}}`: Upper-right corner. Outputs the bilingual title (e.g., `August 2027 / 二零二七年八月`).
* **Date Tags**:
  * Format: `{{DAY_WEEKNUMBER}}` (e.g., `{{SAT_1}}`, `{{MON_2}}`, `{{FRI_5}}`).
  * Replaced with exact `M/D` formatting (e.g., `8/15`). If a month has only 4 weeks, the 5th tags are cleared automatically, leaving the column blank.
* **Weekly Ranges (For Bible Reading)**:
  * Format: `{{WEEK_NUMBER}}` (e.g., `{{WEEK_1}}`, `{{WEEK_2}}`).
  * Generates date ranges for the week starting on Sunday (e.g., `8/2-8/8`).
* **Council on Duty**:
  * `{{COUNCIL_MEMBER}}`: Injected with the Council on Duty for that month from the `Announcements` tab.

### The `MappingConfig` Sheet
The `MappingConfig` tab is the translation matrix between the quarterly sheets and the template. If roles are added or rows/columns shift, update this sheet:

| Column | Description | Example |
| :--- | :--- | :--- |
| **Day Type** | The day of the week this mapping applies to (`SAT`, `MON`, `TUE`, `WED`, `FRI`). | `SAT` |
| **Role Name** | Descriptive label for the role. | `PRAYER` |
| **Section** | `AM` for morning slots, `PM` for afternoon slots, or `CLEANING`. | `AM` |
| **Master Column** | The exact column letter in the Quarterly sheet (e.g., `D` for Praise). | `D` |
| **Template Row** | The exact row number in `MonthlyTemplate` where this worker belongs. | `14` |

*(Note: If `MappingConfig` is ever accidentally deleted, the script will automatically regenerate it with default settings on the next run).*

### Date Matching & Stability
The script matches dates by evaluating JavaScript `Date` objects first, with a fallback to raw text strings (`"M/D"`). To preserve merged cells, borders, and complex formatting without risking Google Sheets API crashes, data is written cell-by-cell. An execution time of ~60–100 seconds is normal and expected.

### In-Place Monthly Generation & `gid` Preservation
Historically, regenerating or nightly re-syncing a monthly schedule deleted and recreated the tab from `MonthlyTemplate`. This changed the sheet's internal Google Sheets `gid`, breaking published web links (`pubhtml?gid=...`) and erasing any last-minute manual cell overrides recorded in `MonthlyOverridesLog`.

Starting in version `2027.9.0`, `generateMonthlySheet(year, month, options)` supports **In-Place Update Mode**:
* **Preserving Published Web Links (`gid`)**: When updating an existing monthly sheet, the generator updates cells in-place without deleting the tab. As a result, the tab's `gid` remains identical, ensuring "Publish to Web" URLs (`PublishedLinks`) distributed to church members never break.
* **Manual Override Protection (`MonthlyOverridesLog`)**: During nightly background syncs (`autoSyncScheduledMonths()`), `preserveManualOverrides: true` is enabled. The generator reads the `MonthlyOverridesLog` tab for that target month and skips overwriting any cell that was manually modified by an administrator, automatically re-applying their edits over incoming quarterly sync data.
* **Interactive Generator Options**: When generating a month interactively via **📅 Schedule Assistant > Generate Monthly Schedule** (`promptMonthlyGenerator()`) for a tab that already exists:
  * **`YES` (Sync In-Place)**: Updates all dates, Bible plan entries, and worker assignments in-place while keeping existing manual overrides and preserving the published `gid`.
  * **`NO` (Full Rebuild from Template)**: Clears and copies `MonthlyTemplate` fresh, resetting manual overrides while **still preserving the sheet tab and `gid`** (via `copyTo` rather than tab deletion).

### 5.1 PDF Export & Email Distribution Engine
The system provides an automated engine (`ExportSchedulePDF.js`) to export the final monthly schedule as a PDF and distribute it via email.

* **Trigger**: Accessed manually by coordinators via **📅 Schedule Assistant > 📧 Email > Export & Email Next Month PDF**.
* **PDF Generation**: Automatically hits the Google Sheets export URL (`.../export?format=pdf&gid=...`) to download the target month's tab as a landscape, fit-to-width PDF. The PDF is archived securely in the same Google Drive folder as the spreadsheet.
* **Configuration Tabs**: The engine relies on five hidden configuration tabs. To prevent spreadsheet bloat, these tabs are **lazy-loaded**—they are programmatically created and populated with default templates the first time a user clicks **📅 Schedule Assistant > 📧 Email > Edit Email Template**.
  1. `EmailSettings`: Stores key-value system settings such as language flags (`en,zh`), sender name, and the templated email subject. *(Note: The actual recipient distribution list is saved securely to `PropertiesService.getScriptProperties()` via the "Set PDF Distribution List" menu, not here, to prevent accidental email leaking or modification).*
  2. `EmailTemplate`: The primary message body. Supports Markdown syntax (bolding, lists, links) and dynamic `{{tokens}}`. The engine parses this tab row-by-row, translating it into clean HTML.
  3. `EmailLinks`: A dynamic list of helpful URLs (Master Calendar, Bible Reading Plan) appended as bullet points. Supports Markdown link syntax.
  4. `Coordinators`: Centralized master directory of ministry coordinators. Controls both email reminder footers (via `include_in_email` = `Y`/`N`) and auto-populates Row 3 (`COORDINATOR:`) across `Q1`–`Q4` master sheets (via `q_columns` mapping).
  5. `PublishedLinks`: A registry cache of "Publish to Web" URLs for each month's generated sheet.
* **Publishing Validation**: Google Apps Script cannot programmatically execute a "Publish to web" action for an individual tab. If a month's sheet hasn't been published yet, the script prompts the coordinator to manually publish the tab, paste the link, and strictly validates that the URL matches the required `pubhtml?gid=...&single=true` format and returns HTTP 200 before saving it to the `PublishedLinks` registry.
* **Email Testing**: Coordinators can run **📅 Schedule Assistant > 📧 Email > Preview Email (send to me)** to receive a test copy of the email before broadcasting to the entire church.

### 5.2 Coordinator Missing-Slot Audit & Alert Tool
Located under **`📅 Schedule Assistant > 📧 Email > Check Missing Slots for Next Month`**, this tool audits the upcoming month's schedule across `Q1–Q4` to catch unfilled duties before monthly publication:
* **Active Cell Inspection**: Scans all active service rows (Monday, Friday, Saturday AM/PM) for dates in the upcoming calendar month. Ignores header rows, inactive gray `#D9D9D9` slots, and intentionally skipped slots marked with `'X'`.
* **Coordinator Grouping**: Automatically groups unfilled roles by ministry coordinator based on the `q_columns` mapping defined in the `Coordinators` tab.
* **Interactive Summary & Notification**: Displays a clean modal summary dialog listing all missing positions and dates per coordinator. If any slots remain open, administrators can click **Send Email Reminders to Coordinators** to dispatch an automated email notice directly to each coordinator with their specific open roles.

---

## 6. Automated Email Reminders

The **`Service_Reminders.js`** script sends reminder emails to scheduled workers **5 days prior** to Monday, Friday, and Saturday services.

### Trigger Execution Logic
The `sendAssignmentReminders()` function runs daily between 8:00 AM and 9:00 AM. Each morning, it calculates the target date (today + 5 days). If the target date falls on Monday, Friday, or Saturday, it processes assignments; on other days, it skips cleanly.

### Email Layout & Content
Reminder emails are styled in responsive HTML with church branding:
1. **Header & Bilingual Greeting**: True Jesus Church logo, service name, and date.
2. **📅 Service Schedule Table**: Times for hymnal worship, prayer, and message.
3. **👥 Holy Work Assignments Table**: Speaker, Interpreter, Pianist, Hymn Leader, AV, Usher, etc.
4. **🧹 Cleaning Duties (Friday/Saturday)**: Restrooms, Love Hall, Chapel, Trash.
5. **📘 Religious Education (Saturday)**: Teachers and Assistants by class level.
6. **📋 Action Items & Reminders**: Direct link to the **Sermon Topics & Hymns** Google Sheet, instructions for finding replacements, and early arrival guidance.

---

## 7. Administrator Maintenance & Annual Rollover Tools

All administrative functions are accessible via the **⚙️ Admin Tools** menu in the top toolbar, organized into these groups: **🚀 Setup Wizard / 🌱 Seed New Church / 📊 Generate Annual Summary** (top level), **🗂️ Tabs**, **🔗 Connections**, **🔄 Sync Data**, **⏰ Automation**, **📅 Annual Rollover**, and **🧰 Repair & Maintenance (Advanced)**. ESSC and email tools live under the **📅 Schedule Assistant** menu.

### 7.0 Multi-Church Onboarding & Directory Tools
* **🚀 Open Setup Wizard (`SetupWizard.html`)**: An interactive 3-step Visual Column Mapper for bulk importing and updating worker rosters:
  - **Source Selection**: Ingest directly from an existing sheet tab or paste a raw clipboard table (tab/comma-delimited).
  - **Visual Column Mapping**: Dropdowns with smart auto-detection for split First/Last Names (auto-combines), Chinese Name, Email, Phone, Status, and Qualifications.
  - **Preloaded Garden Grove Teams Catalog**: Includes all 34+ teams across Worship, Fellowships, RE Classes, Cooking, and Cleaning.
  - **Matrix Checkbox Mode**: Auto-matches incoming columns to standard teams and detects `✓`, `x`, `yes`, `1`, or `true` cell markers.
  - **`+ Add Custom Team`**: Allows adding custom church-specific ministries dynamically.
  - **Checkbox Auto-Bridging**: Automatically checks corresponding team checkboxes on the `Workers` tab upon commit, keeping `=FILTER()` formulas on `Teams` synchronized.
* **🌱 Seed New Church**: Instantiates `ChurchProfile`, `ServiceLayoutConfig`, and `RoleConfig` tabs with local church settings and compiles matching validation rules.
* **Compile Role Configurations (`compileRoleConfig`)** *(⚙️ Admin Tools > 🔄 Sync Data)*: Validates `RoleConfig` (with collision and RE column checks) and auto-generates `ValidationConfig` (dropdowns) and `MappingConfig` (monthly printout layout).

#### Worker Roster Maintenance Runbook (Adding, Updating & Deactivating Workers)
> **📘 Complete Reference:** See [Workers_Rosters_and_Configuration_Tabs_Guide.md](Workers_Rosters_and_Configuration_Tabs_Guide.md) for full column-by-column specifications.

1. **Adding New Workers**:
   * **Direct Entry (1–5 workers)**: Add a row on the `Workers` tab (copying formulas in `Col E` Schedule Name, `Col M` Roster Status, and `Col N` Unavailability Reminder from the row above). Fill in `First Name`, `Last Name`, `Available Mondays`/`Fridays`, `Email`, set `Col K` = `Individual`, `Col L` = `Active`, and check (`TRUE`) the approved ministry columns in `Cols O–BA`.
   * **Bulk Import**: Run **`⚙️ Admin Tools > 🚀 Open Setup Wizard`** to map and commit multiple workers at once.
2. **Post-Update Checklist (1-Click Unified Action vs. Manual Steps)** *(Run after any roster addition, team checkbox change, or status change)*:
   * **⚡ 1-Click Unified Action (Recommended)**: Click **`⚙️ Admin Tools > 🔄 Sync Data > ⚡ Apply All Worker Roster Changes (1-Click)`** or click the **"Refresh Rosters"** button in `Sidebar.html`. This executes the complete 4-step sequence automatically in ~5 seconds:
     1. **Safe Roster Sync & Cache Invalidation (`syncRosters`)**: Checks if `Teams` uses formulas or wide columns to preserve live `=FILTER()` formulas while safely purging chunked `CacheService` keys (`teams_and_workers`, `workers_day_avail`, `workers_data`).
     2. **Re-Apply In-Cell Dropdowns (`applyAllValidation`)**: Rebuilds data validation across all active columns in `Q1–Q4`.
     3. **Sync Unavailability Matrix (`restructureAndSyncUnavailabilityMatrix`)**: Syncs active rostered workers to the external matrix if linked.
     4. **Quarter Version Invalidation (`bumpAllQuarterVersions`)**: Bumps shared quarter version tokens so all active coordinator sidebars refresh instantly.
   * **Manual Step-by-Step Alternative**:
     1. **Rebuild `Q1–Q4` In-Cell Dropdowns**: Click **`⚙️ Admin Tools > 🧰 Repair & Maintenance (Advanced) > Re-apply Data Validation Dropdowns`**.
     2. **Sync External Unavailability Matrix**: Click **`⚙️ Admin Tools > 🔄 Sync Data > Sync Workers to Unavailability Matrix`**.
     3. **Refresh Sidebar Cache**: Click **"Refresh Rosters"** (or **🔄**) at the top of the Assignment Sidebar.
3. **Modifying, Deactivating, or Renaming Existing Workers**:
   * **Changing Ministry Qualifications**: Check/uncheck `Cols O–BA` on `Workers`, then run **`⚡ Apply All Worker Roster Changes (1-Click)`** (or `Re-apply Data Validation Dropdowns` and click **Refresh Rosters** in the Sidebar).
   * **Deactivating a Worker (Moved, On Leave, or Stepped Down)**: **Never delete rows from `Workers`**. Change `Col L` (`Active`) from `Active` to **`Inactive`**, then run **`⚡ Apply All Worker Roster Changes (1-Click)`** (or `Re-apply Data Validation Dropdowns` and `🔄 Sync Workers to Unavailability Matrix`).
   * **Renaming a Worker**: Run `renameWorker("Old Name", "New Name")` in Apps Script (or update `First Name`/`Last Name` on `Workers` and update their existing cells in `Q1–Q4`), then re-apply dropdowns and sync the unavailability matrix.

### 7.1 Annual Rollover & Year Folder Provisioner
* **🚀 Provision New Year (Full Rollover & Year Folder)**: The master annual rollover provisioner. Prompts for the target year (e.g., `2027`), locates or creates the year folder in Google Drive (`Holy Work Schedules > [Year]`), and provisions the complete 3-sheet annual suite:
  1. **Holy Work Schedule**: Duplicates the schedule into the year folder, updates Q1–Q4 anchor dates (`A2`) to the new year, clears worker names (preserving formulas, headers, and `'X'` marks), automatically enforces the **Universal 14-Row Architecture** (`restructureQuarterSheetsToUniversal14`), clears logs, and links to the twin sheets.
  2. **Unavailability Sheet**: Duplicates the unavailability matrix into the year folder, sets permissions to **"Anyone with the link can edit"**, sets anchor date `A2` to `1/1/[Year]` and `B2` to `Q1`, automatically enforces the **Universal 14-Column Architecture** (56 weeks with quarter-bounded formulas), synchronizes and alphabetizes active workers into four sections (Monday, Friday, Saturday AM, Saturday PM), resets checkboxes to `false`, and auto-links to the master schedule.
  3. **Sermon Topics & Hymns Sheet**: Duplicates the sermon topics workbook into the year folder, sets permissions to **"Anyone with the link can edit"**, updates tab names (e.g., `Spring ESSC [Year]`), rewrites `=IMPORTRANGE(...)` formulas to reference Column AY of the newly created Holy Work Schedule (`AY6:AY19`, `AY51:AY64`, `AY66:AY79`, `AY81:AY94`), clears previous topics/hymns, and auto-links into the reminder system.

> **💡 All Three Sheets Are Automatically Linked:** Administrators do NOT need to manually run linking commands during annual rollover.
>
> **Mandatory Post-Provisioning Steps:**
> 1. Open the newly provisioned Holy Work Schedule and click **⚙️ Admin Tools > ⏰ Automation > Activate All Automation Triggers**.
> 2. Open the new **Sermon Topic & Hymns** sheet, click cell **B4**, and click **"Allow access"** if prompted by Google Sheets to authorize the IMPORTRANGE link.
> 3. Once the regional committee publishes the new year's SWR Speaker sheet, run **⚙️ Admin Tools > ⏰ Automation > Setup Daily SWR Sync URL**.
> 4. Verify that the church annual theme at the top of `MonthlyTemplate` is updated for the new year.

---

### 7.2 Automation Triggers (Mandatory Annual Activation)
* **Activate All Automation Triggers** *(⚙️ Admin Tools > ⏰ Automation)*: One-click programmatic setup for all automated background tasks:
  1. `sendAssignmentReminders` (Daily at 8:00 AM)
  2. `syncSWRSpeakers` (Daily at 1:00 AM)
  3. `syncUnavailabilityFromMatrix` (Daily at 2:00 AM)
  4. `autoSyncRESchedules` (Daily at 3:00 AM)
  5. `autoSyncScheduledMonths` (Daily at 4:00 AM)
  6. `generateAnnualSummary` (Monthly on 1st at 5:00 AM)
  7. `sendMonthlyUnavailabilityReminder` (Monthly on 1st at 8:00 AM)

> **⚠️ Why Triggers Do Not Copy Automatically:** Google Apps Script deliberately disables and clears all time-driven triggers when a spreadsheet is duplicated in Google Drive (to prevent duplicate emails and unintended execution). **The administrator must run `Activate All Automation Triggers` once in the new year's spreadsheet.**

---

### 7.3 External Synchronizations
* **🔄 Sync SWR Speakers**: Immediately pulls the external SWR Speaker sheet.
* **🔄 Sync Unavailability Matrix**: Ingests blackout dates and reasons from the linked external Unavailability spreadsheet and flattens them into the local `Unavailability` tab.
* **🔄 Sync Workers to Unavailability Matrix** *(also restructures the matrix to the Universal 14-Column layout)*: Evaluates all active rostered Garden Grove members in the `Workers` tab and synchronizes them into four alphabetized sections in the external matrix, preserving all existing checkmarks.
* **🔄 Sync ESSC Unavailability Rosters (`promptSyncESSCUnavailabilityRosters`)** *(available only under 📅 Schedule Assistant > 🎪 ESSC Schedule Generator)*: Automatically provisions or refreshes the `'Spring ESSC'` or `'Fall ESSC'` tabs in the external Unavailability spreadsheet. Populates active rostered Garden Grove members (where `Resident Church` in Column J is blank) with 11-shift checkboxes (`WED_EVE` to `SUN_PM`), displays dynamic dates for each shift block in Row 2 linked to anchor Wednesday in B2, and preserves existing member responses.
* **🔗 Link Unavailability Sheet**: Prompts for the URL or File ID of the external Unavailability spreadsheet. *(Auto-linked during rollover; used only for initial setup or mid-year relinking).*
* **🔗 Link Sermon Topics & Hymns Sheet**: Prompts for the URL or File ID of the Sermon Topics spreadsheet. *(Auto-linked during rollover; used only for initial setup or mid-year relinking).*
* **⏰ Setup Daily SWR Sync URL**: Saves and validates the regional SWR Speaker URL in `PropertiesService`.

---

### 7.4 Standalone ESSC Scheduling Engine & Generator (`GenerateESSC.js`)

The **Standalone ESSC Scheduling System** (`GG ESSC Schedule.xlsx` / Google Sheet) provides a 3-tier decoupled architecture for managing semi-annual Evangelical Service & Spiritual Convocation (Spring ESSC & Fall ESSC) events without cluttering weekly quarterly sheets. The standalone ESSC workbook contains **zero scripts**, driven by centralized push automation from the Master HWS and real-time native spreadsheet formulas.

#### Unified Sidebar (`Sidebar.html`) & Safe 1500ms Polling
* **Unified Sidebar Assistant**: The standalone ESSC workbook serves the exact same unified, full-featured `Sidebar.html` UI (`showSidebarESSC()`), completely eliminating the legacy inline assistant (`getSidebarHtml()`).
* **Safe Polling (`1500ms`)**: Polling operates at a safe `1500ms` base interval (with adaptive back-off to `3000ms`/`5000ms` and background tab sleeping), permanently retiring the legacy CPU-intensive `80ms` loop.
* **Full Grid Row Coverage (Rows 1–45)**: Context polling and selection triggers now support the full ESSC session range up to row 45 (`ESSC_CONFIG.GRID.LAST_ROW`), ensuring that late Saturday and Sunday evening sessions are fully active.
* **Master Unavailability Auto-Fallback**: If the local `Unavailability` tab is not present in the standalone ESSC workbook, `getUnavailabilityData()` automatically falls back to reading the `Unavailability` tab from `getMasterHWSSpreadsheet()`. All ESSC conflict badges (`ESSC Unavail`, `Choir`, `Choir Buffer`, `Cooking Duty`) function seamlessly and identically across workbooks.

#### Key Administrator Functions (Master HWS Control Tower)
All ESSC administration is centralized in the Master Holy Work Schedule under **`📅 Schedule Assistant > 🎪 ESSC Schedule Generator`**:
* **🔗 Connect Standalone ESSC Workbook (`promptConnectESSC`)**:
  - Menu: `Schedule Assistant > 🎪 ESSC Schedule Generator > 🔗 Connect Standalone ESSC Workbook...`
  - Prompts admin for the URL or Spreadsheet ID of that year's ESSC workbook (e.g. `2027 GG ESSC Schedule`). Saves `ESSC_SPREADSHEET_ID` into `MasterConfig` and `PropertiesService`.
* **📅 Set ESSC Event Dates (`promptSetESSCDates`)**:
  - Menu: `Schedule Assistant > 🎪 ESSC Schedule Generator > 📅 Set ESSC Event Dates...`
  - Prompts for Spring Wednesday (`SPRING_ESSC_WEDNESDAY_START`) and Fall Wednesday (`FALL_ESSC_WEDNESDAY_START`) start dates. Updates `MasterConfig` in both workbooks.
* **🔄 Push Dropdowns & Resolve Placeholders (`syncESSCFromMaster`)**:
  - Menu: `Schedule Assistant > 🎪 ESSC Schedule Generator > 🔄 Push Dropdowns & Resolve Placeholders`
  - Automatically extracts active worker teams from Master HWS (`Teams` / `Workers`), pushes them to the ESSC `Rosters` tab, sets dropdown validation rules, whitelists visiting ministers (clearing red warning chips), clears empty meal row chips, and **arms dynamic formulas in Column A and Column E**.
* **✨ Inject Dynamic Date & Speaker Formulas (`injectDynamicGridFormulasMenu`)**:
  - Menu: `Schedule Assistant > 🎪 ESSC Schedule Generator > ✨ Inject Dynamic Date & Speaker Formulas`
  - Injects native calendar formulas (`=IFERROR(DATEVALUE(TO_TEXT(XLOOKUP(...)))), "")`, `=A$2+1`, etc.) into Column A and `XLOOKUP` formulas into Column E. This enables coordinators to change dates or visiting ministers in `MasterConfig` of the ESSC sheet and see the entire grid update in real time with **zero script execution**!
* **🔄 Sync ESSC Unavailability Rosters (`promptSyncESSCUnavailabilityRosters`)** *(available only under 📅 Schedule Assistant > 🎪 ESSC Schedule Generator)*:
  - Menu: `Schedule Assistant > 🎪 ESSC Schedule Generator > 🔄 Sync ESSC Unavailability Rosters`
  - Provisions or refreshes the dedicated external Unavailability sheet tab (`Spring ESSC` or `Fall ESSC`) with active rostered Garden Grove members (where `Resident Church` in Column J is blank), 11-shift checkboxes, and Row 2 dynamic shift block dates.
* **🚀 Generate & Lock Presentation Schedules (Remote)**:
  - Menu: `Schedule Assistant > 🎪 ESSC Schedule Generator > 🚀 Generate Spring ESSC Schedule (Remote)` / `🚀 Generate Fall ESSC Schedule (Remote)`
  - Reads data from `Spring_ESSC_Grid` or `Fall_ESSC_Grid` in the ESSC sheet, evaluates `ESSC_MappingConfig` expressions, auto-populates formatted multi-line text into presentation tabs (`Spring_ESSC_Worship`, `Spring_ESSC_Work`, `Fall_ESSC_Worship`, `Fall_ESSC_Work`), and **enforces sheet protection (locking)** on the output tabs so coordinators can only edit the data grid.

#### ESSC Conflict Detection & Shift Matching Architecture
When coordinators schedule workers on `Spring_ESSC_Grid` or `Fall_ESSC_Grid` using the Assignment Sidebar (`Sidebar.html`), the system performs specialized real-time checks:
1. **11-Shift Block Dynamic Matching**: Evaluates shift keys across relative calendar day offsets:
   - `WED_EVE` (Offset 0), `THU_EVE` (Offset 1)
   - `FRI_AM`, `FRI_PM1`, `FRI_PM2`, `FRI_EVE` (Offset 2)
   - `SAT_AM`, `SAT_PM`, `SAT_EVE` (Offset 3)
   - `SUN_AM`, `SUN_PM` (Offset 4)
   Dates are parsed at 12:00:00 noon to avoid Daylight Saving Time shifts, and mapped to grid rows via regex session matching.
2. **Choir Presentation Conflict & Pre-Slot Warm-up Buffer**:
   - Detects if a worker is singing in the Choir (Column P) during the session (`Choir` badge).
   - For Saturday AM (Evangelical Service) and Sunday PM (Holy Communion / Choir), choir members require the preceding row (`actualRow - 1`) for vocal warm-up, robing, and staging. The engine flags them as unavailable for support duties during the preceding slot (`Choir Buffer` badge). Children's choir presentations are exempt.
3. **Cooking Team Meal Prep & Cleanup Window Conflicts**:
   - Identifies scheduled cooking groups (`Cooking Team 1–8`, regional fellowships `Cypress / Buena Park`, `Anaheim Hills / Yorba Linda`, `Irvine / Lake Forest / Orange`, and fellowships `Agape`, `EBF`, `Sisters`, `Vessel`).
   - Flags workers on duty during their meal windows:
     - Lunch Window: ~10:30 AM–1:30 PM (630–810 mins)
     - Dinner Window: ~4:15 PM–7:00 PM (975–1140 mins)
     - Concurrent meal sessions
4. **Visual Conflict Badges**: Renders color-coded badges in the Assignment Sidebar: `Choir` (Purple `#6d28d9`), `Choir Buffer` (Indigo `#4338ca`), `Cooking Duty` (Orange `#c2410c`), and `ESSC Unavail` (Rose `#be123c`).

---

### 7.5 Auditing & Performance Tools

* **📊 Generate Annual Summary**: Compiles a master tally of all assignments per worker across Q1–Q4 (excluding Cooking Team and non-English character roles) into an `Annual Summary` tab.
* **Run Master Audit (Check Duplicates)**: Scans across all Q sheets for same-day overlapping assignments.
* **🛠️ Restructure Q1-Q4 (Universal 14-Row)**: Standardizes Q1–Q4 into the uniform 14-row architecture (Rows 1–94), injecting quarter-bounded formulas in Column A, dimming rules, and Column AY speaker lookup.
* **Sync Workers to Unavailability Matrix (Universal 14-Col restructure)**: Standardizes the external Unavailability matrix into 56 date columns (14 per quarter, Cols B–BE) with self-terminating formulas while preserving member checkboxes.
* **🧹 Remove Column(s) (Interactive)**: Interactive admin tool that prompts for target column letter(s) or ranges (e.g. `AS, AT, AU` or `45-47`). Displays a safety preview of role titles, deletes columns across Q1–Q4 and `MonthlyTemplate` in descending order, shifts `Coordinators` `q_columns` mapping and `MappingConfig`, and re-applies dropdown validation and universal layout formatting.
* **Remove Yellow Cell Highlights (Fix Lag)**: Strips heavy `COUNTIF` conditional formatting rules that cause cell editing lag.

---

### 7.6 Data Validation Applicator (`Validation.js`)

#### Why It Is Necessary (Even With the Sidebar)
While the **Schedule Assistant Sidebar** (`Sidebar.html`) handles intelligent multi-factor decision-making (unavailability, proximity warnings, double-booking checks, and same-day session conflicts), the **Data Validation Applicator** (`Validation.js`) provides native, sheet-level guardrails:

1. **Direct In-Sheet Editing & Autocompletion**: Coordinators do not always have the sidebar open for quick single-name edits. Native in-cell dropdowns allow users to click a cell directly and pick from a filtered dropdown arrow or use native autocomplete.
2. **Mobile Device Support (iOS & Android)**: The Google Sheets mobile app **does not support HTML Apps Script sidebars**. Church workers reviewing or updating assignments on phones or tablets rely exclusively on native in-cell dropdowns.
3. **Accidental Entry & Copy-Paste Protection**: Native validation rules reject invalid input (`setAllowInvalid(false)`), preventing accidental typos, misspellings, or unauthorized names.
4. **Repairing Lost Dropdowns (Google Sheets Cut/Paste Bug)**: In Google Sheets, cutting (`Cmd+X` / `Ctrl+X`) a cell and pasting it elsewhere **permanently strips the data validation rule** from the source cell. The Applicator serves as a 1-click healing tool that rebuilds every dropdown rule across Q1–Q4 in seconds.
5. **Defense-in-Depth for the Sidebar**: When the sidebar executes an assignment via `assignWorkerExact()`, it inspects the cell's native `getDataValidation()` rule as a pre-flight sanity check before writing to the sheet.

#### When to Use It
* **After Annual Rollover**: Once a new year is provisioned, run it to ensure all Q1–Q4 sheets have fresh, intact dropdowns.
* **When Dropdown Arrows Disappear**: Whenever coordinators report missing dropdown arrows (typically caused by cut/paste or formatting overwrites).
* **After Restructuring Rows or Columns**: If row architecture is modified or layout changes are made.
* **After Updating the `Teams` Sheet**: If team roster range boundaries change.

#### How to Use It
1. In the top menu, click **`⚙️ Admin Tools > 🧰 Repair & Maintenance (Advanced) > Re-apply Data Validation Dropdowns`**.
2. The script processes ~1,370 cell rules across Q1, Q2, Q3, and Q4 simultaneously in under 5 seconds.
3. A success dialog will confirm:
   > `Success! 🎉 Successfully applied [N] dropdown validation rules across Q1, Q2, Q3, and Q4.`

---

### 7.7 Automated Tab Hygiene & Workbook Organization Suite

To preserve spreadsheet speed and keep the workbook clean and approachable for coordinators, the system retains the **multi-tab single-purpose architecture** (where each configuration sheet is a simple 2D table that travels intact with `File > Make a Copy`) while introducing **automated tab hygiene**:

#### 1. Auto-Hide Config & Log Tabs on Creation
* **Configuration Sheets**: All system configuration sheets (`ChurchProfile`, `ServiceLayoutConfig`, `RoleConfig`, `ValidationConfig`, `MappingConfig`) and email template tabs (`EmailSettings`, `EmailTemplate`, `EmailLinks`, `Coordinators`) automatically invoke `sheet.hideSheet()` upon initial creation or compilation.
* **System Log Sheets**: Audit sheets (`SWR Sync Log`, `RE Sync Log`, `MonthlyGeneratorLog`, `MonthlyOverridesLog`, `Email Log`, `WorkerImportLog`) are automatically hidden on creation.
* **Monthly Template Decoupling**: `MonthlyTemplate` remains permanently hidden. When `generateMonthlySheet(year, month)` creates a monthly printout, it immediately unhides the cloned tab (`newSheet.showSheet()`) so coordinators immediately see their generated schedule.
* **Email Template Editing**: When clicking **`📅 Schedule Assistant > 📧 Email > Edit Email Template`**, the script automatically unhides `EmailTemplate` and jumps directly to it.

#### 2. One-Click "🗂️ Organize & Clean Up Tabs" Tool
Located under **`⚙️ Admin Tools > 🗂️ Tabs > Organize & Clean Up Tabs`**, this tool executes an automated cleanup in seconds:
* **Enforces Workflow Tab Order**:
  $$\text{Q1–Q4} \longrightarrow \text{Active Monthly Tabs (Chronological)} \longrightarrow \text{Workers} \longrightarrow \text{Teams} \longrightarrow \text{Announcements} \longrightarrow \text{Annual Summary} \longrightarrow \text{Unavailability}$$
* **Applies Color-Coded Tab Bottom Borders**:
  * 🟦 **Blue (`#1A73E8`)**: Quarterly master schedules (`Q1`–`Q4`).
  * 🟨 **Gold (`#F9AB00`)**: Active monthly schedules.
  * 🟩 **Green (`#1E8E3E`)**: Worker directories and team rosters (`Workers`, `Teams`).
  * ⬜ **Gray (`#757575`)**: Admin, announcements, summary, and configuration sheets.
* **Automated Past-Month Hiding**: Inspects tab names formatted as `[Month] [Year]` (e.g. `January 2027`). Completed past-month tabs are automatically hidden based on the current calendar date, keeping only active and upcoming monthly tabs visible.
* **Hides Tier 2 & Tier 3 Tabs**: Ensures all config tabs, logs, staging tabs, and backups are tucked away.
* **Anchor Safety**: Automatically activates `Q1` prior to reorganization to guarantee Google Sheets never throws an exception attempting to hide an active sheet.

#### 3. One-Click "👁️ Toggle Admin Config & Log Tabs" Menu Item
Located under **`⚙️ Admin Tools > 🗂️ Tabs > Toggle Admin Config & Log Tabs`**, this tool acts as a single-click switch:
* **Unhide All Admin Tabs**: Reveals all configuration, log, staging, and backup sheets in one click with gray tab indicators so administrators can audit or update tables directly.
* **Hide All Admin Tabs**: If admin tabs are already visible, clicking the toggle hides them all again, returning the spreadsheet to a tidy view for holy work coordinators.

#### 4. Automatic Prior-Year Monthly Tab Purge During Rollover
During annual rollover (`resetForNewYear`), the provisioning engine scans for any generated monthly tabs from prior years (e.g. `January 2026`–`December 2026`) and deletes them automatically, leaving the cloned workbook completely fresh for the new scheduling year.

---

## 8. Performance, Polling & Caching

The Schedule Assistant utilizes an optimized high-concurrency architecture to eliminate spreadsheet lag and support simultaneous coordinator assignments:

### 8.1 Shared Quarter Versioning & Multi-Coordinator Sync
* **Cross-Coordinator Synchronization**: Each quarter sheet (`Q1`–`Q4`) maintains an atomic version token in the shared Script Cache (`q_ver_[Quarter]`).
* **Automated Invalidation**: The version token is automatically bumped upon any grid edit (`onEdit`), sidebar assignment (`assignWorkerExact`), SWR speaker sync (`SyncSWR.js`), RE sync (`SyncRESchedules.js`), or reference roster modification (`Teams`, `Workers`, or `Unavailability` tabs).
* **Zero Redundant Work**: When multiple coordinators work concurrently, their sidebars detect version changes and update automatically without forcing full, expensive sheet reloads.

### 8.2 Chunked Cache Architecture (`CacheService`)
* **100KB Overflow Protection**: Google Apps Script's `CacheService` imposes a strict 100KB limit per cached key. Large datasets (team rosters, unavailability records, and weekday availability) are split into safe 30KB chunks via `cachePutLarge()`, reassembled with integrity verification via `cacheGetLargeJson()`, and cleared via `cacheRemoveLarge()`.
* **Explicit Logging**: Cache write errors are logged to `console.warn` rather than failing silently.

### 8.3 Slim Payload & Browser-Side Quarter Caching
* **Streamlined Payloads**: The server returns compact data payloads. For standard quarterly sheets (`Q1`–`Q4`), redundant arrays (`rowValues`), worker histories, and per-row weekday matrices (`unavailByRow`) are stripped from server transmissions; client-side JavaScript derives worker histories and combines date-specific blackouts with weekly Monday/Friday availability on the fly. For ESSC schedule grids (`isESSC`), the server dynamically computes and delivers `unavailByRow` containing session-specific shift unavailabilities, choir presentation and staging buffers, and cooking team duty window conflicts.
* **In-Memory Quarter Cache**: The sidebar keeps loaded quarters in browser memory for up to 10 minutes (`CHUNK_MAX_AGE_MS = 600000`). Switching back and forth between quarters is instant and requires zero server round-trips.
* **Neighboring Quarter Background Pre-fetch (`prefetchChunk`)**: After loading an active quarter, the sidebar quietly pre-loads adjacent quarters in the background while idle, ensuring instant tab transitions.
* **Full Hard Refresh**: Clicking the **🔄 (Refresh)** button at the top-right of the sidebar clears browser-side caches and forces a full server re-read.

### 8.4 Adaptive Polling & Tab Visibility Optimization
* **Active Polling**: While a coordinator is actively moving the cursor, typing, or making assignments, the sidebar polls at `1.5s` intervals.
* **Idle Back-Off**: After 1 minute of coordinator inactivity, the poll interval smoothly backs off to `3.0s`. After 5 minutes, it relaxes to `5.0s`, drastically reducing Apps Script execution quota consumption. Any keyboard or mouse movement immediately resets the interval to `1.5s`.
* **Background Tab Sleeping**: If the coordinator minimizes or switches browser tabs, polling pauses completely via the HTML5 `visibilitychange` API and resumes immediately when the tab is refocused.

### 8.5 Shorter Script Lock in Assignments (`assignWorkerExact`)
* **Lock Minimization**: In-cell dropdown validation and sanity checks execute *before* acquiring Google Apps Script's `ScriptLock`. The lock is held exclusively for the exact cell write (`setValue`), minimizing lock contention and eliminating timeout errors when multiple coordinators assign workers simultaneously.

### 8.6 Bounded Overrides Audit Log (`MonthlyOverridesLog`)
* **Automatic Batch Trimming**: To prevent sheet bloat and formula recalculation slowdowns on manual edits, the `MonthlyOverridesLog` tab maintains a rolling window of the newest 5,000 entries. When the log exceeds 5,500 rows, it automatically trims the oldest entries in batches.

### 8.7 Load-Balanced Sorting & Quarterly Workload Badges (`Sidebar.html`)
* **Real-Time Workload Pill Badges**: The Available Members list calculates each volunteer's total assignment count in the active quarter from client-side history (`workerHistory[w].length`) and displays a subtle pill badge (e.g. `0x`, `2x`) beside their name.
* **Smart Load-Balanced Sort Toggle**: Coordinators can toggle between **`A–Z`** alphabetical sort and **`Least Assigned`** workload sort via the `#sortToggleBtn` in the Available Members header. `Least Assigned` mode sorts candidates by:
  1. Lowest quarterly assignment count (ascending `0x, 1x, 2x...`).
  2. Proximity fatigue severity (prioritizing workers with zero recent fatigue warnings).
  3. Alphabetical tie-breaker.
  This allows ministry coordinators to easily distribute duties evenly across church members with zero mental calculation.

### 8.8 One-Click Slot Unassignment & Clearing (`Sidebar.html`)
* **Selection Context Clear Button (`✕ Clear`)**: When the active cell has an assigned worker, a prominent `✕ Clear` button appears in the Selection Context card next to `Current: [Worker Name]`, allowing coordinators to unassign the worker in one click.
* **Green Card Toggle**: Clicking the currently assigned worker's green `✓` card in the Available Members list automatically executes an unassignment (`assign("", cardElement)`), cleanly clearing the cell, removing the server-side assignment, and instantly refreshing the UI.

---

## 9. Configuring System Settings

### Modifying Proximity / Fatigue Warnings & Engine Behavior
The Schedule Assistant inspects volunteer service history to alert coordinators about recent service and prevent burnout.

#### How Proximity is Evaluated Technical Architecture:
1. **Client-Side History Construction (`Sidebar.html` - `buildWorkerHistory`):**  
   The sidebar dynamically constructs worker history in the client browser from `rowAssignments` across the entire schedule sheet (all dates and all team columns).
2. **Cross-Team & Cross-Role Scope:**  
   Proximity evaluation is **not** scoped strictly to the current team. It inspects all assignments for a volunteer regardless of team or role. If a member served 3 days ago in Team X, assigning them in Team Y will trigger a 3-day warning displaying their previous duty (`3 days ago (Date - Role/Team)`).
3. **Same-Day Detection:**  
   Assignments on the exact same date but different rows (e.g. Saturday AM vs PM sessions) yield `diffDays === 0` and are surfaced as same-day proximity warnings (`Same day (Date - Role/Team)`).
4. **Color Thresholds (`Sidebar.html`):**
   - 🔴 **Red (`PROXIMITY_RED_DAYS = 1`):** Gap of $\le 1$ day (High Fatigue / Same Day).
   - 🟠 **Orange (`PROXIMITY_ORANGE_DAYS = 7`):** Gap of $\le 7$ days (Medium Fatigue).
   - 🟡 **Yellow (`PROXIMITY_WARNING_DAYS = 14`):** Gap of 8 to 14 days (Moderate Fatigue).

#### Adjusting Configuration Parameters:
- **To change the overall evaluation window (default `14` days):**
  1. Open `Admin_Tools.js` in Apps Script.
  2. In the `CONFIG` object, update `PROXIMITY_WARNING_DAYS: 14` to your desired window (e.g., `10` or `21`).
  3. Also update `var PROXIMITY_WARNING_DAYS = 14;` in `Sidebar.html` (line 247) so client-side filtering matches.
- **To adjust color threshold badges (Red / Orange):**
  1. Open `Sidebar.html` in Apps Script.
  2. Update `PROXIMITY_RED_DAYS` (default `1`) or `PROXIMITY_ORANGE_DAYS` (default `7`) around line 249.
  3. Save the script (💾).

---

## 10. Administrator FAQ & Troubleshooting

### Scenario A: A specific role is pulling the wrong person on the Monthly tab
**Cause**: The Master Column mapping in `MappingConfig` is misaligned with the Quarterly sheet.
1. Check the target Quarterly sheet (e.g., `Q1`) and note the correct column letter for the role.
2. Open the `MappingConfig` tab.
3. Update the `Master Column` letter for that role and Day Type.

### Scenario B: A role maps to the wrong row in the generated Monthly schedule
**Cause**: The Template Row in `MappingConfig` is misaligned.
1. Check the `MonthlyTemplate` tab and note the exact row number for the role.
2. Open the `MappingConfig` tab and update `Template Row` for that role.

### Scenario C: Monthly generation takes ~60–100 seconds
**Context**: This is expected behavior. To preserve merged cells, borders, and complex formatting without triggering Google Sheets API rate-limit errors, the generator writes cell-by-cell.

### Scenario D: Missing assignments for a specific date on the Monthly tab
**Fix**:
1. Verify the date exists in Column A of the target Quarterly sheet (e.g., `Q1`).
2. Ensure the date is formatted as a valid Date or `"M/D"` string.
3. Verify that the placeholder tag (e.g., `{{FRI_1}}`) is present in `MonthlyTemplate`.

### Scenario E: Adding a new weekly role
**Fix**:
1. Add the role row to `MonthlyTemplate`.
2. Open `MappingConfig` and add a new row with the `Day Type`, `Role Name`, `Section`, `Master Column`, and `Template Row`.

### Scenario F: Protecting the `MonthlyTemplate`
**Guidance**: You can freely change fonts, colors, border styles, and column widths in `MonthlyTemplate`. However, **do not delete or move `{{PLACEHOLDER}}` tags** without updating the corresponding row in `MappingConfig`.

### Scenario G: Restoring missing in-cell dropdown arrows
**Fix**: When coordinators cut and paste (`Cmd+X` / `Ctrl+X`), Google Sheets strips validation rules. Run **`⚙️ Admin Tools > 🧰 Repair & Maintenance (Advanced) > Re-apply Data Validation Dropdowns`** to restore all rules across Q1–Q4 in under 5 seconds.

### Scenario H: "Schedule Assistant" menu is missing
**Fix**: The menu is generated by `onOpen()`. Ensure the user is authorized to run scripts. Refresh the browser page. If it still does not appear, check the Apps Script editor for syntax errors or duplicate `onOpen()` declarations.

### Scenario I: Adding new teams to the Sidebar
**Fix**: The sidebar dynamically reads team names from **Row 2** (Category/Team) and **Row 4** (Role Title) of the Quarterly sheets. Add the new role/team into the sheet headers, and the sidebar will recognize it on the next load.

### Scenario J: Coordinator permissions
**Fix**: Coordinators must have **Editor** permissions on the Google Sheet. The first time they open the sidebar, Google will prompt them to authorize the script (`Continue` > `Advanced` > `Go to script`).

---

## 11. Future Enhancement Roadmap

1. **Automated Load-Balancing Suggestions**: An auto-suggest feature in the sidebar that recommends the most suitable available worker based on historical workload and proximity.
2. **Missing Assignment Alerts**: A weekly cron job that scans the upcoming month for unassigned roles and sends alert emails to the relevant coordinators.
3. **Self-Service Volunteer Portal**: A lightweight web portal or Google Form allowing members to submit their own unavailability dates, writing directly into the matrix.
4. **Historical Database Offload**: If historical quarterly tabs eventually slow down spreadsheet calculation, migrate archived years to BigQuery or Cloud SQL while keeping the current active year lightweight.
