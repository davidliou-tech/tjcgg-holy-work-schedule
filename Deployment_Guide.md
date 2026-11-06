# Holy Work Schedule System - Deployment Guide

## Overview
This guide provides a comprehensive roadmap for duplicating and deploying the Holy Work Schedule Assistant for a new church location. By following these steps, you will establish the exact same automated architecture, encompassing the scheduling UI, email reminders, monthly schedule generation, and external data syncing.

---

## System Features Index
This guide outlines the complete process for deploying the Holy Work Schedule Assistant system for a church. By following these steps, you will establish an automated architecture with people in the loop, encompassing the scheduling UI, automated email reminders, monthly schedule generation, and external data syncing. Church Council members and coordinators can utilize this architecture to significantly reduce manual administrative overhead and eliminate scheduling errors.

### 1. Coordinator Scheduling Assistant (Sidebar UI)
A custom Google Sheets sidebar designed to replace manual typing with smart, context-aware assignments.
* **Smart Filtering:** Suggests only the workers assigned to the specific role you are clicking on (derived from the `Teams` tab).
* **Double-Booking Prevention:** Instantly warns if a member is already scheduled for another concurrent duty on the exact same date.
* **Fatigue & Proximity Warnings:** Tracks a member's recent assignments and issues a warning if they are scheduled too closely to a previous duty (e.g., within 14 days), helping distribute the workload evenly.
* **Absence Detection:** Automatically cross-references the `Unavailability` tab and warns coordinators if they try to assign someone who is marked as unavailable on that specific day.

### 2. Automated Weekly Email Reminders
A background time-driven script that prevents "no-shows" by automatically reminding volunteers.
* **Look-Ahead Logic:** Scans the schedule (typically 5 days in advance) to find upcoming services on Mondays, Fridays, and Saturdays.
* **Dynamic Routing:** Matches worker names with their contact info in the `Workers` tab and dispatches a clean, HTML-formatted summary of their upcoming duties.

### 3. Monthly Schedule & PDF Generator
Automates the tedious process of formatting quarterly data for public distribution.
* **Template Mapping:** Uses the `MappingConfig` tab to seamlessly transpose horizontal quarterly data into a clean, vertically-oriented printable layout (`MonthlyTemplate`).
* **One-Click Generation & Syncing:** Creates a perfectly formatted new tab for the requested month. If the month already exists, it can safely overwrite and sync the latest updates.
* **Auto-Locking:** The generated sheets are automatically locked to prevent unauthorized edits.

### 4. Data Synchronizations (Internal & External)
Keeps the local church schedule perfectly synchronized without duplicate data entry.
* **External - SWR Speaker Sync:** Nightly automated sync that pulls regional guest speakers from the external SWR Master Sheet and injects them into the local schedule.
* **External - Unavailability Matrix Sync:** Nightly automated sync that flattens checked boxes from the external annual unavailability spreadsheet into the local `Unavailability` tab for instant Sidebar absence detection.
* **External - R.E. Schedule Sync:** Dynamically detects columns ("Morning", "Afternoon", "Teacher") in individual external R.E. Class spreadsheets and aggregates them into the Master Schedule.
* **External - Bible Reading Plan Sync:** Dynamically pulls the daily Bible reading chapters from an external spreadsheet and injects them into the monthly schedule.
* **Internal - Council on Duty:** Automatically pulls the assigned Council members from the internal `Announcements` tab during monthly generation.
* **Internal - Master to Monthly Sync:** Transposes the horizontal quarterly data from Q1-Q4 into the vertical `MonthlyTemplate`.

### 5. Annual Summary & Reporting
A powerful analytics tool for church leadership.
* **Load Distribution Analysis:** Parses the entire year (Q1-Q4) to calculate a comprehensive frequency count of how many times each member served in each role. 
* **End-of-Year Reviews:** Essential for reviewing volunteer commitment, finding gaps in teams, and making informed decisions for the next year's rostering.

### 6. System Integrity & Maintenance Tools
Built-in utilities for the System Administrator to ensure the spreadsheet doesn't break.
* **Master Audit:** Scans the entire document for overlapping/duplicate assignments that might have bypassed the sidebar and highlights them for correction.
* **Automated Data Validation:** A menu option to re-apply dropdown data validation rules across all cells based on the most up-to-date team rosters, ensuring data consistency even when users type manually.
* **One-Click Annual Provisioner:** Automatically locates or creates the year folder in Google Drive (e.g. `2027`), clones both the Master Schedule and Unavailability sheets, updates anchor dates (Q1-Q4 and B2), clears assignments/checkboxes, synchronizes worker rosters, and links the sheets.
* **One-Click Trigger Activation:** Programmatically configures all background time-driven triggers (Reminders, SWR Sync, Unavailability Sync) in the new year's sheet without manual setup.

---


## Phase 1: Foundational Spreadsheet Setup
The entire system depends on a rigid spreadsheet architecture. Before migrating any code, you must create a Master Schedule Google Sheet with the following mandatory tabs:

### 1. Quarterly Tabs (`Q1`, `Q2`, `Q3`, `Q4`)
* **Purpose:** The source of truth for all scheduling data.
* **Structure:** Standardized on the **Universal 14-Row Architecture (Rows 1–94)** with six fixed 14-row service sections:
  * **Rows 5–19**: Mondays (Row 5 Header, Rows 6–19 data)
  * **Rows 20–34**: Tuesdays (Row 20 Header, Rows 21–34 data)
  * **Rows 35–49**: Wednesdays (Row 35 Header, Rows 36–49 data)
  * **Rows 50–64**: Fridays (Row 50 Header, Rows 51–64 data)
  * **Rows 65–79**: Saturdays AM (Row 65 Header, Rows 66–79 data)
  * **Rows 80–94**: Saturdays PM (Row 80 Header, Rows 81–94 data)
* **Automated Setup**: Can be automatically applied or restored at any time via **⚙️ Admin Tools > 🧰 Repair & Maintenance (Advanced) > 🛠️ Restructure Q1-Q4 (Universal 14-Row)**. Dark gray headers (`#D9D9D9`), centered `'X'` cells, white active cells (`#FFFFFF`), borders, and dynamic self-terminating date formulas are applied automatically.

### 2. Roster Tabs
* **`Teams`:** Defines which workers belong to which roles. Row 1 contains Team names (e.g., A/V, Usher), and Rows 2+ list the workers for that team.
* **`Workers`:** The master directory of all volunteers. It **must** include columns for Name and Email Address (crucial for the email reminder system to function).
* **`Unavailability`:** A log of specific dates when volunteers cannot serve.
  * **Architecture:** Formatted as the **Universal 14-Column Matrix Architecture (56 date columns, B–BE)** corresponding exactly to the Master Schedule's 14-row quarterly sections (14 columns each for Q1, Q2, Q3, and Q4).
  * **Dynamic Date Headers:** Uses self-terminating, quarter-bounded formulas linked directly to the master schedule (`=IF(Q1!A6="","",TEXT(Q1!A6,"M/D"))`). Inactive 13th/14th week columns blank automatically.
  * **Cell A1 Filter:** Includes an interactive worker search dropdown in cell `A1`.
  * **Automated Setup:** Can be automatically structured or repaired at any time via **⚙️ Admin Tools > 🔄 Sync Data > Sync Workers to Unavailability Matrix**.

### 3. Monthly Generation Tabs
* **`MonthlyTemplate`:** A highly stylized layout of what the final printed schedule should look like. It must contain exact `{{PLACEHOLDER}}` tags (e.g., `{{SAT_1}}`, `{{WEEK_1}}`, `{{MONTH_YEAR}}`) where dynamic data will be injected.
* **`MappingConfig`:** The "brain" linking the Q-sheets to the Monthly Template. It maps specific roles (by their Master Column letter in the Q-sheet) to the exact Template Row number in the `MonthlyTemplate`.


### 4. Administrative Tabs
* **`Announcements`:** Contains monthly information such as the Council on Duty. The Monthly Generator script will look at the row corresponding to the month number (e.g., Row 8 for August) and pull the Primary Council member from Column C and the Backup from Column D.

---

## Phase 2: Copying the Codebase (Apps Script)
Once the spreadsheet structure is laid out, the automation code must be ported over.

1. In your new Google Sheet, go to **Extensions > Apps Script**.
2. Create the following files in the editor and copy their respective code from the project source repository into them:
   * `Admin_Tools.js` (Core admin engine: annual provisioner, setup wizard endpoints, menu handlers, SWR sync, and audit tools)
   * `AppConfig.js` (Centralized application configuration, design tokens, and default teams catalog)
   * `SetupWizard.html` (Interactive 3-step Visual Column Mapper & diff review modal UI)
   * `Sidebar.html` (The interactive Coordinator Assignment Assistant UI with instant roster refresh)
   * `GenerateMonthly.js` (The Monthly Printout & PDF Generator dynamic engine)
   * `Service_Reminders.js` (Automated HTML email reminder engine with dynamic column schema and bilingual formatting)
   * `Validation.js` (High-speed batch data validation dropdown applicator with source caching)
   * `SyncSWR.js` (Regional SWR guest speaker sync engine)
   * `SyncRESchedules.js` (Religious Education class schedules sync)
   * `ExportSchedulePDF.js` (Monthly PDF export & email distribution engine)
   * `GenerateESSC.js` (Master HWS control tower engine: manages standalone ESSC workbook remotely, pushes dropdowns, injects dynamic formulas, and generates presentation tabs)
   * `appsscript.json` (Project manifest specifying V8 runtime, time zone, and explicit OAuth scopes)

3. **Configure Project Manifest (`appsscript.json`):**
   In the Apps Script editor, click **Project Settings** (gear icon) and check **"Show 'appsscript.json' manifest file in editor"**. Ensure the `oauthScopes` array includes:
   ```json
   "oauthScopes": [
     "https://www.googleapis.com/auth/spreadsheets",
     "https://www.googleapis.com/auth/drive",
     "https://www.googleapis.com/auth/script.send_mail",
     "https://www.googleapis.com/auth/script.scriptapp",
     "https://www.googleapis.com/auth/script.external_request",
     "https://www.googleapis.com/auth/userinfo.email"
   ]
   ```

4. **Initialize Church Configuration (Config-as-Sheets):**
   Starting in version `2027.8.0`, church-specific settings are decoupled from script code and stored in configuration sheet tabs:
   * **Automated Seeding**: In the Google Sheet top menu, click **⚙️ Admin Tools > 🌱 Seed New Church**. This automatically creates and populates:
     - **`ChurchProfile`**: Set church name, code, contact email, timezone, and SWR location tag.
     - **`ServiceLayoutConfig`**: Sets service days, header rows, and 14-row section boundaries.
     - **`RoleConfig`**: Sets assignable roles, target Q columns, and required skills.
   * **Worker Onboarding**: Click **⚙️ Admin Tools > 🚀 Open Setup Wizard** to ingest your church's roster via the 3-step Visual Column Mapper (supporting split names, Garden Grove's 34+ standard teams, matrix checkboxes, and custom teams).
   * **Compile Roles**: Click **⚙️ Admin Tools > 🔄 Sync Data > Compile Role Configurations** to automatically generate `ValidationConfig` and `MappingConfig`.
   * **Drive Storage Folder**: Set `HOLY_WORK_SCHEDULES_FOLDER_ID` in `Admin_Tools.js` (or via Script Properties) to the Google Drive folder ID for annual archive rollovers.

---

## Phase 3: Activating the Features (Deployment Sequence)

### 1. Coordinator Scheduling Assistant
* **How it works:** Coordinators open a sidebar that reads the `Teams`, `Unavailability`, and the current schedule to suggest valid names, prevent double-booking, and issue fatigue warnings.
* **Activation:** The system is inherently active once the code is pasted and the `onOpen` function runs. When a coordinator clicks **📅 Schedule Assistant > Open Assignment Sidebar**, they will be prompted to grant Google authorization.
* **Validation:** Test the setup by clicking a cell in `Q1`. Ensure the sidebar populates the correct team members and successfully issues a warning if you try to schedule someone back-to-back.

### 2. Syncing with External Sources (SWR, Unavailability, RE & Bible Reading)
* **How it works:** Automates data entry by pulling assigned guest speakers from a regional sheet, worker absences from the annual unavailability matrix, teachers from individual classroom sheets, and daily Bible reading chapters.
* **Activation (SWR Speaker):** Click **⚙️ Admin Tools > ⏰ Automation > Setup Daily SWR Sync URL** and paste the URL of the regional SWR spreadsheet.
  > **⚠️ Annual Update Note:** The regional SWR spreadsheet is updated and published under a brand new URL every calendar year. Whenever a new year's schedule is provisioned, the administrator must rerun this setup step in the new year's spreadsheet once the regional church publishes the new SWR URL.
* **Activation (Unavailability Matrix):** Click **⚙️ Admin Tools > 🔗 Connections > Link Unavailability Sheet** and paste the URL or ID of the annual Holy Work Unavailability spreadsheet. *(Note: This linking step is only needed during initial church deployment. In future years, running `🚀 Provision New Year` automatically duplicates, resets, and links the new year's unavailability sheet without manual intervention!)*
* **Activation (Sermon Topics & Hymns):** Click **⚙️ Admin Tools > 🔗 Connections > Link Sermon Topics & Hymns Sheet** and paste the URL or ID of the annual Sermon Topics & Hymns spreadsheet. *(Note: Only needed during initial church deployment. In future years, running `🚀 Provision New Year` automatically clones the Sermon Topics sheet, sets public edit permissions, updates IMPORTRANGE formulas, clears prior year speaker entries in Column B while preserving 'X' and gray fills across ESSC tabs, and dynamically links it into email reminders).*
* **Activation (RE Classes):** Open `SyncRESchedules.js` and locate the `RE_CLASSES_CONFIG` object. Replace the `spreadsheetId` values with the unique Google Sheet IDs of the RE class spreadsheets. The script utilizes dynamic column detection, so as long as their sheets have recognizable headers like "Teacher" and "Morning", it will map automatically.
* **Activation (Bible Reading Plan):** Open `GenerateMonthly.js` and locate the `BIBLE_PLAN_URL` variable at the top of the file under Configuration. Replace this string with the URL of the annual Bible Reading Plan Google Sheet.
* **Triggering RE Sync:** `autoSyncRESchedules` runs daily at 3:00 AM after **⚙️ Admin Tools > ⏰ Automation > Activate All Automation Triggers**.

### 3. Re-Applying Data Validation Dropdowns (`Validation.js`)
* **How it works:** Injects native Google Sheets `VALUE_IN_RANGE` in-cell dropdown arrows across all quarterly assignment cells (~1,370 cells across Q1–Q4) in under 5 seconds.
* **Why this is critical:** While desktop users utilize the Schedule Assistant Sidebar, mobile users (iOS/Android) cannot open Apps Script sidebars and rely entirely on native in-cell dropdowns. Furthermore, Google Sheets strips data validation whenever a user cuts and pastes cells (`Cmd+X` / `Ctrl+X`).
* **Activation:** Click **⚙️ Admin Tools > 🧰 Repair & Maintenance (Advanced) > Re-apply Data Validation Dropdowns** immediately following deployment or rollover to ensure all dropdown rules are intact.

### 4. One-Click Automated Trigger Activation (Reminders & Daily Syncs)
* **How it works:** The system programmatically registers all recurring background jobs in a single click without manually navigating the Apps Script dashboard.
* **Why this must be run annually:** Google Apps Script deliberately disables and clears all time-driven project triggers whenever a spreadsheet is cloned (to prevent duplicate background emails). Therefore, **this step must be run once upon initial deployment, and once every year inside the newly provisioned Holy Work Schedule**.
* **Activation:** Open the spreadsheet and click **⚙️ Admin Tools > ⏰ Automation > Activate All Automation Triggers**. This automatically schedules:
  1. **Weekly Email Reminders:** Runs `sendAssignmentReminders` daily at **8:00 AM** (dispatches reminders 5 days prior at 8:00 AM for Monday, Friday, and Saturday services).
  2. **SWR Regional Speaker Sync:** Runs `syncSWRSpeakers` nightly between 1:00 AM – 2:00 AM.
  3. **Unavailability Matrix Sync:** Runs `syncUnavailabilityFromMatrix` nightly between 2:00 AM – 3:00 AM.

### 5. Monthly Schedule Generation
* **How it works:** Reads the user's requested MM/YYYY, matches the dates in the Q-sheets, cross-references the `MappingConfig`, injects the Bible Reading Plan from an external URL, and pastes everything into a fresh copy of the `MonthlyTemplate`. The sheet is then locked to authorized editors.
* **Activation:** Run directly via the custom menu **📅 Schedule Assistant > Generate Monthly Schedule**.
* **Automated Sync Trigger (Optional):** In Apps Script > Triggers, you can set a time-driven Daily trigger for the `autoSyncScheduledMonths` function to keep the current and next month's generated tabs perfectly synced with the master Q-sheets in the background.
* **Validation:** Check the auto-generated `MonthlyGeneratorLog` tab to ensure no columns were missed and that the output PDF format is visually correct. Ensure the `Announcements` tab has the correct Council on Duty listed so it auto-populates.

### 6. ESSC Scheduling Workbook (Unified Push Architecture & Dynamic Formulas)
The standalone ESSC workbook (`GG ESSC Schedule.xlsx` / `<Year> GG ESSC Schedule`) manages Semi-Annual Spiritual Convocation and Evangelical Services via an independent horizontal grid, driven by a **centralized push pipeline from the Master HWS Control Tower** and **real-time dynamic formulas**:
1. **Zero Script Overhead:** The standalone ESSC workbook contains **no separate Google Apps Script project**. All administrative commands, roster pushes, formula injections, and presentation generation are executed directly from the Master HWS under **`📅 Schedule Assistant > 🎪 ESSC Schedule Generator`**.
2. **Connect Master HWS to ESSC Workbook:**
   * In Master HWS, click **📅 Schedule Assistant > 🎪 ESSC Schedule Generator > 🔗 Connect Standalone ESSC Workbook...**
   * Paste the Google Spreadsheet ID or URL of that year's ESSC workbook (e.g., `2027 GG ESSC Schedule`). (Note: During annual rollover, `🚀 Provision New Year` automatically creates and binds this link).
3. **Configure Event Dates & Themes:**
   * In Master HWS, run **📅 Schedule Assistant > 🎪 ESSC Schedule Generator > 📅 Set ESSC Event Dates...** to input Wednesday start dates (e.g. `2027-04-14` for Spring and `2027-10-13` for Fall).
   * Or directly enter the start dates and visiting ministers into `MasterConfig` inside the ESSC workbook.
4. **Push Dropdowns & Arm Dynamic Formulas:**
   * In Master HWS, click **📅 Schedule Assistant > 🎪 ESSC Schedule Generator > 🔄 Push Dropdowns & Resolve Placeholders**.
   * **What this executes automatically across workbooks:**
     * **Roster Ingestion:** Pushes active worker teams from Master HWS (`Teams` / `Workers`) into the ESSC `Rosters` tab.
     * **Dynamic Date Formulas (Column A):** Injects native calendar math formulas (`=IFERROR(DATEVALUE(TO_TEXT(XLOOKUP(...)))), "")`, `=A$2+1`, `=A$2+2`, etc.). Any future edit to the Wednesday start date in `MasterConfig` updates all 45 date rows instantly without running any script!
     * **Dynamic Speaker Formulas (Column E):** Injects native `=IFERROR(XLOOKUP("SPRING_VISITING_MINISTER", MasterConfig!$A:$A, MasterConfig!$B:$B, ...))` formulas for visiting, resident, and lay ministers. Changing ministers in `MasterConfig` updates the schedule grid in real time with zero scripts!
     * **Minister Whitelisting:** Adds visiting ministers to the `Speakers` dropdown validation list, eliminating red invalid warning triangles.
     * **Meal Row Cleanup:** Suppresses dropdown chips on empty meal/transition rows.
5. **Staff Assignments & Manual Overrides:**
   * Coordinators edit `Spring_ESSC_Grid` or `Fall_ESSC_Grid` directly. Selecting a worker from the Column E dropdown overrides the formula for that specific slot while leaving all other formula slots intact.
   * Support team duties (`Pianist`, `AV1`, `AV2`, `Usher1`, `Usher2`) are vertically merged across shift blocks for intuitive selection.
6. **Generate & Lock Presentation Sheets (Remote Control):**
   * In Master HWS, run **📅 Schedule Assistant > 🎪 ESSC Schedule Generator > 🚀 Generate Spring ESSC Schedule (Remote)** (or Fall).
   * Generates formatted, bilingual presentation sheets (`Spring_ESSC_Worship`, `Spring_ESSC_Work`) with church logos in the ESSC workbook, and automatically locks them (`LOCKED 🔒`) against accidental edits.

---

## Phase 4: Annual Rollover & Drive Architecture

When provisioning future calendar years (e.g., 2028, 2029):
1. **Google Drive Folder Hierarchy:**
   The Master Schedule parent folder (`HOLY_WORK_SCHEDULES_FOLDER_ID`) hosts nested annual folders (`2027/`, `2028/`, `2029/`).
2. **Automated Subfolder Creation:**
   When running **⚙️ Admin Tools > 🚀 Provision New Year**, the provisioner automatically searches the parent folder for a subfolder matching the target year (e.g. `2028`). If it exists, it uses it; if not, it automatically creates the subfolder and nests all cloned spreadsheets (`2028 Holy Work Schedule`, `2028 Holy Work Unavailability`, and `2028 Sermon Topic & Hymns`) directly inside.
3. **Post-Provisioning Checklist:**
   * Open the newly provisioned Master Schedule in the new year's folder.
   * Run **⚙️ Admin Tools > 🧰 Repair & Maintenance (Advanced) > Re-apply Data Validation Dropdowns**.
   * Run **⚙️ Admin Tools > ⏰ Automation > Activate All Automation Triggers**.
   * Update the SWR URL once published via **⚙️ Admin Tools > ⏰ Automation > Setup Daily SWR Sync URL**.
