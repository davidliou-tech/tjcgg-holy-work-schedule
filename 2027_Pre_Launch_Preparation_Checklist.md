# 2027 Pre-Launch Preparation Checklist

This checklist outlines the step-by-step operational and technical tasks required to prepare an existing **2027 Holy Work Schedule** ecosystem (Master Holy Work Schedule, Sermon Topic & Hymns, Religious Education, and ESSC workbooks) for the `2027.8.0` Multi-Church Config Engine, Automated Tab Hygiene suite, and Coordinator rollout.

---

## Table of Contents
1. [Phase 1: System Administrator — Technical & Structural Upgrade](#1-phase-1-system-administrator--technical--structural-upgrade)
   - [Step 1A: Push Latest Code & Redeploy Web Apps](#step-1a-push-latest-code--redeploy-web-apps)
   - [Step 1B: Seed & Compile the Multi-Church Config Engine (In-Place)](#step-1b-seed--compile-the-multi-church-config-engine-in-place)
   - [Step 1C: Verify 2027 Quarter Sheet Layout & Reset Assignments](#step-1c-verify-2027-quarter-sheet-layout--reset-assignments)
   - [Step 1D: Re-Link the 3 External Workbooks for 2027](#step-1d-re-link-the-3-external-workbooks-for-2027)
   - [Step 1E: Delete Obsolete Legacy Tabs & Run Tab Hygiene](#step-1e-delete-obsolete-legacy-tabs--run-tab-hygiene)
   - [Step 1F: Activate Background Triggers & Sheet Protections](#step-1f-activate-background-triggers--sheet-protections)
2. [Phase 2: Information Department — Rosters, Communications & Portal Setup](#2-phase-2-information-department--rosters-communications--portal-setup)
   - [Step 2A: Update the 2027 `Workers` & `Teams` Rosters](#step-2a-update-the-2027-workers--teams-rosters)
   - [Step 2B: Configure 2027 Coordinators, Emails & Published Links](#step-2b-configure-2027-coordinators-emails--published-links)
   - [Step 2C: Distribute the 2027 Worker Unavailability Form to Members](#step-2c-distribute-the-2027-worker-unavailability-form-to-members)
3. [Phase 3: Ministry Coordinators — Training & 2027 Assignment Workflow](#3-phase-3-ministry-coordinators--training--2027-assignment-workflow)
   - [Part 3A: What to Cover in Coordinator Training (New for 2027)](#part-3a-what-to-cover-in-coordinator-training-new-for-2027)
   - [Part 3B: Coordinator 2027 Pre-Launch Action Items by Ministry](#part-3b-coordinator-2027-pre-launch-action-items-by-ministry)
4. [Recommended Rollout Sequence & Sign-Off Tracker](#4-recommended-rollout-sequence--sign-off-tracker)

---

## 1. Phase 1: System Administrator — Technical & Structural Upgrade

Because the 2027 workbooks have already been created in Google Drive (prior to the `2027.8.0` Multi-Church Config Engine and Tab Hygiene updates) and have not yet been distributed to coordinators, all upgrades below can be completed **in-place** without disrupting active scheduling.

### Step 1A: Push Latest Code & Redeploy Web Apps
- [ ] **Push Apps Script Code**: Run `clasp push` (or copy the updated `.js` and `.html` files into the Google Apps Script editor) for:
  - **2027 Holy Work Schedule** (Master HWS workbook: `Admin_Tools.js`, `Validation.js`, `GenerateMonthly.js`, `Service_Reminders.js`, `ExportSchedulePDF.js`, `WebApp.js`, `Sidebar.html`, `AutoAssignModal.html`, `UnavailabilityModal.html`, `WebApp_Main.html`, `WebApp_Availability.html`, etc.)
  - **2027 Sermon Topic & Hymns** (SWR workbook: `Sermon_Speaker_Sync.js`)
  - **2027 ESSC Schedule** (ESSC workbook: `GenerateESSC.js`, `ESSC_Sidebar.html`, `ESSC_WebApp.html`)
- [ ] **Redeploy Master HWS Web App**:
  - In the **2027 Master HWS** script editor, navigate to **Deploy > Manage deployments**, edit the active deployment, select **New version**, and click **Deploy**.
  - This updates both the **Holy Work Portal** and the `?page=availability` **Worker Unavailability Form**.
- [ ] **Redeploy ESSC Web App**:
  - In the **2027 ESSC Schedule** script editor, navigate to **Deploy > Manage deployments**, create a **New version**, and click **Deploy**.
- [ ] **Reload Workbooks**:
  - Refresh the browser tab for the **2027 Holy Work Schedule** so `onOpen()` loads the updated **⛪ Holy Work Admin** menu items.

---

### Step 1B: Seed & Compile the Multi-Church Config Engine (In-Place)
> [!NOTE]
> Running **🌱 Seed New Church** (`promptSeedNewChurch()`) operates strictly **in-place** inside the currently open 2027 workbook. It does **not** create a new workbook in Google Drive, nor does it overwrite or erase your `Q1–Q4` schedules, `Teams`, or existing member rows in `Workers`.

- [ ] **Run Seed New Church**:
  - Click **⛪ Holy Work Admin > 🛠️ System Setup & Maintenance > 🌱 Seed New Church**.
  - Verify the confirmation dialog reports: *"Seeded initial configurations (ChurchProfile, ServiceLayoutConfig, RoleConfig) and compiled ValidationConfig & MappingConfig successfully."*
  - **What this step does behind the scenes**:
    1. Runs `ensureWorkersSchema()` to append the 7 canonical columns (`Worker ID`, `Status`, `Teams`, `Qualifications`, `Language_Fluency`, `Family_Cluster_ID`, `Max_Assignments_Per_Month`) to Row 2 of `Workers` if not already present, and generates stable `W-001`, `W-002` IDs.
    2. Creates and seeds `ChurchProfile`, `ServiceLayoutConfig` (Universal 14-Row boundaries: Rows 5–94), and `RoleConfig`.
    3. Runs `compileRoleConfig()` to compile `ValidationConfig` and `MappingConfig` (preserving all multi-worker slots on the same template row) and clears the `'church_config'` script cache.
    4. Automatically hides `ChurchProfile`, `ServiceLayoutConfig`, `RoleConfig`, `ValidationConfig`, and `MappingConfig`.
- [ ] **Inspect `ChurchProfile` & `RoleConfig` Settings**:
  - Click **⛪ Holy Work Admin > 🛠️ System Setup & Maintenance > 👁️ Toggle Admin Config & Log Tabs** (`toggleAdminConfigTabs()`) to temporarily reveal the hidden configuration tabs.
  - Open **`ChurchProfile`** and confirm:
    - `church_name` (e.g., `Garden Grove TJC`)
    - `timezone` (e.g., `America/Los_Angeles`)
    - `admin_emails` (authorized administrator email list)
    - `bible_plan_url` (2027 Daily Bible Reading Plan published CSV URL)
    - `monthly_gen_days_ahead` (default `7`)
  - Open **`RoleConfig`** and verify the **Saturday Cleaning** and **Friday Cleaning** column letters match your 2027 `Q1–Q4` Row 5 headers:
    - By default, `seedRoleConfigSheet()` seeds Saturday Cleaning across `AI..AP` (`Chapel 1–3`, `Trash`, `Bros Down`, `Bros Up`, `Sis Down`, `Sis Up`) and Friday Cleaning across `AD..AK`.
    - If your 51-column 2027 layout splits `Bros Down 1 & 2` (`AM`, `AN`), `Sis Down 1 & 2` (`AO`, `AP`), `Bros Up` (`AQ`), and `Sis Up` (`AR`), update those column letters in `RoleConfig` and run **⛪ Holy Work Admin > 🛠️ System Setup & Maintenance > 🔄 Compile Role Configurations**.

---

### Step 1C: Verify 2027 Quarter Sheet Layout & Reset Assignments
- [ ] **Verify `Q1–Q4` Universal 14-Row Structure & 51-Column Schema**:
  - Confirm `Q1`, `Q2`, `Q3`, and `Q4` use the **Universal 14-Row Service Blocks** (Headers at Rows `5, 20, 35, 50, 65, 80`; Data at Rows `6–19, 21–34, 36–49, 51–64, 66–79, 81–94`) and the **51-column (A–AY)** schema.
  - *Only if your 2027 workbook was cloned from an older variable-row 2026 sheet*:
    - Run **⛪ Holy Work Admin > 🛠️ System Setup & Maintenance > 📐 Restructure Q1-Q4 to Universal 14-Row Blocks** (`promptRestructureQuarterSheetsToUniversal14()`).
    - Run **⛪ Holy Work Admin > 🛠️ System Setup & Maintenance > 🧹 Remove Custom Columns (Cols AZ+)** (`promptRemoveCustomColumns()`).
- [ ] **Populate 2027 Dates & Clear Any Leftover 2026 Assignments**:
  - If 2027 dates are not yet generated or 2026 worker names are still present in `Q1–Q4`, run **⛪ Holy Work Admin > 🛠️ System Setup & Maintenance > 🗓️ Reset Schedule for New Year** (`resetForNewYear()`) and enter `2027`.
- [ ] **Remove Legacy Conditional Formatting Rules**:
  - Run **⛪ Holy Work Admin > 🛠️ System Setup & Maintenance > 🧹 Remove Duplicate Warning Formatting** (`removeDuplicateWarningFormatting()`) so cell conflict highlighting is managed cleanly by the script engine.

---

### Step 1D: Re-Link the 3 External Workbooks for 2027
Because 2027 uses new Google Sheet files with new Spreadsheet IDs, update all cross-workbook connections so 2026 data is not accidentally pulled into 2027:

- [ ] **1. Sermon Topic & Hymns (SWR) $\leftrightarrow$ 2027 Master HWS**:
  - Open the **2027 Sermon Topic & Hymns** workbook:
    - Run **⛪ Holy Work Sync > 🔗 Set Target Schedule Workbook** and paste the **2027 Master HWS URL or Spreadsheet ID**.
    - Run **⛪ Holy Work Sync > ⚡ Enable Auto-Sync Trigger** to activate the real-time edit sync trigger.
  - Open the **2027 Master HWS** workbook:
    - Run **⛪ Holy Work Admin > 🛠️ System Setup & Maintenance > 📖 Update Sermon Topics Speaker Formulas** (`promptUpdateSermonTopicsSpeakerFormulas()`) and enter the **2027 SWR URL or Spreadsheet ID**.
    - Hover over one of the `IMPORTRANGE` cells in `Q1` (e.g., Column B or D) and click **Allow access** in Google Sheets so the cross-workbook formula link is authorized.
- [ ] **2. Religious Education (RE) $\rightarrow$ 2027 Master HWS**:
  - In the **2027 Master HWS** workbook, run **⛪ Holy Work Admin > 🔗 Setup RE Sheets > Connect RE Sheet (Set Source)** for each 2027 RE workbook (`J2`, `E2`, `E1`, `K-1st`, `Pre-K`, plus any custom class keys).
  - Run **⛪ Holy Work Admin > 🔗 Setup RE Sheets > Configure Column Mappings** to verify the column mappings for 2027.
  - Test the connection by running **⛪ Holy Work Admin > 🔄 Sync RE Classes Now**.
- [ ] **3. 2027 ESSC Schedule $\leftrightarrow$ 2027 Master HWS**:
  - Open the **2027 ESSC Schedule** workbook.
  - Run **🌸 ESSC Admin > 🔗 Connect Master HWS Workbook** (`connectMasterHWSWorkbookPrompt()`) and paste the **2027 Master HWS URL or Spreadsheet ID**.

---

### Step 1E: Delete Obsolete Legacy Tabs & Run Tab Hygiene
- [ ] **Delete the 8 Obsolete Legacy Tabs**:
  - If any of the following legacy tabs exist in your 2027 Master HWS workbook, right-click and **Delete** them. They have zero script or formula dependencies, and deleting them prevents `organizeAndCleanUpTabs()` from placing them as visible gray tabs at the far right of the tab bar:
    1. `ESSC_Worship_View`
    2. `ESSC_Work_View`
    3. `ESSC_Data`
    4. `Validation Rules Output`
    5. `Team Emails`
    6. `Email`
    7. `Usher Team Emails`
    8. `Sermon Speaker List`
- [ ] **Run Automated Tab Hygiene**:
  - Click **⛪ Holy Work Admin > 🛠️ System Setup & Maintenance > 🗂️ Organize & Clean Up Tabs** (`organizeAndCleanUpTabs()`).
  - Verify that:
    - All **14 Configuration tabs** (`ChurchProfile`, `ServiceLayoutConfig`, `RoleConfig`, `ValidationConfig`, `MappingConfig`, `MonthlyTemplate`, `EmailSettings`, `EmailTemplate`, `EmailLinks`, `Coordinators`, `PublishedLinks`, `MasterConfig`, `ESSC_MappingConfig`, `Rosters`) and **7 Log/Staging tabs** (`SWR Sync Log`, `RE Sync Log`, `MonthlyGeneratorLog`, `MonthlyOverridesLog`, `Email Log`, `WorkerImport_Staging`, `WorkerImportLog`), plus any `_bak_*` tabs, are hidden.
    - Visible tabs are sorted and color-coded:
      - **Blue (`#1F4E79`)**: `Q1`, `Q2`, `Q3`, `Q4`
      - **Gold (`#F4B400`)**: Active Monthly Schedules (e.g., `January 2027`)
      - **Green (`#0F9D58`)**: `Workers`, `Teams`
      - **Gray (`#757575`)**: `Announcements`, `Annual Summary`, `Unavailability`
  - *(Optional)*: Right-click `Unavailability` $\rightarrow$ **Hide sheet** if you want coordinators to interact with blackout dates exclusively through the **📅 Manage Worker Unavailability** dialog.

---

### Step 1F: Activate Background Triggers & Sheet Protections
- [ ] **Activate System Triggers**:
  - Run **⛪ Holy Work Admin > 🛠️ System Setup & Maintenance > ⚡ Activate All System Triggers** (`activateAllTriggers()`).
  - This installs the 6 core automation triggers:
    1. `onEdit` real-time conflict & qualification validator
    2. Daily Monthly Schedule Generator (`checkAndGenerateNextMonthSchedule`)
    3. Weekly PDF Schedule Distribution (`emailCurrentMonthSchedulePDF`)
    4. Weekly Service Reminders (`sendWeeklyServiceReminders`)
    5. Daily RE Schedule Sync (`syncRESchedules`)
    6. 6-Hour SWR Speaker & Topic Fallback Sync (`syncSWRSpeakers`)
- [ ] **Verify Range & Sheet Protections**:
  - Open **Data > Protect sheets and ranges** and confirm that structural headers (`Q1–Q4` Rows 1–5 and Date Column A) and `MonthlyTemplate` are restricted to Administrators so coordinators only edit assignment cells.

---

## 2. Phase 2: Information Department — Rosters, Communications & Portal Setup

Once Phase 1 is complete, the **Information Department** prepares the 2027 member rosters, team rotations, email templates, and unavailability collection before coordinators begin scheduling.

### Step 2A: Update the 2027 `Workers` & `Teams` Rosters
- [ ] **Update Member Directory (`Workers` Tab)**:
  - Add newly baptized members, transfer members, and newly eligible youth workers.
  - Ensure **Full Name (Column A)**, **Email (Column B)**, and **Status (`Active` / `Inactive`)** are accurate.
  - *Best Practice*: Mark members who moved away or are taking an extended leave as `Inactive` in the `Status` column rather than deleting their row.
  - *(Optional Bulk Import)*: Use **⛪ Holy Work Admin > 👥 Worker Roster Import > 📥 Stage Workers from CSV** to stage, validate, and commit new workers in bulk.
- [ ] **Verify Ministry Qualification Matrix (`Workers` Row 2 Columns)**:
  - Place an `"x"` under each ministry role column that a member is approved to serve in for 2027.
  - **Speaker Language Fluency**: Verify every speaker has an `"x"` in the **Speaker** column and their language fluency code (`E`, `M`, `T`, or combinations like `E, M`) in the **Language** column so the dynamic `Speaker_EL`, `Speaker_ML`, and `Speaker_TL` validation lists populate automatically.
  - **Optional Scheduling Guardrails**: Populate `Family_Cluster_ID` (to schedule family members on the same day) or `Max_Assignments_Per_Month` where applicable.
- [ ] **Update Rotation Groups (`Teams` Tab)**:
  - Update 2027 team rosters (`Team A`, `Team B`, `Team 1`, `Team 2`, etc.) and team email lists for Audio, Video, Security, Dishwashing, Kitchen Cleanup, and Sabbath Lunch.
- [ ] **Refresh Dropdowns & Sync Unavailability Roster**:
  - Run **⛪ Holy Work Admin > 🔄 Apply Validation & Sync (`Workers`/`Teams`)** (`applyAllValidation()`) to rebuild all `Q1–Q4` dropdown menus and speaker language lists.
  - Run **⛪ Holy Work Admin > 🛠️ System Setup & Maintenance > 🔄 Sync Workers to Unavailability** (`promptSyncWorkersToUnavailability()`) so the `Unavailability` tab matches the updated 2027 `Workers` roster.

---

### Step 2B: Configure 2027 Coordinators, Emails & Published Links
- [ ] **Update Coordinator Directory**:
  - Run **⛪ Holy Work Admin > 📧 Set Coordinator Emails** to verify 2027 ministry coordinator names, role groups, and email addresses in the `Coordinators` table (used for automated vacant-role alerts and monthly schedule discrepancy notifications).
- [ ] **Update PDF Distribution List & Email Template**:
  - Run **⛪ Holy Work Admin > 🛠️ System Setup & Maintenance > 📧 Set PDF Distribution List** (`promptSetPdfDistributionList()`).
  - Run **⛪ Holy Work Admin > 🛠️ System Setup & Maintenance > 📝 Edit Email Template** (`promptEditEmailTemplate()`) to update any year-specific text or links for 2027.
  - Run **⛪ Holy Work Admin > 🛠️ System Setup & Maintenance > 👁️ Preview Schedule Email** (`promptPreviewScheduleEmail()`) to send a test preview and verify formatting.
- [ ] **Update Published Links & Announcements**:
  - Click **⛪ Holy Work Admin > 🛠️ System Setup & Maintenance > 👁️ Toggle Admin Config & Log Tabs**, update `PublishedLinks` with any 2027 Google Drive PDF folder links or external schedule URLs shown on the **Holy Work Portal**, and click **👁️ Toggle Admin Config & Log Tabs** again to re-hide config tabs.
  - Add a 2027 welcome message and Q1 scheduling deadline in the **`Announcements`** tab.

---

### Step 2C: Distribute the 2027 Worker Unavailability Form to Members
- [ ] **Collect Member Blackout Dates**:
  - Open **⛪ Holy Work Admin > 📅 Manage Worker Unavailability**, copy the **2027 Self-Service Web Form URL** (`?page=availability`), and distribute it to church members/workers.
  - Announce a deadline for members to submit known 2027 Q1 travel/blackout dates **prior** to coordinator auto-assignment.

---

## 3. Phase 3: Ministry Coordinators — Training & 2027 Assignment Workflow

Once the 2027 rosters and initial blackout dates are ready, conduct a brief walkthrough with Ministry Coordinators and share the 2027 workbook link along with `Coordinator_Guide.md`.

### Part 3A: What to Cover in Coordinator Training (New for 2027)
1. **Clean, Color-Coded Tab Bar**:
   - All internal configuration and log tabs are hidden by default so coordinators can focus on scheduling:
     - **Blue Tabs (`Q1–Q4`)**: Master quarterly planning grids.
     - **Gold Tabs (`January 2027`, etc.)**: Active monthly printable schedules.
     - **Green Tabs (`Workers`, `Teams`)**: Member qualification directory and team rotation rosters.
     - **Gray Tabs (`Announcements`, `Annual Summary`)**: Portal announcements and full-year workload statistics.
2. **Managing Worker Unavailability (`📅 Manage Worker Unavailability`)**:
   - Show coordinators how to open **⛪ Holy Work Admin > 📅 Manage Worker Unavailability** to view, add, or edit blackout date ranges (`YYYY-MM-DD to YYYY-MM-DD`) and share the self-service form link with their ministry workers.
3. **Using the Unified 5-Step Auto-Assign Wizard (`🪄 Auto-Assign Wizard`)**:
   - Walk through **⛪ Holy Work Admin > 🪄 Auto-Assign Wizard**:
     - **Step 1 (Timeframe)**: Choose a Quarter (`Q1`–`Q4`) or Custom Date Range.
     - **Step 2 (Roles)**: Check specific role columns or use a Ministry Preset (*Sermon & Hymns*, *AV & Piano*, *Holy Communion*, *Culinary & Lunch*, *Cleaning & Security*).
     - **Step 3 (Strategy & Guardrails)**: Set Target Fill % (`100%` for full assignment, or partial % to reserve slots for manual trainee pairing) and review active guardrails (unavailability dates, no same-day double booking, max assignments per month, family cluster grouping).
     - **Step 4 & 5 (Dry-Run Preview & Commit)**: Click **Generate Preview** to inspect proposed assignments and fairness distribution counts before clicking **Commit to Sheet**.
4. **Using the Interactive Assignment Sidebar (`📋 Open Assignment Sidebar`)**:
   - Show coordinators how to use **⛪ Holy Work Admin > 📋 Open Assignment Sidebar** when manually assigning or swapping workers. The sidebar displays real-time badges for **Unavailable**, **Conflict (Same-Day Double Booking)**, and **Proximity/Fatigue (Assigned Last Week)**, plus total quarterly assignment counts.
5. **Handling Last-Minute Substitutions on Active Monthly Tabs**:
   - Explain the two ways to handle swaps after a monthly tab (e.g., `January 2027`) has been generated:
     - Edit the cell in `Q1` and run **⛪ Holy Work Admin > 🔄 Sync Monthly Schedule from Quarter**, **or**
     - Edit the cell directly on `January 2027` (which logs the change in `MonthlyOverridesLog` so future syncs will not overwrite it).

---

### Part 3B: Coordinator 2027 Pre-Launch Action Items by Ministry

| Coordinator Group | 2027 Pre-Launch Checklist |
| :--- | :--- |
| **All Ministry Coordinators** | - [ ] Verify their ministry's column in `Workers` (or `Teams`) and notify Info Dept of any roster updates.<br>- [ ] Enter any blackout dates received directly from workers via **⛪ Holy Work Admin > 📅 Manage Worker Unavailability**.<br>- [ ] Run **⛪ Holy Work Admin > 🪄 Auto-Assign Wizard** for **Q1 2027** (or Q1–Q4 for team rotations).<br>- [ ] Review and fine-tune special services (Evangelical Services, Spiritual Convocations, Holy Communion) using **📋 Open Assignment Sidebar**.<br>- [ ] Run **⛪ Holy Work Admin > 📊 Generate Annual Summary** to verify balanced workload distribution. |
| **Sermon & Hymns (SWR) Coordinator** | - [ ] Populate 2027 Q1 sermon speakers, topics, and hymns in the **2027 Sermon Topic & Hymns** workbook.<br>- [ ] Confirm speakers and topics automatically appear in `Q1` Columns B–L of the 2027 Master HWS workbook. |
| **Religious Education (RE) Coordinators** | - [ ] Populate 2027 Q1 teachers, assistants, and lessons in the **2027 RE Workbooks** (`J2`, `E2`, `E1`, `K-1st`, `Pre-K`).<br>- [ ] Run **⛪ Holy Work Admin > 🔄 Sync RE Classes Now** in the Master HWS workbook and verify RE columns in `Q1`. |
| **ESSC Coordinator** | - [ ] Open the **2027 ESSC Schedule** workbook.<br>- [ ] Run **🌸 ESSC Admin > 📅 Set ESSC Dates** and **🎨 Set ESSC Themes** for 2027.<br>- [ ] Verify worker rosters in the `Rosters` tab.<br>- [ ] Run **🌸 ESSC Admin > ✨ Generate Schedule** to create the 2027 ESSC tabs and push assignments to the 2027 Master HWS workbook. |

---

## 4. Recommended Rollout Sequence & Sign-Off Tracker

| Step | Owner | Milestone | Status |
| :---: | :--- | :--- | :---: |
| **1** | **System Admin** | Push updated `.js`/`.html` code and redeploy Web Apps across 2027 Master HWS, SWR, and ESSC workbooks. | [ ] |
| **2** | **System Admin** | Run **🌱 Seed New Church** in-place, verify `ChurchProfile` & `RoleConfig`, and re-link 2027 SWR, RE, and ESSC workbooks. | [ ] |
| **3** | **System Admin** | Delete the 8 obsolete legacy tabs and run **🗂️ Organize & Clean Up Tabs**. | [ ] |
| **4** | **Info Dept** | Finalize 2027 `Workers` & `Teams` rosters, run **🔄 Apply Validation & Sync**, and run **🔄 Sync Workers to Unavailability**. | [ ] |
| **5** | **Info Dept** | Update `Coordinators`, `EmailSettings`, `EmailTemplate`, `PublishedLinks`, and `Announcements`, then distribute the **2027 Worker Unavailability Web Form** to members. | [ ] |
| **6** | **Admin + Coordinators** | Conduct a 20-minute coordinator orientation covering the color-coded tab layout, **📅 Manage Worker Unavailability**, **🪄 Auto-Assign Wizard**, and **📋 Open Assignment Sidebar**. | [ ] |
| **7** | **Coordinators** | Complete Q1 2027 (and annual team rotation) assignments across Master HWS, SWR, RE, and ESSC workbooks. | [ ] |
| **8** | **Admin / Info Dept** | Run **📅 Generate Monthly Schedule** for `January 2027` and confirm **⚡ Activate All System Triggers** is active for automated reminders and weekly PDF distribution. | [ ] |
