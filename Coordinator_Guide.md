# Holy Work Schedule - Coordinator Guide

Welcome to the Holy Work Coordinator’s Guide to the Holy Work Schedule system. This guide is for Ministry Coordinators who are responsible for assigning workers for each quarter’s master schedule (Q1, Q2, Q3, Q4). A custom Assignment Sidebar is embedded directly in the master Google Sheet that makes scheduling fast, prevents double-booking, and provides you with context around the workload of workers on your team.

---

## 🛠 Getting Started: Step-by-Step

### 1. Open the Sidebar
> **🔐 First-Time Authorization:** The very first time you use this tool, Google will ask for permission to run the script. When prompted, click **Continue**, select your Google account, click **Advanced** (at the bottom), and then click **Go to script** to grant access. You only have to do this once!

1. Open the Master Schedule Google Sheet.
2. In the top menu bar, click on **📅 Schedule Assistant**.
3. Select **Open Assignment Sidebar**.
4. The sidebar will open on the right side of your screen. Wait a moment for the status indicator at the top to say **"Synced"**.

### 2. Select a Cell
1. Navigate to the current quarter's tab (e.g., `Q1`).
2. Click on any empty cell where you need to assign a worker (Note: you must click inside the schedule area, row 6 or below).
3. **The sidebar will instantly update** to show you the "Selection Context", including the Target Date, the Role/Slot, and the Team you are scheduling for.

### 3. Assign a Worker
1. Under the **Available Members** section, you will see a list of workers for your team.
2. Simply **click a person's name** to assign them.
3. You will see a brief loading spinner, and then the name will appear in the spreadsheet cell automatically!

### 4. Check Workload Balance
1. Click the **Load Balance** tab in the sidebar.
2. The sidebar displays each member's assignment count for the quarter.
3. You can click on a member to see a breakdown of their assignments by role and date.

---

## 🔄 Two Ways to Schedule: Sidebar vs. In-Cell Dropdowns

The Holy Work Schedule supports two methods for entering assignments. Here is when to use each:

| Method | When to Use | Key Capabilities |
| :--- | :--- | :--- |
| **Schedule Assistant Sidebar** *(Primary)* | **Quarterly planning and multi-role scheduling sessions** on a desktop/laptop computer. | Evaluates team eligibility, detects double-booking, checks the annual Unavailability matrix, warns about volunteer fatigue (14-day proximity), and alerts for same-day Saturday session conflicts. |
| **Native In-Cell Dropdowns** *(Quick Edits / Mobile)* | **Quick single-cell edits, name autocompletion**, or when editing via the **Google Sheets mobile app (iOS / Android)** where Apps Script sidebars cannot open. | Built directly into the spreadsheet grid. Restricts selections to authorized team members and prevents misspellings or unauthorized worker entry. |

> **⚠️ If Dropdown Arrows Disappear (Cut & Paste Warning):**  
> In Google Sheets, using **Cut & Paste (`Cmd+X` / `Ctrl+X`)** permanently strips the native dropdown validation from the source cell. If you notice missing dropdown arrows or accidental entries, notify the administrator to run **`⚙️ Admin Tools > 🧰 Repair & Maintenance (Advanced) > Re-apply Data Validation Dropdowns`**, which restores all ~1,370 dropdown rules across Q1–Q4 in seconds!

---

## ⭐ Smart Features

### Conflict Detection (Double-Booking & Member Unavailability)
The system actively prevents you from accidentally double-booking members or scheduling volunteers who are unavailable:
- **Mondays & Fridays:** If a team member is assigned to *any* role on that day, they will appear as "Unavailable" and be removed from the available list.
- **Saturdays:** A team member will appear as "Unavailable" if they are already assigned during that **same session** (Morning or Afternoon).
- **Annual Unavailability & Vacations:** If a member has indicated unavailability on the annual Holy Work Unavailability sheet for that specific service date, or has general weekday restrictions, the sidebar instantly marks them as `Unavailable (Marked Unavailable)`.
- **ESSC Shift-Level Unavailability & Conflict Badges:** When scheduling for semi-annual convocation events (`Spring_ESSC_Grid` or `Fall_ESSC_Grid`), the sidebar dynamically detects shift-specific blackout requests from the dedicated `Spring ESSC` or `Fall ESSC` sheets, choir presentations, and cooking duties:
  - 🟪 **Choir**: Member is singing in the Choir presentation during this service.
  - 🟦 **Choir Buffer**: Member is singing in the Choir on Saturday AM (Evangelical Service) or Sunday PM (Holy Communion / Choir) and requires the preceding slot for vocal warm-up, robes, and staging.
  - 🟧 **Cooking Duty**: Member is assigned to cooking duty whose team meal prep or cleanup window overlaps with this session (Lunch ~10:30 AM–1:30 PM, Dinner ~4:15 PM–7:00 PM, or concurrent meal session).
  - 🌹 **ESSC Unavail**: Member submitted a blackout request for this specific ESSC shift block (e.g. `WED_EVE`, `SAT_AM`, `SUN_PM`).
- You can click **Conflicts / Unavailable** at the bottom of the sidebar to see exactly *why* someone isn't available (e.g., already serving as Pianist, on cooking duty, singing in choir, or taking personal leave).

### Proximity Warnings (Burnout Prevention & Same-Day Assignments)
To prevent volunteer burnout and accidental over-scheduling, the system tracks each member's service history across a 14-day window (`PROXIMITY_WARNING_DAYS = 14`).

- **Cross-Team & Cross-Role Evaluation:**  
  The 14-day proximity check evaluates **all teams and roles** across the entire schedule sheet—not just your specific team. If a volunteer served 3 days ago in *Ushering*, assigning them today in *Audio* will still trigger a 3-day proximity warning displaying their previous duty context (e.g., `3 days ago (Mar 15 - Usher)`).

- **Color-Coded Proximity Notes in Sidebar:**
  Members who served recently show a `⏱️` line under their name, e.g. `⏱️ 3 days ago (Mar 15 - Usher)` or `⏱️ Same day (Greeter)`:
  - 🔴 Red (`PROXIMITY_RED_DAYS = 1`): 1 day or less
  - 🟠 Orange (`PROXIMITY_ORANGE_DAYS = 7`): 7 days or less
  - 🟡 Amber: 8–14 days
  - ⚪ No note: nothing within 14 days

- **Same-Day, Different Session (Saturdays):** The volunteer stays in **Available Members** with a red `⏱️ Same day (Greeter)` note.

- **Coordinator Judgment:** The sidebar does not block or ask for confirmation. Clicking a name assigns it immediately, so read the note first. To undo, clear the cell.

### Smart Speaker-to-Interpreter Language Filtering
When you click an **Interpreter** cell (`Column C`) in `Q1–Q4`, the Sidebar automatically checks who is assigned as the **Speaker** on that service row, looks up whether that speaker preaches in Chinese (`Sermon Speaker (Chinese)`) or English (`Sermon Speaker (English)`) on the `Workers` tab, and **automatically filters your Sidebar candidate list** to show only **Chinese-to-English (`C-to-E`)** or **English-to-Chinese (`E-to-C`)** interpreters!

---

## 👥 How the `Workers` & `Teams` Tabs and Qualifications Work

> **📘 Full Reference:** See [Workers_Rosters_and_Configuration_Tabs_Guide.md](Workers_Rosters_and_Configuration_Tabs_Guide.md) for the complete guide.

### 1. Clean, Color-Coded Tab Bar
To keep the workbook fast and easy to navigate, all internal configuration and log tabs are automatically hidden. The visible tabs at the bottom of your screen are color-coded by purpose:
* 🟦 **Blue (`Q1`, `Q2`, `Q3`, `Q4`)**: Master quarterly scheduling grids where you assign workers.
* 🟨 **Gold (`January 2027`, etc.)**: Active and upcoming monthly printable schedules (completed past months are automatically hidden).
* 🟩 **Green (`Workers`, `Teams`)**: Master volunteer directory, ministry qualifications, and team rosters.
* ⬜ **Gray (`Announcements`, `Annual Summary`, `Unavailability`)**: Council on Duty schedule, full-year workload statistics, and synced blackout dates.

### 2. How Workers Qualify for Your Team (`Workers` $\rightarrow$ `Teams`)
* The **`Workers`** tab is the single source of truth for all volunteers:
  * **Col C (`Available Mondays`) & Col D (`Available Fridays`)**: Checkboxes indicating whether a member is generally available on Monday or Friday evenings.
  * **Col L (`Active`)**: Must be set to **`Active`**. Members marked **`Inactive`** are automatically hidden from `Teams`, in-cell dropdowns, and the Sidebar.
  * **Cols O–BA (39 Ministry Checkboxes)**: Each column represents a ministry team (e.g., `Hymn Leader`, `Pianist`, `Usher`, `AV1`, `AV2`, `Interpreter (Chinese to English)`, `K`–`Sr`, Fellowships, `Cooking Team 1–8`, and Cleaning teams).
* Checking a ministry box (`TRUE`) for an `Active` member on the `Workers` tab automatically adds their `Schedule Name` (`Col E`) to the corresponding column on the **`Teams`** tab via live `=FILTER(...)` formulas.

### 3. What to Do When Your Ministry Roster Changes
If you need to add a new volunteer to your ministry, remove someone who stepped down, or update a member's availability:
1. **Update or Notify the Information Department**:
   * Either check/uncheck your ministry's checkbox column (`Cols O–BA`) on the **`Workers`** tab (if authorized) or ask the **Information Department** to add/update the member on `Workers`.
   * **Important:** Never delete a member's row from `Workers` if they move away or take a leave—instead, set **Col L (`Active`)** to **`Inactive`** so historical schedule records remain intact.
2. **Click "Refresh Rosters" in the Sidebar**:
   * Once `Workers` is updated, click the **"Refresh Rosters"** link at the top of the Assignment Sidebar. Your Sidebar candidate list will update immediately without reloading the spreadsheet!
3. **Ask Admin / Info Dept to Refresh Dropdowns & Unavailability**:
   * After adding new workers or changing team checkboxes, the Admin or Info Department will run **`Re-apply Data Validation Dropdowns`** (so the `Q1–Q4` in-cell dropdown arrows include the new worker) and **`🔄 Sync Workers to Unavailability Matrix`** (so the worker appears on the external Unavailability sheet).

---

## ⚠️ Important Rules for Synced Columns

Assignments for Sermon Speakers and Religious Education Classes (RE) are automatically imported from the SWR Coordinator and Religious Education class schedules.

Please **do not** manually type names into these specific columns on the Master Schedule. If a change needs to be made (e.g., an RE teacher swaps dates), that change must be made in the original RE classroom spreadsheet or regional SWR spreadsheet. The Master Schedule will automatically update itself during the next sync to reflect the correct data. Any manual edits made directly to the Master Schedule for these roles will be overwritten during the nightly automated sync.

---

## 🔄 Personnel Swaps & Monthly Schedule Auto-Sync

When team members swap duties or make schedule changes:
1. **Always Make Swaps on Master Sheets (`Q1`–`Q4`)**: Always enter swaps or substitutions directly on the quarterly sheets (`Q1`–`Q4`) using the **Schedule Assistant Sidebar** or native in-cell dropdowns. Never edit generated monthly printout tabs directly.
2. **Nightly Automated Synchronization**: Every night at **4:00 AM**, the system runs an automated background task (`autoSyncScheduledMonths`) that re-syncs the current active month and next month's generated schedule tabs directly from `Q1`–`Q4`.
3. **Manual Instant Re-Sync**: If a last-minute swap needs to be reflected on a monthly schedule tab immediately, select **`📅 Schedule Assistant > 🔄 Auto-Sync Active & Next Month`** from the menu.

---

## ⚡ High-Speed Performance & Multi-Coordinator Scheduling

The Schedule Assistant is built for high speed and seamless multi-coordinator collaboration:
* **Live Multi-Coordinator Sync**: When multiple coordinators are scheduling at the same time, any assignment or spreadsheet edit made by someone else automatically syncs to your sidebar without you having to reload the page.
* **Instant Quarter Switching**: Once you open a quarter (e.g., `Q1`), the sidebar quietly pre-loads the adjacent quarters in the background. Navigating between `Q1`, `Q2`, `Q3`, and `Q4` tabs is virtually instantaneous!
* **Smart Sleep & Adaptive Polling**: When you switch browser tabs or step away from your computer, the sidebar automatically pauses or slows its background activity to conserve memory and keep your browser fast. It wakes back up the moment you return.
* **Instant "Refresh Rosters"**: If an administrator adds a new church member, updates team assignments, or imports workers via the Setup Wizard, click the **"Refresh Rosters"** link in the sidebar header. This instantly purges the server-side script cache and re-fetches the newest team rosters.
* **One-Click Hard Refresh (🔄)**: If you ever want to force a clean re-read of the active quarter from the spreadsheet, click the **Refresh (🔄)** button at the top right of the sidebar. This clears all temporary browser memory and fetches fresh data.

---

## ❓ FAQs & Troubleshooting

**Q: An administrator added a new member or updated team assignments, but they aren't appearing in my sidebar candidate list.**
*Fix:* Click the **"Refresh Rosters"** link at the top of the sidebar. This invalidates the server-side roster cache and pulls the newest personnel immediately.

**Q: The sidebar is stuck on "Initializing..." or says "Connection timeout."**
*Fix:* This usually happens if you are logged into multiple Google Accounts at the same time in your browser. Google gets confused about which account is running the script. Try opening the sheet in an "Incognito/Private" window, or log out of all Google accounts except the one you need.

**Q: The sidebar says "Click an assignment cell..." even though I clicked one.**
*Fix:* Ensure you are clicking a valid assignment cell inside the main grid (Row 6 or lower). Clicking on headers, dates, or empty space outside the schedule will show this message. 

**Q: The wrong team is showing up for a column.**
*Fix:* Use the **Select Scheduling Team** dropdown menu at the top of the sidebar to manually override and pick your specific team. The list of available members will update immediately.

**Q: I assigned someone, but the spreadsheet didn't update.**
*Fix:* Check the status dot at the top of the sidebar. If it's red (Error), there might have been a temporary internet disconnect. Click the **Refresh (🔄)** button at the top right of the sidebar to reload fresh data from the server, and try clicking the name again.

**Q: The spreadsheet feels laggy when typing directly into cells.**
*Fix:* If native in-cell typing feels sluggish, notify the administrator to run **`⚙️ Admin Tools > 🧰 Repair & Maintenance (Advanced) > Remove Yellow Cell Highlights (Fix Lag)`**. This strips heavy conditional formatting rules that slow down spreadsheet calculations.

---

## 📅 Scheduling ESSC Events

The semi-annual ESSC schedule workbook (`GG ESSC Schedule.xlsx` / `<Year> GG ESSC Schedule`) is designed for fast, frictionless coordinator data entry with **zero scripts needed**:

1. **Dynamic Dates & Visiting Ministers:**
   - The dates (Column A) and speaker assignments (Column E) are connected dynamically to the `MasterConfig` tab via native spreadsheet formulas.
   - If a visiting minister or event date changes, updating the value in `MasterConfig` updates the schedule grids (`Spring_ESSC_Grid` and `Fall_ESSC_Grid`) immediately in real time.
2. **Assigning Workers via In-Cell Dropdown or Assignment Sidebar:**
   - **In-Cell Dropdown:** Click any cell in columns E–P in `Spring_ESSC_Grid` or `Fall_ESSC_Grid` to select an authorized volunteer from the dropdown.
   - **Assignment Sidebar:** Open the Assignment Sidebar (**📅 Schedule Assistant > Open Assignment Sidebar**). Clicking any scheduling cell instantly displays eligible workers, highlights live conflict badges (`Choir`, `Choir Buffer`, `Cooking Duty`, `ESSC Unavail`), and allows 1-click safe assignment.
   - Support team duties (`Pianist`, `AV1`, `AV2`, `Usher1`, `Usher2`) are **vertically merged across shift blocks**, so you only need to make one selection per block.
3. **Manual Overrides for Speakers:**
   - If a specific session needs a different speaker or local preacher than the default visiting minister, simply pick the preacher from the Column E dropdown. That cell will be overridden without affecting other sessions.
4. **Finalizing & Publishing:**
   - When all slots are filled, notify the church administrator or Information Department. They will trigger the remote schedule generator from the Master Holy Work Schedule to lock and publish the official bilingual presentation sheets (`_Worship` and `_Work`).
