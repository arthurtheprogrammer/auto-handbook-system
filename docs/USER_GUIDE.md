# User Guide — Auto Handbook System

A guide for team members who manage the data sources and run the system. No coding knowledge required.

---

## Table of Contents

- [First-Time Setup (Excel Trust & Calculation Settings)](#first-time-setup-excel-trust--calculation-settings)
- [How It Works (Simple Version)](#how-it-works-simple-version)
- [Running the System](#running-the-system)
- [What You Need to Maintain](#what-you-need-to-maintain)
- [Data Source Reference](#data-source-reference)
- [Dashboard Parameters](#dashboard-parameters)
- [Understanding the Output](#understanding-the-output)
- [Refreshing Lecturer Data](#refreshing-lecturer-data)
- [Common Issues & Fixes](#common-issues--fixes)

---

## First-Time Setup (Excel Trust & Calculation Settings)

Before using the system for the first time, you need to configure three settings in Excel. These only need to be done **once per computer**.

> [!IMPORTANT]
> All three settings below are **required**. Without them, macros won't run, enrolment numbers won't load, and calculations may not update.

---

### 1. Enable VBA Macros (Trust Center)

This allows the system's macros to run — including the **LecturerRefresh** module that is embedded in the exported calculation file.

#### Windows

1. Open Excel
2. Go to **File** → **Options**
3. In the left sidebar, click **Trust Center**
4. Click the **Trust Center Settings...** button
5. In the left sidebar, click **Macro Settings**
6. Select **Enable VBA macros (not recommended; potentially dangerous code can run)**
   - Alternatively, select **Disable VBA macros except digitally signed macros** if your organisation signs macros
7. Also tick **Trust access to the VBA project object model** (required for exporting the LecturerRefresh module into the calculation file)
8. Click **OK** → **OK**

#### Mac

1. Open Excel
2. Go to **Excel** (menu bar) → **Preferences**
3. Click **Security & Privacy**
4. Under **Macro Security**, select **Enable all macros**
5. Close the preferences window

---

### 2. Enable External Links & Data Connections (Trust Center)

This allows the workbook to pull enrolment numbers from the Enrolment Tracker file via external connection formulas (e.g., linked workbook references).

#### Windows

1. Open Excel
2. Go to **File** → **Options**
3. In the left sidebar, click **Trust Center**
4. Click the **Trust Center Settings...** button
5. In the left sidebar, click **External Content**
6. Under **Security settings for Workbook Links**, select **Enable automatic update for all Workbook Links**
7. Under **Security settings for Data Connections**, select **Enable all Data Connections**
8. Click **OK** → **OK**

#### Mac

1. Open the workbook
2. If you see a **Security Warning** bar at the top saying "Automatic update of links has been disabled", click **Enable Content**
3. If prompted about data connections, click **Enable**
4. For permanent trust: go to **Excel** → **Preferences** → **Security & Privacy** and ensure external content is allowed

> [!TIP]
> If enrolment numbers show as `0` or `#REF!` after opening, the link trust settings might not be enabled. Go to **Data** → **Edit Links** and click **Update Values** to force a refresh.

---

### 3. Set Calculation Mode to Automatic

The system requires automatic calculation so that formulas (including enrolment lookups and elapsed-time tracking) update in real time.

#### Windows

1. Open Excel
2. Go to **Formulas** tab on the ribbon
3. Click **Calculation Options** (in the Calculation group)
4. Select **Automatic**

Or via settings:
1. Go to **File** → **Options** → **Formulas**
2. Under **Calculation options**, set **Workbook Calculation** to **Automatic**
3. Click **OK**

#### Mac

1. Open Excel
2. Go to **Formulas** tab on the ribbon
3. Click **Calculation Options**
4. Select **Automatic**

Or via preferences:
1. Go to **Excel** → **Preferences** → **Calculation**
2. Under **Calculation**, select **Automatically**
3. Close the preferences window

> [!NOTE]
> The system sets calculation to Automatic, but good to check for troubleshooting if you notice the numbers are not updating.

---

## How It Works (Simple Version)

Think of this system as a **data assembly line**:

```
📂 Enrolment Tracker  ─┐
                        ├──→  🤖 System processes  ──→  📊 Calculation Spreadsheet
📂 Teaching Matrix    ─┘      everything for you         (ready to use)
🌐 Handbook Website   ─┘
```

1. **You provide**: Year and filenames on the Dashboard
2. **The system fetches**: Subject lists, staff allocations, and assessment details automatically
3. **You receive**: A complete marking & admin support calculation spreadsheet

---

## Running the System

### Before You Start

Make sure these files exist on SharePoint:

| What | Where |
|------|-------|
| **Enrolment Tracker** (`.xlsx`) | `TEACHING MATRIX & ENROLMENT TRACKER` folder |
| **Teaching Matrix** (`.xlsx`) | `TEACHING MATRIX & ENROLMENT TRACKER` folder |
| **Automated Handbook Data System** (`.xlsm`) | `TEACHING SUPPORT > Handbook > Auto Handbook System` folder |

### Step-by-Step

1. **Open** the `Automated Handbook Data System.xlsm` workbook from SharePoint
2. **Go to** the `Dashboard` sheet
3. **Fill in** the required fields:

   | Cell | What to Enter | Example |
   |------|--------------|---------|
   | **C2** | The year | `2026` |
   | **C3** | Enrolment Tracker filename (leave blank for default) | `2026 Enrolment Tracker.xlsx` |
   | **C5** | Teaching Matrix filename (leave blank for default) | `2026 Teaching Matrix.xlsx` |
   | **C8** | Word count benchmark (words/hr) | `3000` |
   | **C9** | Exam benchmark (exams/hr) | `3` |
   | **C10** | Marking support benchmark (hrs/stream) | `20` |
   | **C12** | Email for notification (optional) | `you@unimelb.edu.au` |

4. **Click the Run button** (or run macro `GenerateMarkingSupport` from the Developer tab)
5. **Wait** — the process takes approximately 5–10 minutes. You'll see status updates in column F:

   | Cell | Step | Status |
   |------|------|--------|
   | F2 | Subject List Refresh | Running... → Complete |
   | F3 | Generate Subject Queries | Running... → Complete |
   | F4 | Parse Assessment Data | Running... → Complete |
   | F5 | Teaching Stream Refresh | Running... → Complete |
   | F6 | Generate Calculation Sheets | Running... → Complete |

6. **Done!** The exported file appears in the same SharePoint folder as the source workbook

> **Tip**: If you need to stop the process, run `StopWorkflowMonitoring` from the macro menu.

---

## What You Need to Maintain

### ✅ Things You Should Keep Updated

| Item | Why | How Often |
|------|-----|-----------|
| **Enrolment Tracker** file | Provides subject codes and enrolments | Each semester |
| **Teaching Matrix** file | Provides lecturer assignments | Each semester |
| **Year in C2** | Used to fetch the correct handbook year | Each year |
| **Benchmark values** (C8, C9, C10) | Used in marking hour calculations | When policy changes |

### ⚠️ Things You Should NOT Change

| Item | Why |
|------|-----|
| Sheet names (`SubjectList`, `teaching stream`, `AllSubjectsHTML`, `assessment data parsed`) | The macros look for these exact names |
| Table names in the **source workbook** (`subject_list`, `teaching_stream`, `AllSubjectsHTML`, `progress_bar`) | The macros look for these exact names |
| Table names in the **Enrolment Tracker** (`Enrolment_Tracker`) | Power Automate reads from this table |
| Table names in the **Teaching Matrix** (`Teaching_Data`, `Staff_table`) | Power Automate reads from these tables |
| Table name `Enrolment_Number` in the **Enrolment Tracker** | External connection formula in exported calculation sheets pulls enrolment data from this table |
| Column headers in any table | Changing headers will break the data parsing |
| The `progress_bar` table on Dashboard | Used to track workflow completion |
| File paths in the VBA code | These point to your SharePoint folders |

---

## Data Source Reference

### Source Files (What You Maintain)

These are the SharePoint files you manage. The system reads specific columns from each — columns not listed below are ignored.

> [!IMPORTANT]
> Some column headers contain "DO NOT SORT" or have line breaks in them. **Do not add or remove line breaks from column headers** — the scripts normalise whitespace automatically, but changing column names entirely might break the system.

#### Enrolment Tracker (`.xlsx`)

The system reads these columns (matched by keyword, not exact name):

| Keyword Matched | Example Column Header | What the Script Extracts |
|----------------|----------------------|------------------------|
| `Subject Code` | Subject Code | Subject code (e.g., `MGMT10101`) |
| `Subject Name` | Subject Name | Full subject name |
| `Subject Coordinator` | Subject Coordinator | Coordinator name |
| `Status` | Status | Active/Suspended status |
| `Study Period` | Study Period | e.g., `Semester 1`, `Semester 2`, `Summer Term` |
| `Grouped Period` | Grouped Period | `FHY` or `SHY` |
| `Delivery Mode` | Delivery Mode | e.g., `On Campus`, `Online`, `Offshore` |

Other columns (Quota/Cap, enrolment numbers, predictions, program breakdowns, etc.) are **not read** by the subject list parser.

#### Teaching Matrix — Teaching Data Sheet

| Keyword Matched | Example Column Header | What the Script Extracts |
|----------------|----------------------|------------------------|
| `Subject Code` | Subject Code DO NOT SORT | Subject code |
| `Study Period` | Study Period | e.g., `Semester 1` |
| `Lecturer` | Lecturer DO NOT SORT | Lecturer name |
| `Activity ID` | Activity ID | Activity code (e.g., `MGMT10101_U_1_SM1_2026_S01_01`) |

Other columns (Credit Points, Teaching Hours, Quota/Cap, Day/Start/Finish/Venue, program breakdowns, etc.) are **not read** by the teaching stream parser.

#### Teaching Matrix — Staff Sheet

| Keyword Matched | Example Column Header | What the Script Extracts |
|----------------|----------------------|------------------------|
| `Title Given Name Family Name` | Title Given Name Family Name | Lecturer display name (used to match with teaching data) |
| `Status` | Status DO NOT SORT | Employment status (e.g., `Continuing T&R`, `Continuing T`, `Fixed Term`) |

Other columns (FTE, workload, email, scheduling, etc.) are **not read**.

> [!TIP]
> The scripts match column headers using **keyword contains** (case-insensitive, whitespace-normalised). However, line breaks are trickier for the scripts to handle, so it's **recommended to avoid them**.

---

### Output Tables (What the System Writes)

These tables are populated automatically — you generally don't need to edit them.

#### `SubjectList` Sheet → `subject_list` Table

| Column | Header | What It Contains | Example |
|--------|--------|-----------------|---------|
| A | UID (sorter) | Auto-generated unique ID | `20260101_001` |
| B | Subject Code | Standard subject code | `MGMT10101` |
| C | Subject Name | Full subject name | `Management and Marketing` |
| D | Subject Coordinator | Coordinator name | `Jane Smith` |
| E | Delivery Mode | How the subject is delivered | `On Campus` |
| F | Grouped Period | FHY or SHY | `FHY` |
| G | Study Period | Specific study period | `Semester 1` |
| H | Status | Active or inactive | `Active` |
| I | Handbook Link | Auto-generated URL | `https://handbook.unimelb.edu.au/...` |
| J | Exclude | Checkbox to exclude subject | `TRUE` / `FALSE` |

**Auto-exclusion rules** (applied by the system):
- Subject code contains `FNCE`, `ACCT`, or `ECON`
- Subject name contains "indigenous" or "indigenising"
- Last 5 characters of subject code are not numeric
- Delivery mode contains "online" or "offshore"

#### `teaching stream` Sheet → `teaching_stream` Table

| Column | Header | What It Contains | Example |
|--------|--------|-----------------|---------|
| A | Lecturer Key | Auto-generated unique key | `MGMT10101|Semester 1|Jane Smith` |
| B | Subject Code | Standard subject code | `MGMT10101` |
| C | Study Period | Study period | `Semester 1` |
| D | Lecturer | Lecturer name | `Jane Smith` |
| E | Status | Employment status | `Continuing` |
| F | Activity ID | Teaching activity codes | `S01_01; S02_01` |
| G | Streams | Number of teaching streams | `2` |

#### `AllSubjectsHTML` Sheet (Power Query)

| Column | Header | What It Contains |
|--------|--------|-----------------|
| A | SubjectCode | Subject code|
| B | URL | Handbook assessment page URL |
| C | AssessmentTableHTML | Raw HTML of the assessment table |
| D | HTMLLength | Character count of the HTML |
| E | Status | `Success` or `Failed` |
| F | ErrorMessage | Error details (if failed) |
| G | FetchTime | When the data was fetched |

### Assessment Data → `assessment data parsed` Sheet

| Column | Header | What It Contains |
|--------|--------|-----------------|
| A | Subject Code | Subject code |
| B | Study Period | Study period (or "All") |
| C | Assessment Name | Name of the assessment |
| D | Word Count | Word/page count |
| E | Exam | Whether it's an exam |
| F | Group Size | Group size (if group assessment) |
| G | Quantity | Number of assessment items |

---

## Dashboard Parameters

### Required Parameters

| Cell | Parameter | Description | Validation |
|------|-----------|-------------|-----------|
| **C2** | Year | Academic year being processed | Must be a number ≥ 2025 |

### Optional Parameters

| Cell | Parameter | Description | If Left Blank |
|------|-----------|-------------|--------------|
| **C3** | Enrolment Tracker filename | Override default filename | Uses default filename |
| **C5** | Teaching Matrix filename | Override default filename | Uses default filename |
| **C8** | Word count benchmark | Words per hour rate | Defaults to 3000 |
| **C9** | Exam benchmark | Exams per hour rate | Defaults to 3 |
| **C10** | Marking support benchmark | Hours per teaching stream | Defaults to 20 |
| **C12** | Email address | Notification email on completion | No email sent |

---

## Understanding the Output

The system produces a file like: `[YEAR] Marking Admin Support Calculations.xlsm`

This file contains two sheets:

| Sheet | Contents |
|-------|----------|
| **FHY Calculations** | First-half year subjects (Summer Term, Semester 1) |
| **SHY Calculations** | Second-half year subjects (Winter Term, Semester 2) |

Each sheet has subjects organised by study period, with:
- **Assessment details** (columns E–K): List of assessments and their characteristics
- **Lecturer information** (columns L–O): Who teaches each stream and their staff status
- **Marking allocation** (columns P–R): Stream enrolment and marking hours *(you fill these in)*
- **Marker blocks** (columns T onwards): Space for 3 markers per subject *(can add more if needed, just copy and paste)*

---

## Refreshing Lecturer Data

The exported calculation file has a **Refresh** button in cell L2 on each sheet. This lets you update lecturer assignments without re-running the entire system.

### How to Use

1. Open the exported calculation file (e.g., `2026 Marking Admin Support Calculations.xlsm`)
2. Click the **Refresh** button in cell L2 on either sheet
3. The system will:
   - Read the latest Teaching Matrix data
   - Update lecturer names, status, and activity codes (columns L–O)
   - **Preserve** your notes and enrolments (columns P and S)
4. Wait about 1–2 minutes for completion

> **Important**: Your edits in columns P (Stream Enrolment) and S (Lecturer Notes) are always preserved during a refresh. Only columns L–O are updated.

---

## Common Issues & Fixes

### "Please enter a valid year" error
- **Cause**: Cell C2 on the Dashboard is empty or contains text
- **Fix**: Enter a number like `2026` in cell C2

### Process gets stuck at "Running..."
- **Cause**: The cloud workflow didn't report back as complete
- **Fix**: 
  1. Check your internet connection
  2. Wait a few more minutes (sometimes SharePoint syncs slowly)
  3. If stuck for >10 minutes, run `StopWorkflowMonitoring` and try again

### Assessment data shows "Failed" for many subjects
- **Cause**: The handbook year may not be published yet, or the year in C2 is wrong
- **Fix**: 
  1. Check that the year in C2 matches an existing handbook year
  2. Try opening one of the handbook URLs manually in your browser to verify

### "Required sheets are missing" error
- **Cause**: One of the data sheets was renamed or deleted
- **Fix**: Ensure these sheets exist with **exactly** these names:
  - `Dashboard`
  - `SubjectList`
  - `assessment data parsed`
  - `teaching stream`

### Lecturer Refresh button doesn't work
- **Cause**: The source workbook path may have changed
- **Fix**: Contact the developer to update the file path in the macro

### Changes to the Enrolment Tracker or Teaching Matrix aren't reflected
- **Cause**: The system pulls fresh data each run, but file names must match
- **Fix**: 
  1. Make sure the filename in C3/C5 matches the actual file on SharePoint
  2. If you renamed the file, update C3/C5 accordingly

### Something broke and you're not sure why
- **Cause**: SharePoint folders may have been moved/renamed, or sheet/table/column names were changed
- **Fix** — run through this checklist:
  1. **File locations**: Confirm the Enrolment Tracker and Teaching Matrix files are still in the `TEACHING MATRIX & ENROLMENT TRACKER` folder on SharePoint and haven't been moved or renamed
  2. **SharePoint paths**: If folders were reorganised, the VBA file paths and Power Automate flow URLs may need updating — contact the developer
  3. **Sheet names**: Verify these sheets exist with exact names: `Dashboard`, `SubjectList`, `assessment data parsed`, `teaching stream`, `AllSubjectsHTML`
  4. **Source workbook table names**: Verify: `subject_list`, `teaching_stream`, `AllSubjectsHTML`, `progress_bar`
  5. **Enrolment Tracker table names**: Verify: `Enrolment_Tracker`, `Enrolment_Number`
  6. **Teaching Matrix table names**: Verify: `Teaching_Data`, `Staff_table`
  7. **Column headers**: Ensure no column headers have been renamed in the source files — the scripts match by keyword (see [Source Files](#source-files-what-you-maintain) for the exact keywords)

---

## Quick Reference Card

| Action | How |
|--------|-----|
| Run the full process | Fill in Dashboard → Click Run button |
| Stop a running process | Run `StopWorkflowMonitoring` macro |
| Reset status after a crash | Run `ResetStatus` macro |
| Refresh lecturers in exported file | Click **Refresh** button on the sheet |
| Check for errors | Look at the `Process Log` sheet |
| Exclude a subject | Set the `Exclude` checkbox to TRUE in `SubjectList` |
