# User Guide — Auto Handbook System

A guide for team members who manage the data sources and run the system. No coding knowledge required.

---

## Table of Contents

- [⚠️ Important Notice — Handbook Data Download](#%EF%B8%8F-important-notice--handbook-data-download)
- [How It Works (Simple Version)](#how-it-works-simple-version)
- [Running the System](#running-the-system)
- [What You Need to Maintain](#what-you-need-to-maintain)
- [Data Source Reference](#data-source-reference)
- [Dashboard Parameters](#dashboard-parameters)
- [Understanding the Output](#understanding-the-output-marking--admin-support-calculations)
- [Using the Marking Support Output](#using-the-marking-support-output)
- [Backing Up](#backing-up)
- [Refreshing Lecturer Data](#refreshing-lecturer-data)
- [First-Time Setup](#first-time-setup-excel-trust--calculation-settings)
- [Handing Over Power Automate Flows](#handing-over-power-automate-flows)
- [Common Issues & Troubleshooting](#common-issues--troubleshooting)

---

## ⚠️ Important Notice — Handbook Data Download

> [!WARNING]
> **Temporary limitation (as of April 2026):** Handbook data can currently only be downloaded from a **Windows computer connected to the University VPN**.

### What's happening?

The system normally fetches assessment details from the University handbook website automatically. However, due to a university cybersecurity restriction, the cloud-based download method (Power Automate) is currently **unable to access the handbook website** directly.

This means:

- ✅ **Windows + VPN** — The system works normally using the built-in Power Query method. You must be connected to the university VPN.
- ❌ **Mac** — The cloud-based fallback (Power Automate) is temporarily unavailable for handbook data. Mac users will need to arrange access to a Windows machine or use Remote Desktop.
- ❌ **Windows without VPN** — The handbook website may not be reachable. Connect to the VPN first.

### What do I need to do?

1. **Connect to the University VPN** before running the system. If you haven't set up the VPN, follow the university's [VPN setup guide](https://uomservicehub.service-now.com/esc?id=kb_article&sysparm_article=KB0206807)
2. **Run the system on a Windows computer** while connected to the VPN
3. Everything else works the same — just follow the normal steps below

> [!NOTE]
> We are working with the university's cybersecurity team to resolve this. Once the restriction is lifted, the cloud-based method will be re-enabled and Mac users will be able to run the system as before. This notice will be removed when the issue is resolved.

---

## How It Works (Simple Version)

Think of this system as a **data assembly line**:

```text
📂 Enrolment Tracker  ─┐
                       ├──→  🤖 System processes  ──→  📊 Calculation Spreadsheet
📂 Teaching Matrix    ─┤      everything for you         (ready to use)
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
| ---- | ----- |
| **Enrolment Tracker** (`.xlsx`) | `TEACHING MATRIX & ENROLMENT TRACKER` folder |
| **Teaching Matrix** (`.xlsx`) | `TEACHING MATRIX & ENROLMENT TRACKER` folder |
| **Automated Handbook Data System** (`.xlsm`) | `TEACHING SUPPORT > Handbook > Auto Handbook System` folder |

### Step-by-Step

1. **Open** the `Automated Handbook Data System.xlsm` workbook from SharePoint
2. **Go to** the `Dashboard` sheet
3. **Fill in** the required fields:

   | Cell | What to Enter | Example |
   | ---- | ------------- | ------- |
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
   | ---- | ---- | ------ |
   | F2 | Subject List Refresh | Running... → Complete |
   | F3 | Generate Subject Queries | Running... → Complete (or **Skipped** on Mac) |
   | F4 | Parse Assessment Data | Running... → Complete |
   | F5 | Teaching Stream Refresh | Running... → Complete |
   | F6 | Generate Calculation Sheets | Running... → Complete |

6. **Done!** The exported file appears in the same SharePoint folder as the source workbook

> [!NOTE]
> **Running on a Mac?** The Subject Queries step (F3) uses Power Query on Windows but is not available natively on Mac. Instead, on Mac you will be prompted to trigger a **cloud-based HTML download** via Power Automate. This takes a few minutes longer but produces the same result. You can also choose to skip and use existing data (check if the existing data is not from previous year).

<!-- -->

> [!TIP]
> If you need to stop the process, run `StopWorkflowMonitoring` from the macro menu.

---

## What You Need to Maintain

### ✅ Things You Should Keep Updated

| Item | Why | How Often |
| ---- | --- | --------- |
| **Enrolment Tracker** file | Provides subject codes and enrolments | Each semester |
| **Teaching Matrix** file | Provides lecturer assignments | Each semester |
| **Year in C2** | Used to fetch the correct handbook year | Each year |
| **Benchmark values** (C8, C9, C10) | Used in marking hour calculations | When policy changes |

### ⚠️ Things You Should NOT Change

| Item | Why |
| ---- | --- |
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
| --------------- | --------------------- | ------------------------ |
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
| --------------- | --------------------- | ------------------------ |
| `Subject Code` | Subject Code DO NOT SORT | Subject code |
| `Study Period` | Study Period | e.g., `Semester 1` |
| `Lecturer` | Lecturer DO NOT SORT | Lecturer name |
| `Activity ID` | Activity ID | Activity code (e.g., `MGMT10101_U_1_SM1_2026_S01_01`) |

Other columns (Credit Points, Teaching Hours, Quota/Cap, Day/Start/Finish/Venue, program breakdowns, etc.) are **not read** by the teaching stream parser.

#### Teaching Matrix — Staff Sheet

| Keyword Matched | Example Column Header | What the Script Extracts |
| --------------- | --------------------- | ------------------------ |
| `Title Given Name Family Name` | Title Given Name Family Name | Lecturer display name (used to match with teaching data) |
| `Status` | Status DO NOT SORT | Employment status (e.g., `Continuing T&R`, `Continuing T`, `Fixed Term`) |

Other columns (FTE, workload, email, scheduling, etc.) are **not read**.

> [!TIP]
> The scripts match column headers using **keyword contains** (case-insensitive, whitespace-normalised). However, line breaks are trickier for the scripts to handle, so it's **recommended to avoid them**.

---

### Working Tables (Individual deliverables to track each process)

These tables are populated automatically to build towards the final output spreadsheet. You generally don't need to edit them.

#### `SubjectList` Sheet → `subject_list` Table

| Column | Header | What It Contains | Example |
| ------ | ------ | ---------------- | ------- |
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
| ------ | ------ | ---------------- | ------- |
| A | Lecturer Key | Auto-generated unique key | `MGMT10101\|Semester 1\|Jane Smith` |
| B | Subject Code | Standard subject code | `MGMT10101` |
| C | Study Period | Study period | `Semester 1` |
| D | Lecturer | Lecturer name | `Jane Smith` |
| E | Status | Employment status | `Continuing` |
| F | Activity ID | Teaching activity codes | `S01_01; S02_01` |
| G | Streams | Number of teaching streams | `2` |

#### `AllSubjectsHTML` Sheet (Power Query / Power Automate)

| Column | Header | What It Contains |
| ------ | ------ | ---------------- |
| A | SubjectCode | Subject code |
| B | URL | Handbook assessment page URL |
| C | AssessmentTableHTML | Raw HTML of the assessment table |
| D | HTMLLength | Character count of the HTML |
| E | Status | `Success` or `Failed` |
| F | ErrorMessage | Error details (if failed) |
| G | FetchTime | When the data was fetched |

> [!NOTE]
> **Windows users** fetch this data directly using Power Query. **Mac users** fetch this using the **Assessment Query Workflow** in Power Automate which acts as a fallback to download the HTML.

### Assessment Data → `assessment data parsed` Sheet

| Column | Header | What It Contains |
| ------ | ------ | ---------------- |
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
| ---- | --------- | ----------- | ---------- |
| **C2** | Year | Academic year being processed | Must be a number ≥ 2025 |

### Optional Parameters

| Cell | Parameter | Description | If Left Blank |
| ---- | --------- | ----------- | ------------- |
| **C3** | Enrolment Tracker filename | Override default filename | Uses default filename |
| **C5** | Teaching Matrix filename | Override default filename | Uses default filename |
| **C8** | Word count benchmark | Words per hour rate | Defaults to 3000 |
| **C9** | Exam benchmark | Exams per hour rate | Defaults to 3 |
| **C10** | Marking support benchmark | Hours per teaching stream | Defaults to 20 |
| **C12** | Email address | Notification email on completion | No email sent |

---

## Understanding the Output (`Marking & Admin Support Calculations`)

The system produces a file like: `[YEAR]_M&M_Marking Admin Support Calculations.xlsm`

This file contains two sheets:

| Sheet | Contents |
| ----- | -------- |
| **FHY Calculations** | First-half year subjects (Summer Term, Semester 1) |
| **SHY Calculations** | Second-half year subjects (Winter Term, Semester 2) |

### Column Layout

Each sheet has subjects organised by study period. Here is the full column breakdown:

🔒 = locked (formula or system-generated) &emsp; ✏️ = editable (working space for users)

#### Subject Info (A–D)

| Col | Section | | Description |
| --- | ------- | :-: | ----------- |
| A | UID | 🔒 | **Hidden** - can be used to filter, but recommended to just **search by subject code**) |
| B | Subject Code | 🔒 | e.g., `MGMT10001` |
| C | Study Period | 🔒 | e.g., `Semester 1` |
| D | Enrolment | 🔒 | **Formula** — `INDEX/SUMPRODUCT` linking to the Enrolment Tracker on SharePoint to pull live enrolment count |

#### Assessment Data (E–K)

Columns E–H & J are locked because they contain formulas referencing the `assessment data parsed` sheet and handbook data. Editing them would break the automatic calculations.

| Col | Section | | Description |
| --- | ------- | :-: | ----------- |
| E | Assessment Details | 🔒 | Individual assessment details — looked up from the `assessment data parsed` sheet |
| F | Word Count | 🔒 | Word count per assessment, parsed from handbook data |
| G | Exam | 🔒 | Exam duration (`Y`/`N`), parsed from handbook data |
| H | Group Size | 🔒 | Group size if applicable, parsed from handbook data |
| I | Assessment Quantity | ✏️ | **Formula** — calculates quantity per student (`enrolment ÷ group size` if group work). Set to `0` for in-class assessments detected by keyword. Unlocked so users can override the calculated value if needed |
| J | Marking Hours | 🔒 | **Formula** — `quantity × word count ÷ benchmark` (written) or `quantity ÷ exam benchmark` (exams). Benchmarks set in J2/J3. The total row sums all assessments for the subject |
| K | Assessment Notes | ✏️ | Free-text notes about specific assessments |

#### Lecturer & Streams (L–S)

| Col | Section | | Description |
| --- | ------- | :-: | ----------- |
| L | Lecturer/Instructors | ✏️ | Lecturer names — auto-populated from teaching stream data, refreshable via the Refresh button |
| M | Status | ✏️ | Staff status (`Continuing T&R` highlighted to indicate eligibility for marking support) — auto-populated from staff data |
| N | Stream # | ✏️ | Stream number |
| O | Activity Code | ✏️ | aggregated activity code from Allocate+ (e.g., `S01_01; S01_02`) |
| P | Stream(s) Enrolment | ✏️ | **Manual Entry** - enrolment count per stream |
| Q | Allocated Marking | 🔒 | **Formula** — `Stream # x 20 hrs/stream (default/benchmark set on Dashboard C10)` |
| R | Marking Support Hours Available | 🔒 | **Formula** — `Total Marking Hours for the subject * (Stream(s) Enrolment / Total Enrolment) - Allocated Marking` |
| S | Lecturer Notes | ✏️ | Free-text notes about the lecturer/stream |

#### Marker Blocks (T–AW) — 3 identical blocks of 10 columns each

Each subject has space for **3 markers**. Markers 2 (AD–AM) and 3 (AN–AW) follow the exact same 10-column structure as Marker 1 below.

| Col | Section | | Description |
| --- | ------- | :-: | ----------- |
| T | Marker 1 Name | ✏️ | **Manual Entry** — who is marking |
| U | Assessment Details | ✏️ | **Manual Entry** — user-selected Assessment Details (column E), or self-defined marking arrangements (e.g. `Other Casual Academic Activity (2 hours per task)`) |
| V | Word Count | ✏️ | **Formula** — `INDEX/MATCH` pulling the matching word count from column F |
| W | Exam | ✏️ | **Formula** — `INDEX/MATCH` pulling the matching exam duration from column G |
| X | Group Size | ✏️ | **Formula** — `INDEX/MATCH` pulling the matching group size from column H |
| Y | Assessment Quantity | ✏️ | **Manual Entry** — how many assessment/hours this marker handles |
| Z | Marking Allocation | ✏️ | **Formula** — calculates hours from quantity and benchmarks (same logic as column J) |
| AA | Email | ✏️ | Marker's email address |
| AB | Arrangement Notes | ✏️ | Notes about the marking arrangement |
| AC | Contract Requested | ✏️ | Dropdown: `Y` / `N` |

This section is **fully customisable**. You can select only some assessments (not all), overwrite any formula in the marker blocks, or create entirely custom marking arrangements. The formulas are provided as a starting point — feel free to replace them with manual values where needed.

> [!TIP]
> If you need **more than 3 markers** for a subject, copy and paste an existing marker block to the right. The formulas will adjust automatically.

### Overwriting Protected Cells

The sheets are **protected** to prevent accidental changes to handbook-derived and certain formula columns. The protection allows formatting but **blocks inserting or deleting rows**.

To make changes to locked cells or add/remove rows, you need to **unprotect the sheet** first:

1. Go to the **Review** tab on the ribbon
2. Click **Unprotect Sheet** (no password is required)
3. Make your changes
4. Re-protect the sheet when done: **Review** → **Protect Sheet** → **OK**

#### When you might need to do this

**Incorrectly parsed word count or group size** — The system extracts word counts and group sizes from handbook assessment descriptions using pattern matching (e.g., looking for "1500 words" or "groups of 4"). If the handbook uses unusual phrasing, the parsed value may be wrong or missing. Unprotect the sheet, correct columns F/H, and the marking hours in column J will recalculate automatically.

**In-class assessment detection** — The system tries to identify in-class assessments (which don't require traditional marking) by scanning the assessment description for these keywords:

`participation` · `presentation` · `attendance` · `pitch` · `online` · `ongoing` · `in class`

When detected, word count and exam duration are set to `0` and the assessment quantity formula adjusts accordingly. However, because handbook descriptions are not standardised, **the system cannot catch every case**. If an assessment was incorrectly classified (or missed), unprotect the sheet and manually adjust the Assessment Quantity (column I) or Marking Hours (column J) as needed.

**Adding more rows** — If you need extra rows beyond what the system generated (e.g., for additional lecturers or notes), unprotect the sheet first, then insert rows as needed.

---

## Using the Marking Support Output

This section walks you through how to actually **use** the generated calculation spreadsheet once it's been produced. Think of the output as a _starting point_ — the system auto-fills as much as it can from the handbook, but you'll need to review and adjust based on what each academic provides.

### Step 1: Enter Stream Enrolments (Column P)

For each **Teaching & Research (T&R) academic**, enter the **total number of students across all streams** they are teaching for that subject in the **Stream(s) Enrolment** column (P).

For example, if an academic teaches Stream 1 (120 students) and Stream 2 (95 students), enter `215` in column P.

This number is used to calculate how many marking hours are available to that lecturer.

### Step 2: Understand the Marking Support Hours Calculation

The **Marking Support Hours Available** (column R) tells you how many marking support hours a T&R academic can request. It is calculated as:

```
Marking Support Hours Available = Total Marking Hours × (Stream Enrolment / Total Enrolment) − Allocated Marking
```

Where:

| Component | Column | Description |
| --------- | ------ | ----------- |
| **Total Marking Hours** | J (total row) | Sum of marking hours across all assessments for the subject, based on word count and exam benchmarks |
| **Stream Enrolment** | P | The number you entered — total students across the academic's streams |
| **Total Enrolment** | D | Live enrolment count pulled from the Enrolment Tracker |
| **Allocated Marking** | Q | Baseline marking the academic is expected to cover themselves. Formula: `Number of Streams × 20 hrs/stream` (benchmark set on Dashboard C10) |

In plain terms: the system works out the academic's **share** of the total marking (proportional to their students), then **subtracts** the marking they're already expected to do as part of their role.

### Step 3: Verify Assessment Details Against the Academic's Information

The assessment data (columns E–J) is **automatically parsed from the handbook website**. While the parsing is generally accurate, it can only predict so much from the handbook's language. **You should fact-check the assessment details for each subject** before processing.

Common things to verify:

- **Word counts** — Does the word count in column F match what the academic says?
- **Group sizes** — Does the group size in column H match the academic's actual class setup?
- **Exam details** — Is the exam flag (column G) correct?
- **Assessment quantity** — Does the calculated quantity in column I make sense?

> [!IMPORTANT]
> **Check assessment details against what the academic is providing you.** Unexpected word counts or group sizes are the most common source of discrepancies in the final marking hours.

### Step 4: Cross-Check with the Subject Coordinator's Guidelines

The Subject Coordinator (SC) may have their own assessment marking guidelines. Be aware that:

- The system typically uses the **minimum group size** from the handbook to calculate the **maximum marking hours available**
- The SC's actual group sizes may be **larger or different** from the handbook minimum — in that case, the group assessment estimates will not match
- If there's a mismatch, you'll need to correct it (see [Adjusting Assessment Details](#adjusting-assessment-details-columns-fh) below)

### Adjusting Assessment Details (Columns F–H)

If the assessment details don't match what the academic or SC has provided:

1. **Unprotect the sheet**: Go to **Review** → **Unprotect Sheet** (no password needed)
2. **Edit the relevant cells** in columns F (Word Count), G (Exam), or H (Group Size)
3. Changes will automatically flow through to the **Assessment Quantity** (column I) and **Marking Hours** (column J)
4. **Re-protect the sheet** when done: **Review** → **Protect Sheet** → **OK**

### Step 5: Handle Special Cases

Some assessments can't be fully captured by the auto-parsing and need manual attention:

| Scenario | What to Do |
| -------- | ---------- |
| **Class participation with a written/marking component** | If there's a word-equivalent marking component (e.g., "500-word reflection per tutorial"), manually enter the word count in the relevant assessment row |
| **Midterm exams or tests not captured** | Some midterms may not appear in the handbook data — manually add the word count or exam flag |
| **Assessments excluded from or not captured with word count** | If the handbook description doesn't mention a word count but the assessment does require marking, manually enter the word count |
| **Non-standard assessments** | For any assessment the parser didn't capture correctly, unprotect the sheet and adjust columns F–H as needed |

### Step 6: Log and Verify Against University Benchmarks

When the academic provides their own marking calculations:

1. **Log their calculations** in an available **Marker Block** (columns T–AC, AD–AM, or AN–AW) — enter the academic's proposed arrangement in the Assessment Details and Quantity fields
2. **Compare** their calculation against the system's output and the university's benchmarks (3,000 words/hr, 3 exams/hr, 20 hrs/stream)
3. **Check compliance** — sometimes academics may miscalculate by not understanding the benchmarks correctly, or provide numbers that don't align with the uni's standards
4. **The same applies for non-T&R staff** in special cases — enter their total stream enrolment and fill in the arrangement in an available Marker Block to verify compliance

> [!TIP]
> Using a Marker Block to log the academic's proposed arrangement alongside the system's calculation makes it easy to spot discrepancies and confirm compliance in one view.

### Step 7: Account for Extra Marking Commitments

If an academic indicates they are **willing to mark beyond their allocated marking hours**, make sure to adjust the baseline:

- The default allocated marking is **20 hrs × number of streams** (column Q)
- Add the **agreed extra marking hours** to this baseline
- For example: if an academic has 2 streams and agrees to an extra 10 hours, the effective allocated marking is `(2 × 20) + 10 = 50 hours`
- Unprotect the sheet, update the Allocated Marking value in column Q, and re-protect

### ⚠️ Don't Leave Notes in Columns L–O

> [!CAUTION]
> **Do not put notes or comments in the Lecturer (L), Status (M), Stream # (N), or Activity Code (O) columns.** Running the **Lecturer Refresh** button will overwrite everything in columns L–O. Use the **Lecturer Notes** column (S) or **Assessment Notes** column (K) instead — these are preserved during a refresh.

---

## Backing Up

> [!IMPORTANT]
> **Always use the original working document** (`Automated Handbook Data System.xlsm`) for generating sheets — this is the file with all VBA macros and Power Automate flows linked to it. Do **not** run the system from a backup copy.

After each successful run, **save a backup** of both files to preserve that semester's data:

1. **The source workbook** — copy `Automated Handbook Data System.xlsm` and rename it (e.g., `Automated Handbook Data System - 2026 S1 backup.xlsm`)
2. **The exported calculation file** — this is already a separate file (e.g., `2026_M&M_Marking Admin Support Calculations.xlsm`)

Store backups in the `backups` folder within the **Auto Handbook System** folder on SharePoint so the team can refer to previous years' outputs if needed.

> **Why back up?** Each run overwrites the data tables in the source workbook (SubjectList, teaching stream, assessment data). If you want to keep a snapshot of a particular semester's data — for example, to compare year-over-year — you need to save a copy before running again.

---

## Refreshing Lecturer Data

The exported calculation file has a **Refresh** button in cell L2 on each sheet. This lets you update lecturer assignments without re-running the entire system.

### How to Use

1. Open the exported calculation file (e.g., `2026_M&M_Marking Admin Support Calculations.xlsm`)
2. Click the **Refresh** button in cell L2 on either sheet
3. The system will:
   - Read the latest Teaching Matrix data
   - Update lecturer names, status, and activity codes (columns L–O)
   - **Preserve** your notes and enrolments (columns P and S)
4. Wait about 1–2 minutes for completion

> **Important**: Your edits in columns P (Stream Enrolment) and S (Lecturer Notes) are always preserved during a refresh. Only columns L–O are updated.

---

## First-Time Setup (Excel Trust & Calculation Settings)

These three settings need to be configured **once per computer** before first use.

> [!IMPORTANT]
> All three are **required**. Without them, macros might not run properly, enrolment numbers might not load, and calculations might not update.

<details>
<summary>Windows Setup</summary>

#### 1. Enable VBA Macros

1. Open Excel
2. Go to **File** → **Options**
3. In the left sidebar, click **Trust Center**
4. Click the **Trust Center Settings...** button
5. In the left sidebar, click **Macro Settings**
6. Select **Enable VBA macros**
7. Also tick **Trust access to the VBA project object model** (required for the LecturerRefresh module to export into the calculation file)
8. Click **OK** → **OK**

#### 2. Enable External Links & Data Connections

1. Open Excel
2. Go to **File** → **Options**
3. In the left sidebar, click **Trust Center**
4. Click the **Trust Center Settings...** button
5. In the left sidebar, click **External Content**
6. Under **Security settings for Workbook Links**, select **Enable automatic update for all Workbook Links**
7. Under **Security settings for Data Connections**, select **Enable all Data Connections**
8. Click **OK** → **OK**

#### 3. Set Calculation Mode to Automatic

1. Open Excel
2. Go to the **Formulas** tab on the ribbon
3. Click **Calculation Options** (in the Calculation group)
4. Select **Automatic**

Or via settings:

1. Go to **File** → **Options** → **Formulas**
2. Under **Calculation options**, set **Workbook Calculation** to **Automatic**
3. Click **OK**

</details>

<details>
<summary>Mac Setup</summary>

#### 1. Enable VBA Macros

1. Open Excel
2. Go to **Excel** (menu bar) → **Preferences**
3. Click **Security & Privacy**
4. Under **Macro Security**, select **Enable all macros**
5. Close the preferences window

#### 2. Enable External Links & Data Connections

1. Open the workbook
2. If you see a **Security Warning** bar at the top saying "Automatic update of links has been disabled", click **Enable Content**
3. If prompted about data connections, click **Enable**
4. For permanent trust: go to **Excel** → **Preferences** → **Security & Privacy** and ensure external content is allowed

> [!TIP]
> If enrolment numbers show as `0` or `#REF!` after opening, the link trust settings may not be enabled. Go to **Data** → **Edit Links** and click **Update Values** to force a refresh.

#### 3. Set Calculation Mode to Automatic

1. Open Excel
2. Go to the **Formulas** tab on the ribbon
3. Click **Calculation Options**
4. Select **Automatic**

Or via preferences:

1. Go to **Excel** → **Preferences** → **Calculation**
2. Under **Calculation**, select **Automatically**
3. Close the preferences window

</details>

---

## Handing Over Power Automate Flows

If ownership of the Power Automate flows needs to move to a new person (e.g. a staff change), follow these steps. The flow URLs and VBA modules **do not need to change** — only the internal connections and Office Script references need to be updated.

> [!IMPORTANT]
> The new person must have access to the same SharePoint sites as the previous owner (Enrolment Tracker, Teaching Matrix, and the main workbook folder) before starting.

### Step 1 — Create Connections

In [Power Automate](https://make.powerautomate.com), the new person creates their own connections:

1. Go to **Data → Connections → New connection**
2. Create connections for each of the following (authenticate with your university account):
   - **Excel Online (Business)**
   - **SharePoint**
   - **Microsoft Teams** (if applicable)
   - **Office 365 Outlook** (if applicable)

### Step 2 — Copy Office Scripts to Your OneDrive

Power Automate can only use Office Scripts that belong to the authenticated connection's account. The new person needs their own copies:

1. Open each Office Script file shared during handover (`.osts` files)
2. In **Excel Online** → **Automate** tab → **New Script**
3. Paste the script code in and save it — use the same name as the original

Repeat for each script (Subject List Parser, Teaching Stream Parser, and any others shared during handover).

### Step 3 — Update the Solution Connection Reference

The flows are part of a Power Automate solution, so connections are managed centrally:

1. Open the solution in Power Automate
2. Find the **Connection References** section
3. Switch the **Excel Online (Business)** reference to your newly created connection
4. Do the same for SharePoint and any other connections listed
5. Save the solution — this applies to all flows at once

> [!NOTE]
> If you see a "ReadAccess" error when trying to update the connection reference, ask the previous owner to add you as a **co-owner** of the solution first (solution → **Share** → add your name).

### Step 4 — Remap "Run Script" Actions

For any flow that runs an Office Script, the script reference must be updated to point to your copy:

1. Open the flow in edit mode
2. **Before changing anything** — screenshot the "Run Script" action so you have a record of the script name and all parameter values
3. Switch the connection to your connection
4. The **Script** dropdown will go blank — this is expected. Click it and select your copy of the script
5. Re-enter the parameters exactly as they were in the screenshot
6. Save the flow

> [!TIP]
> Repeat Step 4 for every flow that has a "Run Script" action — there may be more than one.

### Step 5 — Test

1. Manually trigger each flow from Power Automate to confirm it runs successfully
2. The **first 1–3 runs** after a connection switch may be slower than usual (up to 5 minutes) — this is normal as the new connection warms up
3. After a few runs, performance should return to normal

---

## Common Issues & Troubleshooting

### "Please enter a valid year" error

**Cause:** Cell C2 on the Dashboard is empty or contains text.

**Fix:** Enter a number like `2026` in cell C2.

---

### Process gets stuck at "Running..."

**Cause:** The cloud workflow didn't report back as complete.

**Fix:**

1. Check your internet connection
2. Wait a few more minutes (sometimes SharePoint syncs slowly)
3. If stuck for >10 minutes, run `StopWorkflowMonitoring` and try again

---

### Assessment data shows "Failed" for many subjects

**Cause:** The handbook year may not be published yet, or the year in C2 is wrong.

**Fix:**

1. Check that the year in C2 matches an existing handbook year
2. Try opening one of the handbook URLs manually in your browser to verify

---

### "VBA Access Required" error or Export/Refresh fails

**Cause:** Excel doesn't have permission to modify its own VBA code, which is needed to attach the lecturer refresh script to the exported calculation file.

**Fix:** Follow the [First-Time Setup](#first-time-setup-excel-trust--calculation-settings) instructions to enable "Trust access to the VBA project object model" in your Excel Settings/Preferences.

---

### "Required sheets are missing" error

**Cause:** One of the data sheets was renamed or deleted.

**Fix:** Ensure these sheets exist with **exactly** these names: `Dashboard`, `SubjectList`, `assessment data parsed`, `teaching stream`.

---

### Lecturer Refresh button doesn't work

**Cause:** The source workbook path may have changed.

**Fix:** Contact the developer to update the file path in the macro.

---

### Changes to the Enrolment Tracker or Teaching Matrix aren't reflected

**Cause:** The system pulls fresh data each run, but file names must match.

**Fix:**

1. Make sure the filename in C3/C5 matches the actual file on SharePoint
2. If you renamed the file, update C3/C5 accordingly

---

### Something broke and you're not sure why

**Cause:** SharePoint folders may have been moved/renamed, or sheet/table/column names were changed.

**Fix** — run through this checklist:

1. **File locations** — Confirm the Enrolment Tracker and Teaching Matrix files are still in the `TEACHING MATRIX & ENROLMENT TRACKER` folder on SharePoint and haven't been moved or renamed
2. **SharePoint paths** — If folders were reorganised, the VBA file paths and Power Automate flow URLs may need updating — contact the developer
3. **Sheet names** — Verify these sheets exist with exact names: `Dashboard`, `SubjectList`, `assessment data parsed`, `teaching stream`, `AllSubjectsHTML`
4. **Source workbook table names** — Verify: `subject_list`, `teaching_stream`, `AllSubjectsHTML`, `progress_bar`
5. **Enrolment Tracker table names** — Verify: `Enrolment_Tracker`, `Enrolment_Number`
6. **Teaching Matrix table names** — Verify: `Teaching_Data`, `Staff_table`
7. **Column headers** — Ensure no column headers have been renamed in the source files — the scripts match by keyword (see [Source Files](#source-files-what-you-maintain) for the exact keywords)

---

## Quick Reference Card

| Action | How |
| ------ | --- |
| Run the full process | Fill in Dashboard → Click Run button |
| Stop a running process | Run `StopWorkflowMonitoring` macro |
| Reset status after a crash | Run `ResetStatus` macro |
| Refresh lecturers in exported file | Click **Refresh** button on the sheet |
| Check for errors | Look at the `Process Log` sheet |
| Exclude a subject | Set the `Exclude` checkbox to TRUE in `SubjectList` |
