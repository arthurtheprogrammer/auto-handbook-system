# Power Automate Flows

This folder contains the two cloud flows that power the data ingestion pipeline. Both are triggered via HTTP POST requests from the VBA `Integration.bas` module.

---

## Flow 1: Subject List Refresh

**File**: [subjectlist.json](subjectlist.json)<br>
**Visual Diagram**: <br>![Subject List flow diagram](subject%20list%20workflow.png)<br>
**Trigger**: HTTP POST from `TriggerSubjectListWorkflow()` in `SubjectListRefresh.bas`<br>
**Purpose**: Reads the Enrolment Tracker from SharePoint, filters active subjects, and runs the `subjectListParser` Office Script to populate the `subject_list` table.

### Input Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `year` | integer | Academic year (e.g., `2026`) |
| `enrolmentTrackerFilename` | string | Override filename (blank = default `YEAR_M&M_Enrolment Tracker.xlsm`) |
| `email` | string | Notification email (optional) |

### Execution Steps

```
1. Initialise File Path         → Initialise `filePath` variable (null)
2. If only YEAR entered         → If filename is blank, use default naming convention
   ├─ True:  Set filePath = "{year}_M&M_Enrolment Tracker.xlsm"
   └─ False: Set filePath = provided filename (append .xlsm if missing)
3. Get YEAR PLANNING table      → Read all rows from `Enrolment_Tracker` table in SharePoint
4. Filter ACTIVE subjects       → Keep only rows where Status contains "Active"
5. Generate SubjectList table   → Run `subjectListParser` Office Script with filtered data
```

### Connectors Used

| Connector | Usage |
|-----------|-------|
| Excel Online (Business) | Read Enrolment Tracker table, run Office Script |

### Completion Signal

The `subjectListParser` Office Script updates the `progress_bar` table in the source workbook, which the VBA `MonitorAndExecute` loop watches (cell F2) for completion.

> [!TIP]
> **`Done`** = the Power Automate flow finished successfully. **`Complete`** = the VBA monitoring loop detected the update and proceeded to the next step. If you only see `Done` but the process stalls, the VBA may have timed out due to heavy computing load on your machine — try closing memory-consuming apps or browser tabs (e.g. Chrome) and re-run the programme again. Alternatively, borrow a spare laptop from reception and run it from that machine.

---

## Flow 2: Teaching Stream Refresh

**File**: [teachingstream.json](teachingstream.json)<br>
**Visual Diagram**: <br>![Teaching Stream flow diagram](teaching%20stream%20workflow.png)<br>
**Trigger**: HTTP POST from `TriggerTeachingStreamWorkflow()` in `TeachingStreamRefresh.bas`<br>
**Purpose**: Reads the Teaching Matrix (two tables: Teaching Data and Staff), filters and transforms the data, then runs the `teachingStreamParser` Office Script to populate the `teaching_stream` table.

### Input Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `year` | integer | Academic year (e.g., `2026`) |
| `teachingMatrixFilename` | string | Override filename (blank = default `YEAR_M&M_Teaching Matrix.xlsm`) |
| `email` | string | Notification email (optional) |

### Execution Steps

```
1. Initialise File Path                    → Initialise `filePath` variable (null)
2. If only YEAR entered                    → If filename is blank, use default naming convention
   ├─ True:  Set filePath = "{year}_M&M_Teaching Matrix.xlsm"
   └─ False: Set filePath = provided filename (append .xlsm if missing)
3. Update progress cell                    → Write "Running..." to progress_bar table
4. List rows present in Teaching Data      ─┐  (parallel)
5. List rows present in Staff table        ─┘  (parallel, pagination: up to 5000 rows)
6. Filter array                            → From Teaching Data: keep rows where Scheduled? = "Open" AND Activity ID ≠ ""
7. Create Staff Lookup Dictionary          → From Staff table: select { Name, Status } for each lecturer
8. Lecturer Stream string                  ─┐  (parallel compose steps)
9. Lecturer Status String                  ─┘
10. Generate Teaching Stream table          → Run `teachingStreamParser` Office Script with both datasets
```

### Connectors Used

| Connector | Connection | Usage |
|-----------|------------|-------|
| Excel Online (Business) | Primary | Read Teaching Data and Staff tables |
| Excel Online (Business) | Secondary | Update progress_bar table, run Office Script |

### Completion Signal

The `teachingStreamParser` Office Script updates the `progress_bar` table in the source workbook, which the VBA `MonitorAndExecute` loop watches (cell F5) for completion.

> [!TIP]
> **`Done`** = the Power Automate flow finished successfully. **`Complete`** = the VBA monitoring loop detected the update and proceeded to the next step. If you only see `Done` but the process stalls, apply the same troubleshooting advice from the [Subject List Refresh](#subject-list-refresh).

---

## SharePoint Paths

Both flows read from the same SharePoint site and document library:

| Item | Path |
|------|------|
| **Site** | SharePoint Group `ad6c8e15-4773-48f0-a918-df5ce6b5a0ec` |
| **Source files** | `/Shared Documents/TEACHING MATRIX & ENROLMENT TRACKER/` |
| **Target workbook** | `/Shared Documents/TEACHING SUPPORT/Handbook (Course & Subject Changes)/Auto Handbook System/Automated Handbook Data System.xlsm` |

## Required Table Names

These table names in the source Excel files must not be changed:

| File | Table Name | Used By |
|------|-----------|---------|
| Enrolment Tracker | `Enrolment_Tracker` | Subject List flow |
| Teaching Matrix | `Teaching_Data` | Teaching Stream flow |
| Teaching Matrix | `Staff_table` | Teaching Stream flow |
| Automated Handbook Data System | `progress_bar` | Both flows (completion signal) |
