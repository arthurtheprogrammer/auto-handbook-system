# Developer Guide

Technical reference for maintaining and extending the Auto Handbook System.

---

## Table of Contents

- [Architecture Overview](#architecture-overview)
- [Data Pipeline](#data-pipeline)
- [Module Reference](#module-reference)
- [Data Sources & Sheet Reference](#data-sources--sheet-reference)
- [Key Cell References](#key-cell-references)
- [Silent Mode](#silent-mode)
- [Cross-Platform Notes](#cross-platform-notes)
- [Troubleshooting](#troubleshooting)

---

## Architecture Overview

The system has **four processing layers** that execute in sequence:

```mermaid
flowchart TD
    subgraph L1["Layer 1: Cloud Data Extraction"]
        PA1["Power Automate Flow 1\n(Subject List)"]
        PA2["Power Automate Flow 2\n(Teaching Stream)"]
        OS1["subjectListParser.osts"]
        OS2["teachingStreamParser.osts"]
        PA1 --> OS1
        PA2 --> OS2
    end

    subgraph L2["Layer 2: Web Scraping"]
        PQ["AllSubjectsHTML\n(Power Query)"]
    end

    subgraph L3["Layer 3: VBA Data Processing"]
        HQ["HTMLQuery.bas\n(refresh query)"]
        AD["AssessmentData.bas\n(parse HTML)"]
    end

    subgraph L4["Layer 4: Output Generation"]
        CS["CalculationSheets.bas\n(FHY + SHY sheets)"]
        EXP["Export to new .xlsm"]
        LR["LecturerRefresh.bas\n(in exported file)"]
    end

    L1 --> L2
    L2 --> L3
    L3 --> L4
```

### Execution Order (Integrated Run)

When `GenerateMarkingSupport()` is called:

| Step | Module | What Happens |
|------|--------|-------------|
| 1 | `Integration.bas` | Sets `SilentMode = True`, reads Dashboard params, validates year |
| 2 | `SubjectListRefresh.bas` | HTTP POST to Power Automate → triggers `subjectListParser.osts` → populates `SubjectList` table |
| 3 | `TeachingStreamRefresh.bas` | HTTP POST to Power Automate → triggers `teachingStreamParser.osts` → populates `teaching_stream` table |
| 4 | `Integration.bas` | Monitors Dashboard F2 and F5 cells for "Done" (polling loop, 2s interval, 30min timeout) |
| 5 | `HTMLQuery.bas` | Refreshes `AllSubjectsHTML` Power Query table (fetches assessment HTML from handbook.unimelb.edu.au) |
| 6 | `AssessmentData.bas` | Parses HTML from `AllSubjectsHTML` → writes structured data to `assessment data parsed` sheet |
| 7 | `CalculationSheets.bas` | Generates `FHY Calculations` and `SHY Calculations` sheets, exports to new `.xlsm` with `LecturerRefresh.bas` embedded |
| 8 | `Integration.bas` | Sends email notification, sets `SilentMode = False` |

---

## Data Pipeline

### Layer 1: Cloud Data Extraction

**Power Automate Flows** are triggered via HTTP POST from VBA. Each flow:
1. Reads a SharePoint Excel file (Enrolment Tracker or Teaching Matrix)
2. Extracts relevant data as JSON
3. Passes JSON to an Office Script in the target workbook
4. The Office Script parses and writes data to the appropriate table

#### Flow 1: Subject List Refresh

| Property | Value |
|----------|-------|
| Trigger | HTTP POST from `SubjectListRefresh.bas` |
| Source | Enrolment Tracker (`.xlsx`) on SharePoint |
| Script | `subjectListParser.osts` |
| Target | `subject_list` table in `SubjectList` sheet |
| Status Cell | Dashboard `F2` |

**Payload**: `{ "year": 2026, "enrolmentTrackerFilename": "...", "email": "..." }`

#### Flow 2: Teaching Stream Refresh

| Property | Value |
|----------|-------|
| Trigger | HTTP POST from `TeachingStreamRefresh.bas` |
| Source | Teaching Matrix (`.xlsx`) on SharePoint |
| Script | `teachingStreamParser.osts` |
| Target | `teaching_stream` table in `teaching stream` sheet |
| Status Cell | Dashboard `F5` |

**Payload**: `{ "year": 2026, "teachingMatrixFilename": "...", "email": "..." }`

### Layer 2: Web Scraping (Power Query)

The `AllSubjectsHTML` Power Query:
1. Reads subject codes from a `Parameters` table in the workbook
2. Constructs URLs: `https://handbook.unimelb.edu.au/{year}/subjects/{code}/assessment`
3. Fetches HTML for each subject
4. Extracts the `<div class="assessment-table">` section
5. Stores results with status, HTML length, and fetch time

### Layer 3: VBA Data Processing

**HTMLQuery.bas** (`GenerateSubjectQueries`):
- Refreshes the Power Query table
- Formats columns, sets hyperlinks, applies table style

**AssessmentData.bas** (`ParseAssessmentData`):
- Reads raw HTML from `AllSubjectsHTML` table
- Parses assessment details (name, word count, exam type, group size, quantity)
- Writes structured records to `assessment data parsed` sheet

### Layer 4: Output Generation

**CalculationSheets.bas** (`GenerateCalculationSheets`):
- Filters subjects by grouped period (FHY/SHY) and exclusion rules
- Cross-references with assessment data and teaching stream data
- Generates calculation sheets with benchmarks (word count/hr, exams/hr, marking support hrs/stream)
- Exports to a new `.xlsm` workbook with `LecturerRefresh.bas` module embedded
- Adds "Refresh" buttons to the exported sheets

---

## Module Reference

### Integration.bas (Orchestrator)

| Function | Purpose |
|----------|---------|
| `GenerateMarkingSupport()` | Main entry point. Sets SilentMode, validates params, triggers workflows, monitors, runs macros |
| `MonitorAndExecute()` | Polling loop that watches F2/F5 for completion, then calls `RunAllMacros` |
| `RunAllMacros()` | Sequential execution: HTMLQuery → AssessmentData → CalculationSheets |
| `ForceCloudSync()` | Saves workbook, refreshes all, forces recalculation (for SharePoint sync) |
| `CheckWorkflowComplete()` | Checks if a cell value is "DONE"/"COMPLETE"/"FINISHED" |
| `FinalizeProcess()` | Freezes elapsed time, sends email notification |
| `SendRequestMac()` / `SendRequestWindows()` | Platform-specific HTTP POST functions |
| `StopWorkflowMonitoring()` | User-callable macro to abort monitoring |
| `ResetStatus()` | Clears all status cells and resets state |

**Global Variables:**
- `Public SilentMode As Boolean` — suppresses MsgBox in sub-modules during integrated runs
- `Public StopMonitoring As Boolean` — flag to abort the monitoring loop
- `Public OriginalCalculationMode As XlCalculation` — saved calc mode for restoration

### SubjectListRefresh.bas

| Function | Purpose |
|----------|---------|
| `RefreshSubjectList()` | Standalone entry point (validates year, triggers workflow, shows MsgBox) |
| `TriggerSubjectListWorkflow()` | HTTP POST to Power Automate endpoint (called by Integration or standalone) |

### TeachingStreamRefresh.bas

| Function | Purpose |
|----------|---------|
| `RefreshTeachingStream()` | Standalone entry point |
| `TriggerTeachingStreamWorkflow()` | HTTP POST to Power Automate endpoint |

### HTMLQuery.bas

| Function | Purpose |
|----------|---------|
| `GenerateSubjectQueries()` | Refreshes `AllSubjectsHTML` Power Query table, applies formatting |
| `FormatTableCleanup()` | Standardizes row heights, column widths, hyperlinks, table style |

### AssessmentData.bas

| Function | Purpose |
|----------|---------|
| `ParseAssessmentData()` | Main parser: reads HTML → writes structured assessment records |
| `SetupHeaders()` | Creates column headers on target sheet |
| `ExtractAssessmentDetails()` | Parses individual assessment entries from HTML |
| `FormatOutput()` | Applies formatting to the output sheet |

### CalculationSheets.bas

| Function | Purpose |
|----------|---------|
| `GenerateCalculationSheets()` | Main entry: creates FHY + SHY sheets, exports workbook |
| `GenerateSheet()` | Creates one calculation sheet (FHY or SHY) |
| `PopulateSheetData()` | Fills in subject data, assessments, formulas |
| `ExportCalculationSheets()` | Creates new workbook, copies sheets, embeds VBA, saves |
| `InitializeProcessLog()` | Creates "Process Log" sheet for real-time logging |
| `LogMessage()` | Writes timestamped messages to Process Log + Debug + StatusBar |
| `VerifyRequiredSheets()` | Checks that Dashboard, SubjectList, assessment data parsed, teaching stream all exist |

### LecturerRefresh.bas (Exported File Only)

| Function | Purpose |
|----------|---------|
| `RefreshLecturerData()` | Main entry: reads source params → triggers Workflow 3 → waits → updates lecturer columns |
| `GetSourceParameters()` | Opens source workbook read-only, reads C2/C5/C12 |
| `TriggerWorkflow3()` | HTTP POST to Power Automate (Teaching Stream endpoint) |
| `WaitForWorkflow3Completion()` | Polls source Dashboard F5 every 3s (2min timeout) |
| `IdentifySubjectBlocks()` | Scans FHY/SHY Calculations for subject blocks by UID pattern |
| `UpdateAllLecturers()` | Refreshes columns L–O, preserves P–S user edits, inserts rows if needed |

> **Note**: This module lives in the **exported** calculation workbook, not the source workbook. It has its own copies of `SendRequestMac`, `SendRequestWindows`, and `EscapeJSON`.

---

## Data Sources & Sheet Reference

### Workbook Sheets

| Sheet Name | Table Name | Purpose | Created By |
|-----------|-----------|---------|-----------|
| `Dashboard` | `progress_bar`, `Parameters` | User inputs, status tracking, benchmarks | Manual |
| `SubjectList` | `subject_list` | Filtered subject data from Enrolment Tracker | `subjectListParser.osts` |
| `AllSubjectsHTML` | `AllSubjectsHTML` | Raw assessment HTML scraped from handbook | Power Query |
| `assessment data parsed` | *(range, not table)* | Structured assessment data parsed from HTML | `AssessmentData.bas` |
| `teaching stream` | `teaching_stream` | Lecturer assignments from Teaching Matrix | `teachingStreamParser.osts` |
| `FHY Calculations` | *(generated)* | First-half year calculation sheet | `CalculationSheets.bas` |
| `SHY Calculations` | *(generated)* | Second-half year calculation sheet | `CalculationSheets.bas` |
| `Process Log` | *(generated)* | Timestamped execution log | `CalculationSheets.bas` |

### External SharePoint Files

| File | Location | Purpose |
|------|----------|---------|
| Enrolment Tracker (`.xlsx`) | `/TEACHING MATRIX & ENROLMENT TRACKER/` | Source of subject codes, names, coordinators, study periods |
| Teaching Matrix (`.xlsx`) | `/TEACHING MATRIX & ENROLMENT TRACKER/` | Source of lecturer assignments, activity codes, stream counts |
| Automated Handbook Data System (`.xlsm`) | `/TEACHING SUPPORT/Handbook (.../Auto Handbook System/` | The main workbook containing all the macros |

---

## Key Cell References

### Dashboard Sheet

| Cell | Purpose | Used By |
|------|---------|---------|
| `C2` | **Year** (e.g., 2026) — used in all workflows and handbook URLs | All modules |
| `C3` | Enrolment Tracker filename (optional override) | `SubjectListRefresh.bas` |
| `C5` | Teaching Matrix filename (optional override) | `TeachingStreamRefresh.bas` |
| `C8` | Word count benchmark (words/hr) | `CalculationSheets.bas` |
| `C9` | Exam benchmark (exams/hr) | `CalculationSheets.bas` |
| `C10` | Marking support benchmark (hrs/stream) | `CalculationSheets.bas` |
| `C12` | Email address for completion notification | `Integration.bas` |
| `C15` | Last run date (auto-filled) | `Integration.bas` |
| `C16` | Last run start time (auto-filled) | `Integration.bas` |
| `C17` | Elapsed time (formula, then frozen) | `Integration.bas` |
| `F2` | **Subject List Refresh status** (monitored for "Done") | `Integration.bas`, `subjectListParser.osts` |
| `F3` | GenerateSubjectQueries status | `Integration.bas` |
| `F4` | ParseAssessmentData status | `Integration.bas` |
| `F5` | **Teaching Stream Refresh status** (monitored for "Done") | `Integration.bas`, `teachingStreamParser.osts` |
| `F6` | GenerateCalculationSheets status | `Integration.bas` |

### Calculation Sheet Columns (FHY/SHY)

| Col | Letter | Header |
|-----|--------|--------|
| 1 | A | UID (hidden) |
| 2 | B | Subject Code |
| 3 | C | Study Period |
| 4 | D | Enrolment |
| 5 | E | Assessment Details |
| 6 | F | Word Count |
| 7 | G | Exam |
| 8 | H | Group Size |
| 9 | I | Assessment Quantity |
| 10 | J | Marking Hours |
| 11 | K | Assessment Notes |
| 12 | L | Lecturer/Instructors |
| 13 | M | Status |
| 14 | N | Stream # |
| 15 | O | Activity Code |
| 16 | P | Stream(s) Enrolment |
| 17 | Q | Allocated Marking |
| 18 | R | Marking Support Hours Available |
| 19 | S | Lecturer Notes |
| 20–29 | T–AC | Marker 1 block (Assessment Details → Contract Requested) |
| 30–39 | AD–AM | Marker 2 block |
| 40–49 | AN–AW | Marker 3 block |

> **LecturerRefresh** updates columns **L–O** only and preserves **P–S** (user edits).

---

## Silent Mode

The `SilentMode` global variable controls whether `MsgBox` calls are displayed:

- **`SilentMode = True`**: Set at the start of `GenerateMarkingSupport()`. All sub-module MsgBox calls are suppressed.
- **`SilentMode = False`**: Set when the process completes or is stopped.
- **Individual runs**: When running a module standalone (e.g., `GenerateCalculationSheets` directly), `SilentMode` defaults to `False`, so all MsgBox dialogs appear normally.

### Modules with SilentMode guards

| Module | MsgBox Count Guarded |
|--------|---------------------|
| `AssessmentData.bas` | 1 |
| `HTMLQuery.bas` | 3 |
| `CalculationSheets.bas` | 11 |

---

## Cross-Platform Notes

The system supports both **Mac** and **Windows**:

- HTTP requests use `#If Mac Then` conditional compilation
  - Mac: `MacScript("do shell script ""curl ...""")` via AppleScript
  - Windows: `MSXML2.XMLHTTP` / `MSXML2.ServerXMLHTTP` COM objects
- Path separators handled via `Application.PathSeparator`
- `LecturerRefresh.bas` uses Mac-compatible 2D arrays instead of Collections for return types

---

## Troubleshooting

### Power Automate flow doesn't trigger

1. Check network connectivity
2. Verify the API URL hasn't been regenerated (check `SubjectListRefresh.bas` and `TeachingStreamRefresh.bas` for hardcoded URLs)
3. Ensure the year in `C2` is valid (≥ 2025)
4. Check if the Power Automate flow is turned on in the Power Automate portal

### Monitoring times out (30 minutes)

1. Check if the Power Automate flow ran successfully in the portal
2. Verify the Office Script updated `F2`/`F5` to "Done" on the Dashboard
3. The `progress_bar` table may not have a matching row for "Subject List" or "Teaching Stream"
4. Cloud sync issues — SharePoint may not be pushing updates to the local workbook

### Power Query returns errors / all "Failed" status

1. Verify `C2` (Year) is correct — handbook URLs are built with this year
2. Check if `https://handbook.unimelb.edu.au/{year}/subjects/` is accessible
3. If all subjects fail, the year may not yet be published on the handbook
4. The `Parameters` table (read by Power Query) must have the correct subject codes

### Calculation sheets not generated

1. Check the **Process Log** sheet for error details
2. Verify all required sheets exist: `Dashboard`, `SubjectList`, `assessment data parsed`, `teaching stream`
3. Check benchmark values in `C8`, `C9`, `C10` are positive numbers
4. Ensure no sheets are protected or locked

### LecturerRefresh fails in exported file

1. The source workbook path is hardcoded in `LecturerRefresh.bas` (line 23) — if the SharePoint folder moves, update it
2. The source workbook must be accessible (not locked by another user)
3. Workflow 3 timeout is 2 minutes — if the Teaching Matrix is very large, increase `maxWaitSeconds` in `WaitForWorkflow3Completion`

### MsgBox dialogs appear during integrated run

1. A module may have an unguarded `MsgBox` call — check that `If Not SilentMode Then` wraps it
2. `SilentMode` is declared in `Integration.bas` — other modules reference it as a Public global
3. If running from the exported file, `SilentMode` doesn't exist (only in source workbook)

### Export saves to wrong location

1. The export path is based on `ThisWorkbook.Path` — when running from SharePoint, this resolves to the cloud path
2. If the save fails, it falls back to `Application.DefaultFilePath` (usually Documents)
3. Check the Process Log for the exact save path used
