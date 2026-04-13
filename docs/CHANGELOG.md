# Changelog

All notable changes to this project will be documented in this file.

The format follows [Keep a Changelog](https://keepachangelog.com/).

---

## [1.3.1] — 2026-04-13

### Fixed

- **Lecturer Refresh Logic** — Refactored `LecturerRefresh.bas` (the macro embedded in exported calculation sheets) to correctly consume the synchronous Power Automate workflow. Replaced the redundant 45-second polling delay with a pre-execution blocking warning dialog, and implemented Office Script `ERROR:` prefix evaluation on the synchronous HTTP response payload.

---

## [1.3.0] — 2026-04-10

### Planned

- **Timing Column** — Add a "Timing" column derived from the assessment data sheets to display when an assessment occurs (e.g., "Week 4") next to the "Assessment Details" column in both the base section and Marker Blocks. This will help markers manage their weekly marking capacities relative to their teaching arrangements.

### Changed

- **Synchronous HTTP Response Parsing (`SubjectListRefresh.bas`, `TeachingStreamRefresh.bas`)** — Upgraded from polling arrays to synchronous blocking HTTP calls. The flows now return their payload directly in the HTTP 200 response body, which VBA parses (stripping JSON quotes and checking for exact `"ERROR:"` prefixes) for instant status evaluation, eliminating redundant SharePoint polling.
- **HTML Query Execution (`HTMLQuery.bas`)** — Handled the 120-second API limit for synchronous Power Automate HTTP requests explicitly by preserving the dedicated 10-minute `F3` polling pattern on Mac, ensuring 5-minute workflows complete reliably without 504 Timeouts.
- **Stream # Column Visibility** — Unhid the "Stream #" (Column N) in generated calculation sheets so it is prominently visible to users.
- **Handbook URL Generation** — Subject Codes in output sheets are now dynamically hyperlinked to their actual handbook pages for the current year (maintaining black, bolded text styling).
- **Export Filenames** — Calculation workbook exports now follow the `[YEAR]_M&M_Marking Admin Support Calculations.xlsm` nomenclature.

### Added

- **`TestAsyncPowerAutomate.bas`** — Added a standalone test module demonstrating detached OS-level HTTP execution (`curl &` on Mac, `WScript.Shell` on Windows) for firing hundreds of simultaneous Power Automate queries without freezing Excel.
- **"Using the Marking Support Output" guide (User Guide)** — Step-by-step walkthrough for using the generated calculation spreadsheet: entering stream enrolments, understanding the marking hours formula, verifying assessment details, handling special cases (class participation, midterms, missing word counts), logging academic calculations in Marker Blocks for compliance checks, and adjusting for extra marking commitments.

---

## [1.2.0] — 2026-03-25

### Added

- **`CheckWorkflowError()` in `Integration.bas`** — VBA poll loop now detects `"Error"` status written by Office Scripts, colours the Dashboard cell red, and exits the monitoring loop immediately without waiting for the 30-minute timeout
- **`SubjectListErrored` / `TeachingStreamErrored` globals** — track per-workflow error state; if either is set, `RunAllMacros` is skipped and a `vbCritical` MsgBox is shown listing the failed workflows

### Changed

- **`subjectListParser.osts` and `teachingStreamParser.osts`** — `catch` blocks now write `"Error"` to the `progress_bar` status cell (via a new `writeProgressStatus` helper) before returning. Previously, script failures left the status stuck on `"Running..."` and caused silent 30-minute VBA timeouts
- **`writeProgressStatus` helper** — shared progress-bar writer extracted into a reusable helper in both Office Scripts; success and error paths now use the same function
- `MonitorAndExecute` exits as soon as both workflows are resolved (complete **or** errored) rather than only when both are complete

---

## [1.1.0] — 2026-03-13

### Added

- **Mac HTML download workflow** — Power Automate cloud fallback for assessment HTML scraping on Mac (monitors F3 for completion, 10-min timeout, option to skip)
- Assessment query Power Automate flow definition and workflow diagram
- Subject code payload validation in Office Scripts

### Fixed

- **LecturerRefresh race condition** — stale F5 status from prior runs no longer causes false early completion. F5 is now cleared to "Running..." (orange) before triggering the workflow and set to "Updated" (green) on completion

### Changed

- Standardised status text and completion messages across all VBA modules
- Updated test cases for lecturer refresh F5 status behaviour and Mac cloud workflow

---

## [1.0.0] — 2026-03-13

### Added

- Full end-to-end automated pipeline: subject list extraction, teaching stream extraction, assessment web scraping, output generation
- Power Automate flows with Office Script parsers for SharePoint data extraction
- Power Query for handbook assessment HTML scraping
- VBA orchestrator (`Integration.bas`) with cross-platform Mac/Windows support
- Calculation sheet generation (FHY/SHY) with formulas, sheet protection, and marker blocks
- Exported workbook with embedded `LecturerRefresh.bas` for mid-semester updates
- Dashboard with parameter inputs, status tracking, and email notification
- Documentation: README, Design Doc, User Guide, Developer Guide, test cases
