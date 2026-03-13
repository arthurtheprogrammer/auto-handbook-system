# Changelog

All notable changes to this project will be documented in this file.

The format follows [Keep a Changelog](https://keepachangelog.com/).

---

## [Unreleased]

*No unreleased changes.*

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
