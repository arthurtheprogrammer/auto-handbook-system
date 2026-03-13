# Test Cases — Auto Handbook System

Since this system runs inside Excel (VBA, Power Query, Office Scripts, Power Automate), traditional automated unit testing frameworks are not applicable. Instead, testing is done through **manual verification** against **golden output baselines** stored on SharePoint.

---

## Golden Output Baselines

Backups from previous successful runs (2026) are stored on the department **SharePoint**, not in this repository, as they contain sensitive university data. See [`tests/golden-outputs/README.md`](golden-outputs/README.md) for the SharePoint location and details.

**To test**: Run the system with Year = `2026` and compare your output against the 2026 backup on SharePoint. Key things to check: subject count, assessment parsing accuracy, formula correctness, and sheet structure.

---

## Manual Test Scenarios

### TC-01: Full Pipeline — Happy Path

| Step | Action | Expected Result |
| ---- | ------ | --------------- |
| 1 | Open source workbook, enter valid year in C2 | Year accepted |
| 2 | Click Run button | All status cells (F2–F6) progress to "Complete" |
| 3 | Wait for completion (~5–10 min) | Exported `.xlsm` appears in SharePoint folder |
| 4 | Open exported file | FHY and SHY sheets present with correct subjects |
| 5 | Verify subject count | Matches number of active, non-excluded subjects in Enrolment Tracker |
| 6 | Verify assessment data | Spot-check 5 subjects against handbook website |
| 7 | Verify lecturer assignments | Spot-check 5 subjects against Teaching Matrix |

### TC-02: Lecturer Refresh (Exported File)

| Step | Action | Expected Result |
| ---- | ------ | --------------- |
| 1 | Open exported calculation file | File opens with macros enabled |
| 2 | Add manual notes in column S and stream enrolments in column P | Data entered |
| 3 | Click Refresh button in L2 | Source Dashboard F5 turns **orange** ("Running...") |
| 4 | Wait for workflow to complete | F5 turns **green** ("Updated") when teaching stream data is refreshed |
| 5 | Verify columns L–O updated | Columns L–O update with latest Teaching Matrix data |
| 6 | Verify columns P and S | Manual entries preserved after refresh |
| 7 | Verify new lecturer appears | If Teaching Matrix was updated, new lecturer shows in column L |

### TC-03: Mac Compatibility

| Step | Action | Expected Result |
| ---- | ------ | --------------- |
| 1 | Open source workbook on Mac | File opens normally |
| 2 | Click Run button | Prompt offers cloud HTML download workflow or skip |
| 3a | Choose "Yes" (cloud workflow) | F3 shows "Running...", then "Complete" when workflow finishes (up to 10 min) |
| 3b | Choose "No" (skip) | F3 shows "Skipped" (grey); existing assessment data used |
| 4 | Wait for completion | Exported file generated |
| 5 | Verify exported file | FHY/SHY sheets present; assessment data is current (3a) or from previous run (3b) |

### TC-04: Invalid Year

| Step | Action | Expected Result |
| ---- | ------ | --------------- |
| 1 | Enter non-numeric value in C2 (e.g., "abc") | Error message: "Please enter a valid year" |
| 2 | Leave C2 blank | Error message displayed |
| 3 | Enter year below 2025 | Error message displayed |

### TC-05: Missing Source Files

| Step | Action | Expected Result |
| ---- | ------ | --------------- |
| 1 | Enter incorrect filename in C3 | Power Automate flow fails; status shows error |
| 2 | Enter incorrect filename in C5 | Power Automate flow fails; status shows error |

### TC-06: Assessment Parsing Edge Cases

| Subject Type | What to Check |
| ------------ | ------------- |
| Subject with no assessments on handbook | `assessment data parsed` shows empty row or "Failed" status |
| Subject with group assessment | Group size correctly parsed; assessment quantity = enrolment ÷ group size |
| Subject with exam | Exam column shows `Y`; marking hours use exam benchmark |
| Subject with in-class assessment | Word count and exam set to `0`; keywords detected correctly |
| Subject with unusual word count format (e.g., "1500–2000 words") | Word count captures a reasonable value |

### TC-07: Year Rollover

| Step | Action | Expected Result |
| ---- | ------ | --------------- |
| 1 | Update C2 to new academic year | Year accepted |
| 2 | Ensure new year's Enrolment Tracker and Teaching Matrix exist on SharePoint | Files accessible |
| 3 | Run full pipeline | Output reflects new year's subjects and handbook data |
| 4 | Compare subject count with previous year's golden output | Reasonable variation (±10–20 subjects) |

---

## Verification Checklist (Post-Run)

Use this checklist after any system modification:

- [ ] Exported file contains both FHY and SHY sheets
- [ ] Subject count matches active, non-excluded subjects in Enrolment Tracker
- [ ] Assessment data spot-check passes (5 subjects vs handbook)
- [ ] Lecturer assignments spot-check passes (5 subjects vs Teaching Matrix)
- [ ] Formulas in columns D, I, J, Q, R calculate correctly
- [ ] Marker block formulas (V, W, X, Z) pull correct INDEX/MATCH values
- [ ] Sheet protection is active (formula columns locked, working columns editable)
- [ ] Refresh button in exported file works and preserves user edits
- [ ] Process completes on both Mac and Windows (if applicable)
