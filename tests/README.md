# Archived Outputs

Archived outputs from previous successful runs serve as **regression baselines** for verifying that system modifications don't break output correctness.

> **⚠️ These files are NOT stored in this repository** — they contain sensitive university data (student enrolments, staff names). They are stored on the department SharePoint.

## Where to Find Them

Previous year backups are stored in the **Auto Handbook System** folder on SharePoint:

```
SharePoint > TEACHING SUPPORT > Handbook > Auto Handbook System > backups
```

If you have access to the department SharePoint, you can compare your output against these baselines.

## What Gets Backed Up

For each run, a backup should include:

1. **Exported calculation file** — e.g., `2026 Marking Admin Support Calculations.xlsm`
2. **Source workbook snapshot** — make a copy of `Automated Handbook Data System.xlsm`after run (contains all populated data tables). Include the year in the filename (e.g., `2026 Automated Handbook Data System.xlsm`) and move it to the `backups` folder

## How to Use for Verification

After modifying any VBA module, Office Script, or Power Query:

1. Run the system with Year = `2026` (known good baseline)
2. Compare your new output against the 2026 backup on SharePoint
3. Check: subject count, assessment parsing accuracy, formula results, sheet structure
