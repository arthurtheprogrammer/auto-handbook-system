// rolloverEnrolmentTracker.ts
// Office Script: Year Rollover for the Enrolment Tracker
// Invoked via a button on the Instructions sheet.
//
// What it does:
//   1. Prompts user for the new upcoming year (e.g. 2027)
//   2. Updates YEAR PLANNING sheet cell B1 with the new year
//   3. Renames year-labelled columns in Enrolment_Tracker table:
//      - "[prevYear] Current Enrol"        → "[newYear] Current Enrol"
//      - "[prevYear] Lec / Sem Hrs"        → "[newYear] Lec / Sem Hrs"
//      - "[prevYear] Tutorial Hrs"         → "[newYear] Tutorial Hrs"
//      - "[prevYear] Dev Hrs"              → "[newYear] Dev Hrs"
//   4. Shifts historical enrol columns (keeps only ±2 years from newYear):
//      - Deletes the oldest year column (newYear - 3)
//      - Renames "[newYear] Enrol Est"     → "[newYear] Enrol"    (promoted from estimates column)
//      - Renames "[prevYear] Enrol"        → "[newYear-1] Enrol"  (already correct, from copy)
//      - Hard-copies "[prevYear] Enrol" values (strips formulas) so data survives the date-column clear
//      - Sets "[newYear] Enrol" column to the LET formula (pulls from Enrolment_Number)
//        · Returns blank when Enrolment_Number has no date columns (table shorter than col D)
//      - Adds a new empty "[newYear+2] Enrol Est" column
//      - Deletes the now-stale "[prevYear+2] Enrol Est" which is now +3 years out
//   5. Updates the date-indicator row (row 3, above the table header at row 4) for formula columns:
//      · Enrolment_Number columns: -3 Period Enrol, -2 Period Enrol, -1 Period Enrol, [newYear] Current Enrol, [newYear] Enrol
//      · Prediction_Tool columns:  -3 Period Prediction Tool, -2 Period Prediction Tool, -1 Period Prediction Tool, Current Prediction Tool
//      · Shows the last date column header from the source table, or "no data" when table has no date columns
//   6. Clears all date columns (after Study Period) in Enrolment_Number and Prediction_Tool tables

function main(workbook: ExcelScript.Workbook) {
  // ── 1. Prompt for new year ────────────────────────────────────────────────
  // Office Scripts don't have a native prompt dialog; use an input cell approach.
  // The button should first show instructions via a named range or fixed cell,
  // then read the year from a dedicated input cell (Instructions!B_ROLLOVER or similar).
  // For now we read from a named range "RolloverYear" if it exists, else fall back
  // to a hard-coded cell on the Instructions sheet.

  const instructionsSheet = workbook.getWorksheet("Instructions");
  if (!instructionsSheet) {
    console.log("Error: 'Instructions' sheet not found.");
    return;
  }

  // Read the year the admin typed into the designated input cell.
  // Merged-cell note: getValue() returns null/empty for non-top-left cells of a merged range.
  // If B1 is empty or invalid, scan A1:F2 to find the value within the merged area.
  const YEAR_INPUT_CELL = "B1";
  let newYearRaw = instructionsSheet.getRange(YEAR_INPUT_CELL).getValue();

  if (!newYearRaw || isNaN(Number(newYearRaw))) {
    const scanVals = instructionsSheet.getRange("A1:F2").getValues();
    let found = false;
    for (const row of scanVals) {
      for (const v of row) {
        const n = Number(v);
        if (v !== null && v !== "" && !isNaN(n) && n > 2020 && n < 2100) {
          newYearRaw = v;
          console.log(`Year found by scanning Instructions A1:F2 (merged cell): ${newYearRaw}`);
          found = true;
          break;
        }
      }
      if (found) break;
    }
  }

  if (!newYearRaw || isNaN(Number(newYearRaw))) {
    // Fall back: read current year from YEAR PLANNING and add 1.
    // Also scan a range in case that cell is also merged.
    const yearPlanningFallback = workbook.getWorksheet("YEAR PLANNING");
    if (yearPlanningFallback) {
      let currentYearVal = yearPlanningFallback.getRange("B1").getValue();
      if (!currentYearVal || isNaN(Number(currentYearVal))) {
        const fpScan = yearPlanningFallback.getRange("A1:F2").getValues();
        let found = false;
        for (const row of fpScan) {
          for (const v of row) {
            const n = Number(v);
            if (v !== null && v !== "" && !isNaN(n) && n > 2020 && n < 2100) {
              currentYearVal = v;
              found = true;
              break;
            }
          }
          if (found) break;
        }
      }
      if (currentYearVal && !isNaN(Number(currentYearVal))) {
        newYearRaw = Number(currentYearVal) + 1;
        console.log(`Instructions B1 empty — auto-detected new year as ${newYearRaw} (YEAR PLANNING year + 1)`);
      }
    }
  }

  if (!newYearRaw || isNaN(Number(newYearRaw))) {
    console.log(`Error: Could not determine the new year. Please type the new year (e.g. 2027) into Instructions!${YEAR_INPUT_CELL} and run again.`);
    return;
  }

  const newYear = Number(newYearRaw);
  const prevYear = newYear - 1;
  const oldestYearToDelete = newYear - 3;  // e.g. for 2027: delete 2024
  const futureYear = newYear + 2;           // e.g. for 2027: add 2029

  console.log(`Rollover: prevYear=${prevYear}, newYear=${newYear}, delete=${oldestYearToDelete}, future=${futureYear}`);

  // ── 2. Update YEAR PLANNING sheet ────────────────────────────────────────
  const yearPlanningSheet = workbook.getWorksheet("YEAR PLANNING");
  if (yearPlanningSheet) {
    yearPlanningSheet.getRange("B1").setValue(newYear);
    console.log(`✓ Updated YEAR PLANNING!B1 to ${newYear}`);
  } else {
    console.log("Warning: 'YEAR PLANNING' sheet not found — skipping B1 update.");
  }

  // ── 3. Rename year-labelled columns in Enrolment_Tracker ─────────────────
  const trackerTable = findTable(workbook, "Enrolment_Tracker");
  if (!trackerTable) {
    console.log("Error: Enrolment_Tracker table not found.");
    return;
  }

  const yearColumns = [
    { from: `${prevYear} Current Enrol`,  to: `${newYear} Current Enrol` },
    { from: `${prevYear} Lec / Sem Hrs`,  to: `${newYear} Lec / Sem Hrs` },
    { from: `${prevYear} Tutorial Hrs`,   to: `${newYear} Tutorial Hrs`  },
    { from: `${prevYear} Dev Hrs`,        to: `${newYear} Dev Hrs`        },
  ];

  renameTableColumns(trackerTable, yearColumns);

  // ── 4. Shift historical Enrol columns ────────────────────────────────────
  // Expected columns before rollover (for 2027 rollover, current year = 2026):
  //   2024 Enrol | 2025 Enrol | 2026 Enrol | 2026 Enrol Est | 2027 Enrol Est | 2028 Enrol Est
  // After rollover (new year = 2027):
  //   2025 Enrol | 2026 Enrol | 2027 Enrol [formula] | 2027 Enrol Est | 2028 Enrol Est | 2029 Enrol Est [empty]

  // 4a. Delete oldest year column (e.g. "2024 Enrol")
  deleteTableColumn(trackerTable, `${oldestYearToDelete} Enrol`);

  // 4b. Rename "${newYear} Enrol Est" → "${newYear} Enrol"
  //     The estimates column added during the previous rollover is now promoted to the active
  //     formula column for this year.  Must happen before the hard-copy and formula steps below.
  renameTableColumns(trackerTable, [{ from: `${newYear} Enrol Est`, to: `${newYear} Enrol` }]);

  // 4c. Rename current year Enrol column to carry forward if needed
  //     (If the column is still named "[prevYear] Enrol" from last year's rollover, rename it)
  //     Note: after archiving, the user's copy will already have prev year labelled correctly.

  // 4c.5. Hard-copy [prevYear] Enrol values before the date columns are cleared.
  //       The [prevYear] Enrol column was seeded with the LET formula during the previous
  //       rollover and currently references Enrolment_Number date columns.  Capturing its
  //       computed values now ensures the historical data survives step 6 (date-column clear).
  try {
    const prevEnrolCol = trackerTable.getColumnByName(`${prevYear} Enrol`);
    if (prevEnrolCol) {
      const prevEnrolRange = prevEnrolCol.getRangeBetweenHeaderAndTotal();
      const prevEnrolValues = prevEnrolRange.getValues();
      prevEnrolRange.setValues(prevEnrolValues); // writes back as hard values, strips any formula
      console.log(`✓ Hard copied ${prevYear} Enrol values (formulas stripped)`);
    } else {
      console.log(`Warning: "${prevYear} Enrol" column not found — skipping hard copy.`);
    }
  } catch {
    console.log(`Warning: Could not hard copy "${prevYear} Enrol" values.`);
  }

  // 4d. Set [newYear] Enrol column to the LET formula.
  //     The formula returns blank ("") when:
  //       • No matching row is found in Enrolment_Number (IFERROR catches the MATCH failure)
  //       • The result is 0 or empty
  //       • The result is non-numeric — which happens when Enrolment_Number has no date
  //         columns yet (table shorter than col D) and INDEX falls back to the Study Period
  //         column, returning a text string.  NOT(ISNUMBER(result)) catches this case and
  //         returns blank instead of surfacing the text or a downstream VALUE error.
  const newYearEnrolCol = findColumnIndex(trackerTable, `${newYear} Enrol`);
  if (newYearEnrolCol !== -1) {
    const dataRange = trackerTable.getColumnByName(`${newYear} Enrol`).getRangeBetweenHeaderAndTotal();
    const rowCount = dataRange.getRowCount();
    // Guard: if numCols <= 3 (Subject Code, Subject Name, Study Period only — no date columns yet),
    // return "" immediately. Without this, INDEX resolves to the Study Period column itself.
    const formula =
      `=LET(` +
      `numCols, COLUMNS(Enrolment_Number[#Data]),` +
      `result, IF(numCols <= 3, "", ` +
        `IFERROR(` +
          `INDEX(Enrolment_Number[#All], ` +
            `MATCH([@[Subject Code]]&[@[Study Period]], ` +
                  `Enrolment_Number[Subject Code]&Enrolment_Number[Study Period], 0) + 1, ` +
            `numCols), ` +
          `"")), ` +
      `IF(OR(result="", result=0), "", result))`;

    for (let r = 0; r < rowCount; r++) {
      dataRange.getCell(r, 0).setFormula(formula);
    }
    console.log(`✓ Set ${newYear} Enrol formula in ${rowCount} rows`);
  } else {
    console.log(`Warning: Column "${newYear} Enrol" not found — skipping formula seeding.`);
  }

  // 4e. Add new future year column (e.g. "2029 Enrol Est") — empty
  trackerTable.addColumn(-1);
  const addedHeaders = trackerTable.getHeaderRowRange().getValues()[0] as string[];
  const lastIdx = addedHeaders.length - 1;
  const headerRange = trackerTable.getHeaderRowRange();
  const headerVals = headerRange.getValues();
  headerVals[0][lastIdx] = `${futureYear} Enrol Est`;
  headerRange.setValues(headerVals);
  console.log(`✓ Added empty "${futureYear} Enrol Est" column`);

  // 4f. Delete the now-stale future-3 Enrol Est column (was newYear+2 before rollover = prevYear+2)
  //     e.g. for 2027 rollover: delete "2028 Enrol Est" (it was 2 years out under 2026, now it's only 1 year out — keep it)
  //     Actually: delete what was "+3 years" from old year, i.e. prevYear+3 = newYear+2... but we just added that.
  //     The stale column to delete is "oldestYearToDelete + 2 Enrol Est" = (newYear-3)+2 = newYear-1 Enrol Est = prevYear Enrol Est
  //     e.g. for rollover to 2027: delete "2026 Enrol Est" (now redundant, replaced by actual 2026 Enrol data)
  deleteTableColumn(trackerTable, `${prevYear} Enrol Est`);

  // ── 5. Update date-indicator row (row 3, above table header at row 4) ─────
  const enrolNumberCols = [
    "-3 Period Enrol",
    "-2 Period Enrol",
    "-1 Period Enrol",
    `${newYear} Current Enrol`,
    `${newYear} Enrol`,
  ];
  for (const colName of enrolNumberCols) {
    updateDateIndicatorRow(trackerTable, colName, "Enrolment_Number");
  }
  const predToolCols = [
    "-3 Period Prediction Tool",
    "-2 Period Prediction Tool",
    "-1 Period Prediction Tool",
    "Current Prediction Tool",
  ];
  for (const colName of predToolCols) {
    updateDateIndicatorRow(trackerTable, colName, "Prediction_Tool");
  }

  // ── 6. Clear date columns in Enrolment_Number and Prediction_Tool ─────────
  clearDateColumns(workbook, "Enrolment_Number", ENROL_NUMBER_SHEET);
  clearDateColumns(workbook, "Prediction_Tool", PRED_TOOL_SHEET);

  console.log("✓ Rollover complete!");
  console.log(`  Year updated to: ${newYear}`);
  console.log(`  Reminder: Clear the Instructions!${YEAR_INPUT_CELL} cell after rollover.`);
}

// ── Helpers ───────────────────────────────────────────────────────────────────

const STUDY_PERIOD_COL = "Study Period";
const ENROL_NUMBER_SHEET = "Enrolment Number Tracker";
const PRED_TOOL_SHEET = "Prediction Tool Tracker";

function findTable(workbook: ExcelScript.Workbook, tableName: string): ExcelScript.Table | undefined {
  for (const ws of workbook.getWorksheets()) {
    const tbl = ws.getTable(tableName);
    if (tbl) return tbl;
  }
  return undefined;
}

function findColumnIndex(table: ExcelScript.Table, columnName: string): number {
  const headers = table.getHeaderRowRange().getValues()[0] as string[];
  return headers.indexOf(columnName);
}

function renameTableColumns(
  table: ExcelScript.Table,
  renames: { from: string; to: string }[]
) {
  const headerRange = table.getHeaderRowRange();
  const headers = headerRange.getValues()[0] as string[];

  for (const { from, to } of renames) {
    const idx = headers.indexOf(from);
    if (idx !== -1) {
      headers[idx] = to;
      console.log(`✓ Renamed column "${from}" → "${to}"`);
    } else {
      console.log(`Warning: Column "${from}" not found — skipping rename.`);
    }
  }

  headerRange.setValues([headers]);
}

function deleteTableColumn(table: ExcelScript.Table, columnName: string) {
  try {
    const col = table.getColumnByName(columnName);
    if (col) {
      col.delete();
      console.log(`✓ Deleted column "${columnName}"`);
    }
  } catch {
    console.log(`Warning: Column "${columnName}" not found — skipping delete.`);
  }
}

function clearDateColumns(workbook: ExcelScript.Workbook, tableName: string, sheetFallback?: string) {
  let table = findTable(workbook, tableName);
  if (!table && sheetFallback) {
    const sheet = workbook.getWorksheet(sheetFallback);
    if (sheet) {
      const tables = sheet.getTables();
      if (tables.length > 0) {
        table = tables[0];
        console.log(`Info: Table "${tableName}" not found by name; using first table on "${sheetFallback}".`);
      }
    }
  }
  if (!table) {
    console.log(`Warning: Table "${tableName}" not found — skipping clear.`);
    return;
  }

  const headers = table.getHeaderRowRange().getValues()[0] as string[];
  const studyPeriodIdx = headers.indexOf(STUDY_PERIOD_COL);

  if (studyPeriodIdx === -1) {
    console.log(`Warning: "${STUDY_PERIOD_COL}" column not found in ${tableName} — skipping clear.`);
    return;
  }

  // Delete all columns after Study Period (right to left to avoid index shift)
  const totalCols = headers.length;
  for (let i = totalCols - 1; i > studyPeriodIdx; i--) {
    try {
      table.getColumnByIndex(i).delete();
    } catch {
      // ignore
    }
  }
  console.log(`✓ Cleared all date columns from "${tableName}" (kept up to Study Period)`);
}

/**
 * Writes a formula to the date-indicator row immediately above the table header
 * (row 3 when the table header is at row 4).  The formula shows:
 *   • The last date-column header from sourceTableName  (when date columns exist), or
 *   • "no data"                                         (when the table has no date columns
 *                                                        beyond Study Period)
 *
 * This prevents VALUE errors and misleading text that appear after the source table
 * has been cleared and INDEX falls back to a non-numeric column.
 *
 * Skips silently if the column is not found or there is no row above the header.
 */
function updateDateIndicatorRow(
  table: ExcelScript.Table,
  columnName: string,
  sourceTableName: string
) {
  try {
    const col = table.getColumnByName(columnName);
    if (!col) {
      console.log(`Warning: Column "${columnName}" not found — skipping date indicator update.`);
      return;
    }

    // Step one row up from the header cell into the date-indicator row.
    const colIndex = col.getIndex();
    const headerCell = table.getHeaderRowRange().getCell(0, colIndex);
    const dateIndicatorCell = headerCell.getOffsetRange(-1, 0);

    const dateFormula =
      `=IFERROR(` +
        `IF(` +
          `COLUMNS(${sourceTableName}[#Data])<=MATCH("${STUDY_PERIOD_COL}",${sourceTableName}[#Headers],0),` +
          `"no data",` +
          `INDEX(${sourceTableName}[#Headers],1,COLUMNS(${sourceTableName}[#Data]))` +
        `),` +
        `"no data"` +
      `)`;

    dateIndicatorCell.setFormula(dateFormula);
    console.log(`✓ Updated date indicator for "${columnName}" (shows last date or "no data")`);
  } catch {
    console.log(`Info: No date indicator row above "${columnName}" header — skipping.`);
  }
}
