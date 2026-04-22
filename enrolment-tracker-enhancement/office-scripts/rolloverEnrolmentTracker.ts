// rolloverEnrolmentTracker.ts
// Office Script: Year Rollover for the Enrolment Tracker
// Invoked via a button on the Instructions sheet.
//
// What it does:
//   1. Reads current year from YEAR PLANNING!B1 (merged-cell-safe); newYear = current + 1
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
//   5. Clears all date columns (after Study Period) in Enrolment_Number and Prediction_Tool tables
//      Note: date-indicator row 3 formulas are self-maintaining — no update needed each rollover

function main(workbook: ExcelScript.Workbook) {
  // ── 1. Read current year from YEAR PLANNING ──────────────────────────────
  // B1 contains the current planning year but may be a merged cell; getValue() returns
  // empty for non-top-left cells of a merge, so scan A1:F3 as a fallback.
  const yearPlanningSheet = workbook.getWorksheet("YEAR PLANNING");
  if (!yearPlanningSheet) {
    console.log("Error: 'YEAR PLANNING' sheet not found.");
    return;
  }

  let currentYearRaw = yearPlanningSheet.getRange("B1").getValue();
  if (!currentYearRaw || isNaN(Number(currentYearRaw))) {
    const scanVals = yearPlanningSheet.getRange("A1:F3").getValues();
    let found = false;
    for (const row of scanVals) {
      if (found) break;
      for (const v of row) {
        const n = Number(v);
        if (v !== null && v !== "" && !isNaN(n) && n > 2020 && n < 2100) {
          currentYearRaw = v;
          console.log(`Current year ${currentYearRaw} found in YEAR PLANNING via merged-cell scan.`);
          found = true;
          break;
        }
      }
    }
  } else {
    console.log(`Current year ${currentYearRaw} read from YEAR PLANNING!B1.`);
  }

  if (!currentYearRaw || isNaN(Number(currentYearRaw))) {
    console.log("Error: Could not read current year from YEAR PLANNING. Ensure B1 (or the merged region around it) contains the current year (e.g. 2026).");
    return;
  }

  const newYear = Number(currentYearRaw) + 1;
  const prevYear = newYear - 1;            // = currentYearRaw
  const oldestYearToDelete = newYear - 3;  // e.g. for 2027: delete 2024
  const futureYear = newYear + 2;          // e.g. for 2027: add 2029

  console.log(`Rollover: currentYear=${currentYearRaw}, newYear=${newYear}, delete=${oldestYearToDelete}, future=${futureYear}`);

  // ── 2. Update YEAR PLANNING sheet ────────────────────────────────────────
  yearPlanningSheet.getRange("B1").setValue(newYear);
  console.log(`✓ Updated YEAR PLANNING!B1 to ${newYear}`);

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
    { from: `${prevYear}\nDev Hrs`,       to: `${newYear} Dev Hrs`        }, // handle cell-wrapped header
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

  // 4e. Add new future year column (e.g. "2029 Enrol Est"), then paint format from newYear+1 Est.
  //     Add first, name it, then get both columns fresh by name for the copy — avoids stale
  //     range references that occur when using pre-insert indices after addColumn.
  trackerTable.addColumn(-1);
  const hdrVals: (string | number | boolean)[][] = trackerTable.getHeaderRowRange().getValues();
  hdrVals[0][hdrVals[0].length - 1] = `${futureYear} Enrol Est`;
  trackerTable.getHeaderRowRange().setValues(hdrVals);
  console.log(`✓ Added "${futureYear} Enrol Est" column`);

  try {
    const srcEstCol = trackerTable.getColumnByName(`${newYear + 1} Enrol Est`);
    const dstEstCol = trackerTable.getColumnByName(`${futureYear} Enrol Est`);
    if (srcEstCol && dstEstCol) {
      // copyFrom on freshly-added table columns is unreliable — copy properties manually
      // from the first data cell of the source column.
      const srcCell = srcEstCol.getRangeBetweenHeaderAndTotal().getCell(0, 0);
      const dstData = dstEstCol.getRangeBetweenHeaderAndTotal();
      dstData.setNumberFormat(srcCell.getNumberFormat());
      const srcFont = srcCell.getFormat().getFont();
      const dstFont = dstData.getFormat().getFont();
      dstFont.setColor(srcFont.getColor());
      dstFont.setBold(srcFont.getBold());
      dstFont.setItalic(srcFont.getItalic());
      dstFont.setSize(srcFont.getSize());
      try { dstData.getFormat().getFill().setColor(srcCell.getFormat().getFill().getColor()); } catch { /* transparent fill — skip */ }
      dstEstCol.getRange().getFormat().setColumnWidth(srcEstCol.getRange().getFormat().getColumnWidth());
      console.log(`✓ Format painted from "${newYear + 1} Enrol Est" to "${futureYear} Enrol Est"`);
    } else {
      console.log(`Warning: Could not find Est column(s) for format copy.`);
    }
  } catch (e) {
    console.log(`Warning: Could not copy format to "${futureYear} Enrol Est": ${(e as Error).message ?? e}`);
  }

  // 4f. Delete the now-stale future-3 Enrol Est column (was newYear+2 before rollover = prevYear+2)
  //     e.g. for 2027 rollover: delete "2028 Enrol Est" (it was 2 years out under 2026, now it's only 1 year out — keep it)
  //     Actually: delete what was "+3 years" from old year, i.e. prevYear+3 = newYear+2... but we just added that.
  //     The stale column to delete is "oldestYearToDelete + 2 Enrol Est" = (newYear-3)+2 = newYear-1 Enrol Est = prevYear Enrol Est
  //     e.g. for rollover to 2027: delete "2026 Enrol Est" (now redundant, replaced by actual 2026 Enrol data)
  deleteTableColumn(trackerTable, `${prevYear} Enrol Est`);

  // 4g. Delete any stale "Enrol Est" columns that are not newYear+1 or newYear+2.
  const keepEstCols = new Set([`${newYear + 1} Enrol Est`, `${futureYear} Enrol Est`]);
  const headersNow = trackerTable.getHeaderRowRange().getValues()[0] as string[];
  for (let i = headersNow.length - 1; i >= 0; i--) {
    const h = String(headersNow[i]);
    if (h.toLowerCase().includes("enrol est") && !keepEstCols.has(h)) {
      deleteTableColumn(trackerTable, h);
    }
  }

  // 4h. Unhide columns from newYear-2 through newYear+2 in case any were hidden.
  const colsToUnhide = [
    `${newYear - 2} Enrol`,
    `${newYear - 1} Enrol`,
    `${newYear} Enrol`,
    `${newYear + 1} Enrol Est`,
    `${futureYear} Enrol Est`,
  ];
  for (const colName of colsToUnhide) {
    try {
      const col = trackerTable.getColumnByName(colName);
      if (col) {
        col.getRange().getFormat().setColumnHidden(false);
        console.log(`✓ Unhid "${colName}"`);
      }
    } catch {
      // not found or already visible — skip
    }
  }

  // 4i. Set row text colour to red for: Subject Coordinator, any "Notes" column,
  //     and TE / A+ Configuration (case-insensitive match; skip if not found).
  const redColHeaders = trackerTable.getHeaderRowRange().getValues()[0] as string[];
  for (const hdr of redColHeaders) {
    const h = String(hdr).toLowerCase().trim();
    const isTarget =
      h === "subject coordinator" ||
      h.includes("notes") ||
      h.startsWith("te / a+ configuration");
    if (!isTarget) continue;
    try {
      const col = trackerTable.getColumnByName(String(hdr));
      if (col) {
        col.getRangeBetweenHeaderAndTotal().getFormat().getFont().setColor("#FF0000");
        console.log(`✓ Set red font on "${hdr}"`);
      }
    } catch {
      // not found or inaccessible — skip
    }
  }

  // ── 5. Clear date columns in Enrolment_Number and Prediction_Tool ─────────
  // Note: the date-indicator row (row 3) cells contain self-maintaining formulas that
  // auto-resolve to the last date column header (or "no data" when cleared) — no update needed.
  clearDateColumnsOnSheet(workbook, ENROL_NUMBER_SHEET);
  clearDateColumnsOnSheet(workbook, PRED_TOOL_SHEET);

  console.log("✓ Rollover complete!");
  console.log(`  Year updated to: ${newYear}`);
  console.log(`  Next rollover will read YEAR PLANNING!B1 = ${newYear} and roll to ${newYear + 1}.`);
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

/**
 * Clears all date columns (after Study Period) using raw sheet range deletion.
 * Works regardless of whether the data is in a formal Excel table.
 * Scans the first 6 rows to find the "Study Period" header, then deletes
 * all columns to its right, right-to-left to keep indices stable.
 */
function clearDateColumnsOnSheet(workbook: ExcelScript.Workbook, sheetName: string) {
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    console.log(`Warning: Sheet "${sheetName}" not found — skipping date column clear.`);
    return;
  }

  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    console.log(`Warning: Sheet "${sheetName}" is empty — skipping.`);
    return;
  }

  const usedColCount = usedRange.getColumnCount();
  const usedRowCount = usedRange.getRowCount();
  const usedStartCol = usedRange.getColumnIndex();

  const scanRows = Math.min(6, usedRowCount);
  let studyPeriodSheetCol = -1;

  outerLoop:
  for (let r = 0; r < scanRows; r++) {
    for (let c = 0; c < usedColCount; c++) {
      const val = String(usedRange.getCell(r, c).getValue()).trim();
      if (val === STUDY_PERIOD_COL) {
        studyPeriodSheetCol = usedStartCol + c;
        break outerLoop;
      }
    }
  }

  if (studyPeriodSheetCol === -1) {
    console.log(`Warning: "${STUDY_PERIOD_COL}" not found on "${sheetName}" — skipping.`);
    return;
  }

  const lastSheetCol = usedStartCol + usedColCount - 1;
  const colsToDelete = lastSheetCol - studyPeriodSheetCol;

  if (colsToDelete <= 0) {
    console.log(`No date columns to delete on "${sheetName}".`);
    return;
  }

  // Delete the entire block of date columns in one operation.
  const currentUsed = sheet.getUsedRange();
  const rows = currentUsed ? currentUsed.getRowCount() : 1;
  const firstDateCol = studyPeriodSheetCol + 1;
  sheet.getRangeByIndexes(0, firstDateCol, rows, colsToDelete)
       .delete(ExcelScript.DeleteShiftDirection.left);
  console.log(`✓ Cleared ${colsToDelete} date column(s) from "${sheetName}" (kept up to "${STUDY_PERIOD_COL}") in one operation`);
}
