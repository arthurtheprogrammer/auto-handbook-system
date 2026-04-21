// updatePredictionAndEnrolment.ts
// Office Script: Updates BOTH the Prediction_Tool AND Enrolment_Number tables
// Triggered by Power Automate flow
// Parameters:
//   predictionData: JSON string of prediction/enrolment records
//   dateColumn: The date column header (e.g. "2026-03-16") to write into
//
// Behaviour:
//   - Prediction_Tool: always overwrite existing date column data
//   - Enrolment_Number: skip if the date column already exists (preserves StudentOne data)

function main(workbook: ExcelScript.Workbook, predictionData: string, dateColumn: string) {
    console.log("Script started");
    console.log("Date column parameter: " + dateColumn);
    console.log("Prediction data length: " + predictionData.length);
    console.log("First 200 chars of prediction data: " + predictionData.substring(0, 200));

    try {
        // Parse the JSON data
        let dataArray: object[];
        try {
            dataArray = JSON.parse(predictionData);
            console.log(`Parsed ${dataArray.length} records from JSON`);
        } catch (error) {
            console.log("Error parsing prediction data: " + error);
            return;
        }

        // Update Prediction_Tool table (always overwrite)
        console.log("\n--- Processing Prediction Tool ---");
        updateTable(workbook, dataArray, dateColumn, {
            tableName: "Prediction_Tool",
            worksheetName: "Prediction Tool Tracker",
            dataField: "Total Predicted Enrolments",
            logPrefix: "prediction",
            skipIfColumnExists: false
        });

        // Update Enrolment_Number table (skip if date column already exists)
        console.log("\n--- Processing Actual Enrolments ---");
        updateTable(workbook, dataArray, dateColumn, {
            tableName: "Enrolment_Number",
            worksheetName: "Enrolment Number Tracker",
            dataField: "Actual Current Year Enrolments",
            logPrefix: "actual enrolment",
            skipIfColumnExists: true
        });

    } catch (error) {
        console.log("Script error occurred:");
        console.log("Error message: " + (error as Error).message);
        console.log("Error name: " + (error as Error).name);
        console.log("Full error: " + JSON.stringify(error));
    }
}

interface TableUpdateConfig {
    tableName: string;
    worksheetName: string;
    dataField: string;
    logPrefix: string;
    skipIfColumnExists: boolean;
}

function updateTable(
    workbook: ExcelScript.Workbook,
    dataArray: object[],
    dateColumn: string,
    config: TableUpdateConfig
) {
    try {
        // Find the table - try specific worksheet first, then search all
        let targetWorksheet = workbook.getWorksheet(config.worksheetName);
        let targetTable: ExcelScript.Table | undefined;

        if (targetWorksheet) {
            console.log(`Found worksheet: ${config.worksheetName}`);
            targetTable = targetWorksheet.getTable(config.tableName);
        }

        // If not found, search all worksheets
        if (!targetTable) {
            console.log(`Searching for ${config.tableName} table in all worksheets...`);
            const worksheets = workbook.getWorksheets();
            for (const ws of worksheets) {
                const tables = ws.getTables();
                for (const tbl of tables) {
                    if (tbl.getName() === config.tableName) {
                        targetTable = tbl;
                        targetWorksheet = ws;
                        console.log(`Found ${config.tableName} table in worksheet: ${ws.getName()}`);
                        break;
                    }
                }
                if (targetTable) break;
            }
        }

        if (!targetTable || !targetWorksheet) {
            console.log(`Error: ${config.tableName} table not found in any worksheet`);
            console.log("Available worksheets: " + workbook.getWorksheets().map(ws => ws.getName()).join(", "));
            return;
        }
        console.log(`${config.tableName} table found successfully`);

        // Get table headers
        const headers = targetTable.getHeaderRowRange().getValues()[0] as string[];

        // Check if the date column exists
        let dateColumnIndex = headers.indexOf(dateColumn);

        // If skipIfColumnExists is true and column exists, skip the entire update
        if (config.skipIfColumnExists && dateColumnIndex !== -1) {
            console.log(`Date column '${dateColumn}' already exists in ${config.tableName} at index ${dateColumnIndex}`);
            console.log(`Skipping update for ${config.tableName} to keep StudentOne data`);
            return;
        }

        // If column doesn't exist, create it
        if (dateColumnIndex === -1) {
            targetTable.addColumn(-1);
            const newHeaders = targetTable.getHeaderRowRange().getValues()[0] as string[];
            const newColumnIndex = newHeaders.length - 1;

            const headerRange = targetTable.getHeaderRowRange();
            const headerValues = headerRange.getValues();
            headerValues[0][newColumnIndex] = dateColumn;
            headerRange.setValues(headerValues);

            dateColumnIndex = newColumnIndex;
            console.log(`✓ Added new date column: ${dateColumn} at index ${dateColumnIndex}`);
        } else {
            console.log(`✓ Date column ${dateColumn} exists at index ${dateColumnIndex} - proceeding with update`);
        }

        // Get all table data
        const tableRange = targetTable.getRangeBetweenHeaderAndTotal();
        const tableData = tableRange.getValues();

        // Find key column indices
        const subjectCodeIndex = headers.indexOf("Subject Code");
        const studyPeriodIndex = headers.indexOf("Study Period");
        const subjectNameIndex = headers.indexOf("Subject Name");

        // Validate required columns
        if (subjectCodeIndex === -1) {
            console.log(`Error: 'Subject Code' column not found in ${config.tableName} table`);
            return;
        }

        if (studyPeriodIndex === -1) {
            console.log(`Error: 'Study Period' column not found in ${config.tableName} table`);
            return;
        }

        console.log(`Subject Code column index: ${subjectCodeIndex}`);
        console.log(`Study Period column index: ${studyPeriodIndex}`);
        console.log(`Date column index: ${dateColumnIndex}`);

        // Process each record
        let updatesCount = 0;
        let skippedCount = 0;

        for (const record of dataArray) {
            const recordObj = record as Record<string, unknown>;

            const sourceSubjectCode: string = recordObj["Subject Code"]?.toString().trim() || "";
            const sourceSubjectName: string = recordObj["Subject Name"]?.toString().trim() || "";
            const sourceStudyPeriod: string = recordObj["Study Period"]?.toString().trim() || "";
            const dataValue: unknown = recordObj[config.dataField];

            if (!sourceSubjectCode || !sourceStudyPeriod) {
                continue;
            }

            // Skip if no data value for this field
            if (!dataValue || dataValue === "") {
                skippedCount++;
                continue;
            }

            const numberValue: number = parseInt(dataValue.toString());
            if (isNaN(numberValue)) {
                console.log(`Skipping record with invalid ${config.dataField} for ${sourceSubjectCode} - ${sourceStudyPeriod}: ${dataValue}`);
                continue;
            }

            // Try to find an existing row
            let matchFound = false;
            for (let rowIndex = 0; rowIndex < tableData.length; rowIndex++) {
                const tableSubjectCode: string = tableData[rowIndex][subjectCodeIndex]?.toString().trim() || "";
                const tableStudyPeriod: string = tableData[rowIndex][studyPeriodIndex]?.toString().trim() || "";

                if (tableSubjectCode === sourceSubjectCode && tableStudyPeriod === sourceStudyPeriod) {
                    // Update existing row's data for this date
                    tableData[rowIndex][dateColumnIndex] = numberValue;
                    matchFound = true;
                    updatesCount++;
                    console.log(`Updated ${config.logPrefix} for ${sourceSubjectCode} - ${sourceStudyPeriod}: ${numberValue}`);
                    break;
                }
            }

            // If no existing row, add a new one
            if (!matchFound) {
                console.log(`No row found for: ${sourceSubjectCode} - ${sourceStudyPeriod}. Adding new row.`);

                const newRow: (string | number | boolean)[] = new Array(headers.length).fill("");

                // Always set Subject Code (should be column A)
                newRow[subjectCodeIndex] = sourceSubjectCode;

                // Always set Subject Name (should be column B) if the column exists
                if (subjectNameIndex !== -1 && sourceSubjectName) {
                    newRow[subjectNameIndex] = sourceSubjectName;
                }

                // Always set Study Period
                newRow[studyPeriodIndex] = sourceStudyPeriod;

                // Set the data value for this date column
                newRow[dateColumnIndex] = numberValue;

                targetTable.addRow(-1, newRow);
                updatesCount++;
                console.log(`Added new row: ${sourceSubjectCode} | ${sourceSubjectName} | ${sourceStudyPeriod} with ${config.logPrefix}: ${numberValue}`);
            }
        }

        // Write updated data back for existing rows
        if (updatesCount > 0) {
            tableRange.setValues(tableData);
            console.log(`✓ Successfully updated or added ${updatesCount} ${config.logPrefix} records`);
            if (skippedCount > 0) {
                console.log(`Skipped ${skippedCount} records with no '${config.dataField}' data`);
            }
        } else {
            console.log(`No ${config.logPrefix} updates were made`);
            if (skippedCount > 0) {
                console.log(`All ${skippedCount} records had no '${config.dataField}' field or had empty values`);
            }
        }

    } catch (error) {
        console.log(`Error updating ${config.tableName} table:`);
        console.log("Error message: " + (error as Error).message);
    }
}
