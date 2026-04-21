// updateEnrolmentNumbers.ts
// Office Script: Updates the Enrolment_Number table with weekly enrolment data
// Triggered by Power Automate flow
// Parameters:
//   enrollmentData: JSON string of enrolment records from StudentOne
//   dateColumn: The date column header (e.g. "2026-03-16") to write into

function main(workbook: ExcelScript.Workbook, enrollmentData: string, dateColumn: string) {
  console.log("Script started");
  console.log("Date column parameter: " + dateColumn);
  console.log("Enrollment data length: " + enrollmentData.length);
  console.log("First 200 chars of data: " + enrollmentData.substring(0, 200));

  try {
    // Parse the enrollment data FIRST (like prediction script)
    let enrollmentArray: object[];
    try {
      enrollmentArray = JSON.parse(enrollmentData);
      console.log(`Parsed ${enrollmentArray.length} records from JSON`);
    } catch (error) {
      console.log("Error parsing enrollment data: " + error);
      return;
    }

    // Get the worksheet and table
    console.log("Looking for worksheet: Enrolment Number Tracker");
    const worksheet = workbook.getWorksheet("Enrolment Number Tracker");
    if (!worksheet) {
      console.log("Error: Worksheet 'Enrolment Number Tracker' not found");
      console.log("Available worksheets: " + workbook.getWorksheets().map(ws => ws.getName()).join(", "));
      return;
    }
    console.log("Worksheet found successfully");

    console.log("Looking for table: Enrolment_Number");
    const table = worksheet.getTable("Enrolment_Number");
    if (!table) {
      console.log("Error: Table 'Enrolment_Number' not found");
      console.log("Available tables: " + worksheet.getTables().map(t => t.getName()).join(", "));
      return;
    }
    console.log("Table found successfully");

    // Get table headers FIRST before checking column
    const headers = table.getHeaderRowRange().getValues()[0] as string[];

    // Check if the date column already exists, if not create it
    let dateColumnIndex = headers.indexOf(dateColumn);

    if (dateColumnIndex === -1) {
      // Add new column
      table.addColumn(-1);
      const newHeaders = table.getHeaderRowRange().getValues()[0] as string[];
      const newColumnIndex = newHeaders.length - 1;

      const headerRange = table.getHeaderRowRange();
      const headerValues = headerRange.getValues();
      headerValues[0][newColumnIndex] = dateColumn;
      headerRange.setValues(headerValues);

      dateColumnIndex = newColumnIndex;
      console.log(`✓ Added new date column: ${dateColumn} at index ${dateColumnIndex}`);
    } else {
      console.log(`✓ Date column ${dateColumn} already exists at index ${dateColumnIndex} - will overwrite`);
    }

    // Get all table data
    const tableRange = table.getRangeBetweenHeaderAndTotal();
    const tableData = tableRange.getValues();

    // Find key column indices
    const subjectCodeIndex = headers.indexOf("Subject Code");
    const studyPeriodIndex = headers.indexOf("Study Period");
    const subjectNameIndex = headers.indexOf("Subject Name");

    // Validate required columns
    if (subjectCodeIndex === -1) {
      console.log("Error: 'Subject Code' column not found in table");
      return;
    }

    if (studyPeriodIndex === -1) {
      console.log("Error: 'Study Period' column not found in table");
      return;
    }

    console.log(`Subject Code column index: ${subjectCodeIndex}`);
    console.log(`Study Period column index: ${studyPeriodIndex}`);
    console.log(`Date column index: ${dateColumnIndex}`);

    // Process each enrollment record
    let updatesCount = 0;
    let skippedCount = 0;

    for (const enrollmentRecord of enrollmentArray) {
      const recordObj = enrollmentRecord as Record<string, unknown>;

      // Use consistent field names matching the schema
      const sourceSubjectCode: string = recordObj["Study Package Cd"]?.toString().trim() || "";
      const sourceSubjectName: string = recordObj["Study Package Full Title"]?.toString().trim() || "";
      const sourceStudyPeriod: string = recordObj["Study Period"]?.toString().trim() || "";
      const currentEnrolled: unknown = recordObj["Current Number Enrolled"];

      if (!sourceSubjectCode || !sourceStudyPeriod) {
        continue;
      }

      // Skip if no enrollment data
      if (!currentEnrolled || currentEnrolled === "") {
        skippedCount++;
        continue;
      }

      const enrollmentNumber: number = parseInt(currentEnrolled.toString());
      if (isNaN(enrollmentNumber)) {
        console.log(`Skipping record with invalid enrollment number for ${sourceSubjectCode} - ${sourceStudyPeriod}: ${currentEnrolled}`);
        continue;
      }

      // Try to find existing row
      let matchFound = false;
      for (let rowIndex = 0; rowIndex < tableData.length; rowIndex++) {
        const tableSubjectCode: string = tableData[rowIndex][subjectCodeIndex]?.toString().trim() || "";
        const tableStudyPeriod: string = tableData[rowIndex][studyPeriodIndex]?.toString().trim() || "";

        if (tableSubjectCode === sourceSubjectCode && tableStudyPeriod === sourceStudyPeriod) {
          // Update existing row's enrollment for this date
          tableData[rowIndex][dateColumnIndex] = enrollmentNumber;
          matchFound = true;
          updatesCount++;
          console.log(`Updated enrollment for ${sourceSubjectCode} - ${sourceStudyPeriod}: ${enrollmentNumber}`);
          break;
        }
      }

      // If no existing row, add a new one
      if (!matchFound) {
        console.log(`No row found for: ${sourceSubjectCode} - ${sourceStudyPeriod}. Adding new row.`);

        const newRow: (string | number | boolean)[] = new Array(headers.length).fill("");

        // Always set Subject Code
        newRow[subjectCodeIndex] = sourceSubjectCode;

        // Set Subject Name if column exists
        if (subjectNameIndex !== -1 && sourceSubjectName) {
          newRow[subjectNameIndex] = sourceSubjectName;
        }

        // Always set Study Period
        newRow[studyPeriodIndex] = sourceStudyPeriod;

        // Set enrollment for this date column
        newRow[dateColumnIndex] = enrollmentNumber;

        table.addRow(-1, newRow);
        updatesCount++;
        console.log(`Added new row: ${sourceSubjectCode} | ${sourceSubjectName} | ${sourceStudyPeriod} with enrollment: ${enrollmentNumber}`);
      }
    }

    // Write updated data back for existing rows
    if (updatesCount > 0) {
      tableRange.setValues(tableData);
      console.log(`✓ Successfully updated or added ${updatesCount} enrollment records`);
      if (skippedCount > 0) {
        console.log(`Skipped ${skippedCount} records with no 'Current Number Enrolled' data`);
      }
    } else {
      console.log("No enrollment updates were made");
      if (skippedCount > 0) {
        console.log(`All ${skippedCount} records had no 'Current Number Enrolled' field or had empty values`);
      }
    }

  } catch (error) {
    console.log("Script error occurred:");
    console.log("Error message: " + (error as Error).message);
    console.log("Error name: " + (error as Error).name);
    console.log("Full error: " + JSON.stringify(error));
  }
}
