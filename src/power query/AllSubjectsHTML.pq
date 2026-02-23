// ===============================================================
// Query: AllSubjectsHTML
// Purpose: Fetch the assessment table HTML from each subject's
//          handbook page and return it as a table, ready for
//          VBA-side parsing in AssessmentData.bas
// Trigger: Refreshed by HTMLQuery.RefreshHTMLDataM() in VBA
// Output:  Table with columns: SubjectCode, URL, AssessmentTableHTML,
//          HTMLLength, Status, ErrorMessage, FetchTime
// Dependencies:
//   - Parameters table in this workbook (Year + SubjectCodes rows)
//   - handbook.unimelb.edu.au (public, no auth required)
// ===============================================================
let
    // =========================================================================
    // PARAMETERS — read Year and SubjectCodes from the workbook's Parameters table
    // =========================================================================
    ParametersTable = Excel.CurrentWorkbook(){[Name = "Parameters"]}[Content],
    // Default to 2025 if the Year parameter is missing
    YearValue =
        if Table.HasColumns(ParametersTable, "Value") and Table.RowCount(ParametersTable) > 0 then
            Text.Trim(ParametersTable{0}[Value])
        else
            "2025",
    // Comma-separated subject codes (e.g. "MGMT20001,BUSA30000")
    SubjectCodesText =
        if Table.HasColumns(ParametersTable, "Value") and Table.RowCount(ParametersTable) > 1 then
            Text.Trim(ParametersTable{1}[Value])
        else
            "",
    SubjectCodesList =
        if SubjectCodesText = "" then
            {}
        else
            List.Transform(Text.Split(SubjectCodesText, ","), each Text.Trim(_)),
    BaseUrl = "https://handbook.unimelb.edu.au/" & YearValue & "/subjects/",
    // =========================================================================
    // FETCH & PARSE — isolate the assessment-table <div> from each page's HTML
    // =========================================================================
    FetchSubjectHTML = (subjectCode as text) as record =>
        let
            LowerCode = Text.Lower(Text.Trim(subjectCode)),
            Url = BaseUrl & LowerCode & "/assessment",
            TryFetch =
                try
                    let
                        SourceHtml = Text.FromBinary(Web.Contents(Url)),
                        // Look for the assessment-table div — the only section we need
                        TableStartDelimiter = "<div class=""assessment-table"">",
                        TableStartPosition = Text.PositionOf(SourceHtml, TableStartDelimiter),
                        IsolatedTableHtml =
                            if TableStartPosition > -1 then
                                let
                                    FromStart = Text.Range(SourceHtml, TableStartPosition),
                                    // 10000 chars is enough to capture the full assessment table
                                    LargeChunk = Text.Start(FromStart, 10000),
                                    TableEnd = Text.PositionOf(LargeChunk, "</table>"),
                                    AfterTable = if TableEnd > -1 then Text.Range(LargeChunk, TableEnd) else LargeChunk,
                                    NextDivEnd = Text.PositionOf(AfterTable, "</div>"),
                                    FinalHtml =
                                        if NextDivEnd > -1 then
                                            Text.Start(LargeChunk, TableEnd + NextDivEnd + 6)
                                        else
                                            LargeChunk
                                in
                                    FinalHtml
                            else
                                "",
                        ExtractedHtml = IsolatedTableHtml
                    in
                        ExtractedHtml
                    otherwise null,
            // Anything under 100 chars is unlikely to be a valid assessment table
            Result =
                if TryFetch <> null and Text.Length(TryFetch) > 100 then
                    [
                        SubjectCode = Text.Upper(subjectCode),
                        URL = Url,
                        AssessmentTableHTML = TryFetch,
                        HTMLLength = Text.Length(TryFetch),
                        Status = "Success",
                        ErrorMessage = null,
                        FetchTime = DateTime.LocalNow()
                    ]
                else
                    [
                        SubjectCode = Text.Upper(subjectCode),
                        URL = Url,
                        AssessmentTableHTML = TryFetch,
                        HTMLLength = if TryFetch <> null then Text.Length(TryFetch) else 0,
                        Status = "Failed",
                        ErrorMessage = if TryFetch = null then
                            "Could not fetch URL"
                        else
                            "Assessment table not found or too short",
                        FetchTime = DateTime.LocalNow()
                    ]
        in
            Result,
    ProcessedData = List.Transform(SubjectCodesList, each FetchSubjectHTML(_)),
    // Convert list of records to table
    ResultTable = Table.FromRecords(ProcessedData)
in
    ResultTable
