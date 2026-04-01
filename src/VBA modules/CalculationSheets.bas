'===============================================================
' Module: CalculationSheets
' Purpose: Generate FHY/SHY marking support calculation sheets
'          from subject, assessment, and teaching stream data
' Main Entry: GenerateCalculationSheets() - called by Integration.bas
' Output: Exported .xlsm workbook saved to SharePoint with
'         per-subject assessment blocks, lecturer rows, and formulas
' Author: Arthur Chen
' Repository: github.com/arthurtheprogrammer/auto-handbook-system
' Dependencies:
'   - SubjectList sheet (subject_list table)
'   - assessment data parsed sheet
'   - teaching stream sheet
'   - Dashboard sheet (year, benchmarks, file paths)
'   - Enrolment Tracker on SharePoint (external link formulas)
'===============================================================
Option Explicit

'===============================================================
' SECTION 1: CONFIGURATION & LOGGING
'===============================================================

Public Const ENABLE_REALTIME_LOG As Boolean = True
Public Const LOG_TO_STATUSBAR As Boolean = True

Dim wsLog As Worksheet
Private Const ENROLMENT_TRACKER_BASE As String = "https://unimelbcloud.sharepoint.com/teams/DepartmentofManagementMarketing-DepartmentOperations/Shared Documents/TEACHING MATRIX & ENROLMENT TRACKER/"

'---------------------------------------------------------------
' InitializeProcessLog
' Purpose: Create a fresh "Process Log" sheet to capture all
'          status messages during generation
' Called by: GenerateCalculationSheets
' Inputs:
'   - wb: the workbook to add the log sheet to
'---------------------------------------------------------------
Sub InitializeProcessLog(wb As Workbook)
    Dim processLogSheet As Worksheet
    Dim existingLog As Worksheet
    On Error Resume Next
    wb.Unprotect
    Set wsLog = Nothing
    
    ' Remove stale log sheet from any previous run
    Set existingLog = wb.Sheets("Process Log")
    If Not existingLog Is Nothing Then
        existingLog.Unprotect
        Application.DisplayAlerts = False
        existingLog.Delete
        Application.DisplayAlerts = True
        
        Set existingLog = Nothing
        Set existingLog = wb.Sheets("Process Log")
        If Not existingLog Is Nothing Then
            If Not SilentMode Then MsgBox "CRITICAL: Could not delete existing Process Log. It may be protected or in use.", vbCritical
            End
        End If
    End If
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    Set processLogSheet = wb.Sheets.add(After:=wb.Sheets(wb.Sheets.count))
    
    On Error Resume Next
    processLogSheet.name = "Process Log"
    
    ' Handle naming conflict — fall back to timestamped name
    If Err.Number <> 0 Then
        Dim errNum As Long
        errNum = Err.Number
        Err.Clear
        
        processLogSheet.name = "ProcessLog_" & Format(Now, "hhmmss")
        If Err.Number <> 0 Then
            If Not SilentMode Then MsgBox "CRITICAL: Could not name log sheet (Error " & errNum & "). Process aborted.", vbCritical
            processLogSheet.Delete
            Application.ScreenUpdating = True
            End
        Else
            If Not SilentMode Then MsgBox "Log sheet created with alternative name: " & processLogSheet.name, vbInformation
        End If
    End If
    
    On Error GoTo ErrorHandler
    
    ' Assign to module-level wsLog so LogMessage can write to it
    Set wsLog = Nothing
    Set wsLog = wb.Sheets(processLogSheet.name)
    If wsLog Is Nothing Then
        If Not SilentMode Then MsgBox "CRITICAL: Failed to verify log sheet creation.", vbCritical
        Application.ScreenUpdating = True
        End
    End If
    
    With wsLog
        .Cells(1, 1).Value = "Timestamp"
        .Cells(1, 2).Value = "Message"
        .Cells(1, 3).Value = "Elapsed Time"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 2).Font.Bold = True
        .Cells(1, 3).Font.Bold = True
        .Columns("A:A").ColumnWidth = 20
        .Columns("B:B").ColumnWidth = 80
        .Columns("C:C").ColumnWidth = 15
    End With
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    If Not SilentMode Then MsgBox "CRITICAL ERROR in InitializeProcessLog: " & Err.description & " (" & Err.Number & ")", vbCritical
    End
End Sub

'---------------------------------------------------------------
' LogMessage
' Purpose: Write a timestamped entry to the Process Log sheet
' Called by: Nearly every function in this module
' Inputs:
'   - msg: the message text
'   - elapsed (optional): seconds elapsed for performance logging
'---------------------------------------------------------------
Sub LogMessage(msg As String, Optional elapsedTime As Double = -1)
    ' Always write to Debug first (never fails)
    Debug.Print msg
    
    ' Try status bar
    On Error Resume Next
    If LOG_TO_STATUSBAR Then
        Application.StatusBar = "Processing: " & Left(msg, 100)
    End If
    
    ' Try writing to log sheet (don't crash if it fails)
    If Not wsLog Is Nothing Then
        Dim origScreenUpdating As Boolean
        origScreenUpdating = Application.ScreenUpdating
        
        If ENABLE_REALTIME_LOG Then Application.ScreenUpdating = True
        
        Dim NextRow As Long
        NextRow = wsLog.Cells(wsLog.Rows.count, 1).End(xlUp).Row + 1
        
        wsLog.Cells(NextRow, 1).Value = Now
        wsLog.Cells(NextRow, 2).Value = msg
        If elapsedTime >= 0 Then
            wsLog.Cells(NextRow, 3).Value = Format(elapsedTime, "0.00") & "s"
        End If
        wsLog.Cells(NextRow, 2).WrapText = True
        
        If ENABLE_REALTIME_LOG Then
            DoEvents
            Application.ScreenUpdating = origScreenUpdating
        End If
    End If
    
    On Error GoTo 0
End Sub

'---------------------------------------------------------------
' VerifyRequiredSheets
' Purpose: Confirm that all prerequisite sheets exist before
'          generation begins
' Called by: GenerateCalculationSheets
' Returns: True if all required sheets are present
'---------------------------------------------------------------
Function VerifyRequiredSheets(wb As Workbook) As Boolean
    Dim requiredSheets As Variant
    requiredSheets = Array("Dashboard", "SubjectList", "assessment data parsed", "teaching stream")
    
    Dim i As Integer
    Dim sheetExists As Boolean
    Dim sheetName As String
    Dim targetSheet As Worksheet
    
    LogMessage "=== Verifying Required Sheets ==="
    VerifyRequiredSheets = True
    
    For i = 0 To UBound(requiredSheets)
        sheetName = requiredSheets(i)
        sheetExists = False
        
        On Error Resume Next
        Set targetSheet = wb.Sheets(sheetName)
        If Err.Number = 0 And Not targetSheet Is Nothing Then
            sheetExists = True
            LogMessage "Sheet found: " & sheetName
        Else
            LogMessage "MISSING SHEET: " & sheetName
            VerifyRequiredSheets = False
        End If
        Err.Clear
        On Error GoTo 0
        
        Set targetSheet = Nothing
    Next i
    
    If Not VerifyRequiredSheets Then
        LogMessage "ERROR: Sheet verification failed. Missing required sheets."
        If Not SilentMode Then MsgBox "Required sheets are missing. Check Process Log for details.", vbCritical
    Else
        LogMessage "All required sheets verified."
    End If
End Function

'---------------------------------------------------------------
' SafeArrayIndex
' Purpose: Safely retrieve a field from a data array by position
'          index (e.g. subject arrays: 0=UID, 1=subject code, 2=subject name,
'          3=study period, 4=grouped period).
'          Returns a default value if the index is out of range.
'---------------------------------------------------------------
Function SafeArrayIndex(arr As Variant, idx As Integer, defaultVal As Variant) As Variant
    On Error Resume Next
    SafeArrayIndex = arr(idx)
    If Err.Number <> 0 Then SafeArrayIndex = defaultVal
    On Error GoTo 0
End Function

'===============================================================
' SECTION 2: MAIN WORKFLOW
'===============================================================

'---------------------------------------------------------------
' GenerateCalculationSheets
' Purpose: Main entry point — generates FHY and SHY calculation
'          sheets, then exports them to a new workbook on SharePoint
' Called by: Integration.RunAllMacros
'---------------------------------------------------------------
Sub GenerateCalculationSheets()
    Dim wb As Workbook
    Dim dashboardSheet As Worksheet
    Set wb = ThisWorkbook
    
    On Error Resume Next
    Set dashboardSheet = wb.Sheets("Dashboard")
    On Error GoTo 0
    
    If Not dashboardSheet Is Nothing Then
        With dashboardSheet.Range("F6")
            .Value = "Running..."
            .Interior.Color = RGB(255, 192, 0)
        End With
        DoEvents
    End If
    
    Dim tStart As Double, tPhase As Double
    tStart = Timer
    
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation
    Dim origEnableEvents As Boolean
    
    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation
    origEnableEvents = Application.EnableEvents
    
    On Error GoTo ErrorHandler
    
    wb.Unprotect
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Call InitializeProcessLog(wb)
    
    LogMessage "=== Starting GenerateCalculationSheets ==="
    LogMessage "Workbook: " & wb.name
    
    ' ========================================================================
    ' Clean up any existing calculation sheets BEFORE generation
    ' ========================================================================
    LogMessage "Checking for existing calculation sheets..."
    
    Dim existingFHY As Worksheet, existingSHY As Worksheet
    
    ' Delete FHY if exists
    Set existingFHY = Nothing
    On Error Resume Next
    Set existingFHY = wb.Sheets("FHY Calculations")
    On Error GoTo ErrorHandler
    
    If Not existingFHY Is Nothing Then
        Application.DisplayAlerts = False
        existingFHY.Delete
        Application.DisplayAlerts = True
        LogMessage "Deleted existing FHY Calculations sheet"
    End If
    
    ' Delete SHY if exists
    Set existingSHY = Nothing
    On Error Resume Next
    Set existingSHY = wb.Sheets("SHY Calculations")
    On Error GoTo ErrorHandler
    
    If Not existingSHY Is Nothing Then
        Application.DisplayAlerts = False
        existingSHY.Delete
        Application.DisplayAlerts = True
        LogMessage "Deleted existing SHY Calculations sheet"
    End If
    ' ========================================================================
    
    If Not VerifyRequiredSheets(wb) Then
        GoTo CleanExit
    End If
    
    Dim wordCountBenchmark As Double
    Dim examBenchmark As Double
    Dim markingSupportBenchmark As Double
    
    tPhase = Timer
    LogMessage "Getting benchmarks..."
    wordCountBenchmark = GetBenchmarkValue(wb, "Dashboard", "C8", 3000)
    examBenchmark = GetBenchmarkValue(wb, "Dashboard", "C9", 3)
    markingSupportBenchmark = GetBenchmarkValue(wb, "Dashboard", "C10", 20)
    
    If wordCountBenchmark <= 0 Or examBenchmark <= 0 Or markingSupportBenchmark <= 0 Then
        LogMessage "ERROR: Invalid benchmark values detected"
        LogMessage "  Word Count: " & wordCountBenchmark
        LogMessage "  Exam: " & examBenchmark
        LogMessage "  Marking Support: " & markingSupportBenchmark
        If Not SilentMode Then MsgBox "Invalid benchmark values in Dashboard sheet (cells C8, C9, C10).", vbCritical
        GoTo CleanExit
    End If
    
    LogMessage "Benchmarks loaded successfully", Timer - tPhase
    LogMessage "  Word Count: " & wordCountBenchmark & " words/hr"
    LogMessage "  Exam: " & examBenchmark & " exams/hr"
    LogMessage "  Marking Support: " & markingSupportBenchmark & " hrs/stream"
    
    tPhase = Timer
    LogMessage "Generating FHY sheet..."
    If Not GenerateSheet(wb, "FHY", wordCountBenchmark, examBenchmark, markingSupportBenchmark) Then
        LogMessage "ERROR: FHY sheet generation failed"
        GoTo CleanExit
    End If
    LogMessage "FHY sheet complete", Timer - tPhase
    
    tPhase = Timer
    LogMessage "Generating SHY sheet..."
    If Not GenerateSheet(wb, "SHY", wordCountBenchmark, examBenchmark, markingSupportBenchmark) Then
        LogMessage "ERROR: SHY sheet generation failed"
        GoTo CleanExit
    End If
    LogMessage "SHY sheet complete", Timer - tPhase
    
    If HasErrorsInLog() Then
        LogMessage "=== Errors detected - Export cancelled ==="
        GoTo CleanExit
    End If
    
    LogMessage "=== Generation Complete - Starting Export ==="
    
    tPhase = Timer
    If Not ExportCalculationSheets(wb) Then
        LogMessage "ERROR: Export failed"
        GoTo CleanExit
    End If
    LogMessage "Export complete", Timer - tPhase
    
    LogMessage "=== Total Time: " & Format(Timer - tStart, "0.00") & " seconds ==="

CleanExit:
    Application.Calculation = origCalculation
    Application.ScreenUpdating = origScreenUpdating
    Application.EnableEvents = origEnableEvents
    Application.StatusBar = False
    
    If HasErrorsInLog() Then
        If Not dashboardSheet Is Nothing Then
            With dashboardSheet.Range("F6")
                .Value = "Failed"
                .Interior.Color = RGB(255, 0, 0)
            End With
            DoEvents
        End If
        If Not SilentMode Then MsgBox "Errors were detected during generation." & vbCrLf & vbCrLf & _
               "Please check the 'Process Log' sheet for details." & vbCrLf & _
               "Files were NOT exported.", vbExclamation, "Generation Completed with Errors"
    Else
        If Not dashboardSheet Is Nothing Then
            With dashboardSheet.Range("F6")
                .Value = "Complete"
                .Interior.Color = RGB(146, 208, 80)
            End With
            DoEvents
        End If
        If Not SilentMode Then MsgBox "Calculation sheets generated and exported successfully!" & vbCrLf & vbCrLf & _
               "Total time: " & Format(Timer - tStart, "0.0") & " seconds" & vbCrLf & _
               "Check the 'Process Log' sheet for details.", vbInformation
    End If
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    Application.Calculation = origCalculation
    Application.ScreenUpdating = origScreenUpdating
    Application.EnableEvents = origEnableEvents
    Application.StatusBar = False
    
    If Not dashboardSheet Is Nothing Then
        With dashboardSheet.Range("F6")
            .Value = "Failed"
            .Interior.Color = RGB(255, 0, 0)
        End With
        DoEvents
    End If
    
    Dim errMsg As String
    errMsg = "Error in GenerateCalculationSheets: " & Err.description & " (Error " & Err.Number & ")"
    
    ' Use Debug.Print if LogMessage fails
    Debug.Print errMsg
    LogMessage errMsg
    
    Call CleanupPartialSheets(wb)
    
    If Not SilentMode Then MsgBox errMsg & vbCrLf & vbCrLf & "Check the 'Process Log' sheet for more details.", vbCritical
    On Error GoTo 0
End Sub

'---------------------------------------------------------------
' ExportCalculationSheets
' Purpose: Copy FHY/SHY sheets into a new workbook, attach the
'          LecturerRefresh VBA module, and save to SharePoint
' Called by: GenerateCalculationSheets
' Returns: True on successful export
'---------------------------------------------------------------
Function ExportCalculationSheets(wb As Workbook) As Boolean
    On Error GoTo ErrorHandler
    
    ExportCalculationSheets = False
    Application.ScreenUpdating = False
    LogMessage "ExportCalculationSheets: Starting"
    
    ' ========================================================================
    ' STEP 1: Validate inputs
    ' ========================================================================
    Dim yearValue As Variant
    yearValue = wb.Sheets("Dashboard").Range("C2").Value
    
    If Not IsNumeric(yearValue) Or yearValue = "" Then
        LogMessage "ERROR: Invalid year value"
        Exit Function
    End If
    
    Dim newFileName As String
    newFileName = CStr(yearValue) & "_M&M_Marking Admin Support Calculations"
    
    ' ========================================================================
    ' STEP 2: Find sheets
    ' ========================================================================
    Dim wsFHY As Worksheet, wsSHY As Worksheet
    On Error Resume Next
    Set wsFHY = wb.Sheets("FHY Calculations")
    Set wsSHY = wb.Sheets("SHY Calculations")
    On Error GoTo ErrorHandler
    
    If wsFHY Is Nothing And wsSHY Is Nothing Then
        LogMessage "ERROR: No calculation sheets found"
        Exit Function
    End If
    
    ' ========================================================================
    ' STEP 3: Create and populate new workbook
    ' ========================================================================
    Dim newWB As Workbook
    Set newWB = Workbooks.add
    LogMessage "Created new workbook"
    
    ' Clean up extra sheets
    Application.DisplayAlerts = False
    On Error Resume Next
    Do While newWB.Sheets.count > 1
        newWB.Sheets(newWB.Sheets.count).Delete
    Loop
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler
    
    ' Copy calculation sheets
    If Not wsFHY Is Nothing Then
        wsFHY.Copy Before:=newWB.Sheets(1)
        LogMessage "Copied FHY Calculations"
    End If
    
    If Not wsSHY Is Nothing Then
        wsSHY.Copy After:=newWB.Sheets(newWB.Sheets.count)
        LogMessage "Copied SHY Calculations"
    End If
    
    ' Remove any blank sheets
    Application.DisplayAlerts = False
    On Error Resume Next
    Dim i As Integer
    For i = newWB.Sheets.count To 1 Step -1
        If newWB.Sheets(i).name <> "FHY Calculations" And newWB.Sheets(i).name <> "SHY Calculations" Then
            newWB.Sheets(i).Delete
        End If
    Next i
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler
    
    If newWB.Sheets.count = 0 Then
        LogMessage "ERROR: No sheets in exported workbook"
        newWB.Close SaveChanges:=False
        Exit Function
    End If
    
    ' Force recalculate exported sheets
    Dim ws As Worksheet
    For Each ws In newWB.Worksheets
        ws.Calculate
    Next ws
    
    ' ========================================================================
    ' STEP 4: Copy VBA module FIRST (creates VBA project)
    ' ========================================================================
    LogMessage "Step 4: Checking for VBA module to copy..."
    
    Dim hasVBA As Boolean
    hasVBA = False
    
    If CopyVBAModuleIfExists(wb, newWB, "LecturerRefresh") Then
        LogMessage "Step 4: SUCCESS - VBA module copied"
        hasVBA = True
        
        Call AddRefreshButtonsToSheets(newWB)
    Else
        LogMessage "No VBA module found in source - will save as XLSX"
    End If
    
    ' ========================================================================
    ' STEP 5: Determine save path
    ' ========================================================================
    Dim savePath As String
    Dim folderPath As String
    
    folderPath = wb.Path
    If folderPath = "" Then
        LogMessage "ERROR: Could not determine workbook folder"
        newWB.Close SaveChanges:=False
        Exit Function
    End If
    
    ' Add appropriate path separator (works on both platforms)
    Dim pathSep As String
    pathSep = Application.PathSeparator  ' Automatically "/" on Mac, "\" on Windows
    
    If Right(folderPath, 1) <> pathSep Then
        folderPath = folderPath & pathSep
    End If
    
    ' Use appropriate extension based on whether VBA exists
    Dim fileExt As String
    Dim fileFormat As Long
    
    If hasVBA Then
        fileExt = ".xlsm"
        fileFormat = xlOpenXMLWorkbookMacroEnabled
        LogMessage "Will save as XLSM (has VBA)"
    Else
        fileExt = ".xlsx"
        fileFormat = xlOpenXMLWorkbook
        LogMessage "Will save as XLSX (no VBA)"
    End If
    
    savePath = folderPath & newFileName & fileExt
    LogMessage "Target save path: " & savePath
    
    ' ========================================================================
    ' STEP 6: Save directly to SharePoint
    ' ========================================================================
    LogMessage "Attempting to save..."
    
    On Error Resume Next
    newWB.SaveAs FileName:=savePath, fileFormat:=fileFormat
    
    If Err.Number <> 0 Then
        Dim saveErrNum As Long
        Dim saveErrDesc As String
        saveErrNum = Err.Number
        saveErrDesc = Err.description
        
        LogMessage "ERROR: SaveAs failed - " & saveErrNum & " " & saveErrDesc
        LogMessage "Failed path was: " & savePath
        LogMessage "Attempting fallback to temp location..."
        
        ' Fallback to temp
        Dim tempPath As String
        tempPath = Application.DefaultFilePath
        pathSep = Application.PathSeparator
        
        If Right(tempPath, 1) <> pathSep Then
            tempPath = tempPath & pathSep
        End If
        
        tempPath = tempPath & newFileName & fileExt
        
        Err.Clear
        newWB.SaveAs FileName:=tempPath, fileFormat:=fileFormat
        
        If Err.Number = 0 Then
            If Not SilentMode Then MsgBox "Could not save to SharePoint." & vbCrLf & vbCrLf & _
                   "File saved to:" & vbCrLf & tempPath, vbInformation
            savePath = tempPath
            LogMessage "Saved to temp location: " & tempPath
        Else
            LogMessage "ERROR: Save failed completely - " & Err.Number & " " & Err.description
            newWB.Close SaveChanges:=False
            Application.ScreenUpdating = True
            Exit Function
        End If
    Else
        LogMessage "SUCCESS: Saved to SharePoint"
    End If
    
    On Error GoTo ErrorHandler
    
    ' ========================================================================
    ' STEP 7: Close exported workbook
    ' ========================================================================
    newWB.Close SaveChanges:=False
    LogMessage "Exported workbook closed"
    
    ' ========================================================================
    ' STEP 8: Delete source sheets
    ' ========================================================================
    Application.DisplayAlerts = False
    On Error Resume Next
    If Not wsFHY Is Nothing Then wb.Sheets("FHY Calculations").Delete
    If Not wsSHY Is Nothing Then wb.Sheets("SHY Calculations").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = True
    LogMessage "Export completed successfully"
    ExportCalculationSheets = True
    Exit Function
    
ErrorHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    LogMessage "ERROR in ExportCalculationSheets: " & Err.Number & " - " & Err.description
    
    On Error Resume Next
    If Not newWB Is Nothing Then newWB.Close SaveChanges:=False
    On Error GoTo 0
    
    ExportCalculationSheets = False
End Function


'---------------------------------------------------------------
' GetBenchmarkValue
' Purpose: Read a benchmark number from a Dashboard cell,
'          falling back to a default if the cell is empty/invalid
' Called by: GenerateCalculationSheets
' Returns: The benchmark value (e.g. word count per hour)
'---------------------------------------------------------------
Function GetBenchmarkValue(wb As Workbook, sheetName As String, cellRef As String, defaultValue As Double) As Double
    On Error GoTo ErrorHandler
    
    Dim benchmarkSheet As Worksheet
    Set benchmarkSheet = wb.Sheets(sheetName)
    
    If benchmarkSheet Is Nothing Then
        LogMessage "WARNING: Sheet " & sheetName & " not found, using default: " & defaultValue
        GetBenchmarkValue = defaultValue
        Exit Function
    End If
    
    Dim cellValue As Variant
    cellValue = benchmarkSheet.Range(cellRef).Value
    
    If IsNumeric(cellValue) And cellValue <> "" Then
        GetBenchmarkValue = CDbl(cellValue)
    Else
        LogMessage "WARNING: Invalid value in " & sheetName & "!" & cellRef & ", using default: " & defaultValue
        GetBenchmarkValue = defaultValue
    End If
    Exit Function
    
ErrorHandler:
    LogMessage "ERROR in GetBenchmarkValue: " & Err.description
    GetBenchmarkValue = defaultValue
End Function

'---------------------------------------------------------------
' GenerateSheet
' Purpose: Create a single calculation sheet (FHY or SHY),
'          populate it with subject data, and apply formatting
' Called by: GenerateCalculationSheets
' Returns: True on success
'---------------------------------------------------------------
Function GenerateSheet(wb As Workbook, groupedPeriod As String, wordBench As Double, examBench As Double, markingSupportBench As Double) As Boolean
    On Error GoTo ErrorHandler
    
    GenerateSheet = False
    Dim tStart As Double
    tStart = Timer
    
    Dim sheetName As String
    sheetName = groupedPeriod & " Calculations"
    
    LogMessage "GenerateSheet: Starting for " & groupedPeriod
    
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets(sheetName).Delete
    Err.Clear
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    
    Dim checkSheet As Worksheet
    On Error Resume Next
    Set checkSheet = wb.Sheets(sheetName)
    If Not checkSheet Is Nothing Then
        LogMessage "ERROR: Could not delete existing " & sheetName & " sheet"
        On Error GoTo ErrorHandler
        GenerateSheet = False
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    Dim wsOutput As Worksheet
    Set wsOutput = wb.Sheets.add(After:=wb.Sheets(wb.Sheets.count - 1))
    wsOutput.name = sheetName
    
    Set checkSheet = Nothing
    Set checkSheet = wb.Sheets(sheetName)
    If checkSheet Is Nothing Then
        LogMessage "ERROR: Failed to verify creation of " & sheetName
        GenerateSheet = False
        Exit Function
    End If
    
    Call CreateHeaders(wsOutput)
    
    ' Benchmark labels in header rows
    Dim benchmarkCells As Variant
    benchmarkCells = Array("D2", "(automatic)", "J2", wordBench & " words/hr", "J3", examBench & " exams/hr", "P2", "(manual)", "Q2", markingSupportBench & " hrs/stream")
    
    Dim i As Long
    For i = LBound(benchmarkCells) To UBound(benchmarkCells) Step 2
        wsOutput.Range(benchmarkCells(i)).Value = benchmarkCells(i + 1)
    Next i
    
    ' Marker block benchmarks (repeated 3 times)
    Dim markerBenchmarkCols As Variant
    markerBenchmarkCols = Array("Z", "AI", "AS")
    
    For i = LBound(markerBenchmarkCols) To UBound(markerBenchmarkCols)
        wsOutput.Range(markerBenchmarkCols(i) & "2").Value = wordBench & " words/hr"
        wsOutput.Range(markerBenchmarkCols(i) & "3").Value = examBench & " exams/hr"
    Next i

    
    wsOutput.Columns("A:A").Hidden = True
    wsOutput.Range("E4").Select
    ActiveWindow.FreezePanes = True
    
    Dim subjectData As Collection
    Set subjectData = GetFilteredSubjectsWithAssessments(wb, groupedPeriod)
    
    If subjectData Is Nothing Then
        LogMessage "ERROR: Failed to retrieve subject data for " & groupedPeriod
        GenerateSheet = False
        Exit Function
    End If
    
    If subjectData.count = 0 Then
        LogMessage "WARNING: No subjects found for " & groupedPeriod
        GenerateSheet = True
        Exit Function
    End If
    
    LogMessage "Found " & subjectData.count & " subjects for " & groupedPeriod
    
    If Not PopulateSheetData(wb, wsOutput, subjectData, wordBench, examBench, markingSupportBench) Then
        LogMessage "ERROR: Failed to populate sheet data for " & groupedPeriod
        GenerateSheet = False
        Exit Function
    End If
    
    Call FormatSheet(wsOutput)
    
    LogMessage "GenerateSheet: Completed for " & groupedPeriod, Timer - tStart
    GenerateSheet = True
    Exit Function
    
ErrorHandler:
    LogMessage "ERROR in GenerateSheet (" & groupedPeriod & "): " & Err.description & " (Error " & Err.Number & ")"
    GenerateSheet = False
End Function

'---------------------------------------------------------------
' CreateHeaders
' Purpose: Write the column header row (A–S) for a calculation sheet
' Called by: GenerateSheet
'---------------------------------------------------------------
Sub CreateHeaders(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim headers As Variant
    headers = Array("UID", "Subject Code", "Study Period", "Enrolment", "Assessment Details", _
        "Word Count", "Exam", "Group Size", "Assessment Quantity", "Marking Hours", "Assessment Notes", _
        "Lecturer/Instructors", "Status", "Stream #", "Activity Code", _
        "Stream(s) Enrolment", "Allocated Marking", "Marking Support Hours Available", "Lecturer Notes", _
        "Marker 1", "Assessment Details", "Word Count", "Exam", "Group Size", _
        "Assessment Quantity", "Marking Allocation", "Email", "Arrangement Notes", "Contract Requested", _
        "Marker 2", "Assessment Details", "Word Count", "Exam", "Group Size", _
        "Assessment Quantity", "Marking Allocation", "Email", "Arrangement Notes", "Contract Requested", _
        "Marker 3", "Assessment Details", "Word Count", "Exam", "Group Size", _
        "Assessment Quantity", "Marking Allocation", "Email", "Arrangement Notes", "Contract Requested")
    
    Dim i As Integer
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
        ws.Cells(1, i + 1).Font.Bold = True
        ws.Cells(1, i + 1).WrapText = True
    Next i
    ws.Rows(1).RowHeight = 30
    
    Exit Sub
ErrorHandler:
    LogMessage "ERROR in CreateHeaders: " & Err.description
End Sub

'===============================================================
' SECTION 3: DATA RETRIEVAL
'===============================================================

'---------------------------------------------------------------
' GetFilteredSubjectsWithAssessments
' Purpose: Build a collection of active subjects for a given
'          grouped period (FHY or SHY), excluding flagged entries
' Called by: GenerateSheet
' Returns: Collection of subject arrays (UID, code, name,
'          study period, grouped period, row index)
'---------------------------------------------------------------
Function GetFilteredSubjectsWithAssessments(wb As Workbook, groupedPeriod As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim tStart As Double
    tStart = Timer
    
    Dim wsSubjects As Worksheet
    Dim wsAssessment As Worksheet
    Set wsSubjects = wb.Sheets("SubjectList")
    Set wsAssessment = wb.Sheets("assessment data parsed")
    
    If wsSubjects Is Nothing Or wsAssessment Is Nothing Then
        LogMessage "ERROR: Required sheets not found in GetFilteredSubjectsWithAssessments"
        Set GetFilteredSubjectsWithAssessments = Nothing
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = wsSubjects.Cells(wsSubjects.Rows.count, "B").End(xlUp).Row
    
    If lastRow < 2 Then
        LogMessage "WARNING: No data in SubjectList sheet"
        Set GetFilteredSubjectsWithAssessments = New Collection
        Exit Function
    End If
    
    Dim assessmentDict As Collection
    Set assessmentDict = New Collection
    
    Dim lastAssRow As Long
    lastAssRow = wsAssessment.Cells(wsAssessment.Rows.count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastAssRow
        Dim key As String
        key = wsAssessment.Cells(i, 2).Value & "|" & wsAssessment.Cells(i, 3).Value
        On Error Resume Next
        assessmentDict.add True, key
        Err.Clear
        On Error GoTo ErrorHandler
    Next i
    
    Dim colUID As Integer, colSubjectCode As Integer, colSubjectName As Integer
    Dim colGroupedPeriod As Integer, colStudyPeriod As Integer, colExclude As Integer
    
    colUID = FindColumn(wsSubjects, "UID (sorter)")
    colSubjectCode = FindColumn(wsSubjects, "Subject Code")
    colSubjectName = FindColumn(wsSubjects, "Subject Name")
    colGroupedPeriod = FindColumn(wsSubjects, "Grouped Period")
    colStudyPeriod = FindColumn(wsSubjects, "Study Period")
    colExclude = FindColumn(wsSubjects, "Exclude")
    
    If colSubjectCode = 0 Or colGroupedPeriod = 0 Or colStudyPeriod = 0 Then
        LogMessage "ERROR: Required columns not found in SubjectList"
        LogMessage "  Subject Code col: " & colSubjectCode
        LogMessage "  Grouped Period col: " & colGroupedPeriod
        LogMessage "  Study Period col: " & colStudyPeriod
        Set GetFilteredSubjectsWithAssessments = Nothing
        Exit Function
    End If
    
    Dim subjects As New Collection
    Dim includedCount As Long
    
    For i = 2 To lastRow
        Dim excludeVal As Variant
        excludeVal = wsSubjects.Cells(i, colExclude).Value
        If excludeVal = True Or UCase(CStr(excludeVal)) = "TRUE" Then GoTo NextIteration
        
        Dim groupedPeriodVal As String
        groupedPeriodVal = Trim(CStr(wsSubjects.Cells(i, colGroupedPeriod).Value))
        If groupedPeriodVal <> groupedPeriod Then GoTo NextIteration
        
        Dim subjectCode As String
        subjectCode = Trim(CStr(wsSubjects.Cells(i, colSubjectCode).Value))
        
        If Len(subjectCode) < 5 Then GoTo NextIteration
        
        If Mid(subjectCode, 5, 1) <> "9" And subjectCode <> "BUSA30000" Then GoTo NextIteration
        
        Dim studyPeriod As String
        studyPeriod = Trim(CStr(wsSubjects.Cells(i, colStudyPeriod).Value))
        
        Dim lookupKey As String
        lookupKey = subjectCode & "|" & studyPeriod
        
        Dim keyExists As Boolean
        keyExists = CollectionKeyExists(assessmentDict, lookupKey)
        
        If Not keyExists Then
            lookupKey = subjectCode & "|All"
            keyExists = CollectionKeyExists(assessmentDict, lookupKey)
            If Not keyExists Then GoTo NextIteration
        End If
        
        Dim subjectInfo(0 To 5) As Variant
        subjectInfo(0) = wsSubjects.Cells(i, colUID).Value
        subjectInfo(1) = subjectCode
        subjectInfo(2) = wsSubjects.Cells(i, colSubjectName).Value
        subjectInfo(3) = studyPeriod
        subjectInfo(4) = groupedPeriodVal
        subjectInfo(5) = i
        
        subjects.add subjectInfo
        includedCount = includedCount + 1
NextIteration:
    Next i
    
    LogMessage "Filtered subjects for " & groupedPeriod & ": " & includedCount, Timer - tStart
    Set GetFilteredSubjectsWithAssessments = subjects
    Exit Function
    
ErrorHandler:
    LogMessage "ERROR in GetFilteredSubjectsWithAssessments: " & Err.description & " (Error " & Err.Number & ")"
    Set GetFilteredSubjectsWithAssessments = Nothing
End Function

'===============================================================
' SECTION 4: SHEET POPULATION
'===============================================================

'---------------------------------------------------------------
' PopulateSheetData
' Purpose: Populate a calculation sheet with all subject blocks,
'          organised by study period (Summer/Sem 1 or Winter/Sem 2)
' Called by: GenerateSheet
' Returns: True on success
'---------------------------------------------------------------
Function PopulateSheetData(wb As Workbook, wsOutput As Worksheet, subjectData As Collection, wordBench As Double, examBench As Double, markingSupportBench As Double) As Boolean
    On Error GoTo ErrorHandler
    
    PopulateSheetData = False
    Dim tStart As Double
    tStart = Timer
    
    If subjectData Is Nothing Or subjectData.count = 0 Then
        LogMessage "WARNING: No subjects to populate"
        PopulateSheetData = True
        Exit Function
    End If
    
    Dim wsAssessment As Worksheet
    Set wsAssessment = wb.Sheets("assessment data parsed")
    
    If wsAssessment Is Nothing Then
        LogMessage "ERROR: Assessment sheet not found"
        Exit Function
    End If
    
    Dim groupedPeriod As String
    groupedPeriod = SafeArrayIndex(subjectData(1), 4, "")
    
    Dim currentRow As Long
    currentRow = 4
    
    ' Add category headers
    If groupedPeriod = "FHY" Then
        wsOutput.Rows(currentRow).Interior.Color = RGB(0, 0, 0)
        wsOutput.Cells(currentRow, 2).Value = "SUMMER"
        wsOutput.Cells(currentRow, 2).Font.Color = RGB(255, 255, 255)
        wsOutput.Cells(currentRow, 2).Font.Bold = True
        currentRow = currentRow + 1
    Else
        wsOutput.Rows(currentRow).Interior.Color = RGB(0, 0, 0)
        wsOutput.Cells(currentRow, 2).Value = "WINTER"
        wsOutput.Cells(currentRow, 2).Font.Color = RGB(255, 255, 255)
        wsOutput.Cells(currentRow, 2).Font.Bold = True
        currentRow = currentRow + 1
    End If
    
    ' Categorise subjects
    Dim summerSubjects As New Collection
    Dim semester1Subjects As New Collection
    Dim winterSubjects As New Collection
    Dim semester2Subjects As New Collection
    
    Dim subject As Variant
    For Each subject In subjectData
        Dim studyPeriod As String
        studyPeriod = SafeArrayIndex(subject, 3, "")
        
        Dim category As String
        category = CategoriseStudyPeriod(studyPeriod, groupedPeriod)
        
        Select Case category
            Case "SUMMER": summerSubjects.add subject
            Case "SEMESTER 1": semester1Subjects.add subject
            Case "WINTER": winterSubjects.add subject
            Case "SEMESTER 2": semester2Subjects.add subject
        End Select
    Next subject
    
    ' Collection to store marker formula info for batch processing
    Dim markerFormulaQueue As New Collection
    
    ' Collection to store dropdown cell locations for batch processing
    Dim dropdownQueue As New Collection
    
   ' Process subjects by category
If groupedPeriod = "FHY" Then
    Call ProcessSubjectCollection(wb, wsOutput, summerSubjects, currentRow, wordBench, examBench, markingSupportBench, markerFormulaQueue, dropdownQueue, "", "SUMMER")
    Call ProcessSubjectCollection(wb, wsOutput, semester1Subjects, currentRow, wordBench, examBench, markingSupportBench, markerFormulaQueue, dropdownQueue, "SEMESTER 1", "SEMESTER 1")
Else
    Call ProcessSubjectCollection(wb, wsOutput, winterSubjects, currentRow, wordBench, examBench, markingSupportBench, markerFormulaQueue, dropdownQueue, "", "WINTER")
    Call ProcessSubjectCollection(wb, wsOutput, semester2Subjects, currentRow, wordBench, examBench, markingSupportBench, markerFormulaQueue, dropdownQueue, "SEMESTER 2", "SEMESTER 2")
End If
    
    ' NOW APPLY ALL MARKER FORMULAS IN BATCH
    LogMessage "Applying marker formulas in batch..."
    Dim formulaInfo As Variant
    For Each formulaInfo In markerFormulaQueue
        Call SetMarkerBlockFormulas(wsOutput, CLng(formulaInfo(0)), CLng(formulaInfo(1)), CInt(formulaInfo(2)), CStr(formulaInfo(3)), CStr(formulaInfo(4)))
    Next formulaInfo
    
    ' NOW APPLY ALL CONTRACT DROPDOWNS IN BATCH
    LogMessage "Applying contract dropdowns in batch..."
    Dim dropdownInfo As Variant
    For Each dropdownInfo In dropdownQueue
        Call SetContractDropdown(wsOutput, CLng(dropdownInfo(0)), CInt(dropdownInfo(1)))
    Next dropdownInfo
    
    LogMessage "PopulateSheetData completed", Timer - tStart
    PopulateSheetData = True
    Exit Function
    
ErrorHandler:
    LogMessage "ERROR in PopulateSheetData: " & Err.description & " (Error " & Err.Number & ")"
    PopulateSheetData = False
End Function

'---------------------------------------------------------------
' ProcessSubjectCollection
' Purpose: Iterate over a collection of subjects for one study
'          period category and write each subject block to the sheet
' Called by: PopulateSheetData
'---------------------------------------------------------------
Sub ProcessSubjectCollection(wb As Workbook, wsOutput As Worksheet, subjects As Collection, _
    ByRef currentRow As Long, wordBench As Double, examBench As Double, markingSupportBench As Double, _
    markerFormulaQueue As Collection, dropdownQueue As Collection, headerText As String, categoryName As String)
    
    If subjects.count = 0 Then Exit Sub
    
    ' Add header if specified
    If headerText <> "" Then
        wsOutput.Rows(currentRow).Interior.Color = RGB(0, 0, 0)
        wsOutput.Cells(currentRow, 2).Value = headerText
        wsOutput.Cells(currentRow, 2).Font.Color = RGB(255, 255, 255)
        wsOutput.Cells(currentRow, 2).Font.Bold = True
        currentRow = currentRow + 1
    End If
    
    ' Process all subjects in collection
    Dim subject As Variant
    For Each subject In subjects
        If Not ProcessSubject(wb, wsOutput, subject, currentRow, wordBench, examBench, markingSupportBench, markerFormulaQueue, dropdownQueue) Then
            LogMessage "ERROR: Failed to process subject in " & categoryName
        End If
    Next subject
End Sub


'---------------------------------------------------------------
' ProcessSubject
' Purpose: Write a single subject block (header row, assessment
'          rows, lecturer rows, formulas, and marker blocks)
' Called by: ProcessSubjectCollection
' Returns: True on success
'---------------------------------------------------------------
Function ProcessSubject(wb As Workbook, wsOutput As Worksheet, ByRef subject As Variant, ByRef currentRow As Long, wordBench As Double, examBench As Double, markingSupportBench As Double, ByRef markerFormulaQueue As Collection, ByRef dropdownQueue As Collection) As Boolean
    On Error GoTo ErrorHandler
    
    ProcessSubject = False
    
    Dim wsAssessment As Worksheet
    Set wsAssessment = wb.Sheets("assessment data parsed")
    
    Dim subjectCode As String
    Dim studyPeriod As String
    subjectCode = SafeArrayIndex(subject, 1, "")
    studyPeriod = SafeArrayIndex(subject, 3, "")
    
    If subjectCode = "" Or studyPeriod = "" Then
        LogMessage "ERROR: Invalid subject data"
        Exit Function
    End If
    
    Dim subjectStartRow As Long
    subjectStartRow = currentRow
    
    ' Get assessments and lecturers
    Dim assessments As Collection
    Set assessments = GetAssessmentsForSubject(wsAssessment, subjectCode, studyPeriod)
    
    If assessments Is Nothing Or assessments.count = 0 Then
        ProcessSubject = True
        Exit Function
    End If
    
    Dim lecturers As Collection
    Set lecturers = GetLecturersForSubjectFlexible(wb, subjectCode, studyPeriod)
    If lecturers Is Nothing Then Set lecturers = New Collection
    
    ' Calculate structure
    Dim uidCount As Integer
    uidCount = assessments.count
    
    Dim assessmentRows As Integer
    assessmentRows = uidCount + 2
    
    Dim totalRowsNeeded As Integer
    totalRowsNeeded = Application.WorksheetFunction.Max(assessmentRows, lecturers.count + 1)
    
    ' Build output array
    Dim outputData() As Variant
    ReDim outputData(1 To assessmentRows, 1 To 10)
    
    ' Header row
    outputData(1, 1) = subjectCode & "_" & studyPeriod & "_0"
    outputData(1, 2) = subjectCode
    outputData(1, 3) = studyPeriod
    outputData(1, 4) = 0
    
    ' Assessment rows
    Dim i As Integer
    For i = 1 To assessments.count
        Dim assessment As Variant
        assessment = assessments(i)
        
        outputData(i + 1, 1) = subjectCode & "_" & studyPeriod & "_" & i
        outputData(i + 1, 5) = SafeArrayIndex(assessment, 1, "")
        outputData(i + 1, 6) = SafeArrayIndex(assessment, 3, 0)
        outputData(i + 1, 7) = SafeArrayIndex(assessment, 4, 0)
        outputData(i + 1, 8) = SafeArrayIndex(assessment, 5, 0)
    Next i
    
    ' Total row
    outputData(assessmentRows, 1) = subjectCode & "_" & studyPeriod & "_" & (uidCount + 1)
    outputData(assessmentRows, 5) = "Total"
    
    ' WRITE DATA
    wsOutput.Cells(currentRow, 1).Resize(assessmentRows, 10).Value = outputData
    
    ' Add enrolment formula
    Dim enrolFile As String
    enrolFile = wb.Sheets("Dashboard").Range("C3").Value
    If enrolFile <> "" Then
        enrolFile = Trim(enrolFile)
        If Right(LCase(enrolFile), 5) <> ".xlsm" Then enrolFile = enrolFile & ".xlsm"
        
        Dim enrolPath As String
        enrolPath = "'" & ENROLMENT_TRACKER_BASE & "[" & enrolFile & "]Enrolment Number Tracker'!"
        
        wsOutput.Cells(currentRow, 4).Formula = "=IFERROR(INDEX(" & enrolPath & "$A:$IZ,SUMPRODUCT((" & enrolPath & "$A:$A=B" & currentRow & ")*(" & enrolPath & "$C:$C=C" & currentRow & ")*ROW(" & enrolPath & "$A:$A)),COUNTA(" & enrolPath & "$1:$1)),0)"
    End If
    
    ' Format header row
    wsOutput.Rows(currentRow).Interior.Color = RGB(192, 192, 192)
    wsOutput.Cells(currentRow + assessmentRows - 1, 5).Font.Bold = True
    
    ' Hyperlink subject code to handbook
    Dim yearValueProcess As String
    yearValueProcess = CStr(wb.Sheets("Dashboard").Range("C2").Value)
    Dim handbookUrl As String
    handbookUrl = "https://handbook.unimelb.edu.au/" & yearValueProcess & "/subjects/" & subjectCode & "/assessment"
    
    wsOutput.Hyperlinks.add Anchor:=wsOutput.Cells(currentRow, 2), _
        Address:=handbookUrl, _
        TextToDisplay:=subjectCode
        
    With wsOutput.Cells(currentRow, 2).Font
        .Bold = True
        .Color = RGB(0, 0, 0)
        .Underline = xlUnderlineStyleNone
    End With
    
    ' Add assessment formulas
    Dim formulaRow As Long
    For i = 1 To assessments.count
        formulaRow = currentRow + i
        Dim assessmentItem As Variant
        assessmentItem = assessments(i)
        Call SetAssessmentQuantityFormula(wsOutput, formulaRow, SafeArrayIndex(assessmentItem, 0, ""), subjectCode, studyPeriod, wsAssessment)
        Call SetMarkingHoursFormula(wsOutput, formulaRow)
    Next i
    
    ' Add total formula
    If uidCount > 0 Then
        wsOutput.Cells(currentRow + assessmentRows - 1, 10).Formula = "=SUM(J" & (currentRow + 1) & ":J" & (currentRow + uidCount) & ")"
    End If
    
    Dim subjectEndRow As Long
    Dim totalRowIndex As Long
    subjectEndRow = currentRow + assessmentRows - 1
    totalRowIndex = subjectEndRow
    
    ' Expand if needed
    Dim extraRowsNeeded As Integer
    extraRowsNeeded = totalRowsNeeded - assessmentRows
    
    If extraRowsNeeded > 0 Then
        wsOutput.Rows(subjectEndRow + 1).Resize(extraRowsNeeded).Insert Shift:=xlDown
        
        Dim j As Integer
        For j = 1 To extraRowsNeeded
            wsOutput.Cells(subjectEndRow + j, 1).Value = subjectCode & "_" & studyPeriod & "_" & (uidCount + 1 + j)
        Next j
        
        subjectEndRow = subjectEndRow + extraRowsNeeded
    End If
    
    ' Populate lecturer data
    If lecturers.count > 0 Then
        Call PopulateLecturerData(wsOutput, lecturers, subjectStartRow + 1, totalRowIndex, markingSupportBench)
    End If
    
    ' QUEUE marker formulas for batch processing (don't execute yet)
    Dim mk As Integer
    For mk = 1 To 3
        Dim markerInfo(0 To 4) As Variant
        markerInfo(0) = subjectStartRow
        markerInfo(1) = subjectEndRow
        markerInfo(2) = mk
        markerInfo(3) = subjectCode
        markerInfo(4) = studyPeriod
        markerFormulaQueue.add markerInfo
    Next mk
    
    ' QUEUE dropdown locations for batch processing (don't execute yet)
    Dim dk As Integer
    For dk = 1 To 3
        Dim dropdownItem(0 To 1) As Variant
        dropdownItem(0) = subjectStartRow
        dropdownItem(1) = dk
        dropdownQueue.Add dropdownItem
    Next dk
    
    currentRow = subjectEndRow + 1
    
    ProcessSubject = True
    Exit Function
    
ErrorHandler:
    LogMessage "ERROR in ProcessSubject (" & subjectCode & "): " & Err.description
    ProcessSubject = False
End Function


'---------------------------------------------------------------
' SetContractDropdown
' Purpose: Add a Y/N data validation dropdown to the "Contract
'          Requested" column in a marker block
' Called by: ProcessSubject
'---------------------------------------------------------------
Sub SetContractDropdown(wsOutput As Worksheet, subjectStartRow As Long, markerNum As Integer)
    On Error GoTo ErrorHandler
    
    ' Add data validation dropdown to first assessment row (+1, not header 0)
    Dim targetRow As Long
    targetRow = subjectStartRow + 1  ' First assessment row
    
    Dim contractCol As Integer
    Select Case markerNum
        Case 1
            contractCol = 29  ' Column AC (was AB=28, shifted +1)
        Case 2
            contractCol = 39  ' Column AM (was AL=38, shifted +1)
        Case 3
            contractCol = 49  ' Column AW (was AV=48, shifted +1)
    End Select
    
    ' Add data validation dropdown with Yes/No options
    Dim cellObj As Range
    Set cellObj = wsOutput.Cells(targetRow, contractCol)
    
    On Error Resume Next
    ' Remove any existing validation
    cellObj.Validation.Delete
    Err.Clear
    
    ' Add dropdown list validation
    With cellObj.Validation
        .Delete
        .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="N,Y"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = False
    End With
    
    ' Set default value to No
    If cellObj.Value = "" Then
        cellObj.Value = "N"
    End If
    
    If Err.Number <> 0 Then
        LogMessage "WARNING: Could not add data validation for Marker " & markerNum & " at row " & targetRow
    End If
    
    On Error GoTo ErrorHandler
    Exit Sub
    
ErrorHandler:
    LogMessage "ERROR in SetContractDropdown: " & Err.description
End Sub

'---------------------------------------------------------------
' SetMarkerBlockFormulas
' Purpose: Write word count, exam, group size, quantity, and
'          allocation formulas for one marker block
' Called by: ProcessSubject (via markerFormulaQueue)
'---------------------------------------------------------------
Sub SetMarkerBlockFormulas(wsOutput As Worksheet, subjectStartRow As Long, subjectEndRow As Long, _
    markerNum As Integer, subjectCode As String, studyPeriod As String)
    
    On Error GoTo ErrorHandler
    
    Dim baseCol As Integer, benchmarkCol As Integer
    Select Case markerNum
        Case 1
            baseCol = 20   ' Column T (shifted +1 from S)
            benchmarkCol = 26  ' Column Z (shifted +1 from Y)
        Case 2
            baseCol = 30   ' Column AD (shifted +1 from AC)
            benchmarkCol = 36  ' Column AJ (shifted +1 from AI)
        Case 3
            baseCol = 40   ' Column AN (shifted +1 from AM)
            benchmarkCol = 45  ' Column AS (shifted +1 from AR)
    End Select
    
    Dim detailsCol As Integer, wordCol As Integer, examCol As Integer
    Dim groupCol As Integer, qtyCol As Integer, allocCol As Integer
    
    detailsCol = baseCol + 1
    wordCol = baseCol + 2
    examCol = baseCol + 3
    groupCol = baseCol + 4
    qtyCol = baseCol + 5
    allocCol = baseCol + 6
    
    ' Calculate number of rows EXCLUDING the header row
    Dim numRows As Long
    numRows = subjectEndRow - subjectStartRow
    
    ' BUILD FORMULA ARRAYS in memory
    Dim wordFormulas As Variant
    Dim examFormulas As Variant
    Dim groupFormulas As Variant
    Dim qtyFormulas As Variant
    Dim allocFormulas As Variant
    
    ReDim wordFormulas(1 To numRows, 1 To 1)
    ReDim examFormulas(1 To numRows, 1 To 1)
    ReDim groupFormulas(1 To numRows, 1 To 1)
    ReDim qtyFormulas(1 To numRows, 1 To 1)
    ReDim allocFormulas(1 To numRows, 1 To 1)
    
    Dim r As Long, arrayRow As Long
    Dim detailColLetter As String, wordColLetter As String, examColLetter As String
    Dim qtyColLetter As String, benchmarkColLetter As String
    
    detailColLetter = ColLetter(detailsCol)
    wordColLetter = ColLetter(wordCol)
    examColLetter = ColLetter(examCol)
    qtyColLetter = ColLetter(qtyCol)
    benchmarkColLetter = ColLetter(benchmarkCol)
    
    ' Start from subjectStartRow + 1 (skip header row)
    For r = subjectStartRow + 1 To subjectEndRow
        arrayRow = r - subjectStartRow
        
        If wsOutput.Cells(r, 5).Value = "Total" Then
            ' Total row - write "Total" text and SUM formula
            wsOutput.Cells(r, detailsCol).Value = "Total"
            wsOutput.Cells(r, detailsCol).Font.Bold = True
            allocFormulas(arrayRow, 1) = "=SUM(" & ColLetter(allocCol) & (subjectStartRow + 1) & ":" & ColLetter(allocCol) & (r - 1) & ")"
        Else
            ' Regular assessment rows - all formulas
            wordFormulas(arrayRow, 1) = "=IFERROR(IF(INDEX($F$" & (subjectStartRow + 1) & ":$F$" & (subjectEndRow - 1) & ", MATCH(" & detailColLetter & r & ",$E$" & (subjectStartRow + 1) & ":$E$" & (subjectEndRow - 1) & ",0))="""","""", INDEX($F$" & (subjectStartRow + 1) & ":$F$" & (subjectEndRow - 1) & ", MATCH(" & detailColLetter & r & ",$E$" & (subjectStartRow + 1) & ":$E$" & (subjectEndRow - 1) & ",0))),"""")"
            
            examFormulas(arrayRow, 1) = "=IFERROR(IF(INDEX($G$" & (subjectStartRow + 1) & ":$G$" & (subjectEndRow - 1) & ", MATCH(" & detailColLetter & r & ",$E$" & (subjectStartRow + 1) & ":$E$" & (subjectEndRow - 1) & ",0))="""","""", INDEX($G$" & (subjectStartRow + 1) & ":$G$" & (subjectEndRow - 1) & ", MATCH(" & detailColLetter & r & ",$E$" & (subjectStartRow + 1) & ":$E$" & (subjectEndRow - 1) & ",0))),"""")"
            
            groupFormulas(arrayRow, 1) = "=IFERROR(IF(INDEX($H$" & (subjectStartRow + 1) & ":$H$" & (subjectEndRow - 1) & ", MATCH(" & detailColLetter & r & ",$E$" & (subjectStartRow + 1) & ":$E$" & (subjectEndRow - 1) & ",0))="""","""", INDEX($H$" & (subjectStartRow + 1) & ":$H$" & (subjectEndRow - 1) & ", MATCH(" & detailColLetter & r & ",$E$" & (subjectStartRow + 1) & ":$E$" & (subjectEndRow - 1) & ",0))),"""")"
            
            qtyFormulas(arrayRow, 1) = "=IF(" & detailColLetter & r & "="""","""",0)"
            
            ' UPDATED ALLOCATION FORMULA:
            ' If qty has number AND (word OR exam OR group are all empty) ? return qty value
            ' Otherwise ? calculate as before
            allocFormulas(arrayRow, 1) = "=IF(" & qtyColLetter & r & "="""","""", IF(ISNUMBER(" & qtyColLetter & r & "), IF(AND(" & wordColLetter & r & "="""", " & examColLetter & r & "=""""), " & qtyColLetter & r & ", IF(ISNUMBER(" & wordColLetter & r & "), " & qtyColLetter & r & "*" & wordColLetter & r & "/VALUE(LEFT($" & benchmarkColLetter & "$2,FIND("" "",$" & benchmarkColLetter & "$2)-1)), IF(ISNUMBER(" & examColLetter & r & "), " & qtyColLetter & r & "/VALUE(LEFT($" & benchmarkColLetter & "$3,FIND("" "",$" & benchmarkColLetter & "$3)-1)),""""))), " & qtyColLetter & r & "))"
        End If
    Next r
    
    ' BATCH WRITE - starting from subjectStartRow + 1 (skip header)
    wsOutput.Range(wsOutput.Cells(subjectStartRow + 1, wordCol), wsOutput.Cells(subjectEndRow, wordCol)).Formula = wordFormulas
    wsOutput.Range(wsOutput.Cells(subjectStartRow + 1, examCol), wsOutput.Cells(subjectEndRow, examCol)).Formula = examFormulas
    wsOutput.Range(wsOutput.Cells(subjectStartRow + 1, groupCol), wsOutput.Cells(subjectEndRow, groupCol)).Formula = groupFormulas
    wsOutput.Range(wsOutput.Cells(subjectStartRow + 1, qtyCol), wsOutput.Cells(subjectEndRow, qtyCol)).Formula = qtyFormulas
    wsOutput.Range(wsOutput.Cells(subjectStartRow + 1, allocCol), wsOutput.Cells(subjectEndRow, allocCol)).Formula = allocFormulas
    
    Exit Sub
    
ErrorHandler:
    LogMessage "ERROR in SetMarkerBlockFormulas: " & Err.description
End Sub

'---------------------------------------------------------------
' ColLetter
' Purpose: Convert a column number to its letter equivalent
'---------------------------------------------------------------
Function ColLetter(colNum As Integer) As String
    ColLetter = Split(Cells(1, colNum).Address, "$")(1)
End Function

'---------------------------------------------------------------
' GetAssessmentsForSubject
' Purpose: Retrieve all assessments for a subject/study period,
'          using a cached array for performance
'          Falls back to "All" study period if no exact match
' Called by: ProcessSubject
' Returns: Collection of assessment arrays
'---------------------------------------------------------------
Function GetAssessmentsForSubject(wsAssessment As Worksheet, subjectCode As String, studyPeriod As String) As Collection
    On Error GoTo ErrorHandler
    
    Static assessmentCache As Variant
    Static cacheInitialized As Boolean
    Static cacheSheetName As String
    
    If wsAssessment Is Nothing Then
        LogMessage "ERROR: wsAssessment is Nothing in GetAssessmentsForSubject"
        Set GetAssessmentsForSubject = Nothing
        Exit Function
    End If
    
    If Not cacheInitialized Or cacheSheetName <> wsAssessment.name Then
        Dim lastRow As Long
        lastRow = wsAssessment.Cells(wsAssessment.Rows.count, "A").End(xlUp).Row
        
        If lastRow < 2 Then
            LogMessage "WARNING: No assessment data found"
            Set GetAssessmentsForSubject = New Collection
            Exit Function
        End If
        
        On Error Resume Next
        assessmentCache = wsAssessment.Range(wsAssessment.Cells(2, 1), wsAssessment.Cells(lastRow, 8)).Value
        If Err.Number <> 0 Then
            LogMessage "ERROR: Failed to cache assessment data: " & Err.description
            On Error GoTo ErrorHandler
            Set GetAssessmentsForSubject = Nothing
            Exit Function
        End If
        On Error GoTo ErrorHandler
        
        cacheInitialized = True
        cacheSheetName = wsAssessment.name
    End If
    
    Dim assessments As New Collection
    Dim foundExact As Boolean
    foundExact = False
    
    Dim assessItem(0 To 6) As Variant
    
    Dim i As Long
    For i = 1 To UBound(assessmentCache, 1)
        If assessmentCache(i, 2) = subjectCode And assessmentCache(i, 3) = studyPeriod Then
            assessItem(0) = assessmentCache(i, 1)
            assessItem(1) = assessmentCache(i, 4)
            assessItem(2) = assessmentCache(i, 5)
            assessItem(3) = assessmentCache(i, 6)
            assessItem(4) = assessmentCache(i, 7)
            assessItem(5) = assessmentCache(i, 8)
            assessItem(6) = studyPeriod
            
            assessments.add assessItem
            foundExact = True
        End If
    Next i
    
    If Not foundExact Then
        Dim assessItemAll(0 To 6) As Variant
        
        For i = 1 To UBound(assessmentCache, 1)
            If assessmentCache(i, 2) = subjectCode And Trim(CStr(assessmentCache(i, 3))) = "All" Then
                assessItemAll(0) = assessmentCache(i, 1)
                assessItemAll(1) = assessmentCache(i, 4)
                assessItemAll(2) = assessmentCache(i, 5)
                assessItemAll(3) = assessmentCache(i, 6)
                assessItemAll(4) = assessmentCache(i, 7)
                assessItemAll(5) = assessmentCache(i, 8)
                assessItemAll(6) = studyPeriod
                
                assessments.add assessItemAll
            End If
        Next i
    End If
    
    Set GetAssessmentsForSubject = assessments
    Exit Function
    
ErrorHandler:
    LogMessage "ERROR in GetAssessmentsForSubject: " & Err.description & " (Error " & Err.Number & ")"
    Set GetAssessmentsForSubject = Nothing
End Function

'---------------------------------------------------------------
' GetLecturersForSubjectFlexible
' Purpose: Find lecturers for a subject, trying multiple study
'          period name variations (e.g. "Summer" vs "Summer Term")
' Called by: ProcessSubject
' Returns: Collection of lecturer arrays (name, status, stream, activity)
'---------------------------------------------------------------
Function GetLecturersForSubjectFlexible(wb As Workbook, subjectCode As String, studyPeriod As String) As Collection
    On Error GoTo ErrorHandler
    
    Static teachingCache As Variant
    Static cacheInitialized As Boolean
    Static cacheSheetName As String
    
    Dim wsTeaching As Worksheet
    Set wsTeaching = wb.Sheets("teaching stream")
    
    If wsTeaching Is Nothing Then
        LogMessage "ERROR: teaching stream sheet not found"
        Set GetLecturersForSubjectFlexible = Nothing
        Exit Function
    End If
    
    If Not cacheInitialized Or cacheSheetName <> wsTeaching.name Then
        Dim lastRow As Long
        lastRow = wsTeaching.Cells(wsTeaching.Rows.count, "B").End(xlUp).Row
        
        If lastRow < 2 Then
            LogMessage "WARNING: No teaching data found"
            Set GetLecturersForSubjectFlexible = New Collection
            Exit Function
        End If
        
        On Error Resume Next
        teachingCache = wsTeaching.Range(wsTeaching.Cells(2, 2), wsTeaching.Cells(lastRow, 7)).Value
        If Err.Number <> 0 Then
            LogMessage "ERROR: Failed to cache teaching data: " & Err.description
            On Error GoTo ErrorHandler
            Set GetLecturersForSubjectFlexible = Nothing
            Exit Function
        End If
        On Error GoTo ErrorHandler
        
        cacheInitialized = True
        cacheSheetName = wsTeaching.name
    End If
    
    Dim lecturers As New Collection
    Dim uniqueDict As Collection
    Set uniqueDict = New Collection
    
    Dim studyPeriodVariations As New Collection
    studyPeriodVariations.add studyPeriod
    
    If InStr(LCase(studyPeriod), "summer") > 0 Then
        On Error Resume Next
        studyPeriodVariations.add "Summer Term"
        studyPeriodVariations.add "Summer"
        Err.Clear
        On Error GoTo ErrorHandler
    ElseIf InStr(LCase(studyPeriod), "winter") > 0 Then
        On Error Resume Next
        studyPeriodVariations.add "Winter Term"
        studyPeriodVariations.add "Winter"
        Err.Clear
        On Error GoTo ErrorHandler
    ElseIf InStr(LCase(studyPeriod), "term") > 0 Then
        Dim periodWithoutTerm As String
        periodWithoutTerm = Replace(studyPeriod, " Term", "", 1, -1, vbTextCompare)
        periodWithoutTerm = Replace(periodWithoutTerm, "Term", "", 1, -1, vbTextCompare)
        If Len(Trim(periodWithoutTerm)) > 0 Then
            On Error Resume Next
            studyPeriodVariations.add Trim(periodWithoutTerm)
            Err.Clear
            On Error GoTo ErrorHandler
        End If
    End If
    
    Dim i As Long
    Dim variation As Variant
    Dim lecturerName As String
    Dim lecturerInfo(0 To 3) As Variant
    
    For Each variation In studyPeriodVariations
        For i = 1 To UBound(teachingCache, 1)
            If teachingCache(i, 1) = subjectCode And Trim(CStr(teachingCache(i, 2))) = CStr(variation) Then
                lecturerName = teachingCache(i, 3)
                
                If Not CollectionKeyExists(uniqueDict, lecturerName) Then
                    lecturerInfo(0) = lecturerName
                    lecturerInfo(1) = teachingCache(i, 4)
                    lecturerInfo(2) = teachingCache(i, 5)
                    lecturerInfo(3) = teachingCache(i, 6)
                    
                    lecturers.add lecturerInfo
                    On Error Resume Next
                    uniqueDict.add True, lecturerName
                    Err.Clear
                    On Error GoTo ErrorHandler
                End If
            End If
        Next i
        
        If lecturers.count > 0 Then Exit For
    Next variation
    
    Set GetLecturersForSubjectFlexible = lecturers
    Exit Function
    
ErrorHandler:
    LogMessage "ERROR in GetLecturersForSubjectFlexible: " & Err.description & " (Error " & Err.Number & ")"
    Set GetLecturersForSubjectFlexible = Nothing
End Function

'---------------------------------------------------------------
' PopulateLecturerData
' Purpose: Write lecturer names, statuses, and formulas (cols L-R)
'          into the subject block
' Called by: ProcessSubject
'---------------------------------------------------------------
Sub PopulateLecturerData(wsOutput As Worksheet, lecturers As Collection, startRow As Long, totalRow As Long, markingSupportBench As Double)
    On Error GoTo ErrorHandler
    
    If lecturers Is Nothing Or lecturers.count = 0 Then Exit Sub
    
    ' Build lecturer data array (Columns L-O: Name, Status, Streams, Activity ID)
    Dim lecturerData As Variant
    ReDim lecturerData(1 To lecturers.count, 1 To 4)
    
    Dim i As Integer
    Dim lecItem As Variant
    
    For i = 1 To lecturers.count
        lecItem = lecturers(i)
        lecturerData(i, 1) = SafeArrayIndex(lecItem, 0, "")  ' Name
        lecturerData(i, 2) = SafeArrayIndex(lecItem, 1, "")  ' Status
        lecturerData(i, 3) = SafeArrayIndex(lecItem, 3, "")  ' Streams
        lecturerData(i, 4) = SafeArrayIndex(lecItem, 2, "")  ' Activity ID
    Next i
    
    ' Write lecturer data (Columns L-O: 12-15)
    wsOutput.Cells(startRow, 12).Resize(lecturers.count, 4).Value = lecturerData
    
    ' Bold first lecturer (subject coordinator)
    If lecturers.count > 0 Then
        wsOutput.Cells(startRow, 12).Font.Bold = True
    End If
    
    ' BUILD FORMULA ARRAYS for columns Q and R (shifted from P and Q)
    Dim formulas() As Variant
    ReDim formulas(1 To lecturers.count, 1 To 2)
    
    Dim currentRow As Long
    For i = 1 To lecturers.count
        currentRow = startRow + i - 1
        
        ' Column Q (17): Allocated Marking formula
        formulas(i, 1) = "=IF(M" & currentRow & "=""Continuing T&R"",N" & currentRow & "*VALUE(LEFT($Q$2,FIND("" "",$Q$2)-1)),"""")"
        
        ' Column R (18): Marking Support Hours Available formula (UPDATED)
        ' =IF(OR(P="",Q=""),"",$J$totalRow*(P/D)-Q)
        formulas(i, 2) = "=IF(OR(P" & currentRow & "="""",Q" & currentRow & "=""""),"""",$J$" & totalRow & "*(P" & currentRow & "/D" & (startRow - 1) & ")-Q" & currentRow & ")"
    Next i
    
    ' BATCH WRITE both formulas at once (columns Q and R, shifted from 16-17 to 17-18)
    wsOutput.Cells(startRow, 17).Resize(lecturers.count, 2).Formula = formulas
    
    Exit Sub
    
ErrorHandler:
    LogMessage "ERROR in PopulateLecturerData: " & Err.description & " (Error " & Err.Number & ")"
End Sub

'---------------------------------------------------------------
' SetAssessmentQuantityFormula
' Purpose: Write the assessment quantity formula (col I) and
'          in-class location value (col E) for one assessment row
' Called by: ProcessSubject
'---------------------------------------------------------------
Sub SetAssessmentQuantityFormula(wsOutput As Worksheet, currentRow As Long, originalUID As String, subjectCode As String, studyPeriod As String, wsAssessment As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim locationValue As Variant
    locationValue = ""
    
    On Error Resume Next
    locationValue = Application.WorksheetFunction.Index(wsAssessment.Range("E:E"), _
                    Application.WorksheetFunction.Match(originalUID, wsAssessment.Range("A:A"), 0))
    On Error GoTo ErrorHandler
    
    Dim enrolmentCell As String
    enrolmentCell = FindEnrolmentCell(wsOutput, subjectCode, studyPeriod)
    
    If enrolmentCell = "" Then
        LogMessage "WARNING: Could not find enrolment cell for " & subjectCode & " " & studyPeriod
        Exit Sub
    End If
    
    If locationValue <> "" And Not IsError(locationValue) Then
        wsOutput.Cells(currentRow, 9).Value = locationValue
    Else
        Dim Formula As String
        Formula = "=IF(H" & currentRow & "<>"""", " & enrolmentCell & "/H" & currentRow & ", " & enrolmentCell & ")"
        wsOutput.Cells(currentRow, 9).Formula = Formula
    End If
    
    Exit Sub
    
ErrorHandler:
    LogMessage "ERROR in SetAssessmentQuantityFormula: " & Err.description & " (Error " & Err.Number & ")"
End Sub

'---------------------------------------------------------------
' SetMarkingHoursFormula
' Purpose: Write the marking hours formula (col J) which calculates
'          hours based on word count or exam benchmarks
' Called by: ProcessSubject
'---------------------------------------------------------------
Sub SetMarkingHoursFormula(wsOutput As Worksheet, currentRow As Long)
    On Error GoTo ErrorHandler
    
    Dim Formula As String
    Formula = "=IF(ISNUMBER(I" & currentRow & "), "
    Formula = Formula & "IF(ISNUMBER(F" & currentRow & "), "
    Formula = Formula & "I" & currentRow & "*(F" & currentRow & "/VALUE(LEFT($J$2,FIND("" "",$J$2)-1))), "
    Formula = Formula & "IF(ISNUMBER(G" & currentRow & "), "
    Formula = Formula & "I" & currentRow & "/(VALUE(LEFT($J$3,FIND("" "",$J$3)-1))), "
    Formula = Formula & """"")), """")"
    
    wsOutput.Cells(currentRow, 10).Formula = Formula
    
    Exit Sub
    
ErrorHandler:
    LogMessage "ERROR in SetMarkingHoursFormula: " & Err.description & " (Error " & Err.Number & ")"
End Sub

'---------------------------------------------------------------
' CategoriseStudyPeriod
' Purpose: Map a study period name to its display category
'          (Summer, Semester 1, Winter, or Semester 2)
'---------------------------------------------------------------
Function CategoriseStudyPeriod(studyPeriod As String, groupedPeriod As String) As String
    Dim sp As String
    sp = LCase(studyPeriod)
    
    If InStr(sp, "summer") > 0 Or InStr(sp, "october") > 0 Or InStr(sp, "november") > 0 Or InStr(sp, "term 1") > 0 Then
        CategoriseStudyPeriod = "SUMMER"
        Exit Function
    End If
    
    If InStr(sp, "winter") > 0 Or InStr(sp, "june") > 0 Or InStr(sp, "july") > 0 Then
        CategoriseStudyPeriod = "WINTER"
        Exit Function
    End If
    
    If InStr(sp, "semester 1") > 0 Or InStr(sp, "term 2") > 0 Then
        CategoriseStudyPeriod = "SEMESTER 1"
        Exit Function
    End If
    
    If InStr(sp, "semester 2") > 0 Or InStr(sp, "term 3") > 0 Or InStr(sp, "term 4") > 0 Then
        CategoriseStudyPeriod = "SEMESTER 2"
        Exit Function
    End If
    
    If groupedPeriod = "FHY" Then
        CategoriseStudyPeriod = "SEMESTER 1"
    Else
        CategoriseStudyPeriod = "SEMESTER 2"
    End If
End Function

'===============================================================
' SECTION 5: FORMATTING & UTILITIES
'===============================================================

'---------------------------------------------------------------
' FindColumn
' Purpose: Locate a column by header name in row 1
' Returns: Column number, or 0 if not found
'---------------------------------------------------------------
Function FindColumn(ws As Worksheet, headerName As String) As Integer
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        LogMessage "ERROR: Worksheet is Nothing in FindColumn"
        FindColumn = 0
        Exit Function
    End If
    
    Dim lastCol As Integer
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    If lastCol < 1 Then
        LogMessage "ERROR: No headers found in FindColumn"
        FindColumn = 0
        Exit Function
    End If
    
    Dim i As Integer
    For i = 1 To lastCol
        If Trim(ws.Cells(1, i).Value) = headerName Then
            FindColumn = i
            Exit Function
        End If
    Next i
    
    LogMessage "WARNING: Column '" & headerName & "' not found"
    FindColumn = 0
    Exit Function
    
ErrorHandler:
    LogMessage "ERROR in FindColumn: " & Err.description
    FindColumn = 0
End Function

'---------------------------------------------------------------
' FindEnrolmentCell
' Purpose: Find the cell address of a subject's enrolment value
'          in the output sheet (matched by subject code + study period)
' Returns: Cell address string, or empty if not found
'---------------------------------------------------------------
Function FindEnrolmentCell(wsOutput As Worksheet, subjectCode As String, studyPeriod As String) As String
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    lastRow = wsOutput.Cells(wsOutput.Rows.count, "A").End(xlUp).Row
    
    Dim targetUID As String
    targetUID = subjectCode & "_" & studyPeriod & "_0"
    
    Dim i As Long
    For i = 4 To lastRow
        Dim cellValue As String
        cellValue = wsOutput.Cells(i, 1).Value
        
        If cellValue = targetUID Then
            FindEnrolmentCell = "$D$" & i
            Exit Function
        End If
    Next i
    
    FindEnrolmentCell = ""
    Exit Function
    
ErrorHandler:
    LogMessage "ERROR in FindEnrolmentCell: " & Err.description
    FindEnrolmentCell = ""
End Function

'---------------------------------------------------------------
' FormatSheet
' Purpose: Apply all formatting to a completed calculation sheet:
'          column widths, number formats, alignment, locking,
'          conditional formatting, marker block styling, protection
' Called by: GenerateSheet
'---------------------------------------------------------------
Sub FormatSheet(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        LogMessage "ERROR: Worksheet is Nothing in FormatSheet"
        Exit Sub
    End If
    
    Dim tStart As Double
    tStart = Timer
    
    On Error Resume Next
    ws.Unprotect
    Err.Clear
    On Error GoTo ErrorHandler
    
    ' =================================================================
    ' COLUMN WIDTHS (excluding marker blocks)
    ' =================================================================
    Dim widths As Variant
    widths = Array("A:A", 25, "B:C", 15, "D:D", 11.5, "E:E", 60, "F:H", 7, "I:J", 10.5, "K:L", 30, "M:M", 11.75, "N:N", 6.5, "O:O", 25, "P:R", 13, "S:S", 30)
    
    Dim i As Long
    For i = LBound(widths) To UBound(widths) Step 2
        ws.Columns(widths(i)).ColumnWidth = widths(i + 1)
    Next i
    
    ' =================================================================
    ' NUMBER FORMATS (excluding marker blocks)
    ' =================================================================
    Dim numFormats As Variant
    numFormats = Array("D:D", "0", "F:I", "0", "J:J", "0.00", "N:N", "0", "P:P", "0", "Q:R", "0.00")
    
    For i = LBound(numFormats) To UBound(numFormats) Step 2
        ws.Columns(numFormats(i)).NumberFormat = numFormats(i + 1)
    Next i
    
    ' =================================================================
    ' HORIZONTAL ALIGNMENT (excluding marker blocks)
    ' =================================================================
    Dim centerCols As Variant
    centerCols = Array("D:D", "F:J", "N:N", "P:R")
    
    For i = LBound(centerCols) To UBound(centerCols)
        ws.Columns(centerCols(i)).HorizontalAlignment = xlCenter
    Next i
    
    ' =================================================================
    ' WRAP TEXT (excluding marker blocks)
    ' =================================================================
    Dim wrapCols As Variant
    wrapCols = Array("E:E", "S:S")
    
    For i = LBound(wrapCols) To UBound(wrapCols)
        ws.Columns(wrapCols(i)).WrapText = True
    Next i
    
    ' =================================================================
    ' HIDDEN COLUMNS
    ' =================================================================
    Dim hideCols As Variant
    hideCols = Array("A:A")
    
    For i = LBound(hideCols) To UBound(hideCols)
        ws.Columns(hideCols(i)).Hidden = True
    Next i
    
    ' =================================================================
    ' LOCKED COLUMNS
    ' =================================================================
    ws.Cells.Locked = False
    
    Dim lockedCols As Variant
    lockedCols = Array("A:H", "J:J", "Q:R")
    
    For i = LBound(lockedCols) To UBound(lockedCols)
        ws.Columns(lockedCols(i)).Locked = True
    Next i
    
    ' =================================================================
    ' SPECIAL FORMATTING
    ' =================================================================
    ws.Columns("B:B").Font.Bold = True
    
    ' =================================================================
    ' CONDITIONAL FORMATTING (Continuing T&R)
    ' =================================================================
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    If lastRow > 3 Then
        Dim rng As Range
        Set rng = ws.Range("M4:M" & lastRow)
        
        On Error Resume Next
        rng.FormatConditions.Delete
        Err.Clear
        On Error GoTo ErrorHandler
        
        With rng.FormatConditions.add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Continuing T&R""")
            .Interior.Color = RGB(198, 239, 206)
            .Font.Color = RGB(0, 97, 0)
        End With
    End If
    
    ' =================================================================
    ' MARKER BLOCKS (Consolidated - 3 identical blocks)
    ' =================================================================
    Dim markerBaseCol As Variant
    Dim markerWidths As Variant
    Dim centerOffsets As Variant
    Dim wrapOffsets As Variant
    
    markerBaseCol = Array(20, 30, 40)  ' T, AD, AN
    markerWidths = Array(30, 60, 7, 7, 7, 13, 13, 30, 30, 10)
    centerOffsets = Array(2, 3, 4, 5, 6, 9)
    wrapOffsets = Array(1, 8)
    
    Dim m As Integer, baseCol As Integer, offset As Integer
    
    For m = 0 To 2
        baseCol = markerBaseCol(m)
        
        ' Column widths
        For offset = 0 To UBound(markerWidths)
            ws.Columns(baseCol + offset).ColumnWidth = markerWidths(offset)
        Next offset
        
        ' Number formats (cols 2-6: cols 2-5 are "0", col 6 is "0.00")
        For offset = 2 To 6
            ws.Columns(baseCol + offset).NumberFormat = IIf(offset = 6, "0.00", "0")
        Next offset
        
        ' Centre alignment
        For offset = 0 To UBound(centerOffsets)
            ws.Columns(baseCol + centerOffsets(offset)).HorizontalAlignment = xlCenter
        Next offset
        
        ' Wrap text
        For offset = 0 To UBound(wrapOffsets)
            ws.Columns(baseCol + wrapOffsets(offset)).WrapText = True
        Next offset
    Next m
    
    ' =================================================================
    ' SHEET PROTECTION
    ' =================================================================
    On Error Resume Next
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowInsertingRows:=False, AllowDeletingRows:=False
    
    If Err.Number <> 0 Then
        LogMessage "WARNING: Could not protect sheet " & ws.name & ": " & Err.description
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    LogMessage "FormatSheet completed", Timer - tStart
    Exit Sub
    
ErrorHandler:
    LogMessage "ERROR in FormatSheet: " & Err.description & " (Error " & Err.Number & ")"
End Sub

'---------------------------------------------------------------
' CollectionKeyExists
' Purpose: Check whether a key already exists in a VBA Collection
'---------------------------------------------------------------
Function CollectionKeyExists(col As Collection, key As String) As Boolean
    On Error Resume Next
    Dim temp As Variant
    temp = col(key)
    CollectionKeyExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

'---------------------------------------------------------------
' HasErrorsInLog
' Purpose: Scan the Process Log for ERROR/CRITICAL/FAILED entries
' Returns: True if any errors were logged
'---------------------------------------------------------------
Function HasErrorsInLog() As Boolean
    On Error Resume Next
    HasErrorsInLog = False
    
    If wsLog Is Nothing Then Exit Function
    
    Dim lastRow As Long
    lastRow = wsLog.Cells(wsLog.Rows.count, 2).End(xlUp).Row
    
    If lastRow < 2 Then Exit Function
    
    Dim i As Long
    For i = 2 To lastRow
        Dim msg As String
        msg = UCase(Trim(CStr(wsLog.Cells(i, 2).Value)))
        
        If InStr(msg, "ERROR") > 0 Or _
           InStr(msg, "CRITICAL") > 0 Or _
           InStr(msg, "FAILED") > 0 Then
            HasErrorsInLog = True
            Exit Function
        End If
    Next i
    
    On Error GoTo 0
End Function

'---------------------------------------------------------------
' CleanupPartialSheets
' Purpose: Delete partially generated FHY/SHY sheets after an
'          error to avoid leaving inconsistent state
' Called by: GenerateCalculationSheets (error handler)
'---------------------------------------------------------------
Sub CleanupPartialSheets(wb As Workbook)
    On Error Resume Next
    
    LogMessage "Attempting cleanup of partial sheets..."
    
    Application.DisplayAlerts = False
    
    Dim wsFHY As Worksheet, wsSHY As Worksheet
    Set wsFHY = wb.Sheets("FHY Calculations")
    Set wsSHY = wb.Sheets("SHY Calculations")
    
    If Not wsFHY Is Nothing Then
        wsFHY.Unprotect
        wsFHY.Delete
        LogMessage "Deleted partial FHY Calculations sheet"
    End If
    
    If Not wsSHY Is Nothing Then
        wsSHY.Unprotect
        wsSHY.Delete
        LogMessage "Deleted partial SHY Calculations sheet"
    End If
    
    Application.DisplayAlerts = True
    
    On Error GoTo 0
End Sub

'---------------------------------------------------------------
' CopyVBAModuleIfExists
' Purpose: Copy the LecturerRefresh VBA module from the source
'          workbook into the exported workbook so the Refresh
'          button works standalone
' Called by: ExportCalculationSheets
' Returns: True if the module was successfully copied
'---------------------------------------------------------------
Function CopyVBAModuleIfExists(sourceWB As Workbook, targetWB As Workbook, moduleName As String) As Boolean
    On Error GoTo ErrorHandler
    
    CopyVBAModuleIfExists = False
    LogMessage "=== CopyVBAModuleIfExists START ==="
    
    ' Check if module exists in source
    Dim sourceModule As Object
    Set sourceModule = Nothing
    On Error Resume Next
    Set sourceModule = sourceWB.VBProject.VBComponents(moduleName)
    On Error GoTo ErrorHandler
    
    If sourceModule Is Nothing Then
        LogMessage "Module '" & moduleName & "' not found in source workbook"
        Exit Function
    End If
    
    LogMessage "Found module: " & moduleName
    
    ' Remove from target if exists
    On Error Resume Next
    targetWB.VBProject.VBComponents.Remove targetWB.VBProject.VBComponents(moduleName)
    Err.Clear
    On Error GoTo ErrorHandler
    
    ' Create new standard module in target
    Dim targetModule As Object
    Set targetModule = targetWB.VBProject.VBComponents.add(1) ' 1 = vbext_ct_StdModule
    targetModule.name = moduleName
    
    ' Copy code line by line (NO TEMP FILES - THE KEY FIX!)
    Dim lineCount As Long
    lineCount = sourceModule.CodeModule.CountOfLines
    
    If lineCount > 0 Then
        Dim allCode As String
        allCode = sourceModule.CodeModule.lines(1, lineCount)
        
        ' ? FIXED: Use AddFromString instead of InsertLines
        targetModule.CodeModule.AddFromString allCode
        LogMessage "Copied " & lineCount & " lines"
    End If
    
    LogMessage "=== CopyVBAModuleIfExists SUCCESS ==="
    CopyVBAModuleIfExists = True
    Exit Function
    
ErrorHandler:
    LogMessage "ERROR in CopyVBAModuleIfExists: " & Err.Number & " - " & Err.description
    CopyVBAModuleIfExists = False
End Function

'---------------------------------------------------------------
' AddRefreshButtonsToSheets
' Purpose: Add a "Refresh Lecturer Data" button in row 2 of the
'          "Lecturer/Instructors" column on each calculation sheet,
'          and a bold warning note in row 3 spanning to the
'          "Activity Code" column. Column positions are resolved
'          dynamically from the header row so the code stays
'          correct even if columns are reordered in future.
' Called by: ExportCalculationSheets
'---------------------------------------------------------------

Sub AddRefreshButtonsToSheets(wb As Workbook)
    On Error GoTo ErrorHandler
    
    LogMessage "Adding refresh buttons to sheets..."
    
    Dim calcSheet As Worksheet
    Dim btn As Button
    Dim sheetNames As Variant
    sheetNames = Array("FHY Calculations", "SHY Calculations")
    
    Dim s As Integer
    For s = LBound(sheetNames) To UBound(sheetNames)
        Set calcSheet = Nothing
        On Error Resume Next
        Set calcSheet = wb.Sheets(sheetNames(s))
        On Error GoTo ErrorHandler
        
        If Not calcSheet Is Nothing Then
            ' ----------------------------------------------------------------
            ' Resolve column positions dynamically from the header row (row 1)
            ' ----------------------------------------------------------------
            Dim lecturerCol As Integer
            Dim activityCodeCol As Integer
            lecturerCol = FindColumn(calcSheet, "Lecturer/Instructors")
            activityCodeCol = FindColumn(calcSheet, "Activity Code")
            
            ' Fallback to hardcoded positions if headers not found
            If lecturerCol = 0 Then
                lecturerCol = 12   ' Column L (default)
                LogMessage "WARNING: 'Lecturer/Instructors' header not found — falling back to col 12 (L)"
            End If
            If activityCodeCol = 0 Then
                activityCodeCol = 15  ' Column O (default)
                LogMessage "WARNING: 'Activity Code' header not found — falling back to col 15 (O)"
            End If
            
            Dim lecturerColLetter As String
            lecturerColLetter = Split(calcSheet.Cells(1, lecturerCol).Address(True, False), "$")(0)
            
            ' ----------------------------------------------------------------
            ' Delete any existing buttons, then add the new one in row 2
            ' ----------------------------------------------------------------
            On Error Resume Next
            calcSheet.Buttons.Delete
            Err.Clear
            On Error GoTo ErrorHandler
            
            Dim btnCell As Range
            Set btnCell = calcSheet.Cells(2, lecturerCol)
            
            Set btn = calcSheet.Buttons.Add(btnCell.Left, btnCell.Top, btnCell.Width, btnCell.Height)
            
            With btn
                .OnAction = "RefreshLecturerData"
                .Caption = "Refresh Lecturer Data"
                .Name = "RefreshButton"
            End With
            
            LogMessage "Added refresh button to " & sheetNames(s) & " at " & lecturerColLetter & "2"
            
            ' ----------------------------------------------------------------
            ' Write bold overwrite warning in row 3 of the Lecturer column
            ' ----------------------------------------------------------------
            Dim noteCell As Range
            Set noteCell = calcSheet.Cells(3, lecturerCol)
            
            Dim activityCodeColLetter As String
            activityCodeColLetter = Split(calcSheet.Cells(1, activityCodeCol).Address(True, False), "$")(0)
            
            noteCell.Value = "Notes in col " & lecturerColLetter & " to " & activityCodeColLetter & " will be overwritten by refresh"
            noteCell.Font.Bold = True
            
            LogMessage "Added overwrite warning note to " & sheetNames(s) & " at " & lecturerColLetter & "3"
        End If
    Next s
    
    Exit Sub
    
ErrorHandler:
    LogMessage "ERROR in AddRefreshButtonsToSheets: " & Err.description
End Sub
