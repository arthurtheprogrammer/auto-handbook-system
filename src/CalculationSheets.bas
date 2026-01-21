Option Explicit

' ===== CONFIGURATION =====
Public Const ENABLE_REALTIME_LOG As Boolean = True
Public Const LOG_TO_STATUSBAR As Boolean = True

' ===== LOGGING SETUP =====
Dim wsLog As Worksheet
Private Const ENROLMENT_TRACKER_BASE As String = "https://unimelbcloud.sharepoint.com/teams/DepartmentofManagementMarketing-DepartmentOperations/Shared Documents/TEACHING MATRIX & ENROLMENT TRACKER/"

Sub InitializeProcessLog(wb As Workbook)
    Dim tempLog As Worksheet
    Dim existingLog As Worksheet
    Dim deleteSuccess As Boolean
    
    On Error Resume Next
    wb.Unprotect
    Set wsLog = Nothing
    
    Set existingLog = wb.Sheets("Process Log")
    If Not existingLog Is Nothing Then
        existingLog.Unprotect
        Application.DisplayAlerts = False
        existingLog.Delete
        Application.DisplayAlerts = True
        
        Set existingLog = Nothing
        Set existingLog = wb.Sheets("Process Log")
        If Not existingLog Is Nothing Then
            MsgBox "CRITICAL: Could not delete existing Process Log. It may be protected or in use.", vbCritical
            End
        End If
    End If
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    Set tempLog = wb.Sheets.add(After:=wb.Sheets(wb.Sheets.Count))
    
    On Error Resume Next
    tempLog.name = "Process Log"
    
    If Err.Number <> 0 Then
        Dim errNum As Long
        errNum = Err.Number
        Err.Clear
        
        tempLog.name = "ProcessLog_" & Format(Now, "hhmmss")
        If Err.Number <> 0 Then
            MsgBox "CRITICAL: Could not name log sheet (Error " & errNum & "). Process aborted.", vbCritical
            tempLog.Delete
            Application.ScreenUpdating = True
            End
        Else
            MsgBox "Log sheet created with alternative name: " & tempLog.name, vbInformation
        End If
    End If
    
    On Error GoTo ErrorHandler
    
    Set wsLog = Nothing
    Set wsLog = wb.Sheets(tempLog.name)
    If wsLog Is Nothing Then
        MsgBox "CRITICAL: Failed to verify log sheet creation.", vbCritical
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
        .Columns("B:B").ColumnWidth = 100
        .Columns("C:C").ColumnWidth = 15
    End With
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "CRITICAL ERROR in InitializeProcessLog: " & Err.description & " (" & Err.Number & ")", vbCritical
    End
End Sub

Sub LogMessage(msg As String, Optional elapsedTime As Double = -1)
    On Error GoTo ErrorHandler
    
    If LOG_TO_STATUSBAR Then
        Application.StatusBar = "Processing: " & Left(msg, 100)
    End If
    
    If Not wsLog Is Nothing Then
        Dim origScreenUpdating As Boolean
        origScreenUpdating = Application.ScreenUpdating
        
        If ENABLE_REALTIME_LOG Then
            Application.ScreenUpdating = True
        End If
        
        Dim nextRow As Long
        nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
        
        wsLog.Cells(nextRow, 1).Value = Now
        wsLog.Cells(nextRow, 2).Value = msg
        If elapsedTime >= 0 Then
            wsLog.Cells(nextRow, 3).Value = Format(elapsedTime, "0.00") & "s"
        End If
        wsLog.Cells(nextRow, 2).WrapText = True
        
        If ENABLE_REALTIME_LOG Then
            DoEvents
            Application.ScreenUpdating = origScreenUpdating
        End If
    End If
    
    Debug.Print msg
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR in LogMessage: " & Err.description
    On Error GoTo 0
End Sub

Function VerifyRequiredSheets(wb As Workbook) As Boolean
    Dim requiredSheets As Variant
    requiredSheets = Array("Dashboard", "SubjectList", "assessment data parsed", "teaching stream")
    
    Dim i As Integer
    Dim sheetExists As Boolean
    Dim sheetName As String
    Dim ws As Worksheet
    
    LogMessage "=== Verifying Required Sheets ==="
    VerifyRequiredSheets = True
    
    For i = 0 To UBound(requiredSheets)
        sheetName = requiredSheets(i)
        sheetExists = False
        
        On Error Resume Next
        Set ws = wb.Sheets(sheetName)
        If Err.Number = 0 And Not ws Is Nothing Then
            sheetExists = True
            LogMessage "Sheet found: " & sheetName
        Else
            LogMessage "MISSING SHEET: " & sheetName
            VerifyRequiredSheets = False
        End If
        Err.Clear
        On Error GoTo 0
        
        Set ws = Nothing
    Next i
    
    If Not VerifyRequiredSheets Then
        LogMessage "ERROR: Sheet verification failed. Missing required sheets."
        MsgBox "Required sheets are missing. Check Process Log for details.", vbCritical
    Else
        LogMessage "All required sheets verified."
    End If
End Function

' ===== HELPER FUNCTIONS =====
Function SafeArrayIndex(arr As Variant, idx As Integer, defaultVal As Variant) As Variant
    On Error Resume Next
    SafeArrayIndex = arr(idx)
    If Err.Number <> 0 Then SafeArrayIndex = defaultVal
    On Error GoTo 0
End Function

Sub GenerateCalculationSheets()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
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
        MsgBox "Invalid benchmark values in Dashboard sheet (cells C8, C9, C10).", vbCritical
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
        MsgBox "Errors were detected during generation." & vbCrLf & vbCrLf & _
               "Please check the 'Process Log' sheet for details." & vbCrLf & _
               "Files were NOT exported.", vbExclamation, "Generation Completed with Errors"
    Else
        MsgBox "Calculation sheets generated and exported successfully!" & vbCrLf & vbCrLf & _
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
    
    Dim errMsg As String
    errMsg = "Error in GenerateCalculationSheets: " & Err.description & " (Error " & Err.Number & ")"
    LogMessage errMsg
    
    Call CleanupPartialSheets(wb)
    
    MsgBox errMsg & vbCrLf & vbCrLf & "Check the 'Process Log' sheet for more details.", vbCritical
    On Error GoTo 0
End Sub

Function ExportCalculationSheets(wb As Workbook) As Boolean
    On Error GoTo ErrorHandler
    
    ExportCalculationSheets = False
    Application.ScreenUpdating = False
    LogMessage "ExportCalculationSheets: Starting"
    
    Dim yearValue As Variant
    yearValue = wb.Sheets("Dashboard").Range("C2").Value
    If Not IsNumeric(yearValue) Or yearValue = "" Then
        LogMessage "ERROR: Invalid year value in Dashboard C2: " & yearValue
        MsgBox "Invalid year value in Dashboard sheet (cell C2).", vbCritical
        Exit Function
    End If
    
    Dim newFileName As String
    newFileName = CStr(yearValue) & " Marking & Admin Support Calculations"
    
    Dim wsFHY As Worksheet, wsSHY As Worksheet
    On Error Resume Next
    Set wsFHY = wb.Sheets("FHY Calculations")
    Set wsSHY = wb.Sheets("SHY Calculations")
    On Error GoTo ErrorHandler
    
    If wsFHY Is Nothing And wsSHY Is Nothing Then
        LogMessage "ERROR: No calculation sheets found to export"
        MsgBox "No calculation sheets found to export.", vbCritical
        Exit Function
    End If
    
    Dim newWB As Workbook
    Set newWB = Workbooks.add
    
    Application.DisplayAlerts = False
    On Error Resume Next
    Do While newWB.Sheets.Count > 1
        newWB.Sheets(newWB.Sheets.Count).Delete
    Loop
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler
    
    If Not wsFHY Is Nothing Then
        wsFHY.Copy Before:=newWB.Sheets(1)
        
        Dim copiedFHY As Worksheet
        Set copiedFHY = newWB.Sheets(1)
        If copiedFHY.name <> "FHY Calculations" Then
            LogMessage "WARNING: FHY sheet name mismatch after copy: " & copiedFHY.name
        Else
            LogMessage "Copied FHY Calculations sheet"
        End If
    End If
    
    If Not wsSHY Is Nothing Then
        wsSHY.Copy After:=newWB.Sheets(newWB.Sheets.Count)
        LogMessage "Copied SHY Calculations sheet"
    End If
    
    Application.DisplayAlerts = False
    On Error Resume Next
    Dim i As Integer
    For i = newWB.Sheets.Count To 1 Step -1
        If newWB.Sheets(i).name <> "FHY Calculations" And _
           newWB.Sheets(i).name <> "SHY Calculations" Then
            newWB.Sheets(i).Delete
        End If
    Next i
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler
    
    If newWB.Sheets.Count = 0 Then
        LogMessage "ERROR: No sheets in exported workbook"
        newWB.Close SaveChanges:=False
        Exit Function
    End If
    
    newWB.Application.Calculation = xlCalculationAutomatic
    
    Dim savePath As String
    Dim sourceFilePath As String
    Dim basePath As String
    sourceFilePath = wb.FullName
    
    If InStr(1, sourceFilePath, "http://", vbTextCompare) > 0 Or _
       InStr(1, sourceFilePath, "https://", vbTextCompare) > 0 Then
        Dim lastSlash As Long
        lastSlash = InStrRev(sourceFilePath, "/")
        If lastSlash > 0 Then
            basePath = Left(sourceFilePath, lastSlash)
        Else
            basePath = wb.Path & "/"
        End If
    Else
        basePath = wb.Path & Application.PathSeparator
    End If
    
    savePath = GetUniqueFilename(basePath, newFileName, ".xlsx")
    LogMessage "Target save path: " & savePath
    
    On Error Resume Next
    newWB.SaveAs FileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    
    If Err.Number <> 0 Then
        Dim saveErr As Long
        saveErr = Err.Number
        LogMessage "Save failure in primary location (Code " & saveErr & ") Path: " & savePath
        
        Dim tempPath As String
        Dim tempBasePath As String
        tempBasePath = Environ("TEMP") & Application.PathSeparator
        tempPath = GetUniqueFilename(tempBasePath, newFileName, ".xlsx")
        
        Err.Clear
        newWB.SaveAs FileName:=tempPath, FileFormat:=xlOpenXMLWorkbook
        
        If Err.Number = 0 Then
            MsgBox "Could not save to SharePoint location." & vbCrLf & vbCrLf & _
                   "File saved to:" & vbCrLf & tempPath, vbInformation
            savePath = tempPath
            LogMessage "Saved to temp location: " & tempPath
        Else
            LogMessage "File save failed everywhere. Path: " & tempPath
            MsgBox "Could not save file. File not saved anywhere.", vbExclamation
            newWB.Close SaveChanges:=False
            Application.ScreenUpdating = True
            Exit Function
        End If
    Else
        LogMessage "Saved new workbook to: " & savePath
    End If
    
    On Error GoTo ErrorHandler
    
    newWB.Close SaveChanges:=False
    
    Application.DisplayAlerts = False
    On Error Resume Next
    If Not wsFHY Is Nothing Then
        wb.Sheets("FHY Calculations").Delete
        Set wsFHY = Nothing
        Set wsFHY = wb.Sheets("FHY Calculations")
        If Not wsFHY Is Nothing Then
            LogMessage "WARNING: Could not delete FHY Calculations from source"
        End If
    End If
    If Not wsSHY Is Nothing Then
        wb.Sheets("SHY Calculations").Delete
        Set wsSHY = Nothing
        Set wsSHY = wb.Sheets("SHY Calculations")
        If Not wsSHY Is Nothing Then
            LogMessage "WARNING: Could not delete SHY Calculations from source"
        End If
    End If
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = True
    LogMessage "Export completed successfully"
    ExportCalculationSheets = True
    Exit Function
    
ErrorHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    LogMessage "ERROR in ExportCalculationSheets: " & Err.description & " (Error " & Err.Number & ")"
    
    On Error Resume Next
    If Not newWB Is Nothing Then newWB.Close SaveChanges:=False
    On Error GoTo 0
    
    ExportCalculationSheets = False
End Function

Function GetUniqueFilename(basePath As String, baseName As String, ext As String) As String
    On Error GoTo ErrorHandler
    
    Dim fullPath As String
    Dim timestamp As String
    Dim counter As Integer
    
    fullPath = basePath & baseName & ext
    
    If Dir(fullPath) = "" Then
        GetUniqueFilename = fullPath
        Exit Function
    End If
    
    timestamp = Format(Now, "_yyyymmdd_hhmmss")
    fullPath = basePath & baseName & timestamp & ext
    
    counter = 1
    Do While Dir(fullPath) <> ""
        fullPath = basePath & baseName & timestamp & "_" & counter & ext
        counter = counter + 1
        
        If counter > 100 Then
            Debug.Print "CRITICAL FAILURE in GetUniqueFilename: Could not generate unique filename after 100 attempts"
            GetUniqueFilename = basePath & baseName & timestamp & "_" & Format(Timer * 1000, "0") & ext
            Exit Function
        End If
    Loop
    
    GetUniqueFilename = fullPath
    Exit Function
    
ErrorHandler:
    Debug.Print "Internal error suppressed in GetUniqueFilename: " & Err.description
    On Error Resume Next
    GetUniqueFilename = basePath & baseName & ext
    On Error GoTo 0
End Function

Function GetBenchmarkValue(wb As Workbook, sheetName As String, cellRef As String, defaultValue As Double) As Double
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = wb.Sheets(sheetName)
    
    If ws Is Nothing Then
        LogMessage "WARNING: Sheet " & sheetName & " not found, using default: " & defaultValue
        GetBenchmarkValue = defaultValue
        Exit Function
    End If
    
    Dim cellValue As Variant
    cellValue = ws.Range(cellRef).Value
    
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
    Set wsOutput = wb.Sheets.add(After:=wb.Sheets(wb.Sheets.Count - 1))
    wsOutput.name = sheetName
    
    Set checkSheet = Nothing
    Set checkSheet = wb.Sheets(sheetName)
    If checkSheet Is Nothing Then
        LogMessage "ERROR: Failed to verify creation of " & sheetName
        GenerateSheet = False
        Exit Function
    End If
    
    Call CreateHeaders(wsOutput)
    
    wsOutput.Range("D2").Value = "(automatic)"
    wsOutput.Range("J2").Value = wordBench & " words/hr"
    wsOutput.Range("J3").Value = examBench & " exams/hr"
    wsOutput.Range("P2").Value = markingSupportBench & " hrs/stream"
    
    wsOutput.Range("Y2").Value = wordBench & " words/hr"
    wsOutput.Range("Y3").Value = examBench & " exams/hr"
    wsOutput.Range("AH2").Value = wordBench & " words/hr"
    wsOutput.Range("AH3").Value = examBench & " exams/hr"
    wsOutput.Range("AR2").Value = wordBench & " words/hr"
    wsOutput.Range("AR3").Value = examBench & " exams/hr"
    
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
    
    If subjectData.Count = 0 Then
        LogMessage "WARNING: No subjects found for " & groupedPeriod
        GenerateSheet = True
        Exit Function
    End If
    
    LogMessage "Found " & subjectData.Count & " subjects for " & groupedPeriod
    
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

Sub CreateHeaders(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim headers As Variant
    headers = Array("UID", "Subject Code", "Study Period", "Enrolment #", "Assessment Details", _
                    "Word Count", "Exam", "Group Size", "Assessment Quantity", "Marking Hours", _
                    "Assessment Notes", "Lecturer/Instructors", "Status", "Stream #", "Activity Code", _
                    "Allocated Marking", "Marking Support Hours Available", "Lecturer Notes", _
                    "Marker 1", "Assessment Details", "Word Count", "Exam", "Group Size", "Assessment Quantity", "Marking Allocation", "Email", "Arrangement Notes", "Contract Requested", _
                    "Marker 2", "Assessment Details", "Word Count", "Exam", "Group Size", "Assessment Quantity", "Marking Allocation", "Email", "Arrangement Notes", "Contract Requested", _
                    "Marker 3", "Assessment Details", "Word Count", "Exam", "Group Size", "Assessment Quantity", "Marking Allocation", "Email", "Arrangement Notes", "Contract Requested")
    
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
    lastRow = wsSubjects.Cells(wsSubjects.Rows.Count, "B").End(xlUp).Row
    
    If lastRow < 2 Then
        LogMessage "WARNING: No data in SubjectList sheet"
        Set GetFilteredSubjectsWithAssessments = New Collection
        Exit Function
    End If
    
    Dim assessmentDict As Collection
    Set assessmentDict = New Collection
    
    Dim lastAssRow As Long
    lastAssRow = wsAssessment.Cells(wsAssessment.Rows.Count, "A").End(xlUp).Row
    
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

Function PopulateSheetData(wb As Workbook, wsOutput As Worksheet, subjectData As Collection, wordBench As Double, examBench As Double, markingSupportBench As Double) As Boolean
    On Error GoTo ErrorHandler
    
    PopulateSheetData = False
    Dim tStart As Double
    tStart = Timer
    
    If subjectData Is Nothing Or subjectData.Count = 0 Then
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
    
    ' Categorize subjects
    Dim summerSubjects As New Collection
    Dim semester1Subjects As New Collection
    Dim winterSubjects As New Collection
    Dim semester2Subjects As New Collection
    
    Dim subject As Variant
    For Each subject In subjectData
        Dim studyPeriod As String
        studyPeriod = SafeArrayIndex(subject, 3, "")
        
        Dim category As String
        category = CategorizeStudyPeriod(studyPeriod, groupedPeriod)
        
        Select Case category
            Case "SUMMER": summerSubjects.add subject
            Case "SEMESTER 1": semester1Subjects.add subject
            Case "WINTER": winterSubjects.add subject
            Case "SEMESTER 2": semester2Subjects.add subject
        End Select
    Next subject
    
    ' Collection to store marker formula info for batch processing
    Dim markerFormulaQueue As New Collection
    
    ' Process subjects by category
    If groupedPeriod = "FHY" Then
        If summerSubjects.Count > 0 Then
            For Each subject In summerSubjects
                If Not ProcessSubject(wb, wsOutput, subject, currentRow, wordBench, examBench, markingSupportBench, markerFormulaQueue) Then
                    LogMessage "ERROR: Failed to process subject in SUMMER"
                End If
            Next subject
        End If
        
        If semester1Subjects.Count > 0 Then
            wsOutput.Rows(currentRow).Interior.Color = RGB(0, 0, 0)
            wsOutput.Cells(currentRow, 2).Value = "SEMESTER 1"
            wsOutput.Cells(currentRow, 2).Font.Color = RGB(255, 255, 255)
            wsOutput.Cells(currentRow, 2).Font.Bold = True
            currentRow = currentRow + 1
            
            For Each subject In semester1Subjects
                If Not ProcessSubject(wb, wsOutput, subject, currentRow, wordBench, examBench, markingSupportBench, markerFormulaQueue) Then
                    LogMessage "ERROR: Failed to process subject in SEMESTER 1"
                End If
            Next subject
        End If
    Else
        If winterSubjects.Count > 0 Then
            For Each subject In winterSubjects
                If Not ProcessSubject(wb, wsOutput, subject, currentRow, wordBench, examBench, markingSupportBench, markerFormulaQueue) Then
                    LogMessage "ERROR: Failed to process subject in WINTER"
                End If
            Next subject
        End If
        
        If semester2Subjects.Count > 0 Then
            wsOutput.Rows(currentRow).Interior.Color = RGB(0, 0, 0)
            wsOutput.Cells(currentRow, 2).Value = "SEMESTER 2"
            wsOutput.Cells(currentRow, 2).Font.Color = RGB(255, 255, 255)
            wsOutput.Cells(currentRow, 2).Font.Bold = True
            currentRow = currentRow + 1
            
            For Each subject In semester2Subjects
                If Not ProcessSubject(wb, wsOutput, subject, currentRow, wordBench, examBench, markingSupportBench, markerFormulaQueue) Then
                    LogMessage "ERROR: Failed to process subject in SEMESTER 2"
                End If
            Next subject
        End If
    End If
    
    ' NOW APPLY ALL MARKER FORMULAS IN BATCH
    LogMessage "Applying marker formulas in batch..."
    Dim formulaInfo As Variant
    For Each formulaInfo In markerFormulaQueue
        Call SetMarkerBlockFormulas(wsOutput, CLng(formulaInfo(0)), CLng(formulaInfo(1)), CInt(formulaInfo(2)), CStr(formulaInfo(3)), CStr(formulaInfo(4)))
    Next formulaInfo
    
    LogMessage "PopulateSheetData completed", Timer - tStart
    PopulateSheetData = True
    Exit Function
    
ErrorHandler:
    LogMessage "ERROR in PopulateSheetData: " & Err.description & " (Error " & Err.Number & ")"
    PopulateSheetData = False
End Function

Function ProcessSubject(wb As Workbook, wsOutput As Worksheet, ByRef subject As Variant, ByRef currentRow As Long, wordBench As Double, examBench As Double, markingSupportBench As Double, ByRef markerFormulaQueue As Collection) As Boolean
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
    
    If assessments Is Nothing Or assessments.Count = 0 Then
        ProcessSubject = True
        Exit Function
    End If
    
    Dim lecturers As Collection
    Set lecturers = GetLecturersForSubjectFlexible(wb, subjectCode, studyPeriod)
    If lecturers Is Nothing Then Set lecturers = New Collection
    
    ' Calculate structure
    Dim uidCount As Integer
    uidCount = assessments.Count
    
    Dim assessmentRows As Integer
    assessmentRows = uidCount + 2
    
    Dim totalRowsNeeded As Integer
    totalRowsNeeded = Application.WorksheetFunction.Max(assessmentRows, lecturers.Count + 1)
    
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
    For i = 1 To assessments.Count
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
        
        wsOutput.Cells(currentRow, 4).Formula = "=IFERROR(INDEX(" & enrolPath & "$I:$I,SUMPRODUCT((" & enrolPath & "$A:$A=B" & currentRow & ")*(" & enrolPath & "$C:$C=C" & currentRow & ")*ROW(" & enrolPath & "$A:$A))),0)"
    End If
    
    ' Format header row
    wsOutput.Rows(currentRow).Interior.Color = RGB(192, 192, 192)
    wsOutput.Cells(currentRow + assessmentRows - 1, 5).Font.Bold = True
    
    ' Add assessment formulas
    Dim formulaRow As Long
    For i = 1 To assessments.Count
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
    If lecturers.Count > 0 Then
        Call PopulateLecturerData(wsOutput, lecturers, subjectStartRow + 1, totalRowIndex, markingSupportBench)
    End If
    
    ' QUEUE marker formulas for batch processing (don't execute yet)
    Dim markerInfo(0 To 4) As Variant
    markerInfo(0) = subjectStartRow
    markerInfo(1) = subjectEndRow
    markerInfo(2) = 1  ' Marker number
    markerInfo(3) = subjectCode
    markerInfo(4) = studyPeriod
    markerFormulaQueue.add markerInfo
    
    Dim markerInfo2(0 To 4) As Variant
    markerInfo2(0) = subjectStartRow
    markerInfo2(1) = subjectEndRow
    markerInfo2(2) = 2
    markerInfo2(3) = subjectCode
    markerInfo2(4) = studyPeriod
    markerFormulaQueue.add markerInfo2
    
    Dim markerInfo3(0 To 4) As Variant
    markerInfo3(0) = subjectStartRow
    markerInfo3(1) = subjectEndRow
    markerInfo3(2) = 3
    markerInfo3(3) = subjectCode
    markerInfo3(4) = studyPeriod
    markerFormulaQueue.add markerInfo3
    
    ' Add checkboxes
    Call AddMarkerCheckboxes(wsOutput, subjectStartRow, 1)
    Call AddMarkerCheckboxes(wsOutput, subjectStartRow, 2)
    Call AddMarkerCheckboxes(wsOutput, subjectStartRow, 3)
    
    currentRow = subjectEndRow + 1
    
    ProcessSubject = True
    Exit Function
    
ErrorHandler:
    LogMessage "ERROR in ProcessSubject (" & subjectCode & "): " & Err.description
    ProcessSubject = False
End Function


' Helper function to build assessment quantity formula
Function BuildAssessmentQuantityFormula(rowNum As Long, assessUID As String, subjectCode As String, studyPeriod As String) As String
    ' This builds the VLOOKUP formula for assessment quantity
    BuildAssessmentQuantityFormula = "=IFERROR(VLOOKUP(""" & assessUID & """,FILTER('assessment data parsed'!$A:$I,('assessment data parsed'!$B:$B=""" & subjectCode & """)*('assessment data parsed'!$C:$C=""" & studyPeriod & """)),9,FALSE),IFERROR(VLOOKUP(""" & assessUID & """,FILTER('assessment data parsed'!$A:$I,('assessment data parsed'!$B:$B=""" & subjectCode & """)*('assessment data parsed'!$C:$C=""All"")),9,FALSE),""""))"
End Function



Sub AddMarkerCheckboxes(wsOutput As Worksheet, subjectStartRow As Long, markerNum As Integer)
    On Error GoTo ErrorHandler
    
    ' Add data validation dropdown to first assessment row (_1), not header (_0)
    Dim targetRow As Long
    targetRow = subjectStartRow + 1  ' First assessment row
    
    Dim contractCol As Integer
    Select Case markerNum
        Case 1
            contractCol = 28  ' Column AB
        Case 2
            contractCol = 38  ' Column AL
        Case 3
            contractCol = 48  ' Column AV
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
        .add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="N,Y"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = False
    End With
    
    ' Set default value to "No"
    If cellObj.Value = "" Then
        cellObj.Value = "N"
    End If
    
    If Err.Number <> 0 Then
        LogMessage "WARNING: Could not add data validation for Marker " & markerNum & " at row " & targetRow
    End If
    
    On Error GoTo ErrorHandler
    Exit Sub
    
ErrorHandler:
    LogMessage "ERROR in AddMarkerCheckboxes: " & Err.description
End Sub

Sub SetMarkerBlockFormulas(wsOutput As Worksheet, subjectStartRow As Long, subjectEndRow As Long, markerNum As Integer, subjectCode As String, studyPeriod As String)
    On Error GoTo ErrorHandler
    
    Dim baseCol As Integer, benchmarkCol As Integer
    Select Case markerNum
        Case 1
            baseCol = 19  ' Column S
            benchmarkCol = 25  ' Column Y
        Case 2
            baseCol = 29  ' Column AC
            benchmarkCol = 35  ' Column AI
        Case 3
            baseCol = 39  ' Column AM
            benchmarkCol = 44  ' Column AR
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
    numRows = subjectEndRow - subjectStartRow  ' Changed from +1
    
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
        arrayRow = r - subjectStartRow  ' Adjusted offset
        
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
            
            allocFormulas(arrayRow, 1) = "=IF(" & qtyColLetter & r & "="""","""", IF(ISNUMBER(" & qtyColLetter & r & "), IF(ISNUMBER(" & wordColLetter & r & "), " & qtyColLetter & r & "*" & wordColLetter & r & "/VALUE(LEFT($" & benchmarkColLetter & "$2,FIND("" "",$" & benchmarkColLetter & "$2)-1)), IF(ISNUMBER(" & examColLetter & r & "), " & qtyColLetter & r & "/VALUE(LEFT($" & benchmarkColLetter & "$3,FIND("" "",$" & benchmarkColLetter & "$3)-1)),"""")), " & qtyColLetter & r & "))"
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



Function ColLetter(colNum As Integer) As String
    ColLetter = Split(Cells(1, colNum).Address, "$")(1)
End Function

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
        lastRow = wsAssessment.Cells(wsAssessment.Rows.Count, "A").End(xlUp).Row
        
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
        lastRow = wsTeaching.Cells(wsTeaching.Rows.Count, "B").End(xlUp).Row
        
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
        
        If lecturers.Count > 0 Then Exit For
    Next variation
    
    Set GetLecturersForSubjectFlexible = lecturers
    Exit Function
    
ErrorHandler:
    LogMessage "ERROR in GetLecturersForSubjectFlexible: " & Err.description & " (Error " & Err.Number & ")"
    Set GetLecturersForSubjectFlexible = Nothing
End Function

Sub PopulateLecturerData(wsOutput As Worksheet, lecturers As Collection, startRow As Long, totalRow As Long, markingSupportBench As Double)
    On Error GoTo ErrorHandler
    
    If lecturers Is Nothing Or lecturers.Count = 0 Then Exit Sub
    
    ' Build lecturer data array (Columns L-O: Name, Status, Streams, Activity ID)
    Dim lecturerData As Variant
    ReDim lecturerData(1 To lecturers.Count, 1 To 4)
    
    Dim i As Integer
    Dim lecItem As Variant
    
    For i = 1 To lecturers.Count
        lecItem = lecturers(i)
        lecturerData(i, 1) = SafeArrayIndex(lecItem, 0, "")  ' Name
        lecturerData(i, 2) = SafeArrayIndex(lecItem, 1, "")  ' Status
        lecturerData(i, 3) = SafeArrayIndex(lecItem, 3, "")  ' Streams
        lecturerData(i, 4) = SafeArrayIndex(lecItem, 2, "")  ' Activity ID
    Next i
    
    ' Write lecturer data (Columns L-O: 12-15)
    wsOutput.Cells(startRow, 12).Resize(lecturers.Count, 4).Value = lecturerData
    
    ' Bold first lecturer (subject coordinator)
    If lecturers.Count > 0 Then
        wsOutput.Cells(startRow, 12).Font.Bold = True
    End If
    
    ' BUILD FORMULA ARRAYS for columns P and Q
    Dim formulas() As Variant
    ReDim formulas(1 To lecturers.Count, 1 To 2)
    
    Dim currentRow As Long
    For i = 1 To lecturers.Count
        currentRow = startRow + i - 1
        
        ' Column P (16): Allocated Marking
        formulas(i, 1) = "=IF(M" & currentRow & "=""Continuing T&R"",N" & currentRow & "*VALUE(LEFT($P$2,FIND("" "",$P$2)-1)),"""")"
        
        ' Column Q (17): Marking Support Hours Available
        formulas(i, 2) = "=IF(P" & currentRow & "="""","""",$J$" & totalRow & "-P" & currentRow & ")"
    Next i
    
    ' BATCH WRITE both formulas at once
    wsOutput.Cells(startRow, 16).Resize(lecturers.Count, 2).Formula = formulas
    
    Exit Sub
    
ErrorHandler:
    LogMessage "ERROR in PopulateLecturerData: " & Err.description & " (Error " & Err.Number & ")"
End Sub

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

Function CategorizeStudyPeriod(studyPeriod As String, groupedPeriod As String) As String
    Dim sp As String
    sp = LCase(studyPeriod)
    
    If InStr(sp, "summer") > 0 Or InStr(sp, "october") > 0 Or InStr(sp, "november") > 0 Or InStr(sp, "term 1") > 0 Then
        CategorizeStudyPeriod = "SUMMER"
        Exit Function
    End If
    
    If InStr(sp, "winter") > 0 Or InStr(sp, "june") > 0 Or InStr(sp, "july") > 0 Then
        CategorizeStudyPeriod = "WINTER"
        Exit Function
    End If
    
    If InStr(sp, "semester 1") > 0 Or InStr(sp, "term 2") > 0 Then
        CategorizeStudyPeriod = "SEMESTER 1"
        Exit Function
    End If
    
    If InStr(sp, "semester 2") > 0 Or InStr(sp, "term 3") > 0 Or InStr(sp, "term 4") > 0 Then
        CategorizeStudyPeriod = "SEMESTER 2"
        Exit Function
    End If
    
    If groupedPeriod = "FHY" Then
        CategorizeStudyPeriod = "SEMESTER 1"
    Else
        CategorizeStudyPeriod = "SEMESTER 2"
    End If
End Function

Function FindColumn(ws As Worksheet, headerName As String) As Integer
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        LogMessage "ERROR: Worksheet is Nothing in FindColumn"
        FindColumn = 0
        Exit Function
    End If
    
    Dim lastCol As Integer
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
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

Function FindEnrolmentCell(wsOutput As Worksheet, subjectCode As String, studyPeriod As String) As String
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    lastRow = wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).Row
    
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
    
    ' ... (Keep all column width settings - unchanged)
    ws.Columns("A:A").ColumnWidth = 25
    ws.Columns("B:C").ColumnWidth = 15
    ws.Columns("D:D").ColumnWidth = 11.5
    ws.Columns("E:E").ColumnWidth = 60
    ws.Columns("F:H").ColumnWidth = 7
    ws.Columns("I:J").ColumnWidth = 10.5
    ws.Columns("K:K").ColumnWidth = 30
    ws.Columns("L:L").ColumnWidth = 30
    ws.Columns("M:M").ColumnWidth = 11.75
    ws.Columns("N:N").ColumnWidth = 6.5
    ws.Columns("O:O").ColumnWidth = 25
    ws.Columns("P:Q").ColumnWidth = 14
    ws.Columns("R:R").ColumnWidth = 30
    
    ws.Columns("S:S").ColumnWidth = 30
    ws.Columns("T:T").ColumnWidth = 60
    ws.Columns("U:W").ColumnWidth = 7
    ws.Columns("X:Y").ColumnWidth = 13
    ws.Columns("Z:AA").ColumnWidth = 30
    ws.Columns("AB:AB").ColumnWidth = 10
    
    ws.Columns("AC:AC").ColumnWidth = 30
    ws.Columns("AD:AD").ColumnWidth = 60
    ws.Columns("AE:AG").ColumnWidth = 7
    ws.Columns("AH:AI").ColumnWidth = 13
    ws.Columns("AJ:AK").ColumnWidth = 30
    ws.Columns("AL:AL").ColumnWidth = 10
    
    ws.Columns("AM:AM").ColumnWidth = 30
    ws.Columns("AN:AN").ColumnWidth = 60
    ws.Columns("AO:AQ").ColumnWidth = 7
    ws.Columns("AR:AS").ColumnWidth = 13
    ws.Columns("AT:AU").ColumnWidth = 30
    ws.Columns("AV:AV").ColumnWidth = 10
    
    ' ... (Keep wrap text settings - unchanged)
    ws.Columns("E:E").WrapText = True
    ws.Columns("R:R").WrapText = True
    ws.Columns("T:T").WrapText = True
    ws.Columns("AA:AB").WrapText = True
    ws.Columns("AD:AD").WrapText = True
    ws.Columns("AJ:AL").WrapText = True
    ws.Columns("AN:AN").WrapText = True
    ws.Columns("AT:AV").WrapText = True
    
    ws.Columns("B:B").Font.Bold = True
    
    ' ... (Keep number formats - unchanged)
    ws.Columns("D:D").NumberFormat = "0"
    ws.Columns("F:F").NumberFormat = "0"
    ws.Columns("G:G").NumberFormat = "0"
    ws.Columns("H:H").NumberFormat = "0"
    ws.Columns("I:I").NumberFormat = "0"
    ws.Columns("J:J").NumberFormat = "0.00"
    ws.Columns("N:N").NumberFormat = "0"
    ws.Columns("P:Q").NumberFormat = "0.00"
    
    ws.Columns("U:W").NumberFormat = "0"
    ws.Columns("X:X").NumberFormat = "0"
    ws.Columns("Y:Y").NumberFormat = "0.00"
    ws.Columns("AE:AG").NumberFormat = "0"
    ws.Columns("AH:AH").NumberFormat = "0"
    ws.Columns("AI:AI").NumberFormat = "0.00"
    ws.Columns("AO:AQ").NumberFormat = "0"
    ws.Columns("AR:AR").NumberFormat = "0"
    ws.Columns("AS:AS").NumberFormat = "0.00"
    
    ws.Columns("A:A").Hidden = True
    ws.Columns("D:D").HorizontalAlignment = xlCenter
    ws.Columns("F:J").HorizontalAlignment = xlCenter
    ws.Columns("N:N").HorizontalAlignment = xlCenter
    ws.Columns("P:Q").HorizontalAlignment = xlCenter
    
    ws.Columns("U:Y").HorizontalAlignment = xlCenter
    ws.Columns("AE:AI").HorizontalAlignment = xlCenter
    ws.Columns("AO:AS").HorizontalAlignment = xlCenter
    
    ' CENTER ALIGN CONTRACT REQUESTED COLUMNS
    ws.Columns("AB:AB").HorizontalAlignment = xlCenter  ' Marker 1
    ws.Columns("AL:AL").HorizontalAlignment = xlCenter  ' Marker 2
    ws.Columns("AV:AV").HorizontalAlignment = xlCenter  ' Marker 3
    
    ' ... (Keep conditional formatting and protection - unchanged)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
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
    
    ws.Cells.Locked = False
    ws.Columns("B:C").Locked = True
    ws.Columns("E:E").Locked = True
    ws.Columns("F:H").Locked = True
    ws.Columns("J:J").Locked = True
    ws.Columns("P:Q").Locked = True
    
    On Error Resume Next
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, _
                AllowFormattingCells:=True, AllowFormattingColumns:=True, _
                AllowFormattingRows:=True, AllowInsertingRows:=False, _
                AllowDeletingRows:=False
    
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

Function CollectionKeyExists(col As Collection, key As String) As Boolean
    On Error Resume Next
    Dim temp As Variant
    temp = col(key)
    CollectionKeyExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Function HasErrorsInLog() As Boolean
    On Error Resume Next
    HasErrorsInLog = False
    
    If wsLog Is Nothing Then Exit Function
    
    Dim lastRow As Long
    lastRow = wsLog.Cells(wsLog.Rows.Count, 2).End(xlUp).Row
    
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
