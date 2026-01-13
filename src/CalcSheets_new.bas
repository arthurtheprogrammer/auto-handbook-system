'===============================================================
' Module: MarkingSupport
' Purpose: Generates FHY/SHY marking support calculation sheets from assessment and teaching data
' Main Entry: GenerateCalculationSheets() - called from Dashboard button
' Output: Two sheets (FHY Calculations, SHY Calculations) exported as separate workbook
' Author: Arthur Chen
' Created: 2025-10-15
' Last Updated: 2026-01-13
' Dependencies:
'   - Excel 365 with external workbook link support
'   - Source sheets: Dashboard, SubjectList, assessment data parsed, teaching stream
'   - External: M&M_Enrolment Tracker (SharePoint)
'   - External: Automated Handbook Data System (SharePoint)
'===============================================================

Option Explicit

'===== CONFIGURATION =====
Public Const ENABLE_REALTIME_LOG As Boolean = True
Public Const LOG_TO_STATUSBAR As Boolean = True

'===== MODULE-LEVEL VARIABLES =====
Dim wsLog As Worksheet

'===== EXTERNAL FILE PATHS =====
' These paths are configurable via Dashboard sheet
Private Const ENROLMENT_TRACKER_BASE As String = "https://unimelbcloud.sharepoint.com/teams/DepartmentofManagementMarketing-DepartmentOperations/Shared Documents/TEACHING MATRIX & ENROLMENT TRACKER/"
Private Const TEACHING_STREAM_BASE As String = "https://unimelbcloud.sharepoint.com/teams/DepartmentofManagementMarketing-DepartmentOperations/Shared Documents/TEACHING SUPPORT/Handbook (Course & Subject Changes)/Auto Handbook System/"
Private Const TEACHING_STREAM_FILE As String = "Automated Handbook Data System.xlsm"

'===============================================================
' SECTION 1: MAIN WORKFLOW
'===============================================================

'---------------------------------------------------------------
' GenerateCalculationSheets
' Purpose: Main entry point - generates both FHY and SHY calculation sheets
' Called by: Dashboard button click / Integration VBA trigger
' Output: Exports calculation sheets to new workbook
'---------------------------------------------------------------
Sub GenerateCalculationSheets()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim tStart As Double, tPhase As Double
    tStart = Timer
    
    ' Save original Excel state
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation
    Dim origEnableEvents As Boolean
    
    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation
    origEnableEvents = Application.EnableEvents
    
    On Error GoTo ErrorHandler
    
    ' Optimize Excel for batch processing
    wb.Unprotect
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Initialize logging
    Call InitializeProcessLog(wb)
    
    LogMessage "=== Starting GenerateCalculationSheets ==="
    LogMessage "Workbook: " & wb.name
    
    ' Verify all required source sheets exist
    If Not VerifyRequiredSheets(wb) Then
        GoTo CleanExit
    End If
    
    ' Load benchmark values from Dashboard
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
    
    ' Generate FHY sheet
    tPhase = Timer
    LogMessage "Generating FHY sheet..."
    If Not GenerateSheet(wb, "FHY", wordCountBenchmark, examBenchmark, markingSupportBenchmark) Then
        LogMessage "ERROR: FHY sheet generation failed"
        GoTo CleanExit
    End If
    LogMessage "FHY sheet complete", Timer - tPhase
    
    ' Generate SHY sheet
    tPhase = Timer
    LogMessage "Generating SHY sheet..."
    If Not GenerateSheet(wb, "SHY", wordCountBenchmark, examBenchmark, markingSupportBenchmark) Then
        LogMessage "ERROR: SHY sheet generation failed"
        GoTo CleanExit
    End If
    LogMessage "SHY sheet complete", Timer - tPhase
    
    ' Check for errors before export
    If HasErrorsInLog() Then
        LogMessage "=== Errors detected - Export cancelled ==="
        GoTo CleanExit
    End If
    
    ' Export calculation sheets to new workbook
    LogMessage "=== Generation Complete - Starting Export ==="
    
    tPhase = Timer
    If Not ExportCalculationSheets(wb) Then
        LogMessage "ERROR: Export failed"
        GoTo CleanExit
    End If
    LogMessage "Export complete", Timer - tPhase
    
    LogMessage "=== Total Time: " & Format(Timer - tStart, "0.00") & " seconds ==="

CleanExit:
    ' Restore Excel state
    Application.Calculation = origCalculation
    Application.ScreenUpdating = origScreenUpdating
    Application.EnableEvents = origEnableEvents
    Application.StatusBar = False
    
    ' Show completion message
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

'---------------------------------------------------------------
' GenerateSheet
' Purpose: Creates a single calculation sheet (FHY or SHY)
' Inputs:
'   - groupedPeriod: "FHY" or "SHY"
'   - Benchmark values for calculations
' Output: New sheet with formatted subject blocks
'---------------------------------------------------------------
Function GenerateSheet(wb As Workbook, groupedPeriod As String, wordBench As Double, examBench As Double, markingSupportBench As Double) As Boolean
    On Error GoTo ErrorHandler
    
    GenerateSheet = False
    Dim tStart As Double
    tStart = Timer
    
    Dim sheetName As String
    sheetName = groupedPeriod & " Calculations"
    
    LogMessage "GenerateSheet: Starting for " & groupedPeriod
    
    ' Delete existing sheet if present
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets(sheetName).Delete
    Err.Clear
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    
    ' Verify deletion succeeded
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
    
    ' Create new sheet
    Dim wsOutput As Worksheet
    Set wsOutput = wb.Sheets.add(After:=wb.Sheets(wb.Sheets.Count - 1))
    wsOutput.name = sheetName
    
    ' Verify creation
    Set checkSheet = Nothing
    Set checkSheet = wb.Sheets(sheetName)
    If checkSheet Is Nothing Then
        LogMessage "ERROR: Failed to verify creation of " & sheetName
        GenerateSheet = False
        Exit Function
    End If
    
    ' Set up headers and frozen panes
    Call CreateHeaders(wsOutput)
    
    ' Add header notes and labels
    wsOutput.Range("D2").Value = "(automatic)"
    wsOutput.Range("J2").Value = wordBench & " words/hr"
    wsOutput.Range("J3").Value = examBench & " exams/hr"
    wsOutput.Range("P2").Value = markingSupportBench & " hrs/stream"
    
    ' Benchmark labels for marker blocks
    wsOutput.Range("Y2").Value = wordBench & " words/hr"
    wsOutput.Range("Y3").Value = examBench & " exams/hr"
    wsOutput.Range("AH2").Value = wordBench & " words/hr"
    wsOutput.Range("AH3").Value = examBench & " exams/hr"
    wsOutput.Range("AR2").Value = wordBench & " words/hr"
    wsOutput.Range("AR3").Value = examBench & " exams/hr"
    
    ' Lecturer column notes
    wsOutput.Range("L2").Value = "bold = subject coordinator"
    wsOutput.Range("L3").Value = "#SPILL! = more rows need to be added"
    
    ' Hide UID column and freeze panes
    wsOutput.Columns("A:A").Hidden = True
    wsOutput.Range("D4").Select
    ActiveWindow.FreezePanes = True
    
    ' Get filtered subject data
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
    
    ' Populate subject data
    If Not PopulateSheetData(wb, wsOutput, subjectData, wordBench, examBench, markingSupportBench) Then
        LogMessage "ERROR: Failed to populate sheet data for " & groupedPeriod
        GenerateSheet = False
        Exit Function
    End If
    
    ' Apply formatting and protection
    Call FormatSheet(wsOutput)
    
    LogMessage "GenerateSheet: Completed for " & groupedPeriod, Timer - tStart
    GenerateSheet = True
    Exit Function
    
ErrorHandler:
    LogMessage "ERROR in GenerateSheet (" & groupedPeriod & "): " & Err.description & " (Error " & Err.Number & ")"
    GenerateSheet = False
End Function

'---------------------------------------------------------------
' ExportCalculationSheets
' Purpose: Exports FHY/SHY sheets to new workbook with unique filename
' Output: New workbook saved to SharePoint or temp location
' Returns: True if export successful
'---------------------------------------------------------------
Function ExportCalculationSheets(wb As Workbook) As Boolean
    On Error GoTo ErrorHandler
    
    ExportCalculationSheets = False
    Application.ScreenUpdating = False
    LogMessage "ExportCalculationSheets: Starting"
    
    ' Get year value for filename
    Dim yearValue As Variant
    yearValue = wb.Sheets("Dashboard").Range("C2").Value
    If Not IsNumeric(yearValue) Or yearValue = "" Then
        LogMessage "ERROR: Invalid year value in Dashboard C2: " & yearValue
        MsgBox "Invalid year value in Dashboard sheet (cell C2).", vbCritical
        Exit Function
    End If
    
    Dim newFileName As String
    newFileName = CStr(yearValue) & " Marking & Admin Support Calculations"
    
    ' Get calculation sheets
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
    
    ' Create new workbook
    Dim newWB As Workbook
    Set newWB = Workbooks.add
    
    ' Delete default sheets
    Application.DisplayAlerts = False
    On Error Resume Next
    Do While newWB.Sheets.Count > 1
        newWB.Sheets(newWB.Sheets.Count).Delete
    Loop
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler
    
    ' Copy calculation sheets
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
    
    ' Remove any extra sheets
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
    
    ' Set workbook properties for external links
    newWB.Application.Calculation = xlCalculationAutomatic
    
    On Error Resume Next
    With newWB
        .UpdateLinks = xlUpdateLinksAlways
        .CheckCompatibility = False
    End With
    Err.Clear
    On Error GoTo ErrorHandler
    
    ' Determine save path
    Dim savePath As String
    Dim sourceFilePath As String
    Dim basePath As String
    sourceFilePath = wb.FullName
    
    ' Handle SharePoint paths
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
    
    ' Try saving to primary location
    On Error Resume Next
    newWB.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    
    If Err.Number <> 0 Then
        Dim saveErr As Long
        saveErr = Err.Number
        LogMessage "Save failure in primary location (Code " & saveErr & ") Path: " & savePath
        
        ' Fallback to temp location
        Dim tempPath As String
        Dim tempBasePath As String
        tempBasePath = Environ("TEMP") & Application.PathSeparator
        tempPath = GetUniqueFilename(tempBasePath, newFileName, ".xlsx")
        
        Err.Clear
        newWB.SaveAs Filename:=tempPath, FileFormat:=xlOpenXMLWorkbook
        
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
    
    ' Close exported workbook
    newWB.Close SaveChanges:=False
    
    ' Delete calculation sheets from source workbook
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

'===============================================================
' SECTION 2: DATA RETRIEVAL AND FILTERING
'===============================================================

'---------------------------------------------------------------
' GetFilteredSubjectsWithAssessments
' Purpose: Gets subjects for a grouped period (FHY/SHY) that have assessments
' Filters:
'   - Matches grouped period
'   - 5th character = "9" OR subject code = "BUSA30000"
'   - Has matching assessments in assessment data parsed sheet
' Returns: Collection of subject arrays [UID, Code, Name, StudyPeriod, GroupedPeriod, OriginalRow]
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
    
    ' Get subject list data
    Dim lastRow As Long
    lastRow = wsSubjects.Cells(wsSubjects.Rows.Count, "B").End(xlUp).Row
    
    If lastRow < 2 Then
        LogMessage "WARNING: No data in SubjectList sheet"
        Set GetFilteredSubjectsWithAssessments = New Collection
        Exit Function
    End If
    
    ' Build assessment lookup dictionary (SubjectCode|StudyPeriod -> exists)
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
    
    ' Find required columns in SubjectList
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
    
    ' Filter subjects
    Dim subjects As New Collection
    Dim includedCount As Long
    
    For i = 2 To lastRow
        ' Skip if marked as excluded
        Dim excludeVal As Variant
        excludeVal = wsSubjects.Cells(i, colExclude).Value
        If excludeVal = True Or UCase(CStr(excludeVal)) = "TRUE" Then GoTo NextIteration
        
        ' Check grouped period matches
        Dim groupedPeriodVal As String
        groupedPeriodVal = Trim(CStr(wsSubjects.Cells(i, colGroupedPeriod).Value))
        If groupedPeriodVal <> groupedPeriod Then GoTo NextIteration
        
        ' Check subject code validity (5th char = "9" OR code = "BUSA30000")
        Dim subjectCode As String
        subjectCode = Trim(CStr(wsSubjects.Cells(i, colSubjectCode).Value))
        
        If Len(subjectCode) < 5 Then GoTo NextIteration
        If Mid(subjectCode, 5, 1) <> "9" And subjectCode <> "BUSA30000" Then GoTo NextIteration
        
        ' Check if subject has assessments
        Dim studyPeriod As String
        studyPeriod = Trim(CStr(wsSubjects.Cells(i, colStudyPeriod).Value))
        
        Dim lookupKey As String
        lookupKey = subjectCode & "|" & studyPeriod
        
        Dim keyExists As Boolean
        keyExists = CollectionKeyExists(assessmentDict, lookupKey)
        
        ' Fallback to "All" study period
        If Not keyExists Then
            lookupKey = subjectCode & "|All"
            keyExists = CollectionKeyExists(assessmentDict, lookupKey)
            If Not keyExists Then GoTo NextIteration
        End If
        
        ' Add subject to collection
        Dim subjectInfo(0 To 5) As Variant
        subjectInfo(0) = wsSubjects.Cells(i, colUID).Value        ' UID
        subjectInfo(1) = subjectCode                               ' Subject Code
        subjectInfo(2) = wsSubjects.Cells(i, colSubjectName).Value ' Subject Name
        subjectInfo(3) = studyPeriod                               ' Study Period
        subjectInfo(4) = groupedPeriodVal                          ' Grouped Period
        subjectInfo(5) = i                                         ' Original Row
        
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

'---------------------------------------------------------------
' GetAssessmentsForSubject
' Purpose: Retrieves all assessments for a specific subject and study period
' Uses: Static array caching for performance
' Returns: Collection of assessment arrays [UID, Description, Location, WordCount, Exam, GroupSize, ActualStudyPeriod]
'---------------------------------------------------------------
Function GetAssessmentsForSubject(wsAssessment As Worksheet, subjectCode As String, studyPeriod As String) As Collection
    On Error GoTo ErrorHandler
    
    ' Static cache for performance
    Static assessmentCache As Variant
    Static cacheInitialized As Boolean
    Static cacheSheetName As String
    
    If wsAssessment Is Nothing Then
        LogMessage "ERROR: wsAssessment is Nothing in GetAssessmentsForSubject"
        Set GetAssessmentsForSubject = Nothing
        Exit Function
    End If
    
    ' Initialize cache if needed
    If Not cacheInitialized Or cacheSheetName <> wsAssessment.name Then
        Dim lastRow As Long
        lastRow = wsAssessment.Cells(wsAssessment.Rows.Count, "A").End(xlUp).Row
        
        If lastRow < 2 Then
            LogMessage "WARNING: No assessment data found"
            Set GetAssessmentsForSubject = New Collection
            Exit Function
        End If
        
        ' Load all assessment data into array (single read operation)
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
    
    ' Search cached data for matching assessments
    Dim assessments As New Collection
    Dim foundExact As Boolean
    foundExact = False
    
    Dim assessItem(0 To 6) As Variant
    
    Dim i As Long
    ' Try exact match first (SubjectCode + StudyPeriod)
    For i = 1 To UBound(assessmentCache, 1)
        If assessmentCache(i, 2) = subjectCode And assessmentCache(i, 3) = studyPeriod Then
            assessItem(0) = assessmentCache(i, 1)  ' UID
            assessItem(1) = assessmentCache(i, 4)  ' Description
            assessItem(2) = assessmentCache(i, 5)  ' Location
            assessItem(3) = assessmentCache(i, 6)  ' Word Count
            assessItem(4) = assessmentCache(i, 7)  ' Exam
            assessItem(5) = assessmentCache(i, 8)  ' Group Size
            assessItem(6) = studyPeriod            ' Actual Study Period
            
            assessments.add assessItem
            foundExact = True
        End If
    Next i
    
    ' Fallback to "All" study period if no exact matches
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
' GetLecturerCountForSubject
' Purpose: Counts unique lecturers for a subject (used for row allocation)
' Note: Uses local teaching stream sheet (not external file)
'---------------------------------------------------------------
Function GetLecturerCountForSubject(wb As Workbook, subjectCode As String, studyPeriod As String) As Integer
    On Error GoTo ErrorHandler
    
    Dim wsTeaching As Worksheet
    Set wsTeaching = wb.Sheets("teaching stream")
    
    If wsTeaching Is Nothing Then
        LogMessage "WARNING: teaching stream sheet not found for lecturer count"
        GetLecturerCountForSubject = 0
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = wsTeaching.Cells(wsTeaching.Rows.Count, "B").End(xlUp).Row
    
    If lastRow < 2 Then
        GetLecturerCountForSubject = 0
        Exit Function
    End If
    
    ' Build unique lecturer dictionary
    Dim uniqueLecturers As Collection
    Set uniqueLecturers = New Collection
    
    Dim i As Long
    For i = 2 To lastRow
        ' Match subject code and study period
        If wsTeaching.Cells(i, 2).Value = subjectCode And _
           wsTeaching.Cells(i, 3).Value = studyPeriod Then
            
            Dim lecturerName As String
            lecturerName = wsTeaching.Cells(i, 4).Value
            
            ' Add to dictionary if not already present
            On Error Resume Next
            uniqueLecturers.add True, lecturerName
            Err.Clear
            On Error GoTo ErrorHandler
        End If
    Next i
    
    GetLecturerCountForSubject = uniqueLecturers.Count
    Exit Function
    
ErrorHandler:
    LogMessage "ERROR in GetLecturerCountForSubject: " & Err.description
    GetLecturerCountForSubject = 0
End Function

'===============================================================
' SECTION 3: SHEET POPULATION
'===============================================================

'---------------------------------------------------------------
' PopulateSheetData
' Purpose: Populates worksheet with all subject blocks organized by study period
' Organization:
'   - FHY: SUMMER header, then SEMESTER 1 header
'   - SHY: WINTER header, then SEMESTER 2 header
'---------------------------------------------------------------
Function PopulateSheetData(wb As Workbook, wsOutput As Worksheet, subjectData As Collection, wordBench As Double, examBench As Double, markingSupportBench As Double) As Boolean
    On Error GoTo ErrorHandler
    
    ' BATCH OPTIMIZATION: Collect all formula/validation requests in memory
    Dim formulaQueue As New Collection
    Dim validationQueue As New Collection
    
    PopulateSheetData = False
    Dim tStart As Double
    tStart = Timer
    
    If subjectData Is Nothing Then
        LogMessage "ERROR: subjectData is Nothing in PopulateSheetData"
        Exit Function
    End If
    
    If subjectData.Count = 0 Then
        LogMessage "WARNING: No subjects to populate"
        PopulateSheetData = True
        Exit Function
    End If
    
    Dim wsAssessment As Worksheet
    Set wsAssessment = wb.Sheets("assessment data parsed")
    
    If wsAssessment Is Nothing Then
        LogMessage "ERROR: Assessment sheet not found in PopulateSheetData"
        Exit Function
    End If
    
    ' Get grouped period from first subject
    Dim groupedPeriod As String
    groupedPeriod = SafeArrayIndex(subjectData(1), 4, "")
    
    If groupedPeriod = "" Then
        LogMessage "ERROR: Could not determine grouped period"
        Exit Function
    End If
    
    ' Start populating from row 4
    Dim currentRow As Long
    currentRow = 4
    
    ' Add first study period header
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
    
    ' Categorize subjects by study period
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
            Case "SUMMER"
                summerSubjects.add subject
            Case "SEMESTER 1"
                semester1Subjects.add subject
            Case "WINTER"
                winterSubjects.add subject
            Case "SEMESTER 2"
                semester2Subjects.add subject
        End Select
    Next subject
    
    ' Populate subjects in order
    If groupedPeriod = "FHY" Then
        ' Process SUMMER subjects
        If summerSubjects.Count > 0 Then
            For Each subject In summerSubjects
                If Not ProcessSubjectBatch(wb, wsOutput, subject, currentRow, wordBench, examBench, markingSupportBench, formulaQueue, validationQueue) Then
                    LogMessage "ERROR: Failed to process subject in SUMMER category"
                End If
            Next subject
        End If
        
        ' Add SEMESTER 1 header and subjects
        If semester1Subjects.Count > 0 Then
            wsOutput.Rows(currentRow).Interior.Color = RGB(0, 0, 0)
            wsOutput.Cells(currentRow, 2).Value = "SEMESTER 1"
            wsOutput.Cells(currentRow, 2).Font.Color = RGB(255, 255, 255)
            wsOutput.Cells(currentRow, 2).Font.Bold = True
            currentRow = currentRow + 1
            
            For Each subject In semester1Subjects
                If Not ProcessSubjectBatch(wb, wsOutput, subject, currentRow, wordBench, examBench, markingSupportBench, formulaQueue, validationQueue) Then
                    LogMessage "ERROR: Failed to process subject in SEMESTER 1 category"
                End If
            Next subject
        End If
    Else
        ' Process WINTER subjects
        If winterSubjects.Count > 0 Then
            For Each subject In winterSubjects
                If Not ProcessSubjectBatch(wb, wsOutput, subject, currentRow, wordBench, examBench, markingSupportBench, formulaQueue, validationQueue) Then
                    LogMessage "ERROR: Failed to process subject in WINTER category"
                End If
            Next subject
        End If
        
        ' Add SEMESTER 2 header and subjects
        If semester2Subjects.Count > 0 Then
            wsOutput.Rows(currentRow).Interior.Color = RGB(0, 0, 0)
            wsOutput.Cells(currentRow, 2).Value = "SEMESTER 2"
            wsOutput.Cells(currentRow, 2).Font.Color = RGB(255, 255, 255)
            wsOutput.Cells(currentRow, 2).Font.Bold = True
            currentRow = currentRow + 1
            
            For Each subject In semester2Subjects
                If Not ProcessSubjectBatch(wb, wsOutput, subject, currentRow, wordBench, examBench, markingSupportBench, formulaQueue, validationQueue) Then
                    LogMessage "ERROR: Failed to process subject in SEMESTER 2 category"
                End If
            Next subject
        End If
    End If
    
    ' BATCH WRITE: Apply all queued formulas and validations
    LogMessage "Applying " & formulaQueue.Count & " formulas and " & validationQueue.Count & " validations..."
    
    Dim formulaItem As Variant
    For Each formulaItem In formulaQueue
        On Error Resume Next
        wsOutput.Range(formulaItem(0)).Formula = formulaItem(1)
        If Err.Number <> 0 Then
            LogMessage "WARNING: Formula write failed for " & formulaItem(0)
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    Next formulaItem
    
    Dim validationItem As Variant
    For Each validationItem In validationQueue
        On Error Resume Next
        With wsOutput.Range(validationItem(0)).Validation
            .Delete
            .add Type:=xlValidateList, Formula1:="N,Y"
            .InCellDropdown = True
        End With
        If Err.Number = 0 Then
            wsOutput.Range(validationItem(0)).Value = "N"
        Else
            LogMessage "WARNING: Validation failed for " & validationItem(0)
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    Next validationItem
    
    LogMessage "PopulateSheetData completed", Timer - tStart
    PopulateSheetData = True
    Exit Function
    
ErrorHandler:
    LogMessage "ERROR in PopulateSheetData: " & Err.description & " (Error " & Err.Number & ")"
    PopulateSheetData = False
End Function

'---------------------------------------------------------------
' ProcessSubjectBatch
' Purpose: Creates complete subject block with assessments, lecturers, and marker sections
' Row Structure:
'   - Row 0 (header): Subject code, study period, dynamic enrolment formula
'   - Row 1+: Assessment details (one row per assessment)
'   - Last row: "Total" row with sum formulas
'   - Extra rows added if lecturer count > assessment count
' Side effects: Updates currentRow to next available row
'---------------------------------------------------------------
Function ProcessSubjectBatch(wb As Workbook, wsOutput As Worksheet, ByRef subject As Variant, ByRef currentRow As Long, wordBench As Double, examBench As Double, markingSupportBench As Double, formulaQueue As Collection, validationQueue As Collection) As Boolean
    On Error GoTo ErrorHandler
    
    ProcessSubjectBatch = False
    
    Dim wsAssessment As Worksheet
    Set wsAssessment = wb.Sheets("assessment data parsed")
    
    ' Extract subject data using safe array accessor
    Dim subjectCode As String
    Dim studyPeriod As String
    subjectCode = SafeArrayIndex(subject, 1, "")
    studyPeriod = SafeArrayIndex(subject, 3, "")
    
    If subjectCode = "" Or studyPeriod = "" Then
        LogMessage "ERROR: Invalid subject data in ProcessSubject"
        Exit Function
    End If
    
    Dim subjectStartRow As Long
    Dim subjectEndRow As Long
    subjectStartRow = currentRow
    
    ' Get assessments and lecturers
    Dim assessments As Collection
    Set assessments = GetAssessmentsForSubject(wsAssessment, subjectCode, studyPeriod)
    
    If assessments Is Nothing Then
        LogMessage "ERROR: Failed to get assessments for " & subjectCode
        Exit Function
    End If
    
    If assessments.Count = 0 Then
        ProcessSubjectBatch = True
        Exit Function
    End If
    
    Dim lecturerCount As Integer
    lecturerCount = GetLecturerCountForSubject(wb, subjectCode, studyPeriod)
    
    ' Calculate rows needed
    Dim uidCount As Integer
    uidCount = assessments.Count
    
    Dim assessmentRows As Integer
    assessmentRows = uidCount + 2  ' Header + assessments + Total row
    
    Dim totalRowsNeeded As Integer
    totalRowsNeeded = Application.WorksheetFunction.Max(assessmentRows, lecturerCount + 1)
    
    ' Build assessment data array
    Dim outputData() As Variant
    ReDim outputData(1 To assessmentRows, 1 To 10)
    
    ' Header row (row_0)
    Dim headerUID As String
    headerUID = subjectCode & "_" & studyPeriod & "_0"
    outputData(1, 1) = headerUID
    outputData(1, 2) = subjectCode
    outputData(1, 3) = studyPeriod
    outputData(1, 4) = 0  ' Will be replaced with formula
    
    ' Assessment rows
    Dim i As Integer
    For i = 1 To assessments.Count
        Dim assessment As Variant
        assessment = assessments(i)
        
        Dim assessmentUID As String
        assessmentUID = subjectCode & "_" & studyPeriod & "_" & i
        
        outputData(i + 1, 1) = assessmentUID
        outputData(i + 1, 5) = SafeArrayIndex(assessment, 1, "")  ' Description
        outputData(i + 1, 6) = SafeArrayIndex(assessment, 3, "")  ' Word Count
        outputData(i + 1, 7) = SafeArrayIndex(assessment, 4, "")  ' Exam
        outputData(i + 1, 8) = SafeArrayIndex(assessment, 5, "")  ' Group Size
        
        ' Check if Location value exists for Assessment Quantity (Col I)
        Dim locationValue As Variant
        locationValue = SafeArrayIndex(assessment, 2, "")  ' Index 2 = Location
        If locationValue <> "" And Not IsEmpty(locationValue) Then
            outputData(i + 1, 9) = locationValue  ' Use Location value
        End If
        ' Note: If empty, Col I will be filled by formula later
    Next i
    
    ' Total row
    Dim totalUID As String
    totalUID = subjectCode & "_" & studyPeriod & "_" & (uidCount + 1)
    outputData(assessmentRows, 1) = totalUID
    outputData(assessmentRows, 5) = "Total"
    
    ' Write assessment data to sheet
    wsOutput.Cells(currentRow, 1).Resize(assessmentRows, 10).Value = outputData
    
    ' Format header and total rows
    wsOutput.Rows(currentRow).Interior.Color = RGB(192, 192, 192)
    wsOutput.Cells(currentRow + assessmentRows - 1, 5).Font.Bold = True
    
    '' Queue dynamic enrolment formula
    Dim enrolmentFileName As String
    enrolmentFileName = wb.Sheets("Dashboard").Range("C3").Value
    
    ' VALIDATION: Ensure filename ends with .xlsm
    If enrolmentFileName <> "" Then
        enrolmentFileName = Trim(enrolmentFileName)
        If Right(LCase(enrolmentFileName), 5) <> ".xlsm" Then
            enrolmentFileName = enrolmentFileName & ".xlsm"
            LogMessage "  Auto-appended .xlsm to enrolment filename: " & enrolmentFileName
        End If
        
        Dim enrolPath As String
        enrolPath = "'" & ENROLMENT_TRACKER_BASE & "[" & enrolmentFileName & "]Enrolment Number Tracker'!"
        Dim enrolFormula As String
        enrolFormula = "=IFERROR(INDEX(" & enrolPath & "$I:$I,SUMPRODUCT((" & enrolPath & "$A:$A=B" & currentRow & ")*(" & enrolPath & "$C:$C=C" & currentRow & ")*ROW(" & enrolPath & "$A:$A))),0)"
        formulaQueue.add Array("D" & currentRow, enrolFormula)
    End If
    
   ' Queue assessment formulas
    Dim formulaRow As Long
    For i = 1 To assessments.Count
        formulaRow = currentRow + i
        
        ' Assessment Quantity (Col I): Only queue formula if Location was empty
        If outputData(i + 1, 9) = Empty Or outputData(i + 1, 9) = "" Then
            Dim qtyFormula As String
            qtyFormula = "=IF(H" & formulaRow & "<>"""", $D$" & currentRow & "/H" & formulaRow & ", $D$" & currentRow & ")"
            formulaQueue.add Array("I" & formulaRow, qtyFormula)
        End If
        
        ' Marking Hours (Col J): Always use formula
        Dim markFormula As String
        markFormula = "=IF(ISNUMBER(I" & formulaRow & "),IF(ISNUMBER(F" & formulaRow & "),I" & formulaRow & "*(F" & formulaRow & "/VALUE(LEFT($J$2,FIND("" "",$J$2)-1))),IF(ISNUMBER(G" & formulaRow & "),I" & formulaRow & "/(VALUE(LEFT($J$3,FIND("" "",$J$3)-1))),"""")),"""")"
        formulaQueue.add Array("J" & formulaRow, markFormula)
    Next i
    
    ' Queue total formula
    If uidCount > 0 Then
        Dim totalRow As Long
        totalRow = currentRow + assessmentRows - 1
        formulaQueue.add Array("J" & totalRow, "=SUM(J" & (currentRow + 1) & ":J" & (currentRow + uidCount) & ")")
    End If
    
    subjectEndRow = currentRow + assessmentRows - 1
    
    ' Add extra rows if more lecturers than assessments
    Dim extraRowsNeeded As Integer
    extraRowsNeeded = totalRowsNeeded - assessmentRows
    
    If extraRowsNeeded > 0 Then
        wsOutput.Rows(subjectEndRow + 1).Resize(extraRowsNeeded).Insert Shift:=xlDown
        
        ' Add UIDs to extra rows
        Dim j As Integer
        For j = 1 To extraRowsNeeded
            Dim extraUID As String
            extraUID = subjectCode & "_" & studyPeriod & "_" & (uidCount + 1 + j)
            wsOutput.Cells(subjectEndRow + j, 1).Value = extraUID
        Next j
        
        subjectEndRow = subjectEndRow + extraRowsNeeded
        
        LogMessage "  Expanded " & subjectCode & " block by " & extraRowsNeeded & " rows for lecturers"
    End If
    
    ' Queue lecturer formulas (and bold first lecturer = subject coordinator)
    Call QueueLecturerFormulas(currentRow + 1, currentRow, subjectEndRow, subjectCode, studyPeriod, formulaQueue, wsOutput)
    
    ' Queue marker block formulas
    Call QueueMarkerFormulas(currentRow, subjectEndRow, 1, formulaQueue)
    Call QueueMarkerFormulas(currentRow, subjectEndRow, 2, formulaQueue)
    Call QueueMarkerFormulas(currentRow, subjectEndRow, 3, formulaQueue)
    
    ' Queue contract requested validations
    validationQueue.add Array("AB" & (currentRow + 1))  ' Marker 1
    validationQueue.add Array("AL" & (currentRow + 1))  ' Marker 2
    validationQueue.add Array("AV" & (currentRow + 1))  ' Marker 3
    
    ' Move to next subject
    currentRow = subjectEndRow + 1
    
    ProcessSubjectBatch = True
    Exit Function
    
ErrorHandler:
    LogMessage "ERROR in ProcessSubject (" & subjectCode & "): " & Err.description & " (Error " & Err.Number & ")"
    ProcessSubjectBatch = False
End Function

Sub QueueLecturerFormulas(lecturerStartRow As Long, subjectStartRow As Long, subjectEndRow As Long, subjectCode As String, studyPeriod As String, formulaQueue As Collection, wsOutput As Worksheet)
    Dim teachingPath As String
    teachingPath = "'" & TEACHING_STREAM_BASE & "[" & TEACHING_STREAM_FILE & "]teaching stream'!"
    
    ' Queue FILTER formula for first lecturer row (with IFERROR to handle no results)
    Dim lecFormula As String
    lecFormula = "=IFERROR(FILTER(" & teachingPath & "$D:$D,(" & teachingPath & "$B:$B=$B$" & subjectStartRow & ")*(" & teachingPath & "$C:$C=$C$" & subjectStartRow & ")),"""")"
    formulaQueue.add Array("L" & lecturerStartRow, lecFormula)
    
    ' Bold the first lecturer cell (subject coordinator)
    wsOutput.Cells(lecturerStartRow, 12).Font.Bold = True
    
    ' Queue formulas for all lecturer rows
    Dim r As Long
    For r = lecturerStartRow To subjectEndRow - 1
        ' Status
        formulaQueue.add Array("M" & r, "=IF(L" & r & "="""","""",IFERROR(INDEX(" & teachingPath & "$E:$E,SUMPRODUCT((" & teachingPath & "$B:$B=$B$" & subjectStartRow & ")*(" & teachingPath & "$C:$C=$C$" & subjectStartRow & ")*(" & teachingPath & "$D:$D=L" & r & ")*ROW(" & teachingPath & "$B:$B))),""""))")
        
        ' Streams
        formulaQueue.add Array("N" & r, "=IF(L" & r & "="""","""",IFERROR(INDEX(" & teachingPath & "$G:$G,SUMPRODUCT((" & teachingPath & "$B:$B=$B$" & subjectStartRow & ")*(" & teachingPath & "$C:$C=$C$" & subjectStartRow & ")*(" & teachingPath & "$D:$D=L" & r & ")*ROW(" & teachingPath & "$B:$B))),""""))")
        
        ' Activity Code
        formulaQueue.add Array("O" & r, "=IF(L" & r & "="""","""",IFERROR(INDEX(" & teachingPath & "$F:$F,SUMPRODUCT((" & teachingPath & "$B:$B=$B$" & subjectStartRow & ")*(" & teachingPath & "$C:$C=$C$" & subjectStartRow & ")*(" & teachingPath & "$D:$D=L" & r & ")*ROW(" & teachingPath & "$B:$B))),""""))")
        
        ' Allocated Marking
        formulaQueue.add Array("P" & r, "=IF(M" & r & "=""Continuing T&R"",N" & r & "*VALUE(LEFT($P$2,FIND("" "",$P$2)-1)),"""")")
        
        ' Marking Support Hours
        formulaQueue.add Array("Q" & r, "=IFERROR(INDEX($J:$J,MATCH(""Total"",$E:$E,0))-P" & r & ","""")")
    Next r
End Sub

Sub QueueMarkerFormulas(subjectStartRow As Long, subjectEndRow As Long, markerNum As Integer, formulaQueue As Collection)
    Dim baseCol As Integer, benchmarkCol As Integer
    
    Select Case markerNum
        Case 1: baseCol = 19: benchmarkCol = 25
        Case 2: baseCol = 29: benchmarkCol = 35
        Case 3: baseCol = 39: benchmarkCol = 44
    End Select
    
    Dim detailsCol As String, wordCol As String, examCol As String, groupCol As String, qtyCol As String, allocCol As String
    detailsCol = ColLetter(baseCol + 1)
    wordCol = ColLetter(baseCol + 2)
    examCol = ColLetter(baseCol + 3)
    groupCol = ColLetter(baseCol + 4)
    qtyCol = ColLetter(baseCol + 5)
    allocCol = ColLetter(baseCol + 6)
    
    Dim r As Long
    For r = subjectStartRow To subjectEndRow
        If r = subjectEndRow Then
            ' Total row
            formulaQueue.add Array(allocCol & r, "=SUM(" & allocCol & (subjectStartRow + 1) & ":" & allocCol & (subjectEndRow - 1) & ")")
        Else
            ' Regular rows
            formulaQueue.add Array(wordCol & r, "=IFERROR(IF(INDEX($F$" & (subjectStartRow + 1) & ":$F$" & (subjectEndRow - 1) & ",MATCH(" & detailsCol & r & ",$E$" & (subjectStartRow + 1) & ":$E$" & (subjectEndRow - 1) & ",0))="""","""",INDEX($F$" & (subjectStartRow + 1) & ":$F$" & (subjectEndRow - 1) & ",MATCH(" & detailsCol & r & ",$E$" & (subjectStartRow + 1) & ":$E$" & (subjectEndRow - 1) & ",0))),"""")")
            
            formulaQueue.add Array(examCol & r, "=IFERROR(IF(INDEX($G$" & (subjectStartRow + 1) & ":$G$" & (subjectEndRow - 1) & ",MATCH(" & detailsCol & r & ",$E$" & (subjectStartRow + 1) & ":$E$" & (subjectEndRow - 1) & ",0))="""","""",INDEX($G$" & (subjectStartRow + 1) & ":$G$" & (subjectEndRow - 1) & ",MATCH(" & detailsCol & r & ",$E$" & (subjectStartRow + 1) & ":$E$" & (subjectEndRow - 1) & ",0))),"""")")
            
            formulaQueue.add Array(groupCol & r, "=IFERROR(IF(INDEX($H$" & (subjectStartRow + 1) & ":$H$" & (subjectEndRow - 1) & ",MATCH(" & detailsCol & r & ",$E$" & (subjectStartRow + 1) & ":$E$" & (subjectEndRow - 1) & ",0))="""","""",INDEX($H$" & (subjectStartRow + 1) & ":$H$" & (subjectEndRow - 1) & ",MATCH(" & detailsCol & r & ",$E$" & (subjectStartRow + 1) & ":$E$" & (subjectEndRow - 1) & ",0))),"""")")
            
            formulaQueue.add Array(qtyCol & r, "=IF(" & detailsCol & r & "="""","""",0)")
            
            formulaQueue.add Array(allocCol & r, "=IF(" & qtyCol & r & "="""","""",IF(ISNUMBER(" & qtyCol & r & "),IF(ISNUMBER(" & wordCol & r & ")," & qtyCol & r & "*(" & wordCol & r & "/VALUE(LEFT($" & ColLetter(benchmarkCol) & "$2,FIND("" "",$" & ColLetter(benchmarkCol) & "$2)-1))),IF(ISNUMBER(" & examCol & r & ")," & qtyCol & r & "/(VALUE(LEFT($" & ColLetter(benchmarkCol) & "$3,FIND("" "",$" & ColLetter(benchmarkCol) & "$3)-1))),""""))," & qtyCol & r & "))")
        End If
    Next r
End Sub

'===============================================================
' SECTION 4: FORMATTING AND DISPLAY
'===============================================================

'---------------------------------------------------------------
' CreateHeaders
' Purpose: Creates column headers for calculation sheet
' Columns: UID, Subject Code, Study Period, Enrolment, Assessment Details, etc.
'---------------------------------------------------------------
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

'---------------------------------------------------------------
' FormatSheet
' Purpose: Applies column widths, formatting, conditional formatting, and protection
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
    
    ' === COLUMN WIDTHS ===
    Dim widthSettings As Variant
    widthSettings = Array( _
        Array("A:A", 25), Array("B:C", 15), Array("D:D", 11.5), Array("E:E", 60), _
        Array("F:H", 7), Array("I:J", 10.5), Array("K:K", 30), Array("L:L", 35), Array("M:M", 11.75), _
        Array("N:N", 6.5), Array("O:O", 25), Array("P:Q", 14), Array("R:R", 30), _
        Array("S:S", 30), Array("T:T", 60), Array("U:W", 7), Array("X:Y", 13), _
        Array("Z:AA", 30), Array("AB:AB", 10), Array("AC:AC", 30), Array("AD:AD", 60), _
        Array("AE:AG", 7), Array("AH:AI", 13), Array("AJ:AK", 30), Array("AL:AL", 10), _
        Array("AM:AM", 30), Array("AN:AN", 60), Array("AO:AQ", 7), Array("AR:AS", 13), _
        Array("AT:AU", 30), Array("AV:AV", 10))
    
    Dim i As Integer
    For i = 0 To UBound(widthSettings)
        ws.Columns(widthSettings(i)(0)).ColumnWidth = widthSettings(i)(1)
    Next i
    
    ' === WRAP TEXT ===
    Dim wrapTextCols As Variant
    wrapTextCols = Array("E:E", "K:L", "R:T", "Z:AD", "AJ:AN", "AT:AV")
    
    For i = 0 To UBound(wrapTextCols)
        ws.Columns(wrapTextCols(i)).WrapText = True
    Next i
    
    ' === FONT BOLD ===
    Dim boldCols As Variant
    boldCols = Array("B:B")
    
    For i = 0 To UBound(boldCols)
        ws.Columns(boldCols(i)).Font.Bold = True
    Next i
    
    ' === HIDDEN COLUMNS ===
    Dim hiddenCols As Variant
    hiddenCols = Array("A:A")
    
    For i = 0 To UBound(hiddenCols)
        ws.Columns(hiddenCols(i)).Hidden = True
    Next i
    
    ' === NUMBER FORMATS ===
    Dim integerCols As Variant
    integerCols = Array("D:D", "F:I", "N:N", "U:X", "AE:AH", "AO:AR")
    
    For i = 0 To UBound(integerCols)
        ws.Columns(integerCols(i)).NumberFormat = "0"
    Next i
    
    Dim decimalCols As Variant
    decimalCols = Array("J:J", "P:Q", "Y:Y", "AI:AI", "AS:AS")
    
    For i = 0 To UBound(decimalCols)
        ws.Columns(decimalCols(i)).NumberFormat = "0.00"
    Next i
    
    ' === HORIZONTAL ALIGNMENT ===
    Dim centerCols As Variant
    centerCols = Array("D:D", "F:J", "N:N", "P:Q", "U:Y", "AB:AB", "AE:AI", "AL:AL", "AO:AS", "AV:AV")
    
    For i = 0 To UBound(centerCols)
        ws.Columns(centerCols(i)).HorizontalAlignment = xlCenter
    Next i
    
    ' === CONDITIONAL FORMATTING ===
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
    
    ' === LOCK CELLS ===
    ws.Cells.Locked = False
    
    Dim lockedCols As Variant
    lockedCols = Array("B:C", "E:H", "J:J", "P:Q")
    
    For i = 0 To UBound(lockedCols)
        ws.Columns(lockedCols(i)).Locked = True
    Next i
    
    ' === PROTECT SHEET ===
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

'===============================================================
' SECTION 5: HELPER FUNCTIONS
'===============================================================

'---------------------------------------------------------------
' SafeArrayIndex
' Purpose: Safely retrieves value from array with default fallback
' Replaces individual GetSubject*, GetAssessment*, GetLecturer* functions
'---------------------------------------------------------------
Function SafeArrayIndex(arr As Variant, idx As Integer, defaultValue As Variant) As Variant
    On Error Resume Next
    SafeArrayIndex = arr(idx)
    If Err.Number <> 0 Then SafeArrayIndex = defaultValue
    On Error GoTo 0
End Function

'---------------------------------------------------------------
' FindColumn
' Purpose: Finds column number by header name
' Returns: Column number (1-based) or 0 if not found
'---------------------------------------------------------------
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

'---------------------------------------------------------------
' FindEnrolmentCell
' Purpose: Finds the cell reference for enrolment number (row_0 of subject)
' Returns: Absolute cell reference like "$D$5" or empty string
'---------------------------------------------------------------
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

'---------------------------------------------------------------
' CategorizeStudyPeriod
' Purpose: Maps study period strings to category headers
' Returns: "SUMMER", "SEMESTER 1", "WINTER", or "SEMESTER 2"
'---------------------------------------------------------------
Function CategorizeStudyPeriod(studyPeriod As String, groupedPeriod As String) As String
    Dim sp As String
    sp = LCase(studyPeriod)
    
    ' Check for specific keywords
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
    
    ' Default categorization based on grouped period
    If groupedPeriod = "FHY" Then
        CategorizeStudyPeriod = "SEMESTER 1"
    Else
        CategorizeStudyPeriod = "SEMESTER 2"
    End If
End Function

'---------------------------------------------------------------
' ColLetter
' Purpose: Converts column number to letter (e.g., 1 -> "A", 27 -> "AA")
'---------------------------------------------------------------
Function ColLetter(colNum As Integer) As String
    ColLetter = Split(Cells(1, colNum).Address, "$")(1)
End Function

'---------------------------------------------------------------
' GetUniqueFilename
' Purpose: Generates unique filename with timestamp if file exists
' Returns: Full path with unique filename
'---------------------------------------------------------------
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

'---------------------------------------------------------------
' GetBenchmarkValue
' Purpose: Retrieves benchmark value from Dashboard with fallback to default
'---------------------------------------------------------------
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

'---------------------------------------------------------------
' CollectionKeyExists
' Purpose: Checks if key exists in Collection
'---------------------------------------------------------------
Function CollectionKeyExists(col As Collection, key As String) As Boolean
    On Error Resume Next
    Dim temp As Variant
    temp = col(key)
    CollectionKeyExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

'===============================================================
' SECTION 6: VERIFICATION AND CLEANUP
'===============================================================

'---------------------------------------------------------------
' VerifyRequiredSheets
' Purpose: Validates that all required source sheets exist
' Returns: False if any sheets are missing
'---------------------------------------------------------------
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

'---------------------------------------------------------------
' HasErrorsInLog
' Purpose: Checks if Process Log contains any error messages
'---------------------------------------------------------------
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

'---------------------------------------------------------------
' CleanupPartialSheets
' Purpose: Removes partially generated calculation sheets after error
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

'===============================================================
' SECTION 7: LOGGING
'===============================================================

'---------------------------------------------------------------
' InitializeProcessLog
' Purpose: Creates or recreates the Process Log sheet
'---------------------------------------------------------------
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

'---------------------------------------------------------------
' LogMessage
' Purpose: Writes message to Process Log and status bar
'---------------------------------------------------------------
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
