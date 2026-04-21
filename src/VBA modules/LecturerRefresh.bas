'===============================================================
' Module: LecturerRefresh
' Purpose: Refresh lecturer data in exported calculation sheets
'          by triggering the Teaching Matrix workflow,
'          then updating columns L-O with fresh data while
'          preserving user edits in columns P and S
' Main Entry: RefreshLecturerData() - triggered by Refresh button
' Output: Updated lecturer names, statuses, streams, activity
'         codes in FHY/SHY sheets; user notes preserved
' Author: Arthur Chen
' Repository: github.com/arthurtheprogrammer/auto-handbook-system
' Dependencies:
'   - Source file: Automated Handbook Data System.xlsm (SharePoint)
'   - teaching stream sheet in source file
'   - Teaching Matrix Power Automate endpoint
'===============================================================

Option Explicit

'===============================================================
' SECTION 1: CONFIGURATION
'===============================================================
' Source file path (Automated Handbook Data System)
Private Const SOURCE_FILE_PATH As String = "https://unimelbcloud.sharepoint.com/teams/DepartmentofManagementMarketing-DepartmentOperations/Shared Documents/TEACHING SUPPORT/Handbook (Course & Subject Changes)/Auto Handbook System/Automated Handbook Data System.xlsm"

' Source sheets
Private Const TEACHING_STREAM_SHEET As String = "teaching stream"

' Teaching Matrix Power Automate endpoint
Private Const TEACHING_MATRIX_URL As String = "https://default0e5bf3cf1ff446b7917652c538c22a.4d.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/7f198e614c734715bc0153d818de1ef7/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=5uUuhFHyiL37O_ajy-6t2r65nFqc7NA_oJDhYmYFT9g"

' SubjectBlockInfo array indices (0-6)
' Used to reference fields in the block info arrays returned by
' IdentifySubjectBlocks
Private Const SBI_SHEETNAME As Integer = 0
Private Const SBI_SUBJECTCODE As Integer = 1
Private Const SBI_STUDYPERIOD As Integer = 2
Private Const SBI_HEADERROW As Integer = 3
Private Const SBI_TOTALROW As Integer = 4
Private Const SBI_LASTSUBJECTROW As Integer = 5
Private Const SBI_NUMASSESSMENTROWS As Integer = 6

'===============================================================
' SECTION 2: MAIN WORKFLOW
'===============================================================

'---------------------------------------------------------------
' RefreshLecturerData
' Purpose: Main entry point — read params, trigger workflow,
'          wait for completion, identify subject blocks, load
'          fresh data, and update lecturer columns
' Called by: Refresh button on calculation sheets
'---------------------------------------------------------------
Public Sub RefreshLecturerData()
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    ' Validate that we are running from the exported calculation workbook,
    ' not accidentally from the source file. FHY or SHY sheet must exist.
    Dim hasFHY As Boolean, hasSHY As Boolean
    On Error Resume Next
    hasFHY = Not (wb.Sheets("FHY Calculations") Is Nothing)
    hasSHY = Not (wb.Sheets("SHY Calculations") Is Nothing)
    On Error GoTo 0
    
    If Not hasFHY And Not hasSHY Then
        MsgBox "This workbook does not contain FHY or SHY Calculations sheets." & vbCrLf & vbCrLf & _
               "Please run 'Refresh Lecturer Data' from the exported calculation workbook, " & _
               "not from the source (Automated Handbook Data System) file.", _
               vbExclamation, "Wrong Workbook"
        Exit Sub
    End If
    
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation
    Dim origEnableEvents As Boolean
    
    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation
    origEnableEvents = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "Preparing to refresh lecturer data..."
    
    Dim updateCount As Long
    updateCount = 0
    
    ' STEP 1: Get parameters from source file
    ' (sourceWb refers to "Auto handbook system.xlsm" on SharePoint)
    Application.StatusBar = "Reading parameters from source file..."
    
    Dim yearValue As String
    Dim teachingMatrix As String
    Dim emailValue As String
    
    If Not GetSourceParameters(yearValue, teachingMatrix, emailValue) Then
        MsgBox "Could not read parameters from source file." & vbCrLf & vbCrLf & _
               "Source: " & SOURCE_FILE_PATH, vbExclamation, "Connection Error"
        GoTo CleanExit
    End If
    
    ' STEP 1.5: Update progress bar in source file to "Running..." (Orange)
    ' We can safely do this now because the 45-second fixed wait bypasses
    ' Mac caching lock issues that previously ruined instant-polling.
    SetTeachingStreamRunningStatus
    
    ' STEP 2: Trigger Teaching Matrix workflow
    Application.StatusBar = "Triggering Teaching Matrix workflow..."
    
    Application.ScreenUpdating = True
    DoEvents
    
    MsgBox "Excel will now connect to the Teaching Matrix to fetch latest lecturer data." & vbCrLf & vbCrLf & _
           "Please note: Excel will appear frozen/unresponsive for 30–60 seconds while waiting for the workflow to complete." & vbCrLf & vbCrLf & _
           "Press OK to begin.", vbInformation, "Starting Refresh"
           
    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing lecturer assignments... (Excel will freeze momentarily)"
    
    If Not TriggerTeachingMatrixWorkflow(yearValue, teachingMatrix, emailValue) Then
        MsgBox "Failed to refresh Teaching Matrix data." & vbCrLf & vbCrLf & _
               "The connection timed out or the server reported an error.", vbExclamation, "Workflow Error"
        GoTo CleanExit
    End If
    
    ' Give SharePoint 5 seconds to commit the file fully before we try to open it
    Application.StatusBar = "Workflow complete! Opening server file..."
    Application.Wait (Now + TimeValue("0:00:05"))
    
    ' STEP 4: Identify all subject blocks
    Application.StatusBar = "Identifying subject blocks..."
    
    Dim subjectBlocks As Collection
    Set subjectBlocks = New Collection
    
    Call IdentifySubjectBlocks(wb, subjectBlocks)
    
    If subjectBlocks.Count = 0 Then
        MsgBox "No subject blocks found in calculation sheets.", vbExclamation, "No Subjects"
        GoTo CleanExit
    End If
    
    ' STEP 5: Load fresh teaching stream data
    Application.StatusBar = "Loading fresh teaching stream data..."
    
    Dim teachingData As Variant
    teachingData = LoadTeachingStreamData(SOURCE_FILE_PATH)
    
    If IsEmpty(teachingData) Then
        MsgBox "No teaching stream data found in source file." & vbCrLf & vbCrLf & _
               "The 'teaching stream' sheet may be empty or missing.", vbExclamation, "No Data"
        GoTo CleanExit
    End If
    
    ' STEP 6: Update lecturer data (columns L-O only, preserve P-S)
    Application.StatusBar = "Updating lecturer data..."
    
    updateCount = UpdateAllLecturers(wb, teachingData, subjectBlocks)
    
CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = origScreenUpdating
    Application.EnableEvents = origEnableEvents
    DoEvents
    Application.StatusBar = False
    
    If updateCount > 0 Then
        MsgBox "Lecturer data refreshed successfully!" & vbCrLf & vbCrLf & _
               "Updated " & updateCount & " subject(s)." & vbCrLf & vbCrLf & _
               "¥ Lecturer names, status, activity codes refreshed (columns L-O)" & vbCrLf & _
               "¥ Your notes and enrolments preserved (columns P, S)", vbInformation, "Refresh Complete"
    End If
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = origScreenUpdating
    Application.EnableEvents = origEnableEvents
    DoEvents
    Application.StatusBar = False
    
    MsgBox "Error refreshing lecturer data:" & vbCrLf & vbCrLf & _
           Err.description & vbCrLf & _
           "Error " & Err.Number, vbCritical, "Refresh Error"
End Sub

'---------------------------------------------------------------
' GetSourceParameters
' Purpose: Open the source workbook read-only and read year,
'          teaching matrix filename, and email from Dashboard
' Returns: True if valid parameters found
'---------------------------------------------------------------
Private Function GetSourceParameters(ByRef yearValue As String, ByRef teachingMatrix As String, ByRef emailValue As String) As Boolean
    On Error Resume Next
    
    GetSourceParameters = False
    
    ' sourceWb = "Automated Handbook Data System.xlsm" on SharePoint
    Dim sourceWb As Workbook
    Dim weOpenedIt As Boolean
    Set sourceWb = FindAlreadyOpenWorkbook(SOURCE_FILE_PATH)
    
    If sourceWb Is Nothing Then
        Set sourceWb = Workbooks.Open(SOURCE_FILE_PATH, ReadOnly:=True, UpdateLinks:=False, Notify:=False)
        weOpenedIt = True
    End If
    
    If sourceWb Is Nothing Then Exit Function
    
    Dim sourceSheet As Worksheet
    Set sourceSheet = sourceWb.Sheets("Dashboard")
    
    If Not sourceSheet Is Nothing Then
        yearValue = Trim(CStr(sourceSheet.Range("C2").Value))
        teachingMatrix = Trim(CStr(sourceSheet.Range("C5").Value))
        emailValue = Trim(CStr(sourceSheet.Range("C12").Value))
        
        ' Year is the only required parameter
        GetSourceParameters = (yearValue <> "" And IsNumeric(yearValue))
    End If
    
    ' Only close if we opened it — don't close a file the user has open
    If weOpenedIt Then sourceWb.Close SaveChanges:=False
    
    On Error GoTo 0
End Function

'---------------------------------------------------------------
' TriggerTeachingMatrixWorkflow
' Purpose: Send HTTP POST to the Power Automate endpoint with
'          year/teaching matrix/email parameters
' Returns: True if request succeeded
'---------------------------------------------------------------
Private Function TriggerTeachingMatrixWorkflow(yearValue As String, teachingMatrix As String, emailValue As String) As Boolean
    On Error GoTo ErrorHandler
    
    TriggerTeachingMatrixWorkflow = False
    
    ' Build JSON payload
    Dim jsonData As String
    jsonData = "{" & Chr(34) & "year" & Chr(34) & ":" & yearValue & "," & _
               Chr(34) & "teachingMatrixFilename" & Chr(34) & ":" & Chr(34) & EscapeJSON(teachingMatrix) & Chr(34) & "," & _
               Chr(34) & "email" & Chr(34) & ":" & Chr(34) & EscapeJSON(emailValue) & Chr(34) & "}"
    
    ' Send HTTP request
    Dim result As String
    
    #If Mac Then
        result = SendRequestMac(TEACHING_MATRIX_URL, jsonData)
    #Else
        result = SendRequestWindows(TEACHING_MATRIX_URL, jsonData)
    #End If
    
    ' Parse the synchronous HTTP response body from the PA flow.
    Dim scriptResult As String
    scriptResult = Trim(result)
    If Left(scriptResult, 1) = Chr(34) And Right(scriptResult, 1) = Chr(34) Then
        scriptResult = Mid(scriptResult, 2, Len(scriptResult) - 2)
    End If
    
    Dim flowFailed As Boolean
    flowFailed = (result = "ERROR" Or result = "" Or UCase(Left(scriptResult, 5)) = "ERROR")
    
    TriggerTeachingMatrixWorkflow = Not flowFailed
    Exit Function
    
ErrorHandler:
    TriggerTeachingMatrixWorkflow = False
End Function

'---------------------------------------------------------------
' FindAlreadyOpenWorkbook
' Purpose: Check whether a workbook from a given path is already
'          open in this Excel session. Returns Nothing if not found.
'          Prevents accidental double-open or closing user files.
'---------------------------------------------------------------
Private Function FindAlreadyOpenWorkbook(filePath As String) As Workbook
    On Error Resume Next
    
    Dim wb As Workbook
    Dim targetName As String
    
    ' Extract bare filename for fallback name-match (SharePoint paths vary)
    Dim parts() As String
    parts = Split(filePath, "/")
    targetName = parts(UBound(parts))
    
    For Each wb In Workbooks
        If wb.FullName = filePath Then
            Set FindAlreadyOpenWorkbook = wb
            Exit Function
        End If
        If LCase(wb.Name) = LCase(targetName) Then
            Set FindAlreadyOpenWorkbook = wb
            Exit Function
        End If
    Next wb
    
    Set FindAlreadyOpenWorkbook = Nothing
    On Error GoTo 0
End Function

'---------------------------------------------------------------
' ClearTeachingStreamStatus
' Purpose: Reset the progress_bar "Teaching Stream" status to empty
'          before triggering the workflow, so monitoring starts clean
'          and doesn't detect stale "Done" from a previous run.
' Called by: RefreshLecturerData (before TriggerTeachingMatrixWorkflow)
'---------------------------------------------------------------
Private Sub ClearTeachingStreamStatus()
    On Error Resume Next
    
    Dim sourceWb As Workbook
    Dim weOpenedIt As Boolean
    Set sourceWb = FindAlreadyOpenWorkbook(SOURCE_FILE_PATH)
    
    If sourceWb Is Nothing Then
        Set sourceWb = Workbooks.Open(SOURCE_FILE_PATH, ReadOnly:=False, UpdateLinks:=False, Notify:=False)
        weOpenedIt = True
    End If
    
    If sourceWb Is Nothing Then Exit Sub
    
    Dim dashSheet As Worksheet
    Set dashSheet = sourceWb.Sheets("Dashboard")
    If dashSheet Is Nothing Then
        If weOpenedIt Then sourceWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Clear via progress_bar table
    Dim tbl As ListObject
    Set tbl = dashSheet.ListObjects("progress_bar")
    
    If Not tbl Is Nothing And tbl.DataBodyRange.Rows.Count > 0 Then
        Dim progressCol As Long, statusCol As Long
        progressCol = 0
        statusCol = 0
        
        Dim c As Long
        For c = 1 To tbl.HeaderRowRange.Columns.Count
            Dim hdrVal As String
            hdrVal = LCase(Trim(CStr(tbl.HeaderRowRange.Cells(1, c).Value)))
            If InStr(hdrVal, "progress") > 0 Then progressCol = c
            If InStr(hdrVal, "status") > 0 Then statusCol = c
        Next c
        
        If progressCol > 0 And statusCol > 0 Then
            Dim r As Long
            For r = 1 To tbl.DataBodyRange.Rows.Count
                Dim rowLabel As String
                rowLabel = LCase(Trim(CStr(tbl.DataBodyRange.Cells(r, progressCol).Value)))
                If InStr(rowLabel, "teaching stream") > 0 Then
                    tbl.DataBodyRange.Cells(r, statusCol).Value = ""
                    tbl.DataBodyRange.Cells(r, statusCol).Interior.ColorIndex = xlNone
                    Exit For
                End If
            Next r
        End If
    End If
    
    ' Also clear raw F5 for legacy compat
    dashSheet.Range("F5").Value = ""
    dashSheet.Range("F5").Interior.ColorIndex = xlNone
    
    sourceWb.Save
    If weOpenedIt Then sourceWb.Close SaveChanges:=False
    
    On Error GoTo 0
End Sub

'---------------------------------------------------------------
' SetTeachingStreamRunningStatus
' Purpose: Update the progress_bar "Teaching Stream" status to
'          "Running..." (Orange) before triggering the workflow,
'          providing visual feedback on the master dashboard.
' Called by: RefreshLecturerData 
'---------------------------------------------------------------
Private Sub SetTeachingStreamRunningStatus()
    On Error Resume Next
    
    Dim sourceWb As Workbook
    Dim weOpenedIt As Boolean
    Set sourceWb = FindAlreadyOpenWorkbook(SOURCE_FILE_PATH)
    
    If sourceWb Is Nothing Then
        Set sourceWb = Workbooks.Open(SOURCE_FILE_PATH, ReadOnly:=False, UpdateLinks:=False, Notify:=False)
        weOpenedIt = True
    End If
    
    If sourceWb Is Nothing Then Exit Sub
    
    Dim dashSheet As Worksheet
    Set dashSheet = sourceWb.Sheets("Dashboard")
    If dashSheet Is Nothing Then
        If weOpenedIt Then sourceWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Clear via progress_bar table
    Dim tbl As ListObject
    Set tbl = dashSheet.ListObjects("progress_bar")
    
    If Not tbl Is Nothing And tbl.DataBodyRange.Rows.Count > 0 Then
        Dim progressCol As Long, statusCol As Long
        progressCol = 0
        statusCol = 0
        
        Dim c As Long
        For c = 1 To tbl.HeaderRowRange.Columns.Count
            Dim hdrVal As String
            hdrVal = LCase(Trim(CStr(tbl.HeaderRowRange.Cells(1, c).Value)))
            If InStr(hdrVal, "progress") > 0 Then progressCol = c
            If InStr(hdrVal, "status") > 0 Then statusCol = c
        Next c
        
        If progressCol > 0 And statusCol > 0 Then
            Dim r As Long
            For r = 1 To tbl.DataBodyRange.Rows.Count
                Dim rowLabel As String
                rowLabel = LCase(Trim(CStr(tbl.DataBodyRange.Cells(r, progressCol).Value)))
                If InStr(rowLabel, "teaching stream") > 0 Then
                    With tbl.DataBodyRange.Cells(r, statusCol)
                        .Value = "Running..."
                        .Interior.Color = RGB(255, 192, 0) ' Orange
                    End With
                    Exit For
                End If
            Next r
        End If
    End If
    
    ' Also update raw F5 for legacy compat
    With dashSheet.Range("F5")
        .Value = "Running..."
        .Interior.Color = RGB(255, 192, 0)
    End With
    
    sourceWb.Save
    If weOpenedIt Then sourceWb.Close SaveChanges:=False
    
    On Error GoTo 0
End Sub

'===============================================================
' SECTION 3: HTTP REQUESTS
'===============================================================

'---------------------------------------------------------------
' SendRequestMac
' Purpose: Send HTTP POST via AppleScript/curl (Mac only)
' Returns: Response text, or "ERROR"
'---------------------------------------------------------------
Private Function SendRequestMac(url As String, jsonData As String) As String
    Dim scriptCode As String
    Dim result As String
    
    jsonData = Replace(jsonData, "\\", "\\\\")
    jsonData = Replace(jsonData, """", "\""")
    
    scriptCode = "do shell script ""curl -s -X POST '" & url & "' " & _
                 "-H 'Content-Type: application/json' " & _
                 "-d '" & jsonData & "' 2>&1"""
    
    On Error Resume Next
    result = MacScript(scriptCode)
    If Err.Number <> 0 Then result = "ERROR"
    On Error GoTo 0
    
    SendRequestMac = result
End Function

'---------------------------------------------------------------
' SendRequestWindows
' Purpose: Send HTTP POST via MSXML2 (Windows only)
' Returns: Response text, or "ERROR"
'---------------------------------------------------------------
Private Function SendRequestWindows(url As String, jsonData As String) As String
    Dim http As Object
    
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP")
    If http Is Nothing Then Set http = CreateObject("MSXML2.ServerXMLHTTP")
    
    If http Is Nothing Then
        SendRequestWindows = "ERROR"
        Exit Function
    End If
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send jsonData
    
    If Err.Number <> 0 Then
        SendRequestWindows = "ERROR"
    Else
        SendRequestWindows = http.responseText
    End If
    
    Set http = Nothing
    On Error GoTo 0
End Function

'---------------------------------------------------------------
' EscapeJSON
' Purpose: Escape special characters for JSON payload
'---------------------------------------------------------------
Private Function EscapeJSON(text As String) As String
    Dim result As String
    result = text
    
    result = Replace(result, "\\", "\\\\")
    result = Replace(result, Chr(34), "\\" & Chr(34))
    result = Replace(result, vbCr, "")
    result = Replace(result, vbLf, "")
    result = Replace(result, vbTab, " ")
    
    EscapeJSON = result
End Function

'===============================================================
' SECTION 4: SUBJECT BLOCK IDENTIFICATION
'===============================================================

'---------------------------------------------------------------
' IdentifySubjectBlocks
' Purpose: Scan FHY and SHY sheets to find all subject blocks
'          by looking for UIDs ending in "_0" (header rows)
' Called by: RefreshLecturerData
'---------------------------------------------------------------
Private Sub IdentifySubjectBlocks(wb As Workbook, subjectBlocks As Collection)
    On Error Resume Next
    
    Dim calcSheet As Worksheet
    For Each calcSheet In wb.Worksheets
        If calcSheet.name = "FHY Calculations" Or calcSheet.name = "SHY Calculations" Then
            calcSheet.Unprotect
            
            Dim lastRow As Long
            lastRow = calcSheet.Cells(calcSheet.Rows.Count, "A").End(xlUp).Row
            
            Dim i As Long
            i = 4  ' Start from row 4 (the first possible subject header)
            
            Do While i <= lastRow
                Dim uid As String
                uid = calcSheet.Cells(i, 1).Value
                
                ' Header rows have UIDs ending in "_0" (e.g. "MGMT20001_Semester 1_0")
                If Len(uid) > 2 And Right(uid, 2) = "_0" Then
                    ' Create SubjectBlockInfo as array
                    Dim blockInfo(0 To 6) As Variant
                    blockInfo(SBI_SHEETNAME) = calcSheet.name
                    blockInfo(SBI_HEADERROW) = i
                    
                    ' Parse UID for subject code and study period
                    Dim parts() As String
                    parts = Split(uid, "_")
                    
                    If UBound(parts) >= 2 Then
                        blockInfo(SBI_SUBJECTCODE) = parts(0)
                        blockInfo(SBI_STUDYPERIOD) = parts(1)
                        
                        ' Find Total row
                        blockInfo(SBI_TOTALROW) = FindTotalRow(calcSheet, i)
                        
                        ' Find last subject row
                        blockInfo(SBI_LASTSUBJECTROW) = FindLastSubjectRow(calcSheet, i)
                        
                        ' Calculate number of assessment rows
                        blockInfo(SBI_NUMASSESSMENTROWS) = blockInfo(SBI_TOTALROW) - blockInfo(SBI_HEADERROW) - 1
                        
                        subjectBlocks.add blockInfo
                        
                        ' Jump to next subject
                        i = blockInfo(SBI_LASTSUBJECTROW)
                    End If
                End If
                
                i = i + 1
            Loop
        End If
    Next calcSheet
    
    On Error GoTo 0
End Sub

'===============================================================
' SECTION 5: TEACHING DATA LOADING
'===============================================================

'---------------------------------------------------------------
' LoadTeachingStreamData
' Purpose: Open the source workbook and load columns B-G of
'          the teaching stream sheet into a 2D array
' Returns: 2D variant array, or Empty if no data
'---------------------------------------------------------------
Private Function LoadTeachingStreamData(sourcePath As String) As Variant
    On Error GoTo ErrorHandler
    
    ' Close any stale in-memory copy so we read FRESH post-workflow data.
    ' The Office Script updates the server copy; the local open copy
    ' still has pre-refresh teaching stream data.
    On Error Resume Next
    Dim staleWb As Workbook
    Set staleWb = FindAlreadyOpenWorkbook(sourcePath)
    If Not staleWb Is Nothing Then
        If staleWb.Name <> ActiveWorkbook.Name Then
            staleWb.Close SaveChanges:=False
        End If
    End If
    On Error GoTo ErrorHandler
    
    Dim sourceWb As Workbook
    Set sourceWb = Workbooks.Open(sourcePath, ReadOnly:=True, UpdateLinks:=False, Notify:=False)
    
    If sourceWb Is Nothing Then
        LoadTeachingStreamData = Empty
        Exit Function
    End If
    
    Dim sourceSheet As Worksheet
    Set sourceSheet = sourceWb.Sheets(TEACHING_STREAM_SHEET)
    
    If sourceSheet Is Nothing Then
        If sourceWb.Name <> ActiveWorkbook.Name Then sourceWb.Close SaveChanges:=False
        LoadTeachingStreamData = Empty
        Exit Function
    End If
    
    ' Read from the teaching_stream table dynamically to avoid column shifting issues
    On Error Resume Next
    Dim tbl As ListObject
    Set tbl = sourceSheet.ListObjects("teaching_stream")
    On Error GoTo ErrorHandler
    
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then
        If sourceWb.Name <> ActiveWorkbook.Name Then sourceWb.Close SaveChanges:=False
        LoadTeachingStreamData = Empty
        Exit Function
    End If
    
    Dim rowCount As Long
    rowCount = tbl.DataBodyRange.Rows.Count
    
    If rowCount = 0 Then
        If sourceWb.Name <> ActiveWorkbook.Name Then sourceWb.Close SaveChanges:=False
        LoadTeachingStreamData = Empty
        Exit Function
    End If
    
    ' Find column indices dynamically
    Dim colSubj As Long, colPeriod As Long, colLec As Long
    Dim colStatus As Long, colAct As Long, colStreams As Long
    
    Dim c As Long
    For c = 1 To tbl.HeaderRowRange.Columns.Count
        Dim hdr As String
        hdr = LCase(Trim(CStr(tbl.HeaderRowRange.Cells(1, c).Value)))
        
        If hdr = "subject code" Then colSubj = c
        If hdr = "study period" Or hdr = "teaching period" Then colPeriod = c
        If hdr = "lecturer" Or hdr = "name" Or hdr = "display name" Then colLec = c
        If hdr = "status" Then colStatus = c
        If InStr(hdr, "activity id") > 0 Or InStr(hdr, "activity code") > 0 Then colAct = c
        If hdr = "streams" Or hdr = "stream count" Then colStreams = c
    Next c
    
    ' Fallback to old hardcoded offsets if table headers are completely unrecognized
    ' (The Office Script writes array: Key, Subj, Period, Lec, Status, Act, Streams)
    If colSubj = 0 Or colLec = 0 Then
        colSubj = 2
        colPeriod = 3
        colLec = 4
        colStatus = 5
        colAct = 6
        colStreams = 7
    End If
    
    ' Extract data into a standard 6-column array
    ' 1: Subject Code, 2: Study Period, 3: Lecturer, 4: Status, 5: Activity Code, 6: Streams
    Dim outData() As Variant
    ReDim outData(1 To rowCount, 1 To 6)
    
    Dim r As Long
    For r = 1 To rowCount
        outData(r, 1) = tbl.DataBodyRange.Cells(r, colSubj).Value
        
        If colPeriod > 0 Then outData(r, 2) = tbl.DataBodyRange.Cells(r, colPeriod).Value
        outData(r, 3) = tbl.DataBodyRange.Cells(r, colLec).Value
        
        If colStatus > 0 Then outData(r, 4) = tbl.DataBodyRange.Cells(r, colStatus).Value
        If colAct > 0 Then outData(r, 5) = tbl.DataBodyRange.Cells(r, colAct).Value
        If colStreams > 0 Then outData(r, 6) = tbl.DataBodyRange.Cells(r, colStreams).Value
    Next r
    
    LoadTeachingStreamData = outData
    
    If sourceWb.Name <> ActiveWorkbook.Name Then sourceWb.Close SaveChanges:=False
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If Not sourceWb Is Nothing Then
        If sourceWb.Name <> ActiveWorkbook.Name Then sourceWb.Close SaveChanges:=False
    End If
    On Error GoTo 0
    LoadTeachingStreamData = Empty
End Function

'===============================================================
' SECTION 6: LECTURER DATA UPDATE
'===============================================================

'---------------------------------------------------------------
' UpdateAllLecturers
' Purpose: Refresh columns L-O from teaching data for every
'          subject block, adding rows before Total if needed.
'          Preserves user edits in columns P-S.
' Called by: RefreshLecturerData
' Returns: Number of subjects updated
'---------------------------------------------------------------
Private Function UpdateAllLecturers(wb As Workbook, teachingData As Variant, subjectBlocks As Collection) As Long
    On Error Resume Next
    
    UpdateAllLecturers = 0
    
    Dim i As Long
    
    ' =========================================================================
    ' Row-offset tracking — when rows are inserted into a sheet for subject N,
    ' all subsequent subjects on the SAME sheet have stale cached row numbers.
    ' We accumulate the delta here and apply it at read time.
    ' Reset to 0 whenever we switch to a different sheet.
    ' =========================================================================
    Dim rowOffset As Long
    Dim currentSheetName As String
    rowOffset = 0
    currentSheetName = ""
    
    For i = 1 To subjectBlocks.Count
        Dim blockInfo As Variant
        blockInfo = subjectBlocks(i)
        
        Dim calcSheet As Worksheet
        Set calcSheet = wb.Sheets(CStr(blockInfo(SBI_SHEETNAME)))
        calcSheet.Unprotect
        
        ' Reset offset when moving to a new sheet
        If CStr(blockInfo(SBI_SHEETNAME)) <> currentSheetName Then
            currentSheetName = CStr(blockInfo(SBI_SHEETNAME))
            rowOffset = 0
        End If
        
        ' Mac-safe type casting — Variant→String before passing to match function
        Dim blockSubjectCode As String
        Dim blockStudyPeriod As String
        blockSubjectCode = CStr(blockInfo(SBI_SUBJECTCODE))
        blockStudyPeriod = CStr(blockInfo(SBI_STUDYPERIOD))
        
        ' Get updated lecturers for this subject (returns array)
        Dim freshLecturers As Variant
        freshLecturers = GetLecturersFromTeachingData(teachingData, blockSubjectCode, blockStudyPeriod)
        
        Dim lecturerCount As Long
        lecturerCount = 0
        
        If IsArray(freshLecturers) Then
            On Error Resume Next
            lecturerCount = UBound(freshLecturers, 1)
            On Error GoTo 0
        End If
        
        If lecturerCount > 0 Then
            Dim headerRow As Long
            Dim totalRow As Long
            Dim firstLecturerRow As Long
            Dim availableRows As Long
            
            ' Apply cumulative offset so stale cached indices become correct
            headerRow = CLng(blockInfo(SBI_HEADERROW)) + rowOffset
            
            ' Re-verify totalRow live from the sheet — the cached value may be
            ' stale if a prior (buggy) refresh inserted rows at the wrong place.
            ' FindTotalRow scans from headerRow forward looking for "Total" in col E,
            ' so it always locates the real current position regardless of history.
            totalRow = FindTotalRow(calcSheet, headerRow)
            
            firstLecturerRow = headerRow + 1
            availableRows = totalRow - firstLecturerRow  ' Rows between header and Total
            
            ' =========================================================================
            ' INSERT ROWS — add rows before Total if more lecturers than available slots
            ' =========================================================================
            If lecturerCount > availableRows Then
                Dim rowsToAdd As Long
                rowsToAdd = lecturerCount - availableRows
                
                Dim insertRow As Long
                Dim j As Long
                
                For j = 1 To rowsToAdd
                    insertRow = totalRow
                    calcSheet.Rows(insertRow).Insert Shift:=xlDown
                    
                    ' Copy formatting from the row above to keep consistent styling
                    calcSheet.Rows(insertRow - 1).Copy
                    calcSheet.Rows(insertRow).PasteSpecial xlPasteFormats
                    Application.CutCopyMode = False
                    
                    calcSheet.Rows(insertRow).ClearContents
                    
                    ' Generate UID for new row
                    Dim newUID As String
                    Dim uidSuffix As Long
                    uidSuffix = (insertRow - headerRow - 1)
                    newUID = blockSubjectCode & "_" & blockStudyPeriod & "_" & uidSuffix
                    calcSheet.Cells(insertRow, 1).Value = newUID
                    
                    ' Update Total row position (it shifted down)
                    totalRow = totalRow + 1
                Next j
                
                ' Accumulate the offset so later subjects on this sheet are corrected
                rowOffset = rowOffset + rowsToAdd
                
                ' Recalculate available rows
                availableRows = totalRow - firstLecturerRow
            End If
            
            ' =========================================================================
            ' CLEAR OLD DATA — columns L–O values AND bold formatting;
            ' columns P–S are user-edited and NOT touched.
            ' ClearContents only clears values; bold from shifted rows persists
            ' unless we explicitly reset it.
            ' =========================================================================
            Dim Row As Long
            For Row = firstLecturerRow To totalRow - 1
                Dim clearRange As Range
                Set clearRange = calcSheet.Range(calcSheet.Cells(Row, 12), calcSheet.Cells(Row, 15))
                clearRange.ClearContents
                clearRange.Font.Bold = False
            Next Row
            
            ' =========================================================================
            ' WRITE FRESH DATA — populate lecturer columns from teaching stream
            ' =========================================================================
            Dim outputRow As Long
            outputRow = firstLecturerRow
            
            Dim k As Long
            For k = 1 To lecturerCount
                If outputRow < totalRow Then
                    calcSheet.Cells(outputRow, 12).Value = freshLecturers(k, 0)  ' Name
                    calcSheet.Cells(outputRow, 13).Value = freshLecturers(k, 1)  ' Status
                    calcSheet.Cells(outputRow, 14).Value = freshLecturers(k, 3)  ' Streams
                    calcSheet.Cells(outputRow, 15).Value = freshLecturers(k, 2)  ' Activity Code
                    
                    ' Bold first lecturer (subject coordinator)
                    If outputRow = firstLecturerRow Then
                        calcSheet.Cells(outputRow, 12).Font.Bold = True
                    End If
                    
                    outputRow = outputRow + 1
                End If
            Next k
            
            Call ApplyLecturerFormulas(calcSheet, headerRow, totalRow)
            
            UpdateAllLecturers = UpdateAllLecturers + 1
        End If
        
        calcSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, _
                   AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
    Next i
    
    On Error GoTo 0
End Function


'---------------------------------------------------------------
' ApplyLecturerFormulas
' Purpose: Write batch formulas for columns Q (Allocated Marking)
'          and R (Marking Support Hours Available) for all
'          lecturer rows in a subject block
' Called by: UpdateAllLecturers
'---------------------------------------------------------------
Private Sub ApplyLecturerFormulas(calcSheet As Worksheet, headerRow As Long, totalRow As Long)
    On Error Resume Next
    
    Dim firstLecturerRow As Long
    Dim lastLecturerRow As Long
    Dim numLecturers As Long
    
    firstLecturerRow = headerRow + 1
    lastLecturerRow = totalRow - 1
    numLecturers = lastLecturerRow - firstLecturerRow + 1
    
    If numLecturers <= 0 Then Exit Sub
    
    ' Build formula arrays for batch write
    Dim formulas As Variant
    ReDim formulas(1 To numLecturers, 1 To 2)
    
    Dim outputRow As Long
    Dim i As Long
    
    i = 1
    For outputRow = firstLecturerRow To lastLecturerRow
        ' Column Q (17): Allocated Marking
        formulas(i, 1) = "=IF(M" & outputRow & "=""Continuing T&R"",N" & outputRow & "*VALUE(LEFT($Q$2,FIND("" "",$Q$2)-1)),"""")"
        
        ' Column R (18): Marking Support Hours Available
        formulas(i, 2) = "=IF(OR(P" & outputRow & "="""",Q" & outputRow & "=""""),"""",$J$" & totalRow & "*(P" & outputRow & "/D" & headerRow & ")-Q" & outputRow & ")"
        
        i = i + 1
    Next outputRow
    
    ' Batch write both formula columns at once
    calcSheet.Cells(firstLecturerRow, 17).Resize(numLecturers, 2).Formula = formulas
    
    On Error GoTo 0
End Sub

'===============================================================
' SECTION 7: HELPER FUNCTIONS
'===============================================================

'---------------------------------------------------------------
' FindTotalRow
' Purpose: Search for the "Total" row in column E starting
'          from the header row
' Returns: Row number of the Total row
'---------------------------------------------------------------
Private Function FindTotalRow(calcSheet As Worksheet, headerRow As Long) As Long
    Dim Row As Long
    For Row = headerRow + 1 To calcSheet.Cells(calcSheet.Rows.Count, "A").End(xlUp).Row
        Dim uid As String
        Dim cellE As String
        
        uid = calcSheet.Cells(Row, 1).Value
        cellE = Trim(CStr(calcSheet.Cells(Row, 5).Value))
        
        ' Find Total row
        If cellE = "Total" Then
            FindTotalRow = Row
            Exit Function
        End If
        
        ' Hit next subject header — Total row must be just above
        If Len(uid) > 2 And Right(uid, 2) = "_0" Then
            FindTotalRow = Row - 1
            Exit Function
        End If
    Next Row
    
    ' Fallback — last used row
    FindTotalRow = calcSheet.Cells(calcSheet.Rows.Count, "A").End(xlUp).Row
End Function

'---------------------------------------------------------------
' FindLastSubjectRow
' Purpose: Find the last row of a subject block (before the
'          next subject header or category header)
' Returns: Row number of the last row in the block
'---------------------------------------------------------------
Private Function FindLastSubjectRow(calcSheet As Worksheet, headerRow As Long) As Long
    Dim Row As Long
    For Row = headerRow + 1 To calcSheet.Cells(calcSheet.Rows.Count, "A").End(xlUp).Row
        Dim nextUID As String
        Dim cellB As String
        
        nextUID = calcSheet.Cells(Row, 1).Value
        cellB = Trim(CStr(calcSheet.Cells(Row, 2).Value))
        
        If Len(nextUID) > 2 And Right(nextUID, 2) = "_0" Then
            FindLastSubjectRow = Row - 1
            Exit Function
        End If
        
        ' Category headers (SUMMER, WINTER, etc.) mark block boundaries
        If (cellB = "SUMMER" Or cellB = "WINTER" Or cellB = "SEMESTER 1" Or cellB = "SEMESTER 2") And _
           Trim(CStr(calcSheet.Cells(Row, 1).Value)) = "" Then
            FindLastSubjectRow = Row - 1
            Exit Function
        End If
    Next Row
    
    FindLastSubjectRow = calcSheet.Cells(calcSheet.Rows.Count, "A").End(xlUp).Row
End Function

'---------------------------------------------------------------
' GetLecturersFromTeachingData
' Purpose: Extract matching lecturers for a subject from the
'          teaching data array, trying exact then flexible
'          study period matching
' Returns: 2D array (1..N, 0..3) of lecturer data, or empty
'---------------------------------------------------------------
Private Function GetLecturersFromTeachingData(teachingData As Variant, subjectCode As String, studyPeriod As String) As Variant
    On Error Resume Next
    
    ' Temporary storage using Collection (internal only)
    Dim tempLecturers As Collection
    Dim uniqueDict As Collection
    
    Set tempLecturers = New Collection
    Set uniqueDict = New Collection
    
    ' Try exact match first
    Dim i As Long
    For i = 1 To UBound(teachingData, 1)
        If teachingData(i, 1) = subjectCode And Trim(CStr(teachingData(i, 2))) = studyPeriod Then
            Dim lecName As String
            lecName = Trim(CStr(teachingData(i, 3)))
            
            If Not CollectionKeyExists(uniqueDict, lecName) Then
                tempLecturers.add Array(lecName, teachingData(i, 4), teachingData(i, 5), teachingData(i, 6))
                On Error Resume Next
                uniqueDict.add True, lecName
                On Error GoTo 0
            End If
        End If
    Next i
    
    ' If no exact match, try flexible matching
    If tempLecturers.Count = 0 Then
        Dim flexPeriod As String
        flexPeriod = Replace(studyPeriod, " Term", "", 1, -1, vbTextCompare)
        flexPeriod = Trim(flexPeriod)
        
        For i = 1 To UBound(teachingData, 1)
            Dim dataStudyPeriod As String
            dataStudyPeriod = Trim(CStr(teachingData(i, 2)))
            
            If teachingData(i, 1) = subjectCode And _
               (dataStudyPeriod = flexPeriod Or _
                InStr(1, studyPeriod, dataStudyPeriod, vbTextCompare) > 0 Or _
                InStr(1, dataStudyPeriod, studyPeriod, vbTextCompare) > 0) Then
                
                lecName = Trim(CStr(teachingData(i, 3)))
                If Not CollectionKeyExists(uniqueDict, lecName) Then
                    tempLecturers.add Array(lecName, teachingData(i, 4), teachingData(i, 5), teachingData(i, 6))
                    On Error Resume Next
                    uniqueDict.add True, lecName
                    On Error GoTo 0
                End If
            End If
        Next i
    End If
    
    ' Convert Collection to 2D Array (Mac-compatible return type)
    If tempLecturers.Count = 0 Then
        ' Return empty array
        GetLecturersFromTeachingData = Array()
        Exit Function
    End If
    
    ' Build 2D array: lecturers(1 to N, 0 to 3)
    ' Columns: 0=Name, 1=Status, 2=ActivityCode, 3=Streams
    Dim lecturersArray() As Variant
    ReDim lecturersArray(1 To tempLecturers.Count, 0 To 3)
    
    Dim j As Long
    Dim lecItem As Variant
    j = 1
    For Each lecItem In tempLecturers
        lecturersArray(j, 0) = lecItem(0)  ' Name
        lecturersArray(j, 1) = lecItem(1)  ' Status
        lecturersArray(j, 2) = lecItem(2)  ' Activity Code
        lecturersArray(j, 3) = lecItem(3)  ' Streams
        j = j + 1
    Next lecItem
    
    GetLecturersFromTeachingData = lecturersArray
End Function

'---------------------------------------------------------------
' CollectionKeyExists
' Purpose: Test whether a key already exists in a VBA Collection
'          (used for deduplication)
' Returns: True if key exists
'---------------------------------------------------------------
Private Function CollectionKeyExists(col As Collection, key As String) As Boolean
    On Error Resume Next
    Dim temp As Variant
    temp = col(key)
    CollectionKeyExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function
