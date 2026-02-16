'=============================================================================
' MODULE: LecturerRefresh
' PURPOSE: Trigger Power Automate Workflow to refresh data from Teaching Matrix, wait for completion, refresh lecturer data in Marking Support
' LOCATION: Exported calculation sheets only (e.g., 2026 Marking & Admin Support Calculations.xlsm)
'
' WORKFLOW:
'   1. User clicks Refresh button
'   2. Read parameters from source file (year, teaching matrix filename, email)
'   3. Trigger Workflow 3 (Teaching Matrix Processing)
'   4. Poll Dashboard F5 cell every 3 seconds until "DONE" (max 2 minutes)
'   5. Load fresh teaching stream data from source
'   6. Update lecturer columns (L-O) with fresh data ONLY
'   7. Preserve user edits in columns P (Stream Enrolment) and S (Notes)
'   8. Add rows before Total if needed (no orphan tracking)
'=============================================================================

Option Explicit

'------------------------------------------------------------------------------
' CONSTANTS - HARDCODED PATHS
'------------------------------------------------------------------------------
' Source file path (Automated Handbook Data System)
Private Const SOURCE_FILE_PATH As String = "https://unimelbcloud.sharepoint.com/teams/DepartmentofManagementMarketing-DepartmentOperations/Shared Documents/TEACHING SUPPORT/Handbook (Course & Subject Changes)/Auto Handbook System/Automated Handbook Data System.xlsm"

' Source sheets
Private Const TEACHING_STREAM_SHEET As String = "teaching stream"

' Workflow 3 Power Automate endpoint
Private Const WORKFLOW3_URL As String = "https://default0e5bf3cf1ff446b7917652c538c22a.4d.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/7f198e614c734715bc0153d818de1ef7/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=5uUuhFHyiL37O_ajy-6t2r65nFqc7NA_oJDhYmYFT9g"

'------------------------------------------------------------------------------
' ARRAY INDICES - Replace Private Type structures
'------------------------------------------------------------------------------
' SubjectBlockInfo array indices (0-6)
Private Const SBI_SHEETNAME As Integer = 0
Private Const SBI_SUBJECTCODE As Integer = 1
Private Const SBI_STUDYPERIOD As Integer = 2
Private Const SBI_HEADERROW As Integer = 3
Private Const SBI_TOTALROW As Integer = 4
Private Const SBI_LASTSUBJECTROW As Integer = 5
Private Const SBI_NUMASSESSMENTROWS As Integer = 6

'=============================================================================
' MAIN ENTRY POINT - RefreshLecturerData
'=============================================================================
Public Sub RefreshLecturerData()
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
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
    Application.StatusBar = "Reading parameters from source file..."
    
    Dim yearValue As String
    Dim teachingMatrix As String
    Dim emailValue As String
    
    If Not GetSourceParameters(yearValue, teachingMatrix, emailValue) Then
        MsgBox "Could not read parameters from source file." & vbCrLf & vbCrLf & _
               "Source: " & SOURCE_FILE_PATH, vbExclamation, "Connection Error"
        GoTo CleanExit
    End If
    
    ' STEP 2: Trigger Workflow 3
    Application.StatusBar = "Triggering Teaching Matrix workflow..."
    
    If Not TriggerWorkflow3(yearValue, teachingMatrix, emailValue) Then
        MsgBox "Failed to trigger Workflow 3." & vbCrLf & vbCrLf & _
               "Check your network connection and try again.", vbExclamation, "Workflow Error"
        GoTo CleanExit
    End If
    
    MsgBox "Teaching Matrix workflow triggered successfully!" & vbCrLf & vbCrLf & _
           "Monitoring for completion... (usually <1 minute)" & vbCrLf & vbCrLf & _
           "Please wait while the workflow processes.", vbInformation, "Workflow Started"
    
    ' STEP 3: Wait for Workflow 3 completion
    Application.StatusBar = "Waiting for Teaching Matrix workflow to complete..."
    
    If Not WaitForWorkflow3Completion(120) Then  ' 2 minute timeout
        Dim response As VbMsgBoxResult
        response = MsgBox("Workflow 3 did not complete within 2 minutes." & vbCrLf & vbCrLf & _
                         "The workflow may still be running in the background." & vbCrLf & vbCrLf & _
                         "Continue refresh with potentially outdated data?", _
                         vbQuestion + vbYesNo, "Workflow Timeout")
        If response = vbNo Then GoTo CleanExit
    End If
    
    Application.StatusBar = "Workflow complete! Refreshing lecturer data..."
    
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
    Application.Calculation = origCalculation
    Application.ScreenUpdating = origScreenUpdating
    Application.EnableEvents = origEnableEvents
    Application.StatusBar = False
    
    If updateCount > 0 Then
        MsgBox "Lecturer data refreshed successfully!" & vbCrLf & vbCrLf & _
               "Updated " & updateCount & " subject(s)." & vbCrLf & vbCrLf & _
               "¥ Lecturer names, status, activity codes refreshed (columns L-O)" & vbCrLf & _
               "¥ Your notes and enrolments preserved (columns P, S)", vbInformation, "Refresh Complete"
    End If
    Exit Sub
    
ErrorHandler:
    Application.Calculation = origCalculation
    Application.ScreenUpdating = origScreenUpdating
    Application.EnableEvents = origEnableEvents
    Application.StatusBar = False
    
    MsgBox "Error refreshing lecturer data:" & vbCrLf & vbCrLf & _
           Err.description & vbCrLf & _
           "Error " & Err.Number, vbCritical, "Refresh Error"
End Sub

'------------------------------------------------------------------------------
' GET SOURCE PARAMETERS
' Opens source file and reads Dashboard C2, C5, C12
'------------------------------------------------------------------------------
Private Function GetSourceParameters(ByRef yearValue As String, ByRef teachingMatrix As String, ByRef emailValue As String) As Boolean
    On Error Resume Next
    
    GetSourceParameters = False
    
    Dim sourceWB As Workbook
    Set sourceWB = Workbooks.Open(SOURCE_FILE_PATH, ReadOnly:=True, UpdateLinks:=False, Notify:=False)
    
    If sourceWB Is Nothing Then Exit Function
    
    Dim ws As Worksheet
    Set ws = sourceWB.Sheets("Dashboard")
    
    If Not ws Is Nothing Then
        yearValue = Trim(CStr(ws.Range("C2").Value))
        teachingMatrix = Trim(CStr(ws.Range("C5").Value))
        emailValue = Trim(CStr(ws.Range("C12").Value))
        
        ' Year is required
        GetSourceParameters = (yearValue <> "" And IsNumeric(yearValue))
    End If
    
    sourceWB.Close SaveChanges:=False
    
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' TRIGGER WORKFLOW 3
' Sends HTTP POST to Power Automate endpoint
'------------------------------------------------------------------------------
Private Function TriggerWorkflow3(yearValue As String, teachingMatrix As String, emailValue As String) As Boolean
    On Error GoTo ErrorHandler
    
    TriggerWorkflow3 = False
    
    ' Build JSON payload
    Dim jsonData As String
    jsonData = "{" & Chr(34) & "year" & Chr(34) & ":" & yearValue & "," & _
               Chr(34) & "teachingMatrixFilename" & Chr(34) & ":" & Chr(34) & EscapeJSON(teachingMatrix) & Chr(34) & "," & _
               Chr(34) & "email" & Chr(34) & ":" & Chr(34) & EscapeJSON(emailValue) & Chr(34) & "}"
    
    ' Send HTTP request
    Dim result As String
    
    #If Mac Then
        result = SendRequestMac(WORKFLOW3_URL, jsonData)
    #Else
        result = SendRequestWindows(WORKFLOW3_URL, jsonData)
    #End If
    
    TriggerWorkflow3 = (result <> "ERROR")
    Exit Function
    
ErrorHandler:
    TriggerWorkflow3 = False
End Function

'------------------------------------------------------------------------------
' WAIT FOR WORKFLOW 3 COMPLETION
' Polls Dashboard F5 every 3 seconds until "DONE" or timeout
'------------------------------------------------------------------------------
Private Function WaitForWorkflow3Completion(maxWaitSeconds As Long) As Boolean
    On Error Resume Next
    
    WaitForWorkflow3Completion = False
    
    Dim startTime As Double
    Dim elapsedTime As Double
    Dim checkCount As Long
    
    startTime = Timer
    checkCount = 0
    
    Do
        DoEvents
        checkCount = checkCount + 1
        
        ' Check workflow status
        Dim currentStatus As String
        currentStatus = GetWorkflow3Status()
        
        ' Update status every 5 checks (~15 seconds)
        If checkCount Mod 5 = 0 Then
            elapsedTime = Timer - startTime
            If elapsedTime < 0 Then elapsedTime = elapsedTime + 86400
            Application.StatusBar = "Workflow status: " & currentStatus & " (elapsed: " & Format(elapsedTime, "0") & "s)"
        End If
        
        ' Check if complete
        Dim statusUpper As String
        statusUpper = UCase(Trim(currentStatus))
        
        If statusUpper = "DONE" Or statusUpper = "COMPLETE" Or statusUpper = "FINISHED" Or statusUpper = "SUCCESS" Then
            Application.StatusBar = "Workflow 3 completed successfully!"
            WaitForWorkflow3Completion = True
            Exit Function
        End If
        
        ' Check timeout
        elapsedTime = Timer - startTime
        If elapsedTime < 0 Then elapsedTime = elapsedTime + 86400
        
        If elapsedTime > maxWaitSeconds Then
            Application.StatusBar = "Workflow timeout reached"
            Exit Function
        End If
        
        ' Wait 3 seconds before next check
        Application.Wait (Now + TimeValue("0:00:03"))
        
    Loop
    
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' GET WORKFLOW 3 STATUS
' Opens source file read-only and reads Dashboard F5
'------------------------------------------------------------------------------
Private Function GetWorkflow3Status() As String
    On Error Resume Next
    
    GetWorkflow3Status = "Unknown"
    
    Dim sourceWB As Workbook
    Set sourceWB = Workbooks.Open(SOURCE_FILE_PATH, ReadOnly:=True, UpdateLinks:=False, Notify:=False)
    
    If Not sourceWB Is Nothing Then
        Dim ws As Worksheet
        Set ws = sourceWB.Sheets("Dashboard")
        
        If Not ws Is Nothing Then
            Dim cellValue As String
            cellValue = Trim(CStr(ws.Range("F5").Value))
            
            If cellValue <> "" Then
                GetWorkflow3Status = cellValue
            Else
                GetWorkflow3Status = "Not Started"
            End If
        End If
        
        sourceWB.Close SaveChanges:=False
    End If
    
    On Error GoTo 0
End Function

'=============================================================================
' HTTP REQUEST FUNCTIONS
'=============================================================================
Function SendRequestMac(url As String, jsonData As String) As String
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

Function SendRequestWindows(url As String, jsonData As String) As String
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

Function EscapeJSON(text As String) As String
    Dim result As String
    result = text
    
    result = Replace(result, "\\", "\\\\")
    result = Replace(result, Chr(34), "\\" & Chr(34))
    result = Replace(result, vbCr, "")
    result = Replace(result, vbLf, "")
    result = Replace(result, vbTab, " ")
    
    EscapeJSON = result
End Function

'=============================================================================
' SUBJECT BLOCK IDENTIFICATION
'=============================================================================

'------------------------------------------------------------------------------
' IDENTIFY SUBJECT BLOCKS
' Scans FHY and SHY Calculations sheets to find all subject blocks
'------------------------------------------------------------------------------
Private Sub IdentifySubjectBlocks(wb As Workbook, subjectBlocks As Collection)
    On Error Resume Next
    
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.name = "FHY Calculations" Or ws.name = "SHY Calculations" Then
            ws.Unprotect
            
            Dim lastRow As Long
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            
            Dim i As Long
            i = 4  ' Start from row 4
            
            Do While i <= lastRow
                Dim uid As String
                uid = ws.Cells(i, 1).Value
                
                ' Check for header row (UID ends with _0)
                If Len(uid) > 2 And Right(uid, 2) = "_0" Then
                    ' Create SubjectBlockInfo as array
                    Dim blockInfo(0 To 6) As Variant
                    blockInfo(SBI_SHEETNAME) = ws.name
                    blockInfo(SBI_HEADERROW) = i
                    
                    ' Parse UID for subject code and study period
                    Dim parts() As String
                    parts = Split(uid, "_")
                    
                    If UBound(parts) >= 2 Then
                        blockInfo(SBI_SUBJECTCODE) = parts(0)
                        blockInfo(SBI_STUDYPERIOD) = parts(1)
                        
                        ' Find Total row
                        blockInfo(SBI_TOTALROW) = FindTotalRow(ws, i)
                        
                        ' Find last subject row
                        blockInfo(SBI_LASTSUBJECTROW) = FindLastSubjectRow(ws, i)
                        
                        ' Calculate number of assessment rows
                        blockInfo(SBI_NUMASSESSMENTROWS) = blockInfo(SBI_TOTALROW) - blockInfo(SBI_HEADERROW) - 1
                        
                        ' Store subject block info
                        subjectBlocks.add blockInfo
                        
                        ' Jump to next subject
                        i = blockInfo(SBI_LASTSUBJECTROW)
                    End If
                End If
                
                i = i + 1
            Loop
        End If
    Next ws
    
    On Error GoTo 0
End Sub

'=============================================================================
' TEACHING DATA LOADING
'=============================================================================

'------------------------------------------------------------------------------
' LOAD TEACHING STREAM DATA
' Opens source file and loads teaching stream sheet
'------------------------------------------------------------------------------
Private Function LoadTeachingStreamData(sourcePath As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim sourceWB As Workbook
    Set sourceWB = Workbooks.Open(sourcePath, ReadOnly:=True, UpdateLinks:=False, Notify:=False)
    
    If sourceWB Is Nothing Then
        LoadTeachingStreamData = Empty
        Exit Function
    End If
    
    Dim ws As Worksheet
    Set ws = sourceWB.Sheets(TEACHING_STREAM_SHEET)
    
    If ws Is Nothing Then
        sourceWB.Close SaveChanges:=False
        LoadTeachingStreamData = Empty
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    If lastRow < 2 Then
        sourceWB.Close SaveChanges:=False
        LoadTeachingStreamData = Empty
        Exit Function
    End If
    
    ' Load columns B-G (Subject Code, Study Period, Lecturer, Status, Activity Code, Streams)
    LoadTeachingStreamData = ws.Range(ws.Cells(2, 2), ws.Cells(lastRow, 7)).Value
    
    sourceWB.Close SaveChanges:=False
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If Not sourceWB Is Nothing Then sourceWB.Close SaveChanges:=False
    On Error GoTo 0
    LoadTeachingStreamData = Empty
End Function

'=============================================================================
' LECTURER DATA UPDATE - SIMPLIFIED (NO ORPHAN TRACKING)
'=============================================================================

'------------------------------------------------------------------------------
' UPDATE ALL LECTURERS
' Refreshes columns L-O from teaching data, preserves P-S user edits
' Adds rows if needed (before Total row), no orphan tracking
'------------------------------------------------------------------------------
Private Function UpdateAllLecturers(wb As Workbook, teachingData As Variant, subjectBlocks As Collection) As Long
    On Error Resume Next
    
    UpdateAllLecturers = 0
    
    Dim i As Long
    
    For i = 1 To subjectBlocks.Count
        Dim blockInfo As Variant
        blockInfo = subjectBlocks(i)
        
        Dim ws As Worksheet
        Set ws = wb.Sheets(CStr(blockInfo(SBI_SHEETNAME)))
        ws.Unprotect
        
        ' Extract to String variables first (Mac-safe type casting)
        Dim blockSubjectCode As String
        Dim blockStudyPeriod As String
        blockSubjectCode = CStr(blockInfo(SBI_SUBJECTCODE))
        blockStudyPeriod = CStr(blockInfo(SBI_STUDYPERIOD))
        
        ' Get fresh lecturers for this subject (returns array)
        Dim freshLecturers As Variant
        freshLecturers = GetLecturersFromTeachingData(teachingData, blockSubjectCode, blockStudyPeriod)
        
        ' Check if we got any lecturers
        Dim lecturerCount As Long
        lecturerCount = 0
        
        If IsArray(freshLecturers) Then
            On Error Resume Next
            lecturerCount = UBound(freshLecturers, 1)
            On Error GoTo 0
        End If
        
        If lecturerCount > 0 Then
            ' Calculate available rows and Total row position
            Dim headerRow As Long
            Dim totalRow As Long
            Dim firstLecturerRow As Long
            Dim availableRows As Long
            
            headerRow = CLng(blockInfo(SBI_HEADERROW))
            totalRow = CLng(blockInfo(SBI_TOTALROW))
            firstLecturerRow = headerRow + 1
            availableRows = totalRow - firstLecturerRow  ' Rows between header and Total
            
            ' Check if we need to add rows
            If lecturerCount > availableRows Then
                Dim rowsToAdd As Long
                rowsToAdd = lecturerCount - availableRows
                
                ' Insert rows just before Total row
                Dim insertRow As Long
                Dim j As Long
                
                For j = 1 To rowsToAdd
                    insertRow = totalRow  ' Always insert at current Total row position
                    ws.Rows(insertRow).Insert Shift:=xlDown
                    
                    ' Copy format from the row above (last lecturer row)
                    ws.Rows(insertRow - 1).Copy
                    ws.Rows(insertRow).PasteSpecial xlPasteFormats
                    Application.CutCopyMode = False
                    
                    ' Clear contents but keep formatting
                    ws.Rows(insertRow).ClearContents
                    
                    ' Generate UID for new row
                    Dim newUID As String
                    Dim uidSuffix As Long
                    uidSuffix = (insertRow - headerRow - 1)
                    newUID = blockSubjectCode & "_" & blockStudyPeriod & "_" & uidSuffix
                    ws.Cells(insertRow, 1).Value = newUID
                    
                    ' Update Total row position (it shifted down)
                    totalRow = totalRow + 1
                Next j
                
                ' Recalculate available rows
                availableRows = totalRow - firstLecturerRow
            End If
            
            ' Clear columns L-O for all lecturer rows (between header and Total)
            Dim Row As Long
            For Row = firstLecturerRow To totalRow - 1
                ws.Cells(Row, 12).ClearContents  ' Column L: Lecturer Name
                ws.Cells(Row, 13).ClearContents  ' Column M: Status
                ws.Cells(Row, 14).ClearContents  ' Column N: Stream Number
                ws.Cells(Row, 15).ClearContents  ' Column O: Activity Code
                ' Columns P-S are NOT touched (preserve user edits)
            Next Row
            
            ' Write fresh lecturer data to columns L-O
            Dim currentRow As Long
            currentRow = firstLecturerRow
            
            Dim k As Long
            For k = 1 To lecturerCount
                If currentRow < totalRow Then
                    ws.Cells(currentRow, 12).Value = freshLecturers(k, 0)  ' Name
                    ws.Cells(currentRow, 13).Value = freshLecturers(k, 1)  ' Status
                    ws.Cells(currentRow, 14).Value = freshLecturers(k, 3)  ' Streams
                    ws.Cells(currentRow, 15).Value = freshLecturers(k, 2)  ' Activity Code
                    
                    ' Bold first lecturer (coordinator)
                    If currentRow = firstLecturerRow Then
                        ws.Cells(currentRow, 12).Font.Bold = True
                    End If
                    
                    currentRow = currentRow + 1
                End If
            Next k
            
            Call ApplyLecturerFormulas(ws, headerRow, totalRow)
            
            UpdateAllLecturers = UpdateAllLecturers + 1
        End If
        
        ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, _
                   AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
    Next i
    
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' APPLY LECTURER FORMULAS
' Applies formulas to columns Q & R for all lecturer rows
'------------------------------------------------------------------------------
Private Sub ApplyLecturerFormulas(ws As Worksheet, headerRow As Long, totalRow As Long)
    On Error Resume Next
    
    Dim firstLecturerRow As Long
    Dim lastLecturerRow As Long
    Dim numLecturers As Long
    
    firstLecturerRow = headerRow + 1
    lastLecturerRow = totalRow - 1
    numLecturers = lastLecturerRow - firstLecturerRow + 1
    
    If numLecturers <= 0 Then Exit Sub
    
    ' Build formula arrays for batch write (faster than row-by-row)
    Dim formulas As Variant
    ReDim formulas(1 To numLecturers, 1 To 2)
    
    Dim currentRow As Long
    Dim i As Long
    
    i = 1
    For currentRow = firstLecturerRow To lastLecturerRow
        ' Column Q (17): Allocated Marking
        ' =IF(M[row]="Continuing T&R",N[row]*VALUE(LEFT($Q$2,FIND(" ",$Q$2)-1)),"")
        formulas(i, 1) = "=IF(M" & currentRow & "=""Continuing T&R"",N" & currentRow & "*VALUE(LEFT($Q$2,FIND("" "",$Q$2)-1)),"""")"
        
        ' Column R (18): Marking Support Hours Available
        ' =IF(Q[row]="","",$J$[totalRow]*(P[row]/D[headerRow])-Q[row])
        formulas(i, 2) = "=IF(Q" & currentRow & "="""","""",$J$" & totalRow & "*(P" & currentRow & "/D" & headerRow & ")-Q" & currentRow & ")"
        
        i = i + 1
    Next currentRow
    
    ' Batch write both formulas at once (columns Q & R = 17-18)
    ws.Cells(firstLecturerRow, 17).Resize(numLecturers, 2).Formula = formulas
    
    On Error GoTo 0
End Sub

'=============================================================================
' HELPER FUNCTIONS
'=============================================================================

'------------------------------------------------------------------------------
' FIND TOTAL ROW
' Searches for "Total" in Column E
'------------------------------------------------------------------------------
Private Function FindTotalRow(ws As Worksheet, headerRow As Long) As Long
    Dim Row As Long
    For Row = headerRow + 1 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        Dim uid As String
        Dim cellE As String
        
        uid = ws.Cells(Row, 1).Value
        cellE = Trim(CStr(ws.Cells(Row, 5).Value))
        
        ' Found Total row
        If cellE = "Total" Then
            FindTotalRow = Row
            Exit Function
        End If
        
        ' Hit next subject header
        If Len(uid) > 2 And Right(uid, 2) = "_0" Then
            FindTotalRow = Row - 1
            Exit Function
        End If
    Next Row
    
    FindTotalRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
End Function

'------------------------------------------------------------------------------
' FIND LAST SUBJECT ROW
' Finds the actual last row of the subject block
'------------------------------------------------------------------------------
Private Function FindLastSubjectRow(ws As Worksheet, headerRow As Long) As Long
    Dim Row As Long
    For Row = headerRow + 1 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        Dim nextUID As String
        Dim cellB As String
        
        nextUID = ws.Cells(Row, 1).Value
        cellB = Trim(CStr(ws.Cells(Row, 2).Value))
        
        ' Next subject header found
        If Len(nextUID) > 2 And Right(nextUID, 2) = "_0" Then
            FindLastSubjectRow = Row - 1
            Exit Function
        End If
        
        ' Category header found
        If (cellB = "SUMMER" Or cellB = "WINTER" Or cellB = "SEMESTER 1" Or cellB = "SEMESTER 2") And _
           Trim(CStr(ws.Cells(Row, 1).Value)) = "" Then
            FindLastSubjectRow = Row - 1
            Exit Function
        End If
    Next Row
    
    FindLastSubjectRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
End Function

'------------------------------------------------------------------------------
' GET LECTURERS FROM TEACHING DATA
' Extracts lecturers for a specific subject from teaching stream data
' RETURNS: 2D Array instead of Collection (Mac-compatible)
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' COLLECTION KEY EXISTS
' Checks if a key exists in a collection
'------------------------------------------------------------------------------
Private Function CollectionKeyExists(col As Collection, key As String) As Boolean
    On Error Resume Next
    Dim temp As Variant
    temp = col(key)
    CollectionKeyExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function
