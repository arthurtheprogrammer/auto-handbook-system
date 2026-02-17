'===============================================================
' Module: Integration
' Purpose: Orchestrate the full marking support generation:
'          trigger Power Automate workflows, monitor completion,
'          then run HTMLQuery, AssessmentData, and CalcSheets
' Main Entry: GenerateMarkingSupport() - triggered by Dashboard button
' Output: Status updates on Dashboard, completion email sent
' Author: Arthur Chen
' Repository: github.com/arthurtheprogrammer/auto-handbook-system
' Dependencies:
'   - Dashboard sheet (year, file paths, status cells)
'   - SubjectListRefresh.bas, TeachingStreamRefresh.bas (workflows)
'   - HTMLQuery.bas, AssessmentData.bas, CalculationSheets.bas
'===============================================================

Public StopMonitoring As Boolean
Public OriginalCalculationMode As XlCalculation
Public SilentMode As Boolean

'===============================================================
' SECTION 1: ORCHESTRATION
'===============================================================

'---------------------------------------------------------------
' GenerateMarkingSupport
' Purpose: Main entry point — validates inputs, triggers both
'          workflows, monitors for completion, then runs all macros
' Called by: Dashboard "Generate" button
'---------------------------------------------------------------
Sub GenerateMarkingSupport()
    Dim ws As Worksheet
    Dim yearValue As String
    Dim yearNum As Long
    Dim enrolmentTracker As String
    Dim teachingMatrix As String
    Dim emailValue As String
    Dim startDateTime As Date
    
    ' Enable silent mode to suppress MsgBox in sub-modules
    SilentMode = True
    
    ' Always keep screen updating ON
    Application.ScreenUpdating = True
    
    ' Force Automatic calculation
    OriginalCalculationMode = Application.Calculation
    Application.Calculation = xlCalculationAutomatic
    
    StopMonitoring = False
    Set ws = ThisWorkbook.Sheets("Dashboard")
    startDateTime = Now
    
    ' Update Dashboard timing cells
    ws.Range("C15").Value = Format(Date, "yyyy-mm-dd")
    ws.Range("C16").Value = Format(startDateTime, "HH:MM:SS")
    With ws.Range("C17")
        .ClearContents
        .Interior.ColorIndex = xlNone
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
        .Formula = "=TEXT(NOW()-C16,""HH:MM:SS"")"
    End With
    
    ws.Calculate
    DoEvents

    ClearStatusColumn ws
    
    ' Validate year input
    On Error Resume Next
    yearNum = CLng(ws.Range("C2").Value)
    If Err.Number <> 0 Or IsEmpty(ws.Range("C2").Value) Or yearNum < 2025 Then
        On Error GoTo 0
        MsgBox "Please enter a valid year (2025 or later)!", vbExclamation, "Invalid Year"
        SilentMode = False
        Application.Calculation = OriginalCalculationMode
        Exit Sub
    End If
    On Error GoTo 0
    yearValue = CStr(yearNum)
    
    ' Read optional parameters
    enrolmentTracker = GetOptionalValue(ws.Range("C3").Value)
    teachingMatrix = GetOptionalValue(ws.Range("C5").Value)
    emailValue = GetOptionalValue(ws.Range("C12").Value)
    
    ' START WORKFLOWS (calling extracted modules)
    TriggerSubjectListWorkflow ws, yearValue, enrolmentTracker, emailValue
    TriggerTeachingStreamWorkflow ws, yearValue, teachingMatrix, emailValue
    
    ' Start monitoring
    MonitorAndExecute ws, emailValue
End Sub

'---------------------------------------------------------------
' MonitorAndExecute
' Purpose: Poll Dashboard status cells (F2, F5) every 2 seconds
'          until both workflows report completion, then run macros.
'          Times out after 30 minutes.
' Called by: GenerateMarkingSupport
'---------------------------------------------------------------
Sub MonitorAndExecute(ws As Worksheet, emailValue As String)
    Dim maxWaitMinutes As Integer
    Dim startTime As Double
    Dim lastRefreshTime As Double
    Dim refreshInterval As Double
    Dim subjectListComplete As Boolean
    Dim teachingStreamComplete As Boolean
    
    maxWaitMinutes = 30
    startTime = Timer
    lastRefreshTime = Timer
    refreshInterval = 5
    subjectListComplete = False
    teachingStreamComplete = False
    
    Application.StatusBar = "Monitoring workflows..."
    
    MsgBox "Workflows triggered! Monitoring F2 and F5 for completion..." & vbCrLf & _
           "Press ESC or run 'StopWorkflowMonitoring' to stop.", vbInformation
    
    ' MONITORING LOOP
    Do While Not StopMonitoring And Timer - startTime < (maxWaitMinutes * 60)
        DoEvents
        
        ' Force cloud sync every 5 seconds
        If Timer - lastRefreshTime > refreshInterval Then
            ForceCloudSync ws
            lastRefreshTime = Timer
        End If
        
        ws.Calculate
        DoEvents
        
        ' Check Subject List Refresh completion (F2)
        If Not subjectListComplete And CheckWorkflowComplete(ws, "F2") Then
            subjectListComplete = True
            Application.StatusBar = "Subject List Refresh completed!"
            Debug.Print "Subject List Refresh Done"
            Application.Wait (Now + TimeValue("0:00:01"))
        End If
        
        ' Check Teaching Stream Refresh completion (F5)
        If Not teachingStreamComplete And CheckWorkflowComplete(ws, "F5") Then
            teachingStreamComplete = True
            Application.StatusBar = "Teaching Stream Refresh completed!"
            Debug.Print "Teaching Stream Refresh Done"
            Application.Wait (Now + TimeValue("0:00:01"))
        End If
        
        ' If both complete, exit monitoring loop
        If subjectListComplete And teachingStreamComplete Then
            Application.StatusBar = "Both workflows complete! Running macros..."
            Exit Do
        End If
        
        ' Periodic status update
        If Int(Timer) Mod 10 < 2 Then
            Dim statusMsg As String
            statusMsg = "Monitoring: "
            If Not subjectListComplete Then statusMsg = statusMsg & "Subject List "
            If Not teachingStreamComplete Then statusMsg = statusMsg & "Teaching Stream "
            If statusMsg <> "Monitoring: " Then
                Application.StatusBar = statusMsg & "in progress..."
            End If
        End If
        
        ' Wait 2 seconds before next check
        Application.Wait (Now + TimeValue("0:00:02"))
    Loop
    
    Application.StatusBar = False
    
    ' POST-MONITORING EXECUTION
    If Not StopMonitoring And subjectListComplete And teachingStreamComplete Then
        ' Both workflows completed - run all macros
        RunAllMacros ws, emailValue
        FinalizeProcess ws, emailValue
        
        MsgBox "All processes completed!" & vbCrLf & vbCrLf & _
               "? Subject List Refresh: Complete" & vbCrLf & _
               "? GenerateSubjectQueries: Complete" & vbCrLf & _
               "? ParseAssessmentData: Complete" & vbCrLf & _
               "? Teaching Stream Refresh: Complete" & vbCrLf & _
               "? GenerateCalculationSheets: Complete", vbInformation
        
    ElseIf StopMonitoring Then
        MsgBox "Monitoring stopped by user.", vbInformation
        
    Else
        ' Timeout - run macros anyway
        MsgBox "Timeout reached. Running macros with current data...", vbExclamation
        RunAllMacros ws, emailValue
        FinalizeProcess ws, emailValue
    End If
    
    ' Disable silent mode and restore calculation
    SilentMode = False
    Application.Calculation = OriginalCalculationMode
End Sub

'---------------------------------------------------------------
' ForceCloudSync
' Purpose: Force a save/refresh/recalc cycle to pick up any
'          SharePoint file changes between polling intervals
' Called by: MonitorAndExecute
'---------------------------------------------------------------
Sub ForceCloudSync(ws As Worksheet)
    On Error Resume Next
    
    ' Disable alerts to bypass dialog boxes
    Application.DisplayAlerts = False
    
    ThisWorkbook.Save
    ThisWorkbook.RefreshAll
    Application.Calculate
    Application.CalculateFullRebuild
    ws.Calculate
    ws.Range("F2:F6").Calculate
    
    ' Force cache refresh by reading values
    Dim tempVal As Variant
    tempVal = ws.Range("F2").Value
    tempVal = ws.Range("F5").Value
    
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True
    DoEvents
    
    ' Re-enable alerts
    Application.DisplayAlerts = True
    
    On Error GoTo 0
End Sub

'---------------------------------------------------------------
' CheckWorkflowComplete
' Purpose: Check whether a Dashboard status cell contains
'          DONE/COMPLETE/FINISHED and colour it green
' Called by: MonitorAndExecute
' Returns: True if the workflow is complete
'---------------------------------------------------------------
Function CheckWorkflowComplete(ws As Worksheet, cellAddress As String) As Boolean
    Dim cellValue As String
    
    DoEvents
    cellValue = Trim(UCase(CStr(ws.Range(cellAddress).Value)))
    
    If cellValue = "DONE" Or cellValue = "COMPLETE" Or cellValue = "FINISHED" Then
        With ws.Range(cellAddress)
            .Interior.Color = RGB(146, 208, 80)
            .Value = "Complete"
        End With
        DoEvents
        CheckWorkflowComplete = True
    Else
        CheckWorkflowComplete = False
    End If
End Function

'---------------------------------------------------------------
' RunAllMacros
' Purpose: Execute GenerateSubjectQueries, ParseAssessmentData,
'          and GenerateCalculationSheets in sequence, updating
'          Dashboard status cells along the way
' Called by: MonitorAndExecute
'---------------------------------------------------------------
Sub RunAllMacros(ws As Worksheet, emailValue As String)
    
    ' Macro 1: GenerateSubjectQueries
    With ws.Range("F3")
        .Value = "Running..."
        .Interior.Color = RGB(255, 192, 0)
    End With
    DoEvents
    
    On Error Resume Next
    GenerateSubjectQueries
    On Error GoTo 0
    
    #If Mac Then
        With ws.Range("F3")
            .Value = "Skipped"
            .Interior.Color = RGB(191, 191, 191) ' Grey
        End With
    #Else
        With ws.Range("F3")
            .Value = "Complete"
            .Interior.Color = RGB(146, 208, 80)
        End With
    #End If
    DoEvents
    
    ' Macro 2: ParseAssessmentData
    With ws.Range("F4")
        .Value = "Running..."
        .Interior.Color = RGB(255, 192, 0)
    End With
    DoEvents
    
    On Error Resume Next
    ParseAssessmentData
    On Error GoTo 0
    
    With ws.Range("F4")
        .Value = "Complete"
        .Interior.Color = RGB(146, 208, 80)
    End With
    DoEvents
    
    ' Macro 3: GenerateCalculationSheets
    With ws.Range("F6")
        .Value = "Running..."
        .Interior.Color = RGB(255, 192, 0)
    End With
    DoEvents
    
    On Error Resume Next
    GenerateCalculationSheets
    Application.Calculation = xlCalculationAutomatic
    On Error GoTo 0
    
    With ws.Range("F6")
        .Value = "Complete"
        .Interior.Color = RGB(146, 208, 80)
    End With
    DoEvents
End Sub

'---------------------------------------------------------------
' FinalizeProcess
' Purpose: Freeze the elapsed time display and send the
'          completion email notification
' Called by: MonitorAndExecute
'---------------------------------------------------------------
Sub FinalizeProcess(ws As Worksheet, emailValue As String)
    ' Freeze elapsed time
    With ws.Range("C17")
        .Value = .Value
    End With
    
    ' Send email
    SendCompletionNotification ws, emailValue
End Sub

'---------------------------------------------------------------
' SendCompletionNotification
' Purpose: Send a completion email via Outlook with a link
'          to the SharePoint output folder
' Called by: FinalizeProcess
'---------------------------------------------------------------
Sub SendCompletionNotification(ws As Worksheet, emailValue As String)
    On Error Resume Next
    
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim yearValue As String
    
    If emailValue = "" Then Exit Sub
    
    yearValue = ws.Range("C2").Value
    Set OutlookApp = CreateObject("Outlook.Application")
    
    If Not OutlookApp Is Nothing Then
        Set OutlookMail = OutlookApp.CreateItem(0)
        
        With OutlookMail
            .To = emailValue
            .subject = yearValue & " Marking & Admin Support Calculations Complete"
            .HTMLBody = "<p>Hello,</p>" & _
                       "<p>" & yearValue & " Marking & Admin Support Calculations has been generated successfully!</p>" & _
                       "<p>Please navigate to the <a href='https://unimelbcloud.sharepoint.com/:f:/r/teams/DepartmentofManagementMarketing-DepartmentOperations/Shared%20Documents/TEACHING%20SUPPORT/Handbook%20(Course%20%26%20Subject%20Changes)/Auto%20Handbook%20System?csf=1&web=1&e=kKYTrQ'>Auto Handbook System</a> folder on SharePoint to find the Excel spreadsheet.</p>" & _
                       "<p>Best regards,<br>Automated Handbook Data System</p>"
            .Send
        End With
    End If
    
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    
    On Error GoTo 0
End Sub

'---------------------------------------------------------------
' StopWorkflowMonitoring
' Purpose: Allow the user to halt the monitoring loop early
'          (e.g. via the macro menu or ESC)
'---------------------------------------------------------------
Sub StopWorkflowMonitoring()
    StopMonitoring = True
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    With ws.Range("C17")
        .Value = "Stopped"
        .Font.Color = RGB(255, 0, 0)
        .Interior.ColorIndex = xlNone
        .Font.Bold = True
    End With
    
    Application.StatusBar = False
    MsgBox "Monitoring stopped.", vbInformation
    
    SilentMode = False
    Application.Calculation = OriginalCalculationMode
End Sub

'---------------------------------------------------------------
' ResetStatus
' Purpose: Clear all Dashboard status cells and restore defaults
'---------------------------------------------------------------
Sub ResetStatus()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ClearStatusColumn ws
    StopMonitoring = False
    SilentMode = False
    Application.StatusBar = False
    Application.Calculation = OriginalCalculationMode
    
    MsgBox "All status cells reset.", vbInformation
End Sub

'---------------------------------------------------------------
' ClearStatusColumn
' Purpose: Clear cells F2:F6 on the Dashboard
'---------------------------------------------------------------
Sub ClearStatusColumn(ws As Worksheet)
    With ws.Range("F2:F6")
        .Value = ""
        .Interior.ColorIndex = xlNone
    End With
End Sub

'===============================================================
' SECTION 2: HTTP REQUESTS
'===============================================================

'---------------------------------------------------------------
' SendRequestMac
' Purpose: Send an HTTP POST via AppleScript/curl (Mac only)
' Returns: Response text, or "ERROR"
'---------------------------------------------------------------

Function SendRequestMac(url As String, jsonData As String) As String
    Dim scriptCode As String
    Dim result As String
    
    jsonData = Replace(jsonData, "\", "\\")
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
' Purpose: Send an HTTP POST via MSXML2 (Windows only)
' Returns: Response text, or "ERROR"
'---------------------------------------------------------------
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

'===============================================================
' SECTION 3: UTILITY FUNCTIONS
'===============================================================

'---------------------------------------------------------------
' GetOptionalValue
' Purpose: Safely read a Dashboard cell to a string, returning
'          empty string if the cell is blank, null, or errored
'---------------------------------------------------------------

Function GetOptionalValue(cellValue As Variant) As String
    On Error Resume Next
    
    If IsEmpty(cellValue) Or IsError(cellValue) Or IsNull(cellValue) Then
        GetOptionalValue = ""
    ElseIf Trim(CStr(cellValue)) = "" Then
        GetOptionalValue = ""
    Else
        GetOptionalValue = Trim(CStr(cellValue))
    End If
    
    On Error GoTo 0
End Function

'---------------------------------------------------------------
' EscapeJSON
' Purpose: Escape special characters for safe inclusion in a
'          JSON string payload sent to Power Automate
'---------------------------------------------------------------
Function EscapeJSON(text As String) As String
    Dim result As String
    result = text
    
    result = Replace(result, "\", "\\")
    result = Replace(result, Chr(34), "\" & Chr(34))
    result = Replace(result, vbCr, "")
    result = Replace(result, vbLf, "")
    result = Replace(result, vbTab, " ")
    
    EscapeJSON = result
End Function