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
Public SubjectListErrored As Boolean
Public TeachingStreamErrored As Boolean

'===============================================================
' SECTION 1: ORCHESTRATION
'===============================================================

'---------------------------------------------------------------
' CheckVBAAccess
' Purpose: Verify "Trust access to VBA project object model" is enabled
' Returns: True if enabled, False otherwise (and shows a prompt)
'---------------------------------------------------------------
Public Function CheckVBAAccess(wb As Workbook, actionContext As String) As Boolean
    Dim vbaCount As Long
    On Error Resume Next
    vbaCount = wb.VBProject.VBComponents.count
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Excel requires access to the VBA Project to " & actionContext & "." & vbCrLf & vbCrLf & _
               "Please enable it from the menu:" & vbCrLf & _
               "Excel > Settings > Security > Trust access to VBA project object model" & vbCrLf & vbCrLf & _
               "Check the User Guide for detailed instructions.", vbCritical, "VBA Access Required"
        CheckVBAAccess = False
    Else
        On Error GoTo 0
        CheckVBAAccess = True
    End If
End Function

'---------------------------------------------------------------
' GenerateMarkingSupport
' Purpose: Main entry point — validates inputs, triggers both
'          workflows, monitors for completion, then runs all macros
' Called by: Dashboard "Generate" button
'---------------------------------------------------------------
Sub GenerateMarkingSupport()
    Dim dashboardSheet As Worksheet
    Dim yearValue As String
    Dim yearNum As Long
    Dim enrolmentTracker As String
    Dim teachingMatrix As String
    Dim emailValue As String
    Dim startDateTime As Date
    
    ' Enable silent mode to suppress MsgBox in sub-modules
    SilentMode = True
    
    If Not CheckVBAAccess(ThisWorkbook, "export calculation sheets") Then
        SilentMode = False
        Exit Sub
    End If
    
    ' Keep screen updating ON so status cells remain visible
    Application.ScreenUpdating = True
    
    ' Force automatic calc so formulas update in real time
    OriginalCalculationMode = Application.Calculation
    Application.Calculation = xlCalculationAutomatic
    
    StopMonitoring = False
    Set dashboardSheet = ThisWorkbook.Sheets("Dashboard")
    startDateTime = Now
    
    ' =========================================================================
    ' TIMING DISPLAY — record start date/time, set live elapsed-time formula
    ' =========================================================================
    dashboardSheet.Range("C15").Value = Format(Date, "yyyy-mm-dd")
    dashboardSheet.Range("C16").Value = Format(startDateTime, "HH:MM:SS")
    With dashboardSheet.Range("C17")
        .ClearContents
        .Interior.ColorIndex = xlNone
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
        .Formula = "=TEXT(NOW()-C16,""HH:MM:SS"")"
    End With
    
    dashboardSheet.Calculate
    DoEvents

    ClearStatusColumn dashboardSheet
    
    ' Validate year input (must be 2025 or later)
    On Error Resume Next
    yearNum = CLng(dashboardSheet.Range("C2").Value)
    If Err.Number <> 0 Or IsEmpty(dashboardSheet.Range("C2").Value) Or yearNum < 2025 Then
        On Error GoTo 0
        MsgBox "Please enter a valid year (2025 or later)!", vbExclamation, "Invalid Year"
        SilentMode = False
        Application.Calculation = OriginalCalculationMode
        Exit Sub
    End If
    On Error GoTo 0
    yearValue = CStr(yearNum)
    
    ' Read optional parameters
    enrolmentTracker = GetOptionalValue(dashboardSheet.Range("C3").Value)
    teachingMatrix = GetOptionalValue(dashboardSheet.Range("C5").Value)
    emailValue = GetOptionalValue(dashboardSheet.Range("C12").Value)
    
    ' =========================================================================
    ' TRIGGER WORKFLOWS — Power Automate flows for Subject List and Teaching Stream
    ' =========================================================================
    TriggerSubjectListWorkflow dashboardSheet, yearValue, enrolmentTracker, emailValue
    TriggerTeachingStreamWorkflow dashboardSheet, yearValue, teachingMatrix, emailValue
    
    MonitorAndExecute dashboardSheet, emailValue
End Sub

'---------------------------------------------------------------
' MonitorAndExecute
' Purpose: Poll Dashboard status cells (F2, F5) every 2 seconds
'          until both workflows report completion, then run macros.
'          Times out after 30 minutes.
' Called by: GenerateMarkingSupport
'---------------------------------------------------------------
Sub MonitorAndExecute(dashboardSheet As Worksheet, emailValue As String)
    Dim maxWaitMinutes As Integer
    Dim startTime As Double
    Dim lastRefreshTime As Double
    Dim refreshInterval As Double
    Dim subjectListComplete As Boolean
    Dim teachingStreamComplete As Boolean
    
    maxWaitMinutes = 30
    startTime = Timer
    lastRefreshTime = Timer
    refreshInterval = 15
    subjectListComplete = False
    teachingStreamComplete = False
    SubjectListErrored = False
    TeachingStreamErrored = False
    
    Application.StatusBar = "Monitoring workflows..."
    
    MsgBox "Workflows triggered! Monitoring F2 and F5 for completion..." & vbCrLf & _
           "Press ESC or run 'StopWorkflowMonitoring' to stop.", vbInformation
    
    ' =========================================================================
    ' POLLING LOOP — check every 5s, force cloud sync every 15s
    ' =========================================================================
    Do While Not StopMonitoring And Timer - startTime < (maxWaitMinutes * 60)
        DoEvents
        
        ' SharePoint sync needed so Office Scripts status updates appear locally
        If Timer - lastRefreshTime > refreshInterval Then
            ForceCloudSync dashboardSheet
            lastRefreshTime = Timer
        End If
        
        dashboardSheet.Calculate
        DoEvents
        
        ' Check if Subject List workflow completed (cell F2)
        If Not subjectListComplete And Not SubjectListErrored Then
            If CheckWorkflowComplete(dashboardSheet, "F2") Then
                subjectListComplete = True
                Application.StatusBar = "Subject List Refresh completed!"
                Debug.Print "Subject List Refresh Done"
                Application.Wait (Now + TimeValue("0:00:01"))
            ElseIf CheckWorkflowError(dashboardSheet, "F2") Then
                SubjectListErrored = True
                Application.StatusBar = "Subject List Refresh FAILED!"
                Debug.Print "Subject List Refresh Error"
            End If
        End If
        
        ' Check if Teaching Stream workflow completed (cell F5)
        If Not teachingStreamComplete And Not TeachingStreamErrored Then
            If CheckWorkflowComplete(dashboardSheet, "F5") Then
                teachingStreamComplete = True
                Application.StatusBar = "Teaching Stream Refresh completed!"
                Debug.Print "Teaching Stream Refresh Done"
                Application.Wait (Now + TimeValue("0:00:01"))
            ElseIf CheckWorkflowError(dashboardSheet, "F5") Then
                TeachingStreamErrored = True
                Application.StatusBar = "Teaching Stream Refresh FAILED!"
                Debug.Print "Teaching Stream Refresh Error"
            End If
        End If
        
        ' Exit if both workflows are resolved (complete or errored)
        If (subjectListComplete Or SubjectListErrored) And _
           (teachingStreamComplete Or TeachingStreamErrored) Then
            Application.StatusBar = "Both workflows resolved. Evaluating results..."
            Exit Do
        End If
        
        ' Periodic status update — show which workflows are still pending
        If Int(Timer) Mod 10 < 5 Then
            Dim statusMsg As String
            statusMsg = "Monitoring: "
            If Not subjectListComplete Then statusMsg = statusMsg & "Subject List "
            If Not teachingStreamComplete Then statusMsg = statusMsg & "Teaching Stream "
            If statusMsg <> "Monitoring: " Then
                Application.StatusBar = statusMsg & "in progress..."
            End If
        End If
        
        ' Wait 5 seconds before next check
        Application.Wait (Now + TimeValue("0:00:05"))
    Loop
    
    Application.StatusBar = False
    
    ' =========================================================================
    ' POST-MONITORING — run macros or handle error/timeout/cancellation
    ' =========================================================================
    Dim anyErrored As Boolean
    anyErrored = SubjectListErrored Or TeachingStreamErrored
    
    If Not StopMonitoring And Not anyErrored And subjectListComplete And teachingStreamComplete Then
        RunAllMacros dashboardSheet, emailValue
        FinaliseProcess dashboardSheet, emailValue
        
        MsgBox "All processes finished!" & vbCrLf & vbCrLf & _
               "- Subject List Refresh" & vbCrLf & _
               "- Teaching Stream Refresh" & vbCrLf & _
               "- Assessment Data HTML Query" & vbCrLf & _
               "- Assessment Data Parsing" & vbCrLf & _
               "- Calculation Sheets Generation", vbInformation
        
    ElseIf anyErrored Then
        ' Rule 4: one or both workflows reported "Error" — do not proceed with macros
        Dim errMsg As String
        errMsg = "One or more workflows encountered an error:" & vbCrLf
        If SubjectListErrored Then errMsg = errMsg & vbCrLf & "  × Subject List Refresh"
        If TeachingStreamErrored Then errMsg = errMsg & vbCrLf & "  × Teaching Stream Refresh"
        errMsg = errMsg & vbCrLf & vbCrLf & "Check the Office Script run history in Power Automate for details."
        MsgBox errMsg, vbCritical, "Workflow Error"
        
    ElseIf StopMonitoring Then
        MsgBox "Monitoring stopped by user.", vbInformation
        
    Else
        ' Timeout — proceed with whatever data is available
        MsgBox "Timeout reached. Running macros with current data...", vbExclamation
        RunAllMacros dashboardSheet, emailValue
        FinaliseProcess dashboardSheet, emailValue
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
Sub ForceCloudSync(dashboardSheet As Worksheet)
    On Error Resume Next
    
    ' Suppress save/refresh dialogue boxes
    Application.DisplayAlerts = False
    
    ThisWorkbook.Save
    ThisWorkbook.RefreshAll
    Application.Calculate
    Application.CalculateFullRebuild
    dashboardSheet.Calculate
    dashboardSheet.Range("F2:F6").Calculate
    
    ' Read values to flush any SharePoint cache
    Dim tempVal As Variant
    tempVal = dashboardSheet.Range("F2").Value
    tempVal = dashboardSheet.Range("F5").Value
    
    ' Toggle screen updating to force visual refresh
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True
    DoEvents
    
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
Function CheckWorkflowComplete(dashboardSheet As Worksheet, cellAddress As String) As Boolean
    Dim cellValue As String
    
    DoEvents
    cellValue = Trim(UCase(CStr(dashboardSheet.Range(cellAddress).Value)))
    
    If cellValue = "DONE" Or cellValue = "COMPLETE" Or cellValue = "FINISHED" Then
        With dashboardSheet.Range(cellAddress)
            .Interior.Color = RGB(146, 208, 80)
            .Value = "Complete" ' Explicitly overwrite to indicate monitored success
        End With
        DoEvents
        CheckWorkflowComplete = True
    Else
        CheckWorkflowComplete = False
    End If
End Function

'---------------------------------------------------------------
' CheckWorkflowError
' Purpose: Rule 4 — detect when Office Script writes "Error" to
'          a Dashboard status cell. Colours the cell red.
' Called by: MonitorAndExecute
' Returns: True if the workflow reported an error
'---------------------------------------------------------------
Function CheckWorkflowError(dashboardSheet As Worksheet, cellAddress As String) As Boolean
    Dim cellValue As String
    
    DoEvents
    cellValue = Trim(UCase(CStr(dashboardSheet.Range(cellAddress).Value)))
    
    If cellValue = "ERROR" Then
        With dashboardSheet.Range(cellAddress)
            .Interior.Color = RGB(255, 0, 0)  ' Red — matches SetProgress "Error" convention
            .Font.Color = RGB(255, 255, 255)  ' White text for contrast
        End With
        DoEvents
        CheckWorkflowError = True
    Else
        CheckWorkflowError = False
    End If
End Function

'---------------------------------------------------------------
' RunAllMacros
' Purpose: Execute GenerateSubjectQueries, ParseAssessmentData,
'          and GenerateCalculationSheets in sequence, updating
'          Dashboard status cells along the way
' Called by: MonitorAndExecute
'---------------------------------------------------------------
Sub RunAllMacros(dashboardSheet As Worksheet, emailValue As String)
    
    ' MACRO 1: GenerateSubjectQueries
    On Error Resume Next
    GenerateSubjectQueries
    On Error GoTo 0
    
    DoEvents
    
    ' MACRO 2: ParseAssessmentData
    On Error Resume Next
    ParseAssessmentData
    On Error GoTo 0
    
    DoEvents
    
    ' MACRO 3: GenerateCalculationSheets
    On Error Resume Next
    GenerateCalculationSheets
    Application.Calculation = xlCalculationAutomatic
    On Error GoTo 0
    
    DoEvents
End Sub

'---------------------------------------------------------------
' FinaliseProcess
' Purpose: Freeze the elapsed time display and send the
'          completion email notification
' Called by: MonitorAndExecute
'---------------------------------------------------------------
Sub FinaliseProcess(dashboardSheet As Worksheet, emailValue As String)
    ' Freeze elapsed time — replace live formula with its current value
    With dashboardSheet.Range("C17")
        .Value = .Value
    End With
    
    SendCompletionNotification dashboardSheet, emailValue
End Sub

'---------------------------------------------------------------
' SendCompletionNotification
' Purpose: Send a completion email via Outlook with a link
'          to the SharePoint output folder
' Called by: FinaliseProcess
'---------------------------------------------------------------
Sub SendCompletionNotification(dashboardSheet As Worksheet, emailValue As String)
    On Error Resume Next
    
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim yearValue As String
    
    If emailValue = "" Then Exit Sub
    
    yearValue = dashboardSheet.Range("C2").Value
    Set OutlookApp = CreateObject("Outlook.Application")
    
    If Not OutlookApp Is Nothing Then
        Set OutlookMail = OutlookApp.CreateItem(0)
        
        With OutlookMail
            .To = emailValue
            .subject = yearValue & "_M&M_Marking Admin Support Calculations Complete"
            .HTMLBody = "<p>Hello,</p>" & _
                       "<p>" & yearValue & "_M&M_Marking Admin Support Calculations has been generated successfully!</p>" & _
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
    
    Dim dashboardSheet As Worksheet
    Set dashboardSheet = ThisWorkbook.Sheets("Dashboard")
    
    With dashboardSheet.Range("C17")
        .Value = "Stopped"
        .Font.Color = RGB(255, 0, 0)
        .Interior.ColorIndex = xlNone
        .Font.Bold = True
    End With
    
    Application.StatusBar = False
    MsgBox "Monitoring stopped.", vbInformation
    
    SilentMode = False
    Application.Calculation = OriginalCalculationMode
    
    ' Hard stop all VBA execution
    End
End Sub

'---------------------------------------------------------------
' ResetStatus
' Purpose: Clear all Dashboard status cells and restore defaults
'---------------------------------------------------------------
Sub ResetStatus()
    Dim dashboardSheet As Worksheet
    Set dashboardSheet = ThisWorkbook.Sheets("Dashboard")
    
    ClearStatusColumn dashboardSheet
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
Sub ClearStatusColumn(dashboardSheet As Worksheet)
    With dashboardSheet.Range("F2:F6")
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