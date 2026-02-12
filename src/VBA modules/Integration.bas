'==============================================================================
' VBA AUTOMATION SYSTEM WITH MONITORING
' Purpose: Trigger workflows, monitor completion, run macros sequentially
'==============================================================================

'------------------------------------------------------------------------------
' GLOBAL VARIABLES
'------------------------------------------------------------------------------
Public StopMonitoring As Boolean
Public OriginalCalculationMode As XlCalculation

'------------------------------------------------------------------------------
' MAIN ENTRY POINT - TriggerPowerAutomate
' Triggers workflows, monitors for completion, then runs all macros
'------------------------------------------------------------------------------
Sub TriggerPowerAutomate()
    Dim ws As Worksheet
    Dim yearValue As String
    Dim yearNum As Long
    Dim enrolmentTracker As String
    Dim teachingMatrix As String
    Dim emailValue As String
    Dim startDateTime As Date
    
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
        Application.Calculation = OriginalCalculationMode
        Exit Sub
    End If
    On Error GoTo 0
    yearValue = CStr(yearNum)
    
    ' Read optional parameters
    enrolmentTracker = GetOptionalValue(ws.Range("C3").Value)
    teachingMatrix = GetOptionalValue(ws.Range("C5").Value)
    emailValue = GetOptionalValue(ws.Range("C12").Value)
    
    ' START WORKFLOWS
    StartWorkflow1 ws, yearValue, enrolmentTracker, emailValue
    StartWorkflow3 ws, yearValue, teachingMatrix, emailValue
    
    ' Start monitoring
    MonitorAndExecute ws, emailValue
End Sub

'------------------------------------------------------------------------------
' WORKFLOW 1 - Enrolment Tracker Processing
'------------------------------------------------------------------------------
Sub StartWorkflow1(ws As Worksheet, yearValue As String, enrolmentTracker As String, emailValue As String)
    Dim url As String
    Dim jsonData As String
    Dim result As String
    
    url = "https://default0e5bf3cf1ff446b7917652c538c22a.4d.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/98c68237600941c8a349156641a1cc54/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ZuW5DcOAmY9csBoZjhpqTYYo0LjeUS3qxAYpN6S5sKg"
    
    With ws.Range("F2")
        .Value = "Running..."
        .Interior.Color = RGB(255, 192, 0)
    End With
    DoEvents
    
    jsonData = "{" & Chr(34) & "year" & Chr(34) & ":" & yearValue & "," & _
               Chr(34) & "enrolmentTrackerFilename" & Chr(34) & ":" & Chr(34) & EscapeJSON(enrolmentTracker) & Chr(34) & "," & _
               Chr(34) & "email" & Chr(34) & ":" & Chr(34) & EscapeJSON(emailValue) & Chr(34) & "}"
    
    #If Mac Then
        result = SendRequestMac(url, jsonData)
    #Else
        result = SendRequestWindows(url, jsonData)
    #End If
End Sub

'------------------------------------------------------------------------------
' WORKFLOW 3 - Teaching Matrix Processing
'------------------------------------------------------------------------------
Sub StartWorkflow3(ws As Worksheet, yearValue As String, teachingMatrix As String, emailValue As String)
    Dim url As String
    Dim jsonData As String
    Dim result As String
    
    url = "https://default0e5bf3cf1ff446b7917652c538c22a.4d.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/7f198e614c734715bc0153d818de1ef7/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=5uUuhFHyiL37O_ajy-6t2r65nFqc7NA_oJDhYmYFT9g"
    
    With ws.Range("F5")
        .Value = "Running..."
        .Interior.Color = RGB(255, 192, 0)
    End With
    DoEvents
    
    jsonData = "{" & Chr(34) & "year" & Chr(34) & ":" & yearValue & "," & _
               Chr(34) & "teachingMatrixFilename" & Chr(34) & ":" & Chr(34) & EscapeJSON(teachingMatrix) & Chr(34) & "," & _
               Chr(34) & "email" & Chr(34) & ":" & Chr(34) & EscapeJSON(emailValue) & Chr(34) & "}"
    
    #If Mac Then
        result = SendRequestMac(url, jsonData)
    #Else
        result = SendRequestWindows(url, jsonData)
    #End If
End Sub

'------------------------------------------------------------------------------
' MONITOR AND EXECUTE - Wait for workflows, then run macros
'------------------------------------------------------------------------------
Sub MonitorAndExecute(ws As Worksheet, emailValue As String)
    Dim maxWaitMinutes As Integer
    Dim startTime As Double
    Dim lastRefreshTime As Double
    Dim refreshInterval As Double
    Dim workflow1Complete As Boolean
    Dim workflow3Complete As Boolean
    
    maxWaitMinutes = 30
    startTime = Timer
    lastRefreshTime = Timer
    refreshInterval = 5
    workflow1Complete = False
    workflow3Complete = False
    
    Application.StatusBar = "Monitoring workflows..."
    
    MsgBox "Workflows triggered! Monitoring F2 and F5 for completion..." & vbCrLf & _
           "Press ESC or run 'StopWorkflowMonitoring' to stop.", vbInformation
    
    ' MONITORING LOOP
    Do While Not StopMonitoring And Timer - startTime < (maxWaitMinutes * 60)
        DoEvents
        
        ' Force cloud sync every 3 seconds
        If Timer - lastRefreshTime > refreshInterval Then
            ForceCloudSync ws
            lastRefreshTime = Timer
        End If
        
        ws.Calculate
        DoEvents
        
        ' Check Workflow 1 completion (F2)
        If Not workflow1Complete And CheckWorkflowComplete(ws, "F2") Then
            workflow1Complete = True
            Application.StatusBar = "Workflow 1 completed!"
            Debug.Print "Workflow 1 Done"
            Application.Wait (Now + TimeValue("0:00:01"))
        End If
        
        ' Check Workflow 3 completion (F5)
        If Not workflow3Complete And CheckWorkflowComplete(ws, "F5") Then
            workflow3Complete = True
            Application.StatusBar = "Workflow 3 completed!"
            Debug.Print "Workflow 3 Done"
            Application.Wait (Now + TimeValue("0:00:01"))
        End If
        
        ' If both complete, exit monitoring loop
        If workflow1Complete And workflow3Complete Then
            Application.StatusBar = "Both workflows complete! Running macros..."
            Exit Do
        End If
        
        ' Periodic status update
        If Int(Timer) Mod 10 < 2 Then
            Dim statusMsg As String
            statusMsg = "Monitoring: "
            If Not workflow1Complete Then statusMsg = statusMsg & "W1 "
            If Not workflow3Complete Then statusMsg = statusMsg & "W3 "
            If statusMsg <> "Monitoring: " Then
                Application.StatusBar = statusMsg & "in progress..."
            End If
        End If
        
        ' Wait 2 seconds before next check
        Application.Wait (Now + TimeValue("0:00:02"))
    Loop
    
    Application.StatusBar = False
    
    ' POST-MONITORING EXECUTION
    If Not StopMonitoring And workflow1Complete And workflow3Complete Then
        ' Both workflows completed - run all macros
        RunAllMacros ws, emailValue
        FinalizeProcess ws, emailValue
        
        MsgBox "All processes completed!" & vbCrLf & vbCrLf & _
               "? Workflow 1: Complete" & vbCrLf & _
               "? GenerateSubjectQueries: Complete" & vbCrLf & _
               "? ParseAssessmentData: Complete" & vbCrLf & _
               "? Workflow 3: Complete" & vbCrLf & _
               "? GenerateCalculationSheets: Complete", vbInformation
        
    ElseIf StopMonitoring Then
        MsgBox "Monitoring stopped by user.", vbInformation
        
    Else
        ' Timeout - run macros anyway
        MsgBox "Timeout reached. Running macros with current data...", vbExclamation
        RunAllMacros ws, emailValue
        FinalizeProcess ws, emailValue
    End If
    
    Application.Calculation = OriginalCalculationMode
End Sub

'------------------------------------------------------------------------------
' FORCE CLOUD SYNC - Refresh data from SharePoint
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' CHECK WORKFLOW COMPLETE - Detect completion and update status
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' RUN ALL MACROS - Sequential execution
'------------------------------------------------------------------------------
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
    
    With ws.Range("F3")
        .Value = "Complete"
        .Interior.Color = RGB(146, 208, 80)
    End With
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

'------------------------------------------------------------------------------
' FINALIZE PROCESS
'------------------------------------------------------------------------------
Sub FinalizeProcess(ws As Worksheet, emailValue As String)
    ' Freeze elapsed time
    With ws.Range("C17")
        .Value = .Value
    End With
    
    ' Send email
    SendCompletionNotification ws, emailValue
End Sub

'------------------------------------------------------------------------------
' EMAIL NOTIFICATION
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' STOP WORKFLOW MONITORING
'------------------------------------------------------------------------------
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
    
    Application.Calculation = OriginalCalculationMode
End Sub

'------------------------------------------------------------------------------
' RESET STATUS
'------------------------------------------------------------------------------
Sub ResetStatus()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ClearStatusColumn ws
    StopMonitoring = False
    Application.StatusBar = False
    Application.Calculation = OriginalCalculationMode
    
    MsgBox "All status cells reset.", vbInformation
End Sub

'------------------------------------------------------------------------------
' HELPER - Clear Status Column
'------------------------------------------------------------------------------
Sub ClearStatusColumn(ws As Worksheet)
    With ws.Range("F2:F6")
        .Value = ""
        .Interior.ColorIndex = xlNone
    End With
End Sub

'==============================================================================
' HTTP REQUEST FUNCTIONS
'==============================================================================

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

'==============================================================================
' UTILITY FUNCTIONS
'==============================================================================

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