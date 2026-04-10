'===============================================================
' Module: SubjectListRefresh
' Purpose: Trigger a Power Automate workflow to refresh the
'          subject list from the Enrolment Tracker on SharePoint
' Main Entry: RefreshSubjectList() - standalone or via Integration
' Author: Arthur Chen
' Repository: github.com/arthurtheprogrammer/auto-handbook-system
' Dependencies:
'   - Dashboard sheet (year, enrolment tracker filename, email)
'   - Integration.bas (EscapeJSON, GetOptionalValue,
'     SendRequestMac/Windows)
'===============================================================

'---------------------------------------------------------------
' RefreshSubjectList
' Purpose: Validate inputs and trigger the subject list workflow
'          as a standalone action from the VBA editor or a button
'---------------------------------------------------------------
Sub RefreshSubjectList()
    Dim ws As Worksheet
    Dim yearValue As String
    Dim yearNum As Long
    Dim enrolmentTracker As String
    Dim emailValue As String
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' Validate year input
    On Error Resume Next
    yearNum = CLng(ws.Range("C2").Value)
    If Err.Number <> 0 Or IsEmpty(ws.Range("C2").Value) Or yearNum < 2025 Then
        On Error GoTo 0
        MsgBox "Please enter a valid year (2025 or later)!", vbExclamation, "Invalid Year"
        Exit Sub
    End If
    On Error GoTo 0
    yearValue = CStr(yearNum)
    
    ' Read parameters
    enrolmentTracker = GetOptionalValue(ws.Range("C3").Value)
    emailValue = GetOptionalValue(ws.Range("C12").Value)
    
    ' Trigger the workflow and show result
    If TriggerSubjectListWorkflow(ws, yearValue, enrolmentTracker, emailValue) Then
        MsgBox "Subject List Refresh completed successfully!", vbInformation
    Else
        MsgBox "Subject List Refresh failed. Check the status cell on the Dashboard.", vbCritical, "Workflow Error"
    End If
End Sub

'---------------------------------------------------------------
' TriggerSubjectListWorkflow
' Purpose: Build JSON payload and send HTTP POST to the subject
'          list Power Automate endpoint, then read the sync
'          HTTP response to determine success or failure.
' Returns: True if the flow reported success, False otherwise
' Called by: Integration.GenerateMarkingSupport, RefreshSubjectList
'---------------------------------------------------------------
Function TriggerSubjectListWorkflow(ws As Worksheet, yearValue As String, enrolmentTracker As String, emailValue As String) As Boolean
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
    
    ' Parse the synchronous HTTP response from the PA flow.
    ' The flow returns a 200 with the Office Script body on success,
    ' or a non-200 / "ERROR" sentinel on failure.
    If result = "ERROR" Or result = "" Then
        With ws.Range("F2")
            .Value = "Error"
            .Interior.Color = RGB(255, 0, 0)
            .Font.Color = RGB(255, 255, 255)
        End With
        DoEvents
        TriggerSubjectListWorkflow = False
    Else
        With ws.Range("F2")
            .Value = "Complete"
            .Interior.Color = RGB(146, 208, 80)
            .Font.Color = RGB(0, 0, 0)
        End With
        DoEvents
        TriggerSubjectListWorkflow = True
    End If
End Function
