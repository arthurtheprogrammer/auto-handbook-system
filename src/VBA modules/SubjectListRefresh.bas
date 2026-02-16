'==============================================================================
' SUBJECT LIST REFRESH MODULE
' Purpose: Trigger Power Automate workflow to refresh the subject list
'          from the Enrolment Tracker SharePoint file
'==============================================================================

'------------------------------------------------------------------------------
' STANDALONE ENTRY POINT - RefreshSubjectList
' Can be run independently from the VBA editor or a button
'------------------------------------------------------------------------------
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
    
    ' Trigger the workflow
    TriggerSubjectListWorkflow ws, yearValue, enrolmentTracker, emailValue
    
    MsgBox "Subject List Refresh workflow triggered successfully!", vbInformation
End Sub

'------------------------------------------------------------------------------
' TRIGGER SUBJECT LIST WORKFLOW
' Called by Integration.bas or standalone via RefreshSubjectList
'------------------------------------------------------------------------------
Sub TriggerSubjectListWorkflow(ws As Worksheet, yearValue As String, enrolmentTracker As String, emailValue As String)
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
