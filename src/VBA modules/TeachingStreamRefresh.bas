'==============================================================================
' TEACHING STREAM REFRESH MODULE
' Purpose: Trigger Power Automate workflow to refresh the teaching stream
'          data from the Teaching Matrix SharePoint file
'==============================================================================

'------------------------------------------------------------------------------
' STANDALONE ENTRY POINT - RefreshTeachingStream
' Can be run independently from the VBA editor or a button
'------------------------------------------------------------------------------
Sub RefreshTeachingStream()
    Dim ws As Worksheet
    Dim yearValue As String
    Dim yearNum As Long
    Dim teachingMatrix As String
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
    teachingMatrix = GetOptionalValue(ws.Range("C5").Value)
    emailValue = GetOptionalValue(ws.Range("C12").Value)
    
    ' Trigger the workflow
    TriggerTeachingStreamWorkflow ws, yearValue, teachingMatrix, emailValue
    
    MsgBox "Teaching Stream Refresh workflow triggered successfully!", vbInformation
End Sub

'------------------------------------------------------------------------------
' TRIGGER TEACHING STREAM WORKFLOW
' Called by Integration.bas or standalone via RefreshTeachingStream
'------------------------------------------------------------------------------
Sub TriggerTeachingStreamWorkflow(ws As Worksheet, yearValue As String, teachingMatrix As String, emailValue As String)
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
