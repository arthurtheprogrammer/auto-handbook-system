'===============================================================
' Module: TeachingStreamRefresh
' Purpose: Trigger a Power Automate workflow to refresh teaching
'          stream data from the Teaching Matrix on SharePoint
' Main Entry: RefreshTeachingStream() - standalone or via Integration
' Author: Arthur Chen
' Repository: github.com/arthurtheprogrammer/auto-handbook-system
' Dependencies:
'   - Dashboard sheet (year, teaching matrix filename, email)
'   - Integration.bas (EscapeJSON, GetOptionalValue,
'     SendRequestMac/Windows)
'===============================================================

'---------------------------------------------------------------
' RefreshTeachingStream
' Purpose: Validate inputs and trigger the teaching stream
'          workflow as a standalone action
'---------------------------------------------------------------
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
    
    ' Trigger the workflow and show result
    If TriggerTeachingStreamWorkflow(ws, yearValue, teachingMatrix, emailValue) Then
        MsgBox "Teaching Stream Refresh completed successfully!", vbInformation
    Else
        MsgBox "Teaching Stream Refresh failed. Check the status cell on the Dashboard.", vbCritical, "Workflow Error"
    End If
End Sub

'---------------------------------------------------------------
' TriggerTeachingStreamWorkflow
' Purpose: Build JSON payload and send HTTP POST to the teaching
'          stream Power Automate endpoint, then read the sync
'          HTTP response to determine success or failure.
' Returns: True if the flow reported success, False otherwise
' Called by: Integration.GenerateMarkingSupport, RefreshTeachingStream
'---------------------------------------------------------------
Function TriggerTeachingStreamWorkflow(ws As Worksheet, yearValue As String, teachingMatrix As String, emailValue As String) As Boolean
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
    
    ' Parse the synchronous HTTP response body from the PA flow.
    Dim scriptResult As String
    scriptResult = Trim(result)
    If Left(scriptResult, 1) = Chr(34) And Right(scriptResult, 1) = Chr(34) Then
        scriptResult = Mid(scriptResult, 2, Len(scriptResult) - 2)
    End If
    
    Dim flowFailed As Boolean
    flowFailed = (result = "ERROR" Or result = "" Or UCase(Left(scriptResult, 5)) = "ERROR")
    
    If flowFailed Then
        With ws.Range("F5")
            .Value = "Error"
            .Interior.Color = RGB(255, 0, 0)
            .Font.Color = RGB(255, 255, 255)
        End With
        DoEvents
        TriggerTeachingStreamWorkflow = False
    Else
        With ws.Range("F5")
            .Value = "Complete"
            .Interior.Color = RGB(146, 208, 80)
            .Font.Color = RGB(0, 0, 0)
        End With
        DoEvents
        TriggerTeachingStreamWorkflow = True
    End If
End Function
