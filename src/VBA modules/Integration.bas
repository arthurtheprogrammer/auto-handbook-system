'===============================================================
' Module: Integration
' Purpose: Orchestrate the full marking support generation:
'          trigger Power Automate workflows sequentially via sync
'          HTTP response, then run HTMLQuery, AssessmentData, and CalcSheets
' Main Entry: GenerateMarkingSupport() - triggered by Dashboard button
' Output: Status updates on Dashboard, completion email sent
' Author: Arthur Chen
' Repository: github.com/arthurtheprogrammer/auto-handbook-system
' Dependencies:
'   - Dashboard sheet (year, file paths, status cells)
'   - SubjectListRefresh.bas, TeachingStreamRefresh.bas (workflows)
'   - HTMLQuery.bas, AssessmentData.bas, CalculationSheets.bas
'===============================================================

Public OriginalCalculationMode As XlCalculation
Public SilentMode As Boolean

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
'          workflows sequentially (sync HTTP), then runs all macros.
'          Workflows are called one after the other because VBA is
'          single-threaded; concurrent sync HTTP calls would freeze
'          Excel until both resolve with no status feedback.
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
    Dim subjectListOK As Boolean
    Dim teachingStreamOK As Boolean
    
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
    ' STEP 1: Subject List workflow — blocks until PA flow responds (sync HTTP)
    ' =========================================================================
    Application.StatusBar = "Running Subject List workflow..."
    subjectListOK = TriggerSubjectListWorkflow(dashboardSheet, yearValue, enrolmentTracker, emailValue)
    
    If Not subjectListOK Then
        MsgBox "Subject List workflow failed. Check cell F2 on the Dashboard." & vbCrLf & vbCrLf & _
               "Teaching Stream workflow will not run.", vbCritical, "Workflow Error"
        SilentMode = False
        Application.Calculation = OriginalCalculationMode
        Application.StatusBar = False
        Exit Sub
    End If
    
    ' =========================================================================
    ' STEP 2: Teaching Stream workflow — runs after Subject List completes
    ' =========================================================================
    Application.StatusBar = "Running Teaching Stream workflow..."
    teachingStreamOK = TriggerTeachingStreamWorkflow(dashboardSheet, yearValue, teachingMatrix, emailValue)
    
    If Not teachingStreamOK Then
        MsgBox "Teaching Stream workflow failed. Check cell F5 on the Dashboard.", vbCritical, "Workflow Error"
        SilentMode = False
        Application.Calculation = OriginalCalculationMode
        Application.StatusBar = False
        Exit Sub
    End If
    
    ' =========================================================================
    ' STEP 3: Both workflows succeeded — run macros and finalise
    ' =========================================================================
    Application.StatusBar = "Both workflows complete. Running macros..."
    RunAllMacros dashboardSheet, emailValue
    FinaliseProcess dashboardSheet, emailValue
    
    MsgBox "All processes finished!" & vbCrLf & vbCrLf & _
           "- Subject List Refresh" & vbCrLf & _
           "- Teaching Stream Refresh" & vbCrLf & _
           "- Assessment Data HTML Query" & vbCrLf & _
           "- Assessment Data Parsing" & vbCrLf & _
           "- Calculation Sheets Generation", vbInformation
    
    SilentMode = False
    Application.Calculation = OriginalCalculationMode
    Application.StatusBar = False
End Sub



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
' ResetStatus
' Purpose: Clear all Dashboard status cells and restore defaults
'---------------------------------------------------------------
Sub ResetStatus()
    Dim dashboardSheet As Worksheet
    Set dashboardSheet = ThisWorkbook.Sheets("Dashboard")
    
    ClearStatusColumn dashboardSheet
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