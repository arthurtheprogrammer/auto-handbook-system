'===============================================================
' Module: HTMLQuery
' Purpose: Refresh the Power Query connection that pulls handbook
'          HTML for all subjects, then format the results table
' Main Entry: GenerateSubjectQueries() - called by Integration.RunAllMacros
' Output: Refreshed and formatted AllSubjectsHTML table with
'         live hyperlinks and status column
' Author: Arthur Chen
' Repository: github.com/arthurtheprogrammer/auto-handbook-system
' Dependencies:
'   - AllSubjectsHTML sheet and table (Power Query connection)
'   - Power Query definition in workbook (Windows only)
'===============================================================

'---------------------------------------------------------------
' GenerateSubjectQueries
' Purpose: Refresh the Power Query connection to pull fresh
'          handbook HTML. On Mac, skips the refresh and uses
'          existing data. Reports success/failure counts.
' Called by: Integration.RunAllMacros
'---------------------------------------------------------------
Sub GenerateSubjectQueries()
    Dim wb As Workbook
    Dim htmlSheet As Worksheet
    Dim dashboardSheet As Worksheet
    Dim subjectsTable As ListObject
    
    Set wb = ThisWorkbook
    
    On Error Resume Next
    Set dashboardSheet = wb.Sheets("Dashboard")
    On Error GoTo 0
    
    If Not dashboardSheet Is Nothing Then
        With dashboardSheet.Range("F3")
            .Value = "Running..."
            .Interior.Color = RGB(255, 192, 0)
        End With
        DoEvents
    End If
    
    ' Check if the worksheet exists
    On Error Resume Next
    Set htmlSheet = wb.Worksheets("AllSubjectsHTML")
    On Error GoTo 0
    
    If htmlSheet Is Nothing Then
        If Not SilentMode Then MsgBox "Worksheet 'AllSubjectsHTML' not found."
        Exit Sub
    End If
    
    ' Check if the table exists
    On Error Resume Next
    Set subjectsTable = htmlSheet.ListObjects("AllSubjectsHTML")
    On Error GoTo 0
    
    If subjectsTable Is Nothing Then
        If Not SilentMode Then MsgBox "Table 'AllSubjectsHTML' not found on the worksheet."
        Exit Sub
    End If
    
    ' =========================================================================
    ' MAC DETECTION — Power Query is Windows-only, use Power Automate instead
    ' =========================================================================
    #If Mac Then
        Dim macResponse As VbMsgBoxResult
        macResponse = MsgBox("Power Query is not available on Mac." & vbCrLf & vbCrLf & _
                       "Would you like to trigger the cloud HTML download workflow instead?" & vbCrLf & _
                       "(This uses Power Automate and may take a few minutes)" & vbCrLf & vbCrLf & _
                       "Click No to skip and use existing data.", _
                       vbQuestion + vbYesNo, "Mac Detected")
        
        If macResponse = vbYes Then
            ' Trigger the HTML download workflow via Power Automate
            Dim yearValue As String
            Dim emailValue As String
            Dim macWs As Worksheet
            Set macWs = wb.Sheets("Dashboard")
            
            yearValue = CStr(macWs.Range("C2").Value)
            emailValue = ""
            On Error Resume Next
            emailValue = CStr(macWs.Range("C12").Value)
            On Error GoTo 0
            
            ' Build unique subject codes JSON array from SubjectList
            Dim wsSubjects As Worksheet
            Set wsSubjects = wb.Sheets("SubjectList")
            Dim subjectCodes As String
            subjectCodes = "["
            
            If Not wsSubjects Is Nothing Then
                Dim lastSubjectRow As Long
                lastSubjectRow = wsSubjects.Cells(wsSubjects.Rows.Count, "B").End(xlUp).Row
                
                ' Use a Collection to track unique valid codes
                Dim uniqueCodes As New Collection
                Dim r As Long
                For r = 2 To lastSubjectRow
                    Dim code As String
                    code = Trim(CStr(wsSubjects.Cells(r, 2).Value))
                    ' Validate: code must be non-empty, at least 9 chars, last 5 chars must be numeric
                    If Len(code) >= 9 And IsNumeric(Right(code, 5)) Then
                        On Error Resume Next
                        uniqueCodes.Add code, code
                        Err.Clear
                        On Error GoTo 0
                    End If
                Next r
                
                ' Copy to array and sort alphabetically
                If uniqueCodes.Count > 0 Then
                    Dim sortArr() As String
                    ReDim sortArr(1 To uniqueCodes.Count)
                    Dim c As Long
                    For c = 1 To uniqueCodes.Count
                        sortArr(c) = uniqueCodes(c)
                    Next c
                    
                    ' Bubble sort
                    Dim s As Long, t As Long, tmp As String
                    For s = 1 To UBound(sortArr) - 1
                        For t = s + 1 To UBound(sortArr)
                            If sortArr(s) > sortArr(t) Then
                                tmp = sortArr(s)
                                sortArr(s) = sortArr(t)
                                sortArr(t) = tmp
                            End If
                        Next t
                    Next s
                    
                    ' Build JSON array from sorted codes
                    For c = 1 To UBound(sortArr)
                        If c > 1 Then subjectCodes = subjectCodes & ","
                        subjectCodes = subjectCodes & Chr(34) & EscapeJSON(sortArr(c)) & Chr(34)
                    Next c
                End If
            End If
            subjectCodes = subjectCodes & "]"
            
            Dim htmlWorkflowUrl As String
            htmlWorkflowUrl = "https://default0e5bf3cf1ff446b7917652c538c22a.4d.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/ae04d220b067440fb5c56f887cda8541/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=_LTvpCheHaRroFK5ri7nZYH23mmI5pofxhqvimW4LqM"
            
            Dim jsonPayload As String
            jsonPayload = "{" & Chr(34) & "year" & Chr(34) & ":" & yearValue & "," & _
                           Chr(34) & "email" & Chr(34) & ":" & Chr(34) & EscapeJSON(emailValue) & Chr(34) & "," & _
                           Chr(34) & "subjects" & Chr(34) & ":" & subjectCodes & "}"
            
            If Not dashboardSheet Is Nothing Then
                With dashboardSheet.Range("F3")
                    .Value = "Running..."
                    .Interior.Color = RGB(255, 192, 0)
                End With
                DoEvents
            End If
            
            Application.StatusBar = "Triggering HTML download workflow..."
            Dim triggerResult As String
            triggerResult = SendRequestMac(htmlWorkflowUrl, jsonPayload)
            
            If triggerResult = "ERROR" Then
                Application.StatusBar = False
                If Not SilentMode Then MsgBox "Failed to trigger HTML download workflow." & vbCrLf & vbCrLf & _
                       "Check your network connection. Continuing with existing data.", vbExclamation, "Workflow Error"
                
                If Not dashboardSheet Is Nothing Then
                    With dashboardSheet.Range("F3")
                        .Value = "Skipped"
                        .Interior.Color = RGB(191, 191, 191)
                    End With
                End If
                
                Call FormatTableCleanup(htmlSheet, subjectsTable)
                Exit Sub
            End If
            
            ' Poll F3 for "Done" (workflow writes this on completion)
            Application.StatusBar = "Waiting for HTML download workflow to complete..."
            Dim startTime As Double
            Dim maxWaitSeconds As Long
            maxWaitSeconds = 600  ' 10 minute timeout
            startTime = Timer
            
            Do
                DoEvents
                Application.Wait (Now + TimeValue("0:00:05"))
                
                ' Force sync to pick up cloud changes
                On Error Resume Next
                wb.Save
                On Error GoTo 0
                
                Dim f3Status As String
                f3Status = UCase(Trim(CStr(dashboardSheet.Range("F3").Value)))
                
                If f3Status = "DONE" Or f3Status = "COMPLETE" Or f3Status = "FINISHED" Then
                    Application.StatusBar = "Formatting table..."
                    
                    ' Silently format, leave F3 as "Done" (set by Power Automate)
                    Call FormatTableCleanup(htmlSheet, subjectsTable)
                    
                    Application.StatusBar = False
                    Exit Sub
                End If
                
                Dim elapsed As Double
                elapsed = Timer - startTime
                If elapsed < 0 Then elapsed = elapsed + 86400
                Application.StatusBar = "Waiting for HTML download... (" & Format(elapsed, "0") & "s)"
                
                If elapsed > maxWaitSeconds Then
                    Application.StatusBar = False
                    If Not SilentMode Then MsgBox "HTML download workflow did not complete within 10 minutes." & vbCrLf & vbCrLf & _
                           "The workflow may still be running in the background." & vbCrLf & _
                           "Continuing with existing data.", vbExclamation, "Workflow Timeout"
                    
                    If Not dashboardSheet Is Nothing Then
                        With dashboardSheet.Range("F3")
                            .Value = "Timeout"
                            .Interior.Color = RGB(255, 0, 0)
                        End With
                    End If
                    
                    Call FormatTableCleanup(htmlSheet, subjectsTable)
                    Exit Sub
                End If
            Loop
        Else
            ' User chose to skip
            If Not dashboardSheet Is Nothing Then
                With dashboardSheet.Range("F3")
                    .Value = "Skipped"
                    .Interior.Color = RGB(191, 191, 191)
                End With
                DoEvents
            End If
            
            Application.StatusBar = False
            Call FormatTableCleanup(htmlSheet, subjectsTable)
            Exit Sub
        End If
    #End If
    
    ' =========================================================================
    ' WINDOWS — Refresh the Power Query connection
    ' =========================================================================
    Application.StatusBar = "Refreshing AllSubjectsHTML query..."
    
    On Error Resume Next
    Dim refreshError As Long
    ' Try QueryTable refresh
    If Not subjectsTable.QueryTable Is Nothing Then
        subjectsTable.QueryTable.BackgroundQuery = False
        subjectsTable.QueryTable.Refresh BackgroundQuery:=False
        refreshError = Err.Number
    Else
        ' Fall back to named connection if QueryTable is unavailable
        Dim conn As WorkbookConnection
        Dim connFound As Boolean
        connFound = False
        
        For Each conn In wb.Connections
            If InStr(1, conn.Name, "AllSubjectsHTML", vbTextCompare) > 0 Then
                conn.Refresh
                refreshError = Err.Number
                connFound = True
                Exit For
            End If
        Next conn
        
        ' Last resort — should not normally reach here
        If Not connFound Then
            wb.RefreshAll
            refreshError = Err.Number
        End If
        DoEvents
    End If
    On Error GoTo 0
    
    ' Check if refresh actually succeeded
    If refreshError <> 0 Then
        If Not SilentMode Then MsgBox "Power Query refresh encountered an error (" & refreshError & ")." & vbCrLf & vbCrLf & _
               "The table may contain stale data. Check Data > Queries & Connections " & _
               "in Excel to verify the query is configured correctly.", vbExclamation, "Refresh Warning"
    End If
    
    ' Validate that FetchTime is fresh (today) — stale dates mean refresh didn't run
    On Error Resume Next
    Dim fetchTimeCol As Range
    Set fetchTimeCol = subjectsTable.ListColumns("FetchTime").DataBodyRange
    On Error GoTo 0
    
    If Not fetchTimeCol Is Nothing Then
        If subjectsTable.ListRows.Count > 0 Then
            Dim latestFetch As Variant
            latestFetch = fetchTimeCol.Cells(1).Value
            
            If IsDate(latestFetch) Then
                If Int(CDate(latestFetch)) < Int(Now) Then
                    If Not SilentMode Then MsgBox "The AllSubjectsHTML data appears stale — FetchTime " & _
                           "is from " & Format(CDate(latestFetch), "yyyy-mm-dd") & "." & vbCrLf & vbCrLf & _
                           "The Power Query may not have refreshed correctly." & vbCrLf & _
                           "Check Data > Queries & Connections in Excel.", vbExclamation, "Refresh Warning"
                End If
            End If
        End If
    End If
    
    Application.StatusBar = False
    
    Call FormatTableCleanup(htmlSheet, subjectsTable)
    
    ' =========================================================================
    ' POST-REFRESH STATUS CHECK — report success/failure counts to user
    ' =========================================================================
    If Not SilentMode Then
        Dim statusCol As Range
        On Error Resume Next
        Set statusCol = subjectsTable.ListColumns("Status").DataBodyRange
        On Error GoTo 0
        
        If Not statusCol Is Nothing Then
            Dim totalRows As Long, failedRows As Long
            Dim statusCell As Range
            totalRows = statusCol.Rows.Count
            failedRows = 0
            
            For Each statusCell In statusCol
                If UCase(Trim(statusCell.Value)) = "FAILED" Then
                    failedRows = failedRows + 1
                End If
            Next statusCell
            
            If failedRows > 0 Then
                MsgBox "Query refreshed — " & (totalRows - failedRows) & "/" & totalRows & _
                       " succeeded, " & failedRows & " failed." & vbCrLf & vbCrLf & _
                       "Failed subjects may have invalid handbook URLs." & vbCrLf & _
                       "Check the Status column in AllSubjectsHTML for details.", _
                       vbExclamation, "Refresh Complete (with errors)"
            Else
                MsgBox "Query refreshed and formatted successfully — " & totalRows & " succeeded.", vbInformation, "Refresh Complete"
            End If
        Else
            MsgBox "Query refreshed and formatted successfully."
        End If
    End If
    
    If Not dashboardSheet Is Nothing Then
        With dashboardSheet.Range("F3")
            .Value = "Done"
            .Interior.Color = RGB(146, 208, 80)
        End With
        DoEvents
    End If
End Sub


'---------------------------------------------------------------
' FormatTableCleanup
' Purpose: Apply consistent formatting to the AllSubjectsHTML
'          table: row heights, column widths, hyperlinks,
'          date formats, header styling, and freeze panes
' Called by: GenerateSubjectQueries
'---------------------------------------------------------------
Sub FormatTableCleanup(htmlSheet As Worksheet, Optional subjectsTable As ListObject = Nothing)
    On Error Resume Next
    
    Dim usedRange As Range
    Set usedRange = htmlSheet.usedRange
    
    If usedRange Is Nothing Then Exit Sub
    
    ' =========================================================================
    ' SORT — ensure consistent SubjectCode order (A-Z) after concurrent writes
    ' =========================================================================
    If Not subjectsTable Is Nothing Then
        If subjectsTable.ListRows.Count > 0 Then
            subjectsTable.Sort.SortFields.Clear
            subjectsTable.Sort.SortFields.Add2 Key:=subjectsTable.ListColumns("SubjectCode").Range, _
                SortOn:=xlSortOnValues, Order:=xlAscending
            subjectsTable.Sort.Header = xlYes
            subjectsTable.Sort.Apply
        End If
    End If
    
    ' =========================================================================
    ' ROW AND COLUMN SIZING — standardise dimensions for readability
    ' =========================================================================
    usedRange.WrapText = False
    usedRange.Rows.AutoFit
    usedRange.Rows.RowHeight = 15
    htmlSheet.Columns.AutoFit
    
    ' URL and HTML columns set with extra width to show full content
    htmlSheet.Columns("B:C").ColumnWidth = 70
    
    ' Cap remaining columns to prevent excessively wide sheets
    Dim col As Long
    For col = 4 To usedRange.Columns.Count
        If htmlSheet.Columns(col).ColumnWidth > 50 Then
            htmlSheet.Columns(col).ColumnWidth = 50
        End If
    Next col
    
    ' =========================================================================
    ' HYPERLINKS — make URL column clickable for quick handbook access
    ' =========================================================================
    If Not subjectsTable Is Nothing Then
        Dim urlCol As Range
        Set urlCol = subjectsTable.ListColumns("URL").DataBodyRange
        
        Dim cell As Range
        For Each cell In urlCol
            If cell.Value <> "" And Not IsEmpty(cell.Value) Then
                ' Clear old hyperlink first to avoid duplicates
                If cell.Hyperlinks.Count > 0 Then
                    cell.Hyperlinks(1).Delete
                End If
                htmlSheet.Hyperlinks.add Anchor:=cell, Address:=cell.Value, TextToDisplay:=cell.Value
            End If
        Next cell
    End If
    
    ' =========================================================================
    ' ALIGNMENT AND DATE FORMATTING
    ' =========================================================================
    htmlSheet.Columns("D:E").HorizontalAlignment = xlCenter
    
    ' Format FetchTime to ISO-style for consistency across locales
    If Not subjectsTable Is Nothing Then
        On Error Resume Next
        Dim timeCol As Range
        Set timeCol = subjectsTable.ListColumns("FetchTime").DataBodyRange
        
        If Not timeCol Is Nothing Then
            timeCol.NumberFormat = "yyyy-mm-dd hh:mm:ss"
        End If
        On Error GoTo 0
    End If
    
    ' =========================================================================
    ' HEADER STYLING — taller header row with Olive Green theme
    ' =========================================================================
    If Not subjectsTable Is Nothing Then
        subjectsTable.HeaderRowRange.RowHeight = 18
        subjectsTable.HeaderRowRange.Font.Bold = True
        
        ' Left-align all headers, then selectively centre numeric columns
        subjectsTable.HeaderRowRange.HorizontalAlignment = xlLeft
        
        On Error Resume Next
        subjectsTable.ListColumns("HTMLLength").Range.Cells(1).HorizontalAlignment = xlCenter
        subjectsTable.ListColumns("Status").Range.Cells(1).HorizontalAlignment = xlCenter
        On Error GoTo 0
        
        subjectsTable.HeaderRowRange.VerticalAlignment = xlCenter
        subjectsTable.TableStyle = "TableStyleMedium4"
    End If
    
    ' Ensure content cells are top-aligned for multi-line HTML previews
    With usedRange
        .VerticalAlignment = xlTop
    End With
    
    ' Override centre-alignment for text-heavy columns
    htmlSheet.Columns("A:A").HorizontalAlignment = xlLeft  ' SubjectCode
    htmlSheet.Columns("F:F").HorizontalAlignment = xlLeft  ' ErrorMessage
    
    ' Freeze header so column names stay visible while scrolling
    If Not subjectsTable Is Nothing Then
        htmlSheet.Activate
        htmlSheet.Range("A2").Select
        ActiveWindow.FreezePanes = True
    End If
    
    On Error GoTo 0
End Sub