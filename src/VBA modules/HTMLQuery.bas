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
    Dim subjectsTable As ListObject
    
    Set wb = ThisWorkbook
    
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
    ' MAC DETECTION — Power Query is Windows-only
    ' =========================================================================
    #If Mac Then
        If Not SilentMode Then
            MsgBox "Power Query refresh is not compatible with Mac." & vbCrLf & vbCrLf & _
                   "The handbook data will not be refreshed, but the rest of the " & _
                   "process will continue using existing data." & vbCrLf & vbCrLf & _
                   "This is fine after the first run each year — handbook content " & _
                   "rarely changes after semester starts.", _
                   vbInformation, "Mac Detected"
        End If
        
        Application.StatusBar = "Skipped Power Query refresh (Mac)..."
        DoEvents
        Application.StatusBar = False
        
        ' Format existing data
        Call FormatTableCleanup(htmlSheet, subjectsTable)
        Exit Sub
    #End If
    
    ' =========================================================================
    ' WINDOWS — Refresh the Power Query connection
    ' =========================================================================
    Application.StatusBar = "Refreshing AllSubjectsHTML query..."
    
    On Error Resume Next
    ' Try QueryTable refresh
    If Not subjectsTable.QueryTable Is Nothing Then
        subjectsTable.QueryTable.BackgroundQuery = False
        subjectsTable.QueryTable.Refresh BackgroundQuery:=False
    Else
        ' Fall back to named connection if QueryTable is unavailable
        Dim conn As WorkbookConnection
        Dim connFound As Boolean
        connFound = False
        
        For Each conn In wb.Connections
            If InStr(1, conn.Name, "AllSubjectsHTML", vbTextCompare) > 0 Then
                conn.Refresh
                connFound = True
                Exit For
            End If
        Next conn
        
        ' Last resort — should not normally reach here
        If Not connFound Then
            wb.RefreshAll
        End If
        DoEvents
    End If
    On Error GoTo 0
    
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