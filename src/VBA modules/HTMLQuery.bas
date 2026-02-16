Sub GenerateSubjectQueries()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set wb = ThisWorkbook
    
    ' Check if the worksheet exists
    On Error Resume Next
    Set ws = wb.Worksheets("AllSubjectsHTML")
    On Error GoTo 0
    
    If ws Is Nothing Then
        If Not SilentMode Then MsgBox "Worksheet 'AllSubjectsHTML' not found."
        Exit Sub
    End If
    
    ' Check if the table exists
    On Error Resume Next
    Set tbl = ws.ListObjects("AllSubjectsHTML")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        If Not SilentMode Then MsgBox "Table 'AllSubjectsHTML' not found on the worksheet."
        Exit Sub
    End If
    
    ' Refresh the query table (this waits until complete by default)
    Application.StatusBar = "Refreshing AllSubjectsHTML query..."
    
    On Error Resume Next
    ' Try to refresh via QueryTable
    If Not tbl.QueryTable Is Nothing Then
        tbl.QueryTable.BackgroundQuery = False
        tbl.QueryTable.Refresh BackgroundQuery:=False
    Else
        ' Or refresh all queries in the workbook and wait
        wb.RefreshAll
        DoEvents
    End If
    On Error GoTo 0
    
    Application.StatusBar = False
    
    ' Format the table
    Call FormatTableCleanup(ws, tbl)
    
    If Not SilentMode Then MsgBox "Query refreshed and formatted successfully."
End Sub


Sub FormatTableCleanup(ws As Worksheet, Optional tbl As ListObject = Nothing)
    ' Clean up table formatting: standard row heights, no text wrap, autofit columns
    ' Similar to Office Script reset formatting functionality
    
    On Error Resume Next
    
    Dim usedRange As Range
    Set usedRange = ws.usedRange
    
    If usedRange Is Nothing Then Exit Sub
    
    ' Disable text wrapping for all cells
    usedRange.WrapText = False
    
    ' Set all rows to standard height (15 points)
    ' First autofit to ensure content visibility, then standardize
    usedRange.Rows.AutoFit
    usedRange.Rows.RowHeight = 15
    
    ' Autofit columns for readability
    ws.Columns.AutoFit
    
    ' Set specific column widths for columns B and C (URL and HTML columns)
    ws.Columns("B:C").ColumnWidth = 70
    
    ' Optional: Set maximum column width for other columns to prevent excessive widths
    Dim col As Long
    For col = 4 To usedRange.Columns.Count ' Start from column D onwards
        If ws.Columns(col).ColumnWidth > 50 Then
            ws.Columns(col).ColumnWidth = 50
        End If
    Next col
    
    ' Hyperlink column B cells with their own URL values
    If Not tbl Is Nothing Then
        Dim urlCol As Range
        Set urlCol = tbl.ListColumns("URL").DataBodyRange
        
        Dim cell As Range
        For Each cell In urlCol
            If cell.Value <> "" And Not IsEmpty(cell.Value) Then
                ' Remove existing hyperlink if present
                If cell.Hyperlinks.Count > 0 Then
                    cell.Hyperlinks(1).Delete
                End If
                ' Add hyperlink
                ws.Hyperlinks.add Anchor:=cell, Address:=cell.Value, TextToDisplay:=cell.Value
            End If
        Next cell
    End If
    
    ' Center align columns D and E (HTMLLength and Status)
    ws.Columns("D:E").HorizontalAlignment = xlCenter
    
    ' Format FetchTime column (column G) to YYYY-MM-DD HH:MM:SS
    If Not tbl Is Nothing Then
        On Error Resume Next
        Dim timeCol As Range
        Set timeCol = tbl.ListColumns("FetchTime").DataBodyRange
        
        If Not timeCol Is Nothing Then
            timeCol.NumberFormat = "yyyy-mm-dd hh:mm:ss"
        End If
        On Error GoTo 0
    End If
    
    ' If table was provided, ensure header row is slightly taller
    If Not tbl Is Nothing Then
        tbl.HeaderRowRange.RowHeight = 18
        tbl.HeaderRowRange.Font.Bold = True
        
        ' Left align all headers first
        tbl.HeaderRowRange.HorizontalAlignment = xlLeft
        
        ' Then center align only columns D and E headers
        On Error Resume Next
        tbl.ListColumns("HTMLLength").Range.Cells(1).HorizontalAlignment = xlCenter
        tbl.ListColumns("Status").Range.Cells(1).HorizontalAlignment = xlCenter
        On Error GoTo 0
        
        tbl.HeaderRowRange.VerticalAlignment = xlCenter
        
        ' Apply Olive Green Medium table style (TableStyleMedium4)
        tbl.TableStyle = "TableStyleMedium4"
    End If
    
    ' Set standard formatting for all cells
    With usedRange
        .VerticalAlignment = xlTop
    End With
    
    ' Reset horizontal alignment for columns that aren't specifically centered
    ws.Columns("A:A").HorizontalAlignment = xlLeft ' SubjectCode
    ws.Columns("F:F").HorizontalAlignment = xlLeft ' ErrorMessage
    
    ' Freeze the header row if table exists
    If Not tbl Is Nothing Then
        ws.Activate
        ws.Range("A2").Select
        ActiveWindow.FreezePanes = True
    End If
    
    On Error GoTo 0
End Sub