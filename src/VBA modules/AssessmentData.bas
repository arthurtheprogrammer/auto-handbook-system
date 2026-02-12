' Custom type to hold study period data - must be at module level
Type StudyPeriodData
    periodName As String
    assessments() As String
    assessmentCount As Integer
End Type

' Main parsing subroutine
Sub ParseAssessmentData()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentRow As Long
    Dim subjectCode As String
    Dim htmlContent As String
        
    ' Set up source worksheet
    Set wsSource = ThisWorkbook.Sheets("AllSubjectsHTML")
    
    ' Create or clear target worksheet
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets("assessment data parsed")
    If wsTarget Is Nothing Then
        Set wsTarget = ThisWorkbook.Sheets.add
        wsTarget.name = "assessment data parsed"
    Else
        wsTarget.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Set up headers
    Call SetupHeaders(wsTarget)
    
    ' Initialize processing
    currentRow = 2
    lastRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row
    
    ' Process each subject
    For i = 2 To lastRow ' Row 1 has headers
        subjectCode = Trim(wsSource.Cells(i, 1).Value)
        htmlContent = wsSource.Cells(i, 3).Value
            
        If subjectCode <> "" And htmlContent <> "" Then
            Call ProcessSubjectHTML(subjectCode, htmlContent, wsTarget, currentRow)
        End If
    Next i
    
    ' Apply final formatting
    Call FormatOutput(wsTarget, currentRow)
    
    ' Show completion message
    MsgBox "Assessment data parsing completed! " & (currentRow - 2) & " records processed."
End Sub

Sub SetupHeaders(wsTarget As Worksheet)
    With wsTarget
        .Cells(1, 1).Value = "UID"
        .Cells(1, 2).Value = "Subject Code"
        .Cells(1, 3).Value = "Study Period"
        .Cells(1, 4).Value = "Assessment Description"
        .Cells(1, 5).Value = "Location"
        .Cells(1, 6).Value = "Word Count"
        .Cells(1, 7).Value = "Exam Duration"
        .Cells(1, 8).Value = "Group Size"
        .Cells(1, 9).Value = "Timing"
        .Cells(1, 10).Value = "Percentage"
    End With
End Sub

Sub FormatOutput(wsTarget As Worksheet, currentRow As Long)
    With wsTarget
        ' Auto-fit all columns first
        .Columns.AutoFit
            
        ' Set specific width for Assessment Description column
        .Columns("D").ColumnWidth = 150
        
        ' Wrap text for Assessment Description column
        .Columns("D").WrapText = True
        
        ' Set uniform width for Location to Group Size columns (E to H)
        .Columns("E:H").ColumnWidth = 15
        
        ' Center align columns from Location to Percentage (E to H, J)
        If currentRow > 1 Then
            .Range("E1:H" & (currentRow - 1)).HorizontalAlignment = xlCenter
            .Range("J1:J" & (currentRow - 1)).HorizontalAlignment = xlCenter
        End If
        
        ' Left align Timing column (I)
        If currentRow > 1 Then
            .Range("I1:I" & (currentRow - 1)).HorizontalAlignment = xlLeft
        End If
            
        ' Format headers
        .Range("A1:J1").Font.Bold = True
            
        ' Set uniform row height for data rows
        If currentRow > 2 Then
            .Range("2:" & (currentRow - 1)).RowHeight = 15
        End If
    End With
End Sub

Sub ProcessSubjectHTML(subjectCode As String, htmlContent As String, wsTarget As Worksheet, ByRef currentRow As Long)
    Dim studyPeriods() As StudyPeriodData
    Dim periodCount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim assessmentCount As Integer
    
    ' Extract all study periods with their assessments
    Call ExtractStudyPeriods(htmlContent, studyPeriods, periodCount)
    
    ' Process each study period
    For i = 0 To periodCount - 1
        assessmentCount = 1
            
        ' Process each assessment in this study period
        For j = 0 To studyPeriods(i).assessmentCount - 1
            If Trim(studyPeriods(i).assessments(j)) <> "" Then
                Call ProcessAssessmentRow(subjectCode, studyPeriods(i).periodName, assessmentCount, studyPeriods(i).assessments(j), wsTarget, currentRow)
                assessmentCount = assessmentCount + 1
                currentRow = currentRow + 1
            End If
        Next j
    Next i
End Sub

Sub ExtractStudyPeriods(htmlContent As String, ByRef studyPeriods() As StudyPeriodData, ByRef periodCount As Integer)
    Dim h3StartPos As Long
    Dim h3EndPos As Long
    Dim nextH3Pos As Long
    Dim currentPos As Long
    Dim periodName As String
    Dim sectionContent As String
    Dim tempPeriods(10) As StudyPeriodData
    Dim i As Integer
    
    periodCount = 0
    currentPos = 1
    
    ' Check if there are any <h3> tags
    h3StartPos = InStr(1, htmlContent, "<h3>", vbTextCompare)
    
    If h3StartPos = 0 Then
        ' No study period headers - assign everything to "All"
        tempPeriods(0).periodName = "All"
        Call ExtractAssessmentsFromSection(htmlContent, tempPeriods(0).assessments, tempPeriods(0).assessmentCount)
        periodCount = 1
    Else
        ' Process each <h3> section
        Do While currentPos <= Len(htmlContent)
            h3StartPos = InStr(currentPos, htmlContent, "<h3>", vbTextCompare)
            If h3StartPos = 0 Then Exit Do
                
            ' Extract period name
            h3StartPos = h3StartPos + 4
            h3EndPos = InStr(h3StartPos, htmlContent, "</h3>", vbTextCompare)
            If h3EndPos = 0 Then Exit Do
                
            periodName = Trim(Mid(htmlContent, h3StartPos, h3EndPos - h3StartPos))
                
            ' Find section boundary
            nextH3Pos = InStr(h3EndPos + 5, htmlContent, "<h3>", vbTextCompare)
            If nextH3Pos = 0 Then
                sectionContent = Mid(htmlContent, h3EndPos + 5)
            Else
                sectionContent = Mid(htmlContent, h3EndPos + 5, nextH3Pos - (h3EndPos + 5))
            End If
                
            ' Store period data
            If periodCount < 10 Then
                tempPeriods(periodCount).periodName = periodName
                Call ExtractAssessmentsFromSection(sectionContent, tempPeriods(periodCount).assessments, tempPeriods(periodCount).assessmentCount)
                periodCount = periodCount + 1
            End If
                
            ' Move to next section
            currentPos = IIf(nextH3Pos = 0, Len(htmlContent) + 1, nextH3Pos)
        Loop
    End If
    
    ' Copy to output array
    If periodCount > 0 Then
        ReDim studyPeriods(0 To periodCount - 1)
        For i = 0 To periodCount - 1
            studyPeriods(i).periodName = tempPeriods(i).periodName
            studyPeriods(i).assessmentCount = tempPeriods(i).assessmentCount
                
            ' Copy the assessments array
            If tempPeriods(i).assessmentCount > 0 Then
                ReDim studyPeriods(i).assessments(0 To tempPeriods(i).assessmentCount - 1)
                Dim j As Integer
                For j = 0 To tempPeriods(i).assessmentCount - 1
                    studyPeriods(i).assessments(j) = tempPeriods(i).assessments(j)
                Next j
            End If
        Next i
    End If
End Sub

Sub ExtractAssessmentsFromSection(sectionContent As String, ByRef assessments() As String, ByRef assessmentCount As Integer)
    Dim startPos As Long
    Dim endPos As Long
    Dim currentPos As Long
    Dim rowContent As String
    Dim tempAssessments(50) As String
    Dim i As Integer
    
    assessmentCount = 0
    currentPos = 1
    
    ' Extract all table rows, skipping headers
    Do While currentPos <= Len(sectionContent)
        startPos = InStr(currentPos, sectionContent, "<tr>", vbTextCompare)
        If startPos = 0 Then Exit Do
            
        endPos = InStr(startPos, sectionContent, "</tr>", vbTextCompare)
        If endPos = 0 Then Exit Do
            
        rowContent = Mid(sectionContent, startPos, endPos - startPos + 5)
            
        ' Skip header rows and empty content
        If InStr(1, rowContent, "<th>", vbTextCompare) = 0 And Trim(rowContent) <> "" Then
            If assessmentCount < 50 Then
                tempAssessments(assessmentCount) = rowContent
                assessmentCount = assessmentCount + 1
            End If
        End If
            
        currentPos = endPos + 5
    Loop
    
    ' Copy to output array
    If assessmentCount > 0 Then
        ReDim assessments(0 To assessmentCount - 1)
        For i = 0 To assessmentCount - 1
            assessments(i) = tempAssessments(i)
        Next i
    End If
End Sub

Sub ProcessAssessmentRow(subjectCode As String, studyPeriod As String, assessmentCount As Integer, rowContent As String, wsTarget As Worksheet, currentRow As Long)
    Dim uid As String
    Dim description As String
    Dim wordCount As Long
    Dim examDuration As Integer
    Dim groupSize As Integer
    Dim isInClassAssessment As Boolean
    Dim timing As String
    Dim percentage As String
    
    ' Generate UID
    uid = subjectCode & "_" & studyPeriod & "_" & assessmentCount
    
    ' Extract and clean description
    description = ExtractDescription(rowContent)
    
    ' Determine if this is an in-class assessment
    isInClassAssessment = CheckIfInClass(description)
    
    ' Extract numerical data based on assessment type
    If isInClassAssessment Then
        wordCount = 0          ' Don't extract word count for in-class assessments
        examDuration = 0       ' Don't extract exam duration for in-class assessments
    Else
        wordCount = ExtractWordCount(rowContent)
        examDuration = ExtractExamDuration(description)
    End If
    
    ' Always extract group size regardless of assessment type
    groupSize = ExtractGroupSize(description)
    
    ' Extract timing and percentage
    timing = ExtractTiming(rowContent)
    percentage = ExtractPercentage(rowContent)
    
    ' Write data to worksheet
    Call WriteAssessmentData(wsTarget, currentRow, uid, subjectCode, studyPeriod, description, wordCount, examDuration, groupSize, isInClassAssessment, timing, percentage)
End Sub

Function CheckIfInClass(description As String) As Boolean
    Dim lowerDesc As String
    lowerDesc = LCase(description)
    
    CheckIfInClass = (InStr(1, lowerDesc, "participation") > 0) Or _
                     (InStr(1, lowerDesc, "presentation") > 0) Or _
                     (InStr(1, lowerDesc, "attendance") > 0) Or _
                     (InStr(1, lowerDesc, "pitch") > 0) Or _
                     (InStr(1, lowerDesc, "online") > 0) Or _
                     (InStr(1, lowerDesc, "ongoing") > 0) Or _
                     (InStr(1, lowerDesc, "in class") > 0)
End Function

Sub WriteAssessmentData(wsTarget As Worksheet, currentRow As Long, uid As String, subjectCode As String, studyPeriod As String, description As String, wordCount As Long, examDuration As Integer, groupSize As Integer, isInClass As Boolean, timing As String, percentage As String)
    With wsTarget
        .Cells(currentRow, 1).Value = uid
        .Cells(currentRow, 2).Value = subjectCode
        .Cells(currentRow, 3).Value = studyPeriod
        .Cells(currentRow, 4).Value = description
        
        ' Set assessment location for in-class assessments
        If isInClass Then .Cells(currentRow, 5).Value = "in class"
        
        ' Only populate if values are greater than 0
        If wordCount > 0 Then .Cells(currentRow, 6).Value = wordCount
        If examDuration > 0 Then .Cells(currentRow, 7).Value = examDuration
        If groupSize > 0 Then .Cells(currentRow, 8).Value = groupSize
        
        ' Add timing and percentage
        If timing <> "" Then .Cells(currentRow, 9).Value = timing
        If percentage <> "" Then .Cells(currentRow, 10).Value = percentage
    End With
End Sub

Function ExtractDescription(rowContent As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim description As String
    
    ' Find first table cell content
    startPos = InStr(1, rowContent, "<td>", vbTextCompare)
    If startPos > 0 Then
        startPos = startPos + 4
        endPos = InStr(startPos, rowContent, "</td>", vbTextCompare)
        If endPos > 0 Then
            description = Mid(rowContent, startPos, endPos - startPos)
            description = CleanHTMLContent(description)
            ExtractDescription = Trim(description)
        End If
    End If
End Function

Function ExtractTiming(rowContent As String) As String
    Dim tdCount As Integer
    Dim pos As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim timingContent As String
    
    ExtractTiming = ""
    tdCount = 0
    pos = 1
    
    ' Find the second <td> tag (timing column)
    Do While pos <= Len(rowContent)
        startPos = InStr(pos, rowContent, "<td>", vbTextCompare)
        If startPos = 0 Then Exit Do
        
        tdCount = tdCount + 1
        
        If tdCount = 2 Then
            ' Found the timing column
            startPos = startPos + 4
            endPos = InStr(startPos, rowContent, "</td>", vbTextCompare)
            If endPos > 0 Then
                timingContent = Mid(rowContent, startPos, endPos - startPos)
                ExtractTiming = Trim(CleanHTMLContent(timingContent))
            End If
            Exit Do
        End If
        
        pos = startPos + 4
    Loop
End Function

Function ExtractPercentage(rowContent As String) As String
    Dim tdCount As Integer
    Dim pos As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim percentageContent As String
    
    ExtractPercentage = ""
    tdCount = 0
    pos = 1
    
    ' Find the third <td> tag (percentage column)
    Do While pos <= Len(rowContent)
        startPos = InStr(pos, rowContent, "<td>", vbTextCompare)
        If startPos = 0 Then Exit Do
        
        tdCount = tdCount + 1
        
        If tdCount = 3 Then
            ' Found the percentage column
            startPos = startPos + 4
            endPos = InStr(startPos, rowContent, "</td>", vbTextCompare)
            If endPos > 0 Then
                percentageContent = Mid(rowContent, startPos, endPos - startPos)
                ExtractPercentage = Trim(CleanHTMLContent(percentageContent))
            End If
            Exit Do
        End If
        
        pos = startPos + 4
    Loop
End Function

Function CleanHTMLContent(htmlText As String) As String
    Dim result As String
    Dim i As Long
    Dim char As String
    Dim insideTag As Boolean
    
    result = ""
    insideTag = False
    
    ' Remove HTML tags and replace break elements with spaces
    For i = 1 To Len(htmlText)
        char = Mid(htmlText, i, 1)
            
        If char = "<" Then
            insideTag = True
            ' Add space for line break tags
            If i + 3 <= Len(htmlText) Then
                Select Case LCase(Mid(htmlText, i, 4))
                    Case "</p>", "<br>", "<br/"
                        result = result & " "
                End Select
            End If
        ElseIf char = ">" Then
            insideTag = False
        ElseIf Not insideTag Then
            result = result & char
        End If
    Next i
    
    ' Clean up line breaks and multiple spaces
    result = Replace(result, Chr(10), " ")
    result = Replace(result, Chr(13), " ")
    result = Replace(result, vbCrLf, " ")
    
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    CleanHTMLContent = result
End Function

Function ExtractWordCount(rowContent As String) As Long
    Dim wordPos As Long
    Dim pos As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim wordCountStr As String
    
    ExtractWordCount = 0
    
    ' Find rightmost occurrence of "words"
    wordPos = InStrRev(rowContent, "words", -1, vbTextCompare)
    
    If wordPos > 0 Then
        ' Work backwards to find the number
        pos = wordPos - 1
            
        ' Skip spaces
        Do While pos > 0 And Mid(rowContent, pos, 1) = " "
            pos = pos - 1
        Loop
            
        ' Find number boundaries
        endPos = pos
        Do While pos > 0 And IsNumeric(Mid(rowContent, pos, 1))
            pos = pos - 1
        Loop
        startPos = pos + 1
            
        ' Extract and validate number
        If startPos <= endPos Then
            wordCountStr = Mid(rowContent, startPos, endPos - startPos + 1)
            If IsNumeric(wordCountStr) And wordCountStr <> "" Then
                ExtractWordCount = CLng(wordCountStr)
            End If
        End If
    End If
End Function

Function ExtractExamDuration(description As String) As Integer
    Dim examPos As Long
    Dim hoursPos As Long
    Dim pos As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim durationStr As String
    
    ExtractExamDuration = 0
    
    ' Look for "exam" keyword
    examPos = InStr(1, description, "exam", vbTextCompare)
    If examPos > 0 Then
        ' Look for "hours" after "exam"
        hoursPos = InStr(examPos, description, "hours", vbTextCompare)
        If hoursPos > 0 Then
            ' Work backwards from "hours" to find number
            pos = hoursPos - 1
                
            ' Skip spaces
            Do While pos > 0 And Mid(description, pos, 1) = " "
                pos = pos - 1
            Loop
                
            ' Extract number
            endPos = pos
            Do While pos > 0 And IsNumeric(Mid(description, pos, 1))
                pos = pos - 1
            Loop
            startPos = pos + 1
                
            If startPos <= endPos Then
                durationStr = Mid(description, startPos, endPos - startPos + 1)
                If IsNumeric(durationStr) And durationStr <> "" Then
                    ExtractExamDuration = CInt(durationStr)
                    Exit Function
                End If
            End If
        End If
            
        ' Default duration if exam mentioned without specific hours
        ExtractExamDuration = 2
    End If
End Function

Function ExtractGroupSize(description As String) As Integer
    Dim groupsOfPos As Long
    Dim pos As Long
    Dim numberStr As String
    Dim dashPos As Long
    Dim startNum As Integer
    Dim endNum As Integer
    
    ExtractGroupSize = 0
    
    ' Look specifically for "groups of" pattern
    groupsOfPos = InStr(1, description, "groups of", vbTextCompare)
    If groupsOfPos > 0 Then
        ' Move to position after "groups of"
        pos = groupsOfPos + 9
            
        ' Skip spaces
        Do While pos <= Len(description) And Mid(description, pos, 1) = " "
            pos = pos + 1
        Loop
            
        ' Extract number or range
        If pos <= Len(description) And IsNumeric(Mid(description, pos, 1)) Then
            numberStr = ""
            Do While pos <= Len(description) And (IsNumeric(Mid(description, pos, 1)) Or Mid(description, pos, 1) = "-")
                numberStr = numberStr & Mid(description, pos, 1)
                pos = pos + 1
            Loop
                
            ' Parse single number or range
            dashPos = InStr(numberStr, "-")
            If dashPos > 0 Then
                ' Handle range (e.g., "3-5")
                If dashPos > 1 And dashPos < Len(numberStr) Then
                    startNum = CInt(Left(numberStr, dashPos - 1))
                    endNum = CInt(Mid(numberStr, dashPos + 1))
                        
                    ' Use start number for small ranges, middle for larger ranges
                    If endNum - startNum <= 2 Then
                        ExtractGroupSize = startNum
                    Else
                        ExtractGroupSize = Int((startNum + endNum) / 2)
                    End If
                End If
            Else
                ' Handle single number
                If IsNumeric(numberStr) And numberStr <> "" Then
                    ExtractGroupSize = CInt(numberStr)
                End If
            End If
        End If
    End If
End Function