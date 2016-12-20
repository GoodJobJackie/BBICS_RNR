Attribute VB_Name = "BBICS_RNR"

Public reportStart As Date
Public reportEnd As Date
Public ProgramName As String
Public ProgramDescription As String
Public ProgramSD As String
Public SkillName As String
Public mCm As String
Public skipFlag
Public reportStartRow
Public reportEndRow
Public flagFullRest
Public flagFullService
Public startDateRow As Integer
Public endDateRow As Integer
Public programCount As Integer
Public skip As Boolean
Public renameI As Integer
Public prevProgramName As Integer
Public BxDict As New Scripting.Dictionary
Public objFSO As Object
Public objFolder As Object
Public objFile As Object
Public x As Workbook


Dim dataSheetName As String
Dim objWord
Dim objDoc

Sub ARestructureAndGenerateReport()
Attribute ARestructureAndGenerateReport.VB_ProcData.VB_Invoke_Func = "r\n14"

    Dim bottomDate
    
    dataSheetName = ActiveSheet.Name
    bottomDateRow = Cells(5, 1).End(xlDown).Row
    
    If Range("A1:A3").MergeCells = True Then
        UserAction.ActionReformat.Enabled = False
    End If
    If flagFullRest = True Or Cells(4, 1).Value = "" Then
        UserAction.ActionRestuctureFull.Enabled = False
    End If
    If flagFullService = True Then
        UserAction.ActionFullService.Enabled = False
    End If
        
    UserAction.Show

End Sub

Sub NewRestructuring()
'
' NewRestructuring Macro
'
' Keyboard Shortcut: Ctrl+w
'
    Dim headerCol As Long
    Dim nextHeaderCol As Long
    Dim headerCell As Range
    Dim startTime As Double
    Dim minutesElapsed As String
    
    startTime = Timer
    headerCol = 2
    nextHeaderCol = 5
    
    ActiveWindow.Zoom = 90
    ' Create Initials Header A1:A3
    CreateHeader
    ' Check if Column B is empty and, if so, delete
    EmptyBCheck
    ' Select Column A, format and add edgeleft border
    MasterListFormat
    ' Select columns with program dates, format, add edgeleft borders
    FormatProgramDates
    ' Find longest date column and copy to Column A
    FindLastDate
    ' Set FreezePane at B4
    Cells(4, 2).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    ' Cycle through programs and move data to matching date
    MoveData
    ' Calculate elapsed time
    minutesElapsed = Format((Timer - startTime) / 86400, "hh:mm:ss")
    ' Create new sheet for population by current programs
    PopulatePrograms
    Cells(1, 1).Select
    ActiveWindow.Zoom = 90
    ' Write MCM to cells for copy/paste to progress report.
    CreateProgramLists
    ' Import Skill/Programs masters
    ImportSkillsPrograms
    ' Dialog box confirming completion
    MsgBox "Task complete in " & minutesElapsed, vbInformation + vbSystemModal + vbOKOnly, "Success"
End Sub

Sub RequestDates()

    reportStart = ""
    reportEnd = ""
    
    Do While Not IsDate(reportStart)
        reportStart = InputBox("Please input next report start date - m/d/yyyy")
        If IsDate(reportStart) Then
            reportStart = Format(CDate(reportStart), "m/d/yyyy")
        Else
            MsgBox "Please enter valid date."
        End If
    Loop
    
    Do While Not IsDate(reportEnd)
        reportEnd = InputBox("Please input next report end date - m/d/yyyy")
        If IsDate(reportEnd) Then
            reportEnd = Format(CDate(reportEnd), "m/d/yyyy")
        Else
            MsgBox "Please enter valid date."
        End If
    Loop
    
End Sub

Sub EmptyBCheck()

        For i = 1 To 1000
            If Cells(i, 2).Value = "" Then
                If Cells(i, 2).Value = "" Then
                    If i = 1000 Then
                        Columns("B:B").Select
                        Selection.Delete Shift:=xlToLeft
                    End If
                End If
            Else: Exit Sub
            End If
        Next i
        
End Sub

Sub CreateHeader()

    Cells.Select
        Range("BQ21").Activate
        Selection.ColumnWidth = 11
        Range("A1:A3").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Merge
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
        End With
        Selection.Font.Size = 18
        Selection.Font.Bold = True
        Selection.Font.Italic = True
End Sub

Sub MasterListFormat()

    Columns("A").Select
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.NumberFormat = "mm/dd/yyyy"
    
End Sub

Sub FormatProgramDates()

    For i = 1 To 1000
        If Cells(2, i).Value = "" Then
            Else
            Columns(i).Select
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            Selection.Borders(xlEdgeTop).LineStyle = xlNone
            Selection.Borders(xlEdgeBottom).LineStyle = xlNone
            Selection.Borders(xlEdgeRight).LineStyle = xlNone
            Selection.Borders(xlInsideVertical).LineStyle = xlNone
            Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
            Selection.NumberFormat = "mm/dd/yyyy"
        End If
    Next i
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.NumberFormat = "mm/dd/yyyy"
    
End Sub
Sub FindLastDate()

    For i = 1 To 1000
        If Cells(2, i).Value = "" Then
        Else
            For j = 4 To 2000
                If Cells(j, i).Value > abValue Then
                    prevabValue = abValue
                    abValue = Cells(j, i).Value
                    a = i
                    b = j
                End If
            Next j
        End If
    Next i
    
    Range(Cells(4, a), Cells(b, a)).Copy
    Cells(4, 1).Select
    ActiveSheet.Paste
    
End Sub

Sub MoveData()

    ' Find next program chunk
    For col = 2 To 1000
        If Cells(2, col).Value = "" Then
        Else
            If Cells(4, col).Value = "" Then
                Cells(2, col).Cut
                Cells(1, col).Paste
            End If
            ' Assign header cell to variable
            headerCell = Cells(2, col)
            ' Assign date column to variable
            headerCol = col
            ' Assign next header column to variable
            If Cells(2, headerCol).End(xlToRight).Value = "" Then
                nextHeaderCol = headerCol + 15
            Else
                Cells(2, headerCol).End(xlToRight).Select
                nextHeaderCol = Selection.Column
                nextHeaderCol = nextHeaderCol - 2
            End If
        ' Fill cell below header with " " for asthetic purposes
        Cells(3, col).Value = " "
           ' Move down rows and check if date is <, >, = to master date list
            For Row = 4 To 2000
                If Cells(Row, col).Value = "" Then
                    Exit For
                Else
                    ' If > then cut chunk and paste one below
                    If Cells(Row, col).Value > Cells(Row, 1) Then
                        If Cells((Row + 1), col) = "" Then
                            ' Catch last row
                            Range(Cells(Row, headerCol), Cells(Row, (nextHeaderCol))).Select
                            Selection.Cut
                            Cells((Row + 1), col).Select
                            ActiveSheet.Paste
                        Else
                            ' Ctrl+down + (next header column - 2)
                            Range(Cells(Row, headerCol), Cells(Selection.End(xlDown), nextHeaderCol)).Select
                            Selection.Cut
                            Cells((Row + 1), col).Select
                            ActiveSheet.Paste
                        End If
                    Cells(Row, col).Select
                    ' If < then insert new date in master list
                    ElseIf Cells(Row, col).Value < Cells(Row, 1) Then
                        ' Catch last row
                        
                        ' Select all previous program chunks at date and below
                        Range(Cells(Row, 1), Cells(Selection.End(xlDown), (headerCol - 1))).Select
                        ' Cut and paste down one row
                        Selection.Cut
                        Cells((Row + 1), 1).Select
                        ActiveSheet.Paste
                        ' Copy new date to master list
                        Cells(Row, col).Select
                        Selection.Copy
                        Cells(Row, 1).Select
                        ActiveSheet.Paste
                    ' If = then move to next row
                    End If
                End If
                ' Continue until no more dates
            Next Row
        End If
    Next col
    
End Sub

Sub SelectChunk()
    '
    ' SelectChunk Macro
    '
    ' Keyboard Shortcut: Ctrl+m
    '
    chunkHeader = 2
    chunkHeader = Selection.col
    nextHeader = Cells(2, (chunkHeader.End(xlToRight) - 2))
    programChunk = Range(Selection, (Cells(Selection.End(xlDown), nextHeader)))
    programChunk.Select
    
End Sub

Sub PopulatePrograms()

    Dim programRow
    Dim skillCount
    Dim skillStart As Date
    Dim skillEnd As Date
    Dim skillDate
    Dim skillStartRow
    Dim skillEndRow
    Dim lastSkillEnd As Date
    Dim lastSkillRow As Integer
    Dim programSheet As String
    Dim bottomDateRow
    Dim prev1 As Integer
    Dim prev2 As Integer
    Dim prev3 As Integer
    Dim prevRow As Integer

    col = 2
    skillCount = 1
    programRow = 2
    dataSheetName = ActiveSheet.Name
    programSheet = "Programs"
    bottomDateRow = Cells(5, 1).End(xlDown).Row
                
    UserForm_Initialize
    UserForm1.Show
    
    ' Create and format program/skill sheet
    Application.DisplayAlerts = False
    For Each Sheet In Worksheets
        If programSheet = Sheet.Name Then
            Sheet.Delete
        End If
    Next Sheet
    Application.DisplayAlerts = True
    
    Worksheets.Add().Name = programSheet
    Worksheets("Programs").Cells(1, 1).Value = "Program"
    Worksheets("Programs").Cells(1, 2).Value = "Skill"
    Worksheets("Programs").Cells(1, 3).Value = "Mastered"
    Worksheets("Programs").Cells(1, 4).Value = "Continued"
    Worksheets("Programs").Cells(1, 5).Value = "Maintenance"
    Worksheets("Programs").Columns("A:B").ColumnWidth = 60
    Worksheets("Programs").Columns("C:E").ColumnWidth = 12
    Worksheets(dataSheetName).Activate
    ' Look for next program
    For col = 2 To 1000
        If Cells(2, col) <> "" Then
            headerCell = Cells(2, col)
            headerCol = col
            If Cells(2, headerCol).End(xlToRight).Value = "" Then
                If Cells(3, headerCol + 2).Value = "" Then
                    nextHeaderCol = col + 1
                Else
                    nextHeaderCol = Cells(3, (col + 1)).End(xlToRight).Column
                End If
            Else
                Cells(2, headerCol).End(xlToRight).Select
                nextHeaderCol = Selection.Column
                nextHeaderCol = nextHeaderCol - 2
            End If
            'Cycle through skills
            For i = (col + 1) To nextHeaderCol
                ' Check for empty skill columns
                If Cells(3, i).End(xlDown).Value = "" Then GoTo NextIteration
                
                ProgramName = Cells(2, col).Value
                SkillName = Cells(3, i).Value
                Range(Cells(3, i).End(xlDown), Cells(1000, i).End(xlUp)).Select
                ActiveWindow.Zoom = True
                ActiveWindow.Zoom = 90
                
                UserForm_Initialize
                
                ' Store skill start/end dates as variables
                If Cells(4, i).Value = "" Then
                    skillStartRow = Cells(4, i).End(xlDown).Row
                    skillStart = Cells(skillStartRow, col).Value
                Else
                    skillStart = Cells(4, col).Value
                End If
                
                'Check and store most current skill ending date
                Cells(1000, i).End(xlUp).Select
                skillEnd = Selection.End(xlToLeft).Value
                If Cells(4, nextHeaderCol).Value = "" Then
                    
                    lastSkillRow = Cells(1000, nextHeaderCol).End(xlUp).Row
                    lastSkillEnd = Cells(lastSkillRow, col).Value
                Else
                    lastSkillEnd = Cells(4, col).Value
                End If
                
                ' Check for skill within report dates
                If skillEnd < reportStart And reportStart - lastSkillEnd < 182 Then
                ' Do nothing
                ElseIf skillStart > reportEnd Then
                ' Also do nothing
                ElseIf reportStart - lastSkillEnd > 182 Then
                ' Check for new maintenance program and add to maintenance list if not already done
                    If Cells(2, headerCol).Font.Color = RGB(166, 166, 166) Then
                        'Do nothing
                    Else
                        Worksheets("Programs").Cells(programRow, 1).Value = ProgramName
                        Worksheets("Programs").Cells(programRow, 5).Value = "X"
                        programRow = programRow + 1
                        Cells(2, col).Value = ProgramName
                        Cells(2, headerCol).Font.Color = RGB(166, 166, 166)
                    End If
                Else
                    ' Open MCM dialog box
                    Dim k As Integer
                    Dim l As Long
                    k = Cells(1000, i).End(xlUp).Row
                    l = Cells(4, i).End(xlDown).Row
                    Range(Cells(l, i), Cells(k, i)).Activate
                    MCMbox.Show
                    Worksheets("Programs").Cells(programRow, 1).Value = Trim(ProgramName)
                    Select Case mCm
                        Case Is = "Mastered"
                            Worksheets("Programs").Cells(programRow, 2).Value = SkillName
                            Worksheets("Programs").Cells(programRow, 3).Value = "X"
                        Case Is = "Continued"
                            Worksheets("Programs").Cells(programRow, 2).Value = SkillName
                            Worksheets("Programs").Cells(programRow, 4).Value = "X"
                        Case Is = "Maintenance"
                            Worksheets("Programs").Cells(programRow, 5).Value = "X"
                            Cells(2, headerCol).Font.Color = RGB(166, 166, 166)
                    End Select
                    programRow = programRow + 1
                    Cells(2, col).Value = ProgramName
                End If
NextIteration:
            Next i
        End If
    Next col
    
    ' Reset report period borders to black/white
     For i = 4 To bottomDateRow
        If reportStart = Trim(Cells(i, 1).Value) Then
            Rows(i).Select
            With Selection.Borders(xlEdgeTop)
                .Color = RGB(0, 0, 0)
                .LineStyle = xlContinuous
            End With
            startDateRow = i
        End If
        If Trim(Cells(i, 1).Value) = reportEnd Then
            Rows(i).Select
            With Selection.Borders(xlEdgeBottom)
                .Color = RGB(0, 0, 0)
                .LineStyle = xlContinuous
            End With
            endDateRow = i
        End If
    Next i
    Rows(startDateRow & ":" & endDateRow).Select
    Selection.Interior.Color = -4142
    
End Sub

Public Sub UserForm_Initialize()

    Dim suggestStart
    Dim suggestEnd
    Dim placeHolder
    Dim reportDates As String
    
    MCMbox.programNameBox.Value = Trim(ProgramName)
    MCMbox.skillNameBox.Value = Trim(SkillName)
    MCMbox.progMast.Value = True
    MCMbox.NextProgram.SetFocus
    
    
    reportDates = Worksheets("CI").Cells(2, 6).Value & " " & Worksheets("CI").Cells(2, 7).Value & _
        " - " & Worksheets("CI").Cells(2, 8).Value & " " & Worksheets("CI").Cells(2, 9).Value
    UserForm1.reportDates.Value = reportDates
    With UserForm1.ComboBox1
        For i = 4 To Worksheets(dataSheetName).Cells(5, 1).End(xlDown).Row
            .AddItem Worksheets(dataSheetName).Cells(i, 1).Value
        Next i
    End With
    
    With UserForm1.ComboBox2
        For i = 4 To Worksheets(dataSheetName).Cells(5, 1).End(xlDown).Row
            .AddItem Worksheets(dataSheetName).Cells(i, 1).Value
        Next i
    End With
    
    Select Case Worksheets("CI").Cells(2, 6).Value
        Case Is = "January"
            suggestStart = "1/1/" & Worksheets("CI").Cells(2, 7).Value
        Case Is = "February"
            suggestStart = "2/1/" & Worksheets("CI").Cells(2, 7).Value
        Case Is = "March"
            suggestStart = "3/1/" & Worksheets("CI").Cells(2, 7).Value
        Case Is = "April"
            suggestStart = "4/1/" & Worksheets("CI").Cells(2, 7).Value
        Case Is = "May"
            suggestStart = "5/1/" & Worksheets("CI").Cells(2, 7).Value
        Case Is = "June"
            suggestStart = "6/1/" & Worksheets("CI").Cells(2, 7).Value
        Case Is = "July"
            suggestStart = "7/1/" & Worksheets("CI").Cells(2, 7).Value
        Case Is = "August"
            suggestStart = "8/1/" & Worksheets("CI").Cells(2, 7).Value
        Case Is = "September"
            suggestStart = "9/1/" & Worksheets("CI").Cells(2, 7).Value
        Case Is = "October"
            suggestStart = "10/1/" & Worksheets("CI").Cells(2, 7).Value
        Case Is = "November"
            suggestStart = "11/1/" & Worksheets("CI").Cells(2, 7).Value
        Case Is = "December"
            suggestStart = "12/1/" & Worksheets("CI").Cells(2, 7).Value
    End Select
        
    Select Case Worksheets("CI").Cells(2, 8).Value
        Case Is = "January"
            suggestEnd = "1/31/" & Worksheets("CI").Cells(2, 9).Value
        Case Is = "February"
            suggestEnd = "2/28/" & Worksheets("CI").Cells(2, 9).Value
        Case Is = "March"
            suggestEnd = "3/31/" & Worksheets("CI").Cells(2, 9).Value
        Case Is = "April"
            suggestEnd = "4/30/" & Worksheets("CI").Cells(2, 9).Value
        Case Is = "May"
            suggestEnd = "5/31/" & Worksheets("CI").Cells(2, 9).Value
        Case Is = "June"
            suggestEnd = "6/30/" & Worksheets("CI").Cells(2, 9).Value
        Case Is = "July"
            suggestEnd = "7/31/" & Worksheets("CI").Cells(2, 9).Value
        Case Is = "August"
            suggestEnd = "8/31/" & Worksheets("CI").Cells(2, 9).Value
        Case Is = "September"
            suggestEnd = "9/30/" & Worksheets("CI").Cells(2, 9).Value
        Case Is = "October"
            suggestEnd = "10/31/" & Worksheets("CI").Cells(2, 9).Value
        Case Is = "November"
            suggestEnd = "11/30/" & Worksheets("CI").Cells(2, 9).Value
        Case Is = "December"
            suggestEnd = "12/31/" & Worksheets("CI").Cells(2, 9).Value
    End Select

    For i = 4 To Worksheets("Data").Cells(4, 1).End(xlDown).Row
        If Trim(Worksheets("Data").Cells(i, 1).Value) < suggestStart And Trim(Worksheets("Data").Cells(i + 1, 1).Value) > suggestStart Then
            suggestStart = Trim(Worksheets("Data").Cells(i + 1, 1).Value)
        End If
        If Trim(Worksheets("Data").Cells(i, 1).Value) < suggestEnd And Trim(Worksheets("Data").Cells(i + 1, 1).Value) > suggestEnd Then
            suggestEnd = Trim(Worksheets("Data").Cells(i, 1).Value)
        End If
    Next i
      
    UserForm1.ComboBox1.Value = suggestStart
    UserForm1.ComboBox2.Value = suggestEnd
      
End Sub


Sub CreateProgramLists()

    Dim bottomProgramRow
    Dim countMast
    Dim countCont
    Dim countMaint
    
    countMast = 1
    countCont = 1
    countMaint = 1
    
    Worksheets("Programs").Activate
    
    bottomProgramRow = Cells(1, 1).End(xlDown).Row
    For Row = 2 To bottomProgramRow
        If Cells(Row, 3).Value = "X" Then
            If countMast = 1 Then
                Cells(1, 6).Value = Cells(1, 6) & countMast & ") " & Cells(Row, 1).Value & " (" & Cells(Row, 2).Value & ")"
                countMast = countMast + 1
            Else
                Cells(1, 6).Value = Cells(1, 6) & ", " & countMast & ") " & Cells(Row, 1).Value & " (" & Cells(Row, 2).Value & ")"
                countMast = countMast + 1
            End If
        End If
        If Cells(Row, 4).Value = "X" Then
            If countCont = 1 Then
                Cells(1, 7).Value = Cells(1, 7) & countCont & ") " & Cells(Row, 1).Value & " (" & Cells(Row, 2).Value & ")"
                countCont = countCont + 1
            Else
                Cells(1, 7).Value = Cells(1, 7) & ", " & countCont & ") " & Cells(Row, 1).Value & " (" & Cells(Row, 2).Value & ")"
                countCont = countCont + 1
            End If
        End If
        If Cells(Row, 5).Value = "X" Then
            If countMaint = 1 Then
                Cells(1, 8).Value = Cells(1, 8) & countMaint & ") " & Cells(Row, 1).Value
                countMaint = countMaint + 1
            Else
                Cells(1, 8).Value = Cells(1, 8) & ", " & countMaint & ") " & Cells(Row, 1).Value
                countMaint = countMaint + 1
            End If
        End If
    Next Row
    
    If Cells(1, 6).Value = "" Then
        Cells(1, 6).Value = "N/A."
    Else
        Cells(1, 6).Value = Cells(1, 6).Value & "."
    End If
    
    If Cells(1, 7).Value = "" Then
        Cells(1, 7).Value = "N/A."
    Else
        Cells(1, 7).Value = Cells(1, 7).Value & "."
    End If
    If Cells(1, 8).Value = "" Then
        Cells(1, 8).Value = "N/A."
    Else
        Cells(1, 8).Value = Cells(1, 8).Value & "."
    End If

End Sub

Sub SaveReport()

    Dim saveAsFilename As String
    
    saveAsFilename = Worksheets("Data").Cells(1, 1).Value & " - " & _
        Worksheets("CI").Cells(2, 4).Value & " Progress Report [" & _
        Format(reportStart, "yyyy") & "_" & Format(reportStart, "mm") & " - " & _
        Format(reportEnd, "yyyy") & "_" & Format(reportEnd, "mm") & "]"
    answer = MsgBox("Save report as: " & saveAsFilename, vbYesNo + vbQuestion)
    If answer = vbYes Then
        objDoc.SaveAs ("C:\Users\jackie\Documents\Client Files\Progress Reports\" & saveAsFilename)
    End If
    
End Sub

Sub SingleRestructure()
Attribute SingleRestructure.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' SingleRestructure Macro
'
' Keyboard Shortcut: Ctrl+r
'
Dim col

col = ActiveCell.Column

If Cells(2, col).Value = "" Then
    Else
        ' Assign header cell to variable
        headerCell = Cells(2, col)
        ' Assign date column to variable
        headerCol = col
        ' Assign next header column to variable
        If Cells(2, headerCol).End(xlToRight).Value = "" Then
            nextHeaderCol = headerCol + 15
        Else
            Cells(2, headerCol).End(xlToRight).Select
            nextHeaderCol = Selection.Column
            nextHeaderCol = nextHeaderCol - 2
        End If
    ' Fill cell below header with " " for asthetic purposes
    Cells(3, col).Value = " "
       ' Move down rows and check if date is <, >, = to master date list
        For Row = 4 To 2000
            If Cells(Row, col).Value = "" Then
                Exit For
            Else
                ' If > then cut chunk and paste one below
                If Cells(Row, col).Value > Cells(Row, 1) Then
                    If Cells((Row + 1), col) = "" Then
                        ' Catch last row
                        Range(Cells(Row, headerCol), Cells(Row, (nextHeaderCol))).Select
                        Selection.Cut
                        Cells((Row + 1), col).Select
                        ActiveSheet.Paste
                    Else
                        ' Ctrl+down + (next header column - 2)
                        Range(Cells(Row, headerCol), Cells(Selection.End(xlDown), nextHeaderCol)).Select
                        Selection.Cut
                        Cells((Row + 1), col).Select
                        ActiveSheet.Paste
                    End If
                Cells(Row, col).Select
                ' If < then insert new date in master list
                ElseIf Cells(Row, col).Value < Cells(Row, 1) Then
                    ' Select all previous program chunks at date and below
                    Range(Cells(Row, 1), Cells(Selection.End(xlDown), (headerCol - 1))).Select
                    ' Cut and paste down one row
                    Selection.Cut
                    Cells((Row + 1), 1).Select
                    ActiveSheet.Paste
                    ' Copy new date to master list
                    Cells(Row, col).Select
                    Selection.Copy
                    Cells(Row, 1).Select
                    ActiveSheet.Paste
                ' If = then move to next row
                End If
            End If
            ' Continue until no more dates
        Next Row
    End If
    
End Sub

Sub ImportSkillsPrograms()

    Dim z As Workbook
    Dim w As Workbook
    Dim x As Workbook
    Dim y As Workbook
    Dim k As Integer
    
    Worksheets.Add().Name = "CI"
    Worksheets.Add().Name = "SDL"
    Worksheets.Add().Name = "PD"
      
    Set z = ActiveWorkbook
    Set w = Workbooks.Open("C:\Users\jackie\Documents\Client Files\Progress Reports\FMP_DataExport\FMP_DataExport.xlsx")
    Set x = Workbooks.Open("C:\Users\jackie\Documents\Client Files\Progress Reports\FMP_DataExport\SkillDeficitList.xlsx")
    Set y = Workbooks.Open("C:\Users\jackie\Documents\Client Files\Progress Reports\FMP_DataExport\ProgramDescriptions.xlsx")
    
    Application.DisplayAlerts = False
    
    w.Sheets("CI").Range("A1:M2").Copy
    z.Sheets("CI").Range("A1:M2").PasteSpecial
    w.Close
    z.Sheets("CI").Columns("A:M").AutoFit
    
    x.Sheets("SDL").Range("A1:B112").Copy
    z.Sheets("SDL").Range("A1:B112").PasteSpecial
    x.Close
    
    k = Worksheets("PD").Cells(1000, 1).End(xlUp).Row
    y.Sheets("PD").Range("A1:C" & k).Copy
    z.Sheets("PD").Range("A1:C" & k).PasteSpecial
    y.Close
    
    Worksheets("Data").Activate

    Application.DisplayAlerts = True

End Sub

Sub PopulateReport()

    Dim chunks As Integer
    Dim chunk As String
    Dim objRange
    Dim currentBottomRow As Integer
    Dim s As Object
    Dim bx As Variant
    Dim bxCount As Integer
    Dim bxString As String

    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Open("C:\Users\jackie\Documents\Client Files\Progress Reports\FMP_DataExport\PRT.docx")
    objWord.Visible = True
    
    With objDoc
        For Each s In .Sections
            With s.Headers(wdHeaderFooterPrimary).Range.Find
                .Forward = True
                .Text = "[clientName]"
                .Replacement.Text = Worksheets("CI").Cells(2, 1).Value & " " & Worksheets("CI").Cells(2, 2).Value
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
        Next
    End With
    
    With objDoc.Content.Find
        .Forward = True
        .Text = "[clientName]"
        .Replacement.Text = Worksheets("CI").Cells(2, 1).Value & " " & Worksheets("CI").Cells(2, 2).Value
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With objDoc.Content.Find
        .Forward = True
        .Text = "[firstName]"
        .Replacement.Text = Worksheets("CI").Cells(2, 1).Value
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With objDoc.Content.Find
        .Forward = True
        .Text = "[mqba]"
        .Replacement.Text = Worksheets("CI").Cells(2, 4).Value
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With objDoc.Content.Find
        .Forward = True
        .Text = "[begDate]"
        .Replacement.Text = Worksheets("CI").Cells(2, 6).Value & " " & Worksheets("CI").Cells(2, 7).Value
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With objDoc.Content.Find
        .Forward = True
        .Text = "[endDate]"
        .Replacement.Text = Worksheets("CI").Cells(2, 8).Value & " " & Worksheets("CI").Cells(2, 9).Value
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With objDoc.Content.Find
        .Forward = True
        .Text = "[parentGuardian]"
        .Replacement.Text = Worksheets("CI").Cells(2, 3).Value
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With objDoc.Content.Find
        .Forward = True
        .Text = "[initialDate]"
        .Replacement.Text = Worksheets("CI").Cells(2, 10).Value
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With objDoc.Content.Find
        .Forward = True
        .Text = "[assessmentDate]"
        .Replacement.Text = Worksheets("CI").Cells(2, 11).Value
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With objDoc.Content.Find
        .Forward = True
        .Text = "[skillDeficitList]"
        .Replacement.Text = Worksheets("SDL").Cells(1, 1).Value & vbCr & vbCr & "[next]"
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    For i = 2 To 112
        With objDoc.Content.Find
            .Forward = True
            .Text = "[next]"
            .Replacement.Text = Worksheets("SDL").Cells(i, 1).Value & vbCr & vbCr & "[next]"
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    Next i
    With objDoc.Content.Find
        .Forward = True
        .Text = "[next]"
        .Replacement.Text = ""
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With objDoc.Content.Find
        .Forward = True
        .Text = "[parentConcerns]"
        .Replacement.Text = Worksheets("CI").Cells(2, 12).Value
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With objDoc.Content.Find
        .Forward = True
        .Text = "[parentTraining]"
        .Replacement.Text = Worksheets("CI").Cells(2, 13).Value
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    ' Divide up mastered program list into copyable chunks
    With objDoc.Content.Find
        .ClearFormatting
        chunks = Round(Len(Worksheets("Programs").Cells(1, 6).Value) / 250, 0)
        If Len(Worksheets("Programs").Cells(1, 6).Value) Mod 250 > 0 Then chunks = chunks + 1
        If chunks = 1 Then
            .Forward = True
            .Text = "[masteredPrograms]"
            .Replacement.Text = Worksheets("Programs").Cells(1, 6).Value
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        Else
            For i = 1 To chunks
                chunk = Mid(Worksheets("Programs").Cells(1, 6).Value, ((i - 1) * 250) + 1, 250)
                If i = 1 Then
                    .Forward = True
                    .Text = "[masteredPrograms]"
                    .Replacement.Text = chunk & "[nx]"
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                Else
                    .Forward = True
                    .Text = "[nx]"
                    .Replacement.Text = chunk & "[nx]"
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                End If
            Next i
        End If
    End With
    With objDoc.Content.Find
        .Forward = True
        .Text = "[nx]"
        .Replacement.Text = ""
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Divide up continued program list into copyable chunks
    With objDoc.Content.Find
        .ClearFormatting
        chunks = Round(Len(Worksheets("Programs").Cells(1, 7).Value) / 250, 0)
        If Len(Worksheets("Programs").Cells(1, 7).Value) Mod 250 > 0 Then chunks = chunks + 1
        If chunks = 1 Then
            .Forward = True
            .Text = "[continuedPrograms]"
            .Replacement.Text = Worksheets("Programs").Cells(1, 7).Value
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        Else
            For i = 1 To chunks
                chunk = Mid(Worksheets("Programs").Cells(1, 7).Value, ((i - 1) * 250) + 1, 250)
                If i = 1 Then
                    .Forward = True
                    .Text = "[continuedPrograms]"
                    .Replacement.Text = chunk & "[nx]"
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                Else
                    .Forward = True
                    .Text = "[nx]"
                    .Replacement.Text = chunk & "[nx]"
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                End If
            Next i
        End If
    End With
    With objDoc.Content.Find
        .Forward = True
        .Text = "[nx]"
        .Replacement.Text = ""
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Divide up maintenance program list into copyable chunks
    With objDoc.Content.Find
        .ClearFormatting
        chunks = Round(Len(Worksheets("Programs").Cells(1, 8).Value) / 250, 0)
        If Len(Worksheets("Programs").Cells(1, 8).Value) Mod 250 > 0 Then chunks = chunks + 1
        If chunks = 1 Then
            .Forward = True
            .Text = "[maintPrograms]"
            .Replacement.Text = Worksheets("Programs").Cells(1, 8).Value
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        Else
            For i = 1 To chunks
                chunk = Mid(Worksheets("Programs").Cells(1, 8).Value, ((i - 1) * 250) + 1, 250)
                If i = 1 Then
                    .Forward = True
                    .Text = "[maintPrograms]"
                    .Replacement.Text = chunk & "[nx]"
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                Else
                    .Forward = True
                    .Text = "[nx]"
                    .Replacement.Text = chunk & "[nx]"
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                End If
            Next i
        End If
    End With
    With objDoc.Content.Find
        .Forward = True
        .Text = "[nx]"
        .Replacement.Text = ""
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    ProgramDescriptionsList
    ProgramMatch
    
    Worksheets("Current").Activate
    currentBottomRow = Worksheets("Current").Cells(1, 1).End(xlDown).Row
    
    'Write program/descriptions to table
    Set objRange = objDoc.Range
    Set objTable = objDoc.Tables(1)
        For i = 2 To currentBottomRow
            If Worksheets("Current").Cells(i, 4).Value = "skip" Then
            Else
                objTable.Cell(i, 1).Range.Text = Worksheets("Current").Cells(i, 1).Value
                objTable.Cell(i, 2).Range.Text = Worksheets("Current").Cells(i, 2).Value
                objTable.Cell(i, 3).Range.Text = Worksheets("Current").Cells(i, 3).Value
                If i = currentBottomRow Then
                Else
                    objTable.Rows.Add
                End If
            End If
        Next i
        
    'Sort bx data
    BxData
    'Write bx data to report
    bxCount = BxDict.Count
    For Each bx In BxDict
        If bxCount = BxDict.Count Then
            bxString = ", and " & bxCount & ") " & BxDict(bx) & " counts of " & bx & "."
        ElseIf bxCount = 1 Then
            bxString = bxCount & ") " & BxDict(bx) & " counts of " & bx & bxString
        Else
            bxString = ", " & bxCount & ") " & BxDict(bx) & " counts of " & bx & bxString
        End If
        bxCount = bxCount - 1
    Next bx
          
        With objDoc.Content.Find
        .ClearFormatting
        chunks = Round(Len(bxString) / 250, 0)
        If Len(bxString) Mod 250 > 0 Then chunks = chunks + 1
        If chunks = 1 Then
            .Forward = True
            .Text = "[bx]"
            .Replacement.Text = bxString
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        Else
            For i = 1 To chunks
                chunk = Mid(bxString, ((i - 1) * 250) + 1, 250)
                If i = 1 Then
                    .Forward = True
                    .Text = "[bx]"
                    .Replacement.Text = chunk & "[nx]"
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                Else
                    .Forward = True
                    .Text = "[nx]"
                    .Replacement.Text = chunk & "[nx]"
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                End If
            Next i
        End If
    End With
    
    With objDoc.Content.Find
        .Forward = True
        .Text = "[nx]"
        .Replacement.Text = ""
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    TutorHrs
    'SaveReport
    
    Application.DisplayAlerts = False
    Worksheets("PD").Delete
    Worksheets("SDL").Delete
    Worksheets("CI").Delete
    Worksheets("Programs").Delete
    Worksheets("Current").Delete
    Application.DisplayAlerts = True
    
    objWord.Application.Activate
    objWord.Application.WindowState = wdWindowStateMaximize
    'objWord.ActiveDocument.SaveAs fileName:="C:\Users\jackie\Documents\Client Files\Progress Reports\"

End Sub

Sub ProgramDescriptionsList()

    Worksheets.Add().Name = "Current"
    Worksheets("Data").Activate
    dataSheetName = ActiveSheet.Name
    programCount = 1
    
    For i = 2 To 1000
        If Cells(2, i).Value = "" Then
            ' Do nothing
        Else
            ProgramName = Trim(Cells(2, i).Value)
            Worksheets("Current").Cells(programCount, 1).Value = ProgramName
            programCount = programCount + 1
        End If
    Next i

End Sub

Sub ProgramMatch()

    Dim found As Boolean

    For i = 1 To Worksheets("Current").Cells(1, 1).End(xlDown).Row
        ProgramName = Trim(Worksheets("Current").Cells(i, 1).Value)
        found = False
        For scan = 2 To Worksheets("PD").Cells(1000, 1).End(xlUp).Row
            If ProgramName = Trim(Worksheets("PD").Cells(scan, 1).Value) Then
                Worksheets("Current").Cells(i, 2).Value = Worksheets("PD").Cells(scan, 2).Value
                Worksheets("Current").Cells(i, 3).Value = Worksheets("PD").Cells(scan, 3).Value
                found = True
            Else
            End If
        Next scan
        If found = True Then
        Else
            ' Prompt for program description.
            With DescriptionBox.programSuggestion
                For j = 2 To Worksheets("PD").Cells(2000, 1).End(xlUp).Row
                    .AddItem Worksheets("PD").Cells(j, 1).Value
                 Next j
            End With

            skip = False
            DescriptionBox.CommandButton3.Enabled = True
            DescriptionBox.currentProgramName.Value = ProgramName
            DescriptionBox.programNameBox.Value = ProgramName
            DescriptionBox.programDescriptionBox = "The client will "
            DescriptionBox.programSDBox.Value = "The SD will vary depending on the instructions in the workbook, the client�s IEP, and the interventionist."
            For j = 2 To x.Worksheets("Data").Cells(2, 3000).End(xlToLeft).Column
                If Worksheets("Data").Cells(2, j).Value = ProgramName Then
                    Worksheets("Data").Cells(2, j).Activate
                End If
            Next j
            DescriptionBox.Show
            
            If skip = True Then
                    Worksheets("Current").Cells(i, 4).Value = "skip"
                Else
                    Worksheets("Current").Cells(i, 1).Value = ProgramName
                    Worksheets("Current").Cells(i, 2).Value = ProgramDescription
                    Worksheets("Current").Cells(i, 3).Value = ProgramSD
            End If
        End If
    Next i

End Sub

Sub BxData()

    Dim bxRow As Integer
    Dim bxColStart As Integer
    Dim bxColEnd As Integer
    
    Dim Arr() As Variant
    Dim Temp1 As Variant
    Dim Temp2 As Variant
    Dim Txt As String
    Dim i As Long
    Dim j As Long
    
    Worksheets("Bx Data").Activate
    For i = 3 To Cells(4, 1).End(xlDown).Row
        If Cells(i, 1).Value = reportStart Then
            bxRow = i
        End If
    Next i
    bxColStart = (Cells(2, 2).End(xlToRight).Column) + 2
    bxColEnd = Cells(2, bxColStart).End(xlToRight).Column
    
    For i = bxColStart To bxColEnd
        If Cells(bxRow, i).Value = 0 Then
        Else
            BxDict.Add Cells(2, i).Value, Cells(bxRow, i).Value
        End If
    Next i
    
    ReDim Arr(0 To BxDict.Count - 1, 0 To 1)
    
    For i = 0 To BxDict.Count - 1
        Arr(i, 0) = BxDict.Keys(i)
        Arr(i, 1) = BxDict.Items(i)
    Next i
    
    For i = LBound(Arr, 1) To UBound(Arr, 1) - 1
        For j = i + 1 To UBound(Arr, 1)
            If Arr(i, 1) > Arr(j, 1) Then
                Temp1 = Arr(j, 0)
                Temp2 = Arr(j, 1)
                Arr(j, 0) = Arr(i, 0)
                Arr(j, 1) = Arr(i, 1)
                Arr(i, 0) = Temp1
                Arr(i, 1) = Temp2
            End If
        Next j
    Next i
    
    BxDict.RemoveAll
    
    For i = LBound(Arr, 1) To UBound(Arr, 1)
        BxDict.Add Key:=Arr(i, 0), Item:=Arr(i, 1)
    Next i
    
        For i = 0 To BxDict.Count - 1
        Txt = Txt & BxDict.Keys(i) & vbTab & BxDict.Items(i) & vbCrLf
    Next i
    
End Sub

Sub RenamePrograms()
   
   prevProgramName = 2
   
    For renameI = 2 To 1000
        If Cells(2, renameI).Value = "" Then
            ' Do nothing
        Else
            Cells(2, renameI).Activate
            ProgramName = Trim(Cells(2, renameI).Value)
            ProgramNames.currentProgramName.Value = ProgramName
            ProgramNames.NewProgramName.Value = ProgramName
            With ProgramNames.ProgramLists
                For j = 2 To Worksheets("PD").Cells(1000, 1).End(xlUp).Row
                    .AddItem Worksheets("PD").Cells(j, 1).Value
                Next j
            End With
            ProgramNames.NewProgramName.SetFocus
                With ProgramNames.NewProgramName
                    .SelStart = 0
                    .SelLength = Len(.Text)
                End With
            ProgramNames.Show
            Cells(2, renameI).Value = ProgramName
            prevProgramName = renameI
        End If
    Next renameI

End Sub

Sub TutorHrs()

    Dim tutorHrRow As Integer
    Dim tutorHrDate As String
    Dim monthCount As Integer
    
    Worksheets("Tutor Hr Data").Activate
    reportStart = Format(reportStart, "MMM yyyy")
    For i = 3 To Cells(3, 1).End(xlDown).Row
        If Cells(i, 1).Value = reportStart Then
            tutorHrRow = i
        End If
    Next i

    Set objRange = objDoc.Range
    Set objTable = objDoc.Tables(2)
    Select Case Worksheets("CI").Cells(2, 4).Value
        Case Is = "Quarterly"
            monthCount = tutorHrRow + 2
            j = 2
            For i = tutorHrRow To monthCount
                objTable.Cell(j, 1).Range.Text = Format(Cells(i, 1).Value, "MMMM yyyy")
                objTable.Cell(j, 2).Range.Text = Cells(i, 2).Value
                If i = monthCount Then
                Else
                    objTable.Rows.Add
                End If
                j = j + 1
            Next i
        Case Is = "Biannual"
            monthCount = tutorHrRow + 5
            j = 2
            For i = tutorHrRow To monthCount
                objTable.Cell(j, 1).Range.Text = Format(Cells(i, 1).Value, "MMMM yyyy")
                objTable.Cell(j, 2).Range.Text = Cells(i, 2).Value
                If i = monthCount Then
                Else
                    objTable.Rows.Add
                End If
                j = j + 1
            Next i
        Case Is = "Annual"
            monthCount = tutorHrRow + 11
            j = 2
            For i = tutorHrRow To monthCount
                objTable.Cell(j, 1).Range.Text = Format(Cells(i, 1).Value, "MMMM yyyy")
                objTable.Cell(j, 2).Range.Text = Cells(i, 2).Value
                If i = monthCount Then
                Else
                    objTable.Rows.Add
                End If
                j = j + 1
            Next i
        End Select

End Sub

Sub DataEntryPrograms()
       
    DataEntryBox.SessionDate.SetFocus
    DataEntryBox.buttonNextData.Default = True
    DataEntryBox.ProgramList = "Select preexisting program..."
    DataEntryBox.SkillList = "<Please select program first>"
    DataEntryBox.SkillList.Enabled = False
    DataEntryBox.AddProgram.Enabled = False
    DataEntryBox.AddSkill.Enabled = False
    DataEntryBox.buttonNextData.Enabled = False
    
        
        For col = 2 To Worksheets("Data").Cells(2, 1000).End(xlToLeft).Column
            If Worksheets("Data").Cells(2, col).Value = "" Then
            Else
                DataEntryBox.ProgramList.AddItem Cells(2, col).Value
            End If
        Next col
       
       
    DataEntryBox.Show

End Sub

