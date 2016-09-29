Attribute VB_Name = "Module4"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                           '
' BBICS Client Data Restucturing and Report Generation                      '
' Written by Jacqueline Flores                                              '
' goodjobjacqueline@gmail.com                                               '
'                                                                           '
' 2016_09_29  v2.1                                                          '
' -Added dialog box for modularization                                      '
'                                                                           '
' 2016_09_23    v2.0                                                        '
' -Added program population for report generation.                          '
'                                                                           '
' 2016_09_16    v1.2.1                                                      '
' -Cleaned up code                                                          '
'                                                                           '
' 2016_09_06    v1.2                                                        '
' -Added time elapsed message box upon completion.                          '
'                                                                           '
' 2016_08_19    v1.1                                                        '
' -Increased range of row/column checks                                     '
'                                                                           '
' 2016_08_19    v1.0                                                        '
' Successful restructuring                                                  '
'                                                                           '
'                                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public reportStart As String
Public reportEnd As String
Public ProgramName As String
Public SkillName As String
Public mCm As String
Public skipFlag
Public reportStartRow
Public reportEndRow
Public flagFullRest
Public flagFullService

' Dim answer As Integer
Dim dataSheetName As String

Sub ARestructureAndGenerateReport()
Attribute ARestructureAndGenerateReport.VB_ProcData.VB_Invoke_Func = "r\n14"

'Dim answer
'answer = MsgBox("Would you like to restructure data and generate report?", vbYesNo + vbQuestion)
'If answer = vbYes Then
    'NewRestructuring
'End If
If Range("A1:A3").MergeCells = True Then
    UserAction.ActionReformat.Enabled = False
End If
If flagFullRest = True Then
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
    ' Request dates for next report
    'RequestDates
    ' Outline report period
    'OutlineReportPeriod
    ' Create new sheet for population by current programs
    PopulatePrograms
    Cells(1, 1).Select
    ActiveWindow.Zoom = 90
    ' Write MCM to cells for copy/paste to progress report.
    CreateProgramLists
    ' Prompt to Save As
    SaveWorkbook
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
    For Col = 2 To 1000
        If Cells(2, Col).Value = "" Then
        Else
            ' Assign header cell to variable
            headerCell = Cells(2, Col)
            ' Assign date column to variable
            headerCol = Col
            ' Assign next header column to variable
            If Cells(2, headerCol).End(xlToRight).Value = "" Then
                nextHeaderCol = headerCol + 15
            Else
                Cells(2, headerCol).End(xlToRight).Select
                nextHeaderCol = Selection.Column
                nextHeaderCol = nextHeaderCol - 2
            End If
        ' Fill cell below header with " " for asthetic purposes
        Cells(3, Col).Value = " "
           ' Move down rows and check if date is <, >, = to master date list
            For Row = 4 To 2000
                If Cells(Row, Col).Value = "" Then
                    Exit For
                Else
                    ' If > then cut chunk and paste one below
                    If Cells(Row, Col).Value > Cells(Row, 1) Then
                        If Cells((Row + 1), Col) = "" Then
                            ' Catch last row
                            Range(Cells(Row, headerCol), Cells(Row, (nextHeaderCol))).Select
                            Selection.Cut
                            Cells((Row + 1), Col).Select
                            ActiveSheet.Paste
                        Else
                            ' Ctrl+down + (next header column - 2)
                            Range(Cells(Row, headerCol), Cells(Selection.End(xlDown), nextHeaderCol)).Select
                            Selection.Cut
                            Cells((Row + 1), Col).Select
                            ActiveSheet.Paste
                        End If
                    Cells(Row, Col).Select
                    ' If < then insert new date in master list
                    ElseIf Cells(Row, Col).Value < Cells(Row, 1) Then
                        ' Catch last row
                        
                        ' Select all previous program chunks at date and below
                        Range(Cells(Row, 1), Cells(Selection.End(xlDown), (headerCol - 1))).Select
                        ' Cut and paste down one row
                        Selection.Cut
                        Cells((Row + 1), 1).Select
                        ActiveSheet.Paste
                        ' Copy new date to master list
                        Cells(Row, Col).Select
                        Selection.Copy
                        Cells(Row, 1).Select
                        ActiveSheet.Paste
                    ' If = then move to next row
                    End If
                End If
                ' Continue until no more dates
            Next Row
        End If
    Next Col
    
End Sub

Sub SelectChunk()
    '
    ' SelectChunk Macro
    '
    ' Keyboard Shortcut: Ctrl+m
    '
    chunkHeader = 2
    chunkHeader = Selection.Col
    nextHeader = Cells(2, (chunkHeader.End(xlToRight) - 2))
    programChunk = Range(Selection, (Cells(Selection.End(xlDown), nextHeader)))
    programChunk.Select
    
End Sub

Sub PopulatePrograms()

    Dim programRow
    Dim skillCount
    Dim skillStart
    Dim skillEnd

    Col = 2
    skillCount = 1
    programRow = 2
    dataSheetName = ActiveSheet.Name
    
    ' Create and format program/skill sheet
    Worksheets.Add().Name = "Programs"
    Worksheets("Programs").Cells(1, 1).Value = "Program"
    Worksheets("Programs").Cells(1, 2).Value = "Skill"
    Worksheets("Programs").Cells(1, 3).Value = "Mastered"
    Worksheets("Programs").Cells(1, 4).Value = "Continued"
    Worksheets("Programs").Cells(1, 5).Value = "Maintenance"
    Worksheets("Programs").Columns("A:B").ColumnWidth = 60
    Worksheets("Programs").Columns("C:E").ColumnWidth = 12
    Worksheets(dataSheetName).Activate
    ' Look for next program
    For Col = 2 To 1000
        If Cells(2, Col) <> "" Then
            headerCell = Cells(2, Col)
            headerCol = Col
            If Cells(2, headerCol).End(xlToRight).Value = "" Then
                If Cells(3, headerCol + 2).Value = "" Then
                    nextHeaderCol = Col + 1
                Else
                    nextHeaderCol = Cells(3, (Col + 1)).End(xlToRight).Column
                End If
            Else
                Cells(2, headerCol).End(xlToRight).Select
                nextHeaderCol = Selection.Column
                nextHeaderCol = nextHeaderCol - 2
            End If
            'Cycle through skills
            For i = (Col + 1) To nextHeaderCol
                ProgramName = Cells(2, Col).Value
                SkillName = Cells(3, i).Value
                ' Fit skill to window
                'Range(Cells(reportStartRow, i), Cells(reportEndRow, i)).Select
                Range(Cells(3, i).End(xlDown), Cells(1000, i).End(xlUp)).Select
                ActiveWindow.Zoom = True
                ActiveWindow.Zoom = 90
                ' Open MCM dialog box
                UserForm_Initialize
                MCMbox.Show
                If skipFlag = True Then
                ' Do nothing
                Else
                    Worksheets("Programs").Cells(programRow, 1).Value = ProgramName
                    Worksheets("Programs").Cells(programRow, 2).Value = SkillName
                    Select Case mCm
                        Case Is = "Mastered"
                            Worksheets("Programs").Cells(programRow, 3).Value = "X"
                        Case Is = "Continued"
                            Worksheets("Programs").Cells(programRow, 4).Value = "X"
                        Case Is = "Maintenance"
                            Worksheets("Programs").Cells(programRow, 5).Value = "X"
                    End Select
                    programRow = programRow + 1
                    Cells(2, Col).Value = ProgramName
                End If
            Next i
        End If
    Next Col
    
End Sub

Public Sub UserForm_Initialize()
    MCMbox.programNameBox.Value = ProgramName
    MCMbox.skillNameBox.Value = SkillName
    MCMbox.programNameBox.SetFocus
End Sub

Sub OutlineReportPeriod()

    Dim bottomDateRow
    
    bottomDateRow = Cells(4, 1).End(xlDown).Row
    
    For Row = 5 To bottomDateRow
        If Cells(Row, 1).Value > reportStart Then
            If Cells(Row - 1, 1).Value < reportStart Then
                reportStart = Cells(Row, 1).Value
            End If
        End If
        If Cells(Row - 1, 1).Value < reportEnd Then
            If Cells(Row, 1).Value > reportEnd Then
                reportEnd = Cells(Row - 1, 1).Value
            End If
        End If
    Next Row
    
    For Row = 4 To bottomDateRow
        If Cells(Row, 1).Value = reportStart Then
            Rows(Row).Select
            Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
            reportStartRow = Row
        End If
        If Cells(Row, 1).Value = reportEnd Then
            Rows(Row).Select
            Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
            reportEndRow = Row
        End If
    Next Row
    
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
                Cells(1, 6).Value = Cells(1, 6) & countMast & ") " & Cells(Row, 1).Value & ": " & Cells(Row, 2).Value
                countMast = countMast + 1
            Else
                Cells(1, 6).Value = Cells(1, 6) & ", " & countMast & ") " & Cells(Row, 1).Value & ": " & Cells(Row, 2).Value
                countMast = countMast + 1
            End If
        End If
        If Cells(Row, 4).Value = "X" Then
            If countCont = 1 Then
                Cells(1, 7).Value = Cells(1, 7) & countCont & ") " & Cells(Row, 1).Value & ": " & Cells(Row, 2).Value
                countCont = countCont + 1
            Else
                Cells(1, 7).Value = Cells(1, 7) & ", " & countCont & ") " & Cells(Row, 1).Value & ": " & Cells(Row, 2).Value
                countCont = countCont + 1
            End If
        End If
        If Cells(Row, 5).Value = "X" Then
            If countMaint = 1 Then
                Cells(1, 8).Value = Cells(1, 8) & countMaint & ") " & Cells(Row, 1).Value & ": " & Cells(Row, 2).Value
                countMaint = countMaint + 1
            Else
                Cells(1, 8).Value = Cells(1, 8) & ", " & countMaint & ") " & Cells(Row, 1).Value & ": " & Cells(Row, 2).Value
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
        Cells(1, 7).Value = Cells(1, 6).Value & "."
    End If
    
    If Cells(1, 8).Value = "" Then
        Cells(1, 8).Value = "N/A."
    Else
        Cells(1, 8).Value = Cells(1, 6).Value & "."
    End If
    
End Sub

Sub SaveWorkbook()

    Dim saveAsFilename As String
    Dim lastDate
    
    Worksheets(dataSheetName).Activate
    lastDate = Cells(4, 1).End(xlDown).Value
    lastDate = Format(lastDate, "YYYY") & "_" & Format(lastDate, "MM") & "_" & Format(lastDate, "DD")
    saveAsFilename = Cells(1, 1).Value & " - " & lastDate
    answer = MsgBox("Would you like to save as: " & saveAsFilename, vbYesNo + vbQuestion)
    If answer = vbYes Then
        ActiveWorkbook.SaveAs ("C:\Users\jackie\Documents\Client Files\Data\Formatted\" & saveAsFilename)
    End If
End Sub

Sub SingleRestructure()
'
' SingleRestructure Macro
'
' Keyboard Shortcut: Ctrl+r
'
Dim Col

Col = ActiveCell.Column

If Cells(2, Col).Value = "" Then
    Else
        ' Assign header cell to variable
        headerCell = Cells(2, Col)
        ' Assign date column to variable
        headerCol = Col
        ' Assign next header column to variable
        If Cells(2, headerCol).End(xlToRight).Value = "" Then
            nextHeaderCol = headerCol + 15
        Else
            Cells(2, headerCol).End(xlToRight).Select
            nextHeaderCol = Selection.Column
            nextHeaderCol = nextHeaderCol - 2
        End If
    ' Fill cell below header with " " for asthetic purposes
    Cells(3, Col).Value = " "
       ' Move down rows and check if date is <, >, = to master date list
        For Row = 4 To 2000
            If Cells(Row, Col).Value = "" Then
                Exit For
            Else
                ' If > then cut chunk and paste one below
                If Cells(Row, Col).Value > Cells(Row, 1) Then
                    If Cells((Row + 1), Col) = "" Then
                        ' Catch last row
                        Range(Cells(Row, headerCol), Cells(Row, (nextHeaderCol))).Select
                        Selection.Cut
                        Cells((Row + 1), Col).Select
                        ActiveSheet.Paste
                    Else
                        ' Ctrl+down + (next header column - 2)
                        Range(Cells(Row, headerCol), Cells(Selection.End(xlDown), nextHeaderCol)).Select
                        Selection.Cut
                        Cells((Row + 1), Col).Select
                        ActiveSheet.Paste
                    End If
                Cells(Row, Col).Select
                ' If < then insert new date in master list
                ElseIf Cells(Row, Col).Value < Cells(Row, 1) Then
                    ' Catch last row
                    
                    ' Select all previous program chunks at date and below
                    Range(Cells(Row, 1), Cells(Selection.End(xlDown), (headerCol - 1))).Select
                    ' Cut and paste down one row
                    Selection.Cut
                    Cells((Row + 1), 1).Select
                    ActiveSheet.Paste
                    ' Copy new date to master list
                    Cells(Row, Col).Select
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

