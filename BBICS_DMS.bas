Attribute VB_Name = "BBICS_DMS"
Public Const version As String = "v4.7.1"

Public reportStart, reportEnd, current As Date
Public ProgramName, ProgramDescription, ProgramSD, SkillName, mCm, guessText, errorTracking As String
Public startDateRow, endDateRow, programCount, prevProgramName, renameI, editRow, topEditRow, _
    bottomEditRow, rowsIndex, reportStartRow, reportEndRow, number, guess, lower, upper, guesses, bestGuesses As Integer
Public skip, skipFlag As Boolean
Public BxDict As New Scripting.Dictionary
Public objFSO, objFolder, objFile As Object
Public X, Y As Workbook
Public dateRows() As Integer

Dim dataSheetName As String
Dim objWord
Dim objDoc

Sub ARestructureAndGenerateReport()
Attribute ARestructureAndGenerateReport.VB_ProcData.VB_Invoke_Func = "r\n14"

    Dim bottomDateRow
    
    On Error Resume Next
    errorTracking = "ARestructureAndGenerateReport"
    
    dataSheetName = ActiveSheet.Name
    bottomDateRow = Cells(5, 1).End(xlDown).row
    
    UserAction.version.Caption = version
    
    If X Is Nothing Then
        UserAction.ActionDataEntry.Enabled = True
        UserAction.actionSaveWorkbook.Enabled = False
        UserAction.actionCloseWorkbook.Enabled = False
        UserAction.actionIPG.Enabled = False
    Else
        UserAction.ActionDataEntry.Enabled = False
        UserAction.actionSaveWorkbook.Enabled = True
        UserAction.actionCloseWorkbook.Enabled = True
        UserAction.actionIPG.Enabled = True
    End If
    
    ActiveWindow.WindowState = xlMinimized
        
    UserAction.Show
    
End Sub

Sub NewRestructuring()
'
' NewRestructuring Macro
'
' Keyboard Shortcut: Ctrl+w
'
    Dim headerCol, nextHeaderCol As Long
    Dim headerCell As Range
    Dim startTime As Double
    Dim minutesElapsed As String
    
    errorTracking = "NewRestructuring"
    
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

Sub EmptyBCheck()

    errorTracking = "EmptyBCheck"

        For i = 1 To 10000
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

    errorTracking = "CreateHeader"

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

    errorTracking = "MasterListFormat"

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

    errorTracking = "FormatProgramDates"

    For i = 1 To 10000
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

    errorTracking = "FindLastDate"

    For i = 1 To 10000
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

    On Error GoTo ErrorHandling
    errorTracking = "MoveData"
    
    ' Find next program chunk
    For col = 2 To 10000
        If Cells(2, col).Value = "" Then
        Else
            If Cells(4, col).Value = "" Then
                Cells(2, col).Cut
                Cells(1, col).PasteSpecial
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
            For row = 4 To 2000
                If Cells(row, col).Value = "" Then
                    Exit For
                Else
                    ' If > then cut chunk and paste one below
                    If Cells(row, col).Value > Cells(row, 1) Then
                        If Cells((row + 1), col) = "" Then
                            ' Catch last row
                            Range(Cells(row, headerCol), Cells(row, (nextHeaderCol))).Select
                            Selection.Cut
                            Cells((row + 1), col).Select
                            ActiveSheet.Paste
                        Else
                            ' Ctrl+down + (next header column - 2)
                            Range(Cells(row, headerCol), Cells(Selection.End(xlDown), nextHeaderCol)).Select
                            Selection.Cut
                            Cells((row + 1), col).Select
                            ActiveSheet.Paste
                        End If
                    Cells(row, col).Select
                    ' If < then insert new date in master list
                    ElseIf Cells(row, col).Value < Cells(row, 1) Then
                        ' Catch last row
                        
                        ' Select all previous program chunks at date and below
                        Range(Cells(row, 1), Cells(Selection.End(xlDown), (headerCol - 1))).Select
                        ' Cut and paste down one row
                        Selection.Cut
                        Cells((row + 1), 1).Select
                        ActiveSheet.Paste
                        ' Copy new date to master list
                        Cells(row, col).Select
                        Selection.Copy
                        Cells(row, 1).Select
                        ActiveSheet.Paste
                    ' If = then move to next row
                    End If
                End If
                ' Continue until no more dates
            Next row
        End If
    Next col
    
ErrorHandling:
    ErrHandling
    
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

    Dim programRow, skillCount, skillStartRow, skillEndRow, lastSkillRow, bottomDateRow, prev1, prev2, prev3, prevRow As Integer
    Dim skillStart, skillEnd, skillDate, lastSkillEnd As Date
    Dim programSheet, deletedSkill As String
    
    On Error Resume Next
    errorTracking = "PopulatePrograms"
    
    Application.ScreenUpdating = False
    
    col = 2
    skillCount = 1
    programRow = 2
    dataSheetName = ActiveSheet.Name
    programSheet = "Programs"
    bottomDateRow = Cells(5, 1).End(xlDown).row
                
    UserForm_Initialize
    errorTracking = "PopulatePrograms"
    UserForm1.Show
    
    ' Create and format program/skill sheet
    Application.DisplayAlerts = False
    For Each Sheet In Worksheets
        If programSheet = Sheet.Name Then
            Sheet.Delete
        End If
    Next Sheet
    Application.DisplayAlerts = True
    
    For Each Sheet In Worksheets
        If Sheet.Name = "Programs" Then
            Application.DisplayAlerts = False
            Worksheets("Programs").Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next Sheet
    
    Worksheets.Add().Name = programSheet
    
    Worksheets("Programs").Cells(1, 1).Value = "Program"
    Worksheets("Programs").Cells(1, 2).Value = "Skill"
    Worksheets("Programs").Cells(1, 3).Value = "Mastered"
    Worksheets("Programs").Cells(1, 4).Value = "Continued"
    Worksheets("Programs").Cells(1, 5).Value = "Maintenance"
    Worksheets("Programs").Columns("A:B").ColumnWidth = 60
    Worksheets("Programs").Columns("C:E").ColumnWidth = 12
    Worksheets("Programs").Columns("A:E").NumberFormat = "@"
    Worksheets(dataSheetName).Activate
    
    'Delete empty skill columns
    For col = 2 To Cells(2, 10000).End(xlToLeft).Column
        If Cells(2, col) <> "" Then
            nextHeaderCol = Cells(2, col).End(xlToRight).Column
            For i = col + 1 To Cells(3, nextHeaderCol).End(xlToLeft).Column
                If Cells(3, i).End(xlDown) = "" Then
                    deletedSkill = Cells(3, i).Value
                    DeleteBox.lblDeleting.Caption = "'" & Cells(2, col).Value & ": " & deletedSkill & "' skill column(" & col & ") empty." & vbCrLf & vbCrLf & "Deleting..."
                    Application.OnTime Now + TimeSerial(0, 0, 1.5), "UnloadDeleteBox"
                    DeleteBox.Show
                    Columns(i).Delete
                End If
            Next i
        End If
    Next col
    
    'Delete extra empty columns
    For col = 3 To Cells(2, 10000).End(xlToLeft).Column
        If Cells(1, col).End(xlDown).Value = "" And Cells(1, col - 1).End(xlDown).Value = "" Then
            DeleteBox.lblDeleting.Caption = "Deleting extra empty column(" & col & ")..."
            Application.OnTime Now + TimeSerial(0, 0, 1.5), "UnloadDeleteBox"
            DeleteBox.Show
            Columns(col).Delete
        End If
    Next col

    ' Look for next program
    For col = 2 To 10000
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
                                
                ProgramName = Cells(2, col).Value
                SkillName = Cells(3, i).Value
                Range(Cells(3, i).End(xlDown), Cells(1000, i).End(xlUp)).Select
                ActiveWindow.Zoom = True
                ActiveWindow.Zoom = 90
                
                UserForm_Initialize
                errorTracking = "PopulatePrograms"
                
                ' Store skill start/end dates as variables
                If Cells(4, i).Value = "" Then
                    skillStartRow = Cells(4, i).End(xlDown).row
                    skillStart = Cells(skillStartRow, col).Value
                Else
                    skillStart = Cells(4, col).Value
                End If
                
                'Check and store most current skill ending date
                Cells(1000, i).End(xlUp).Select
                skillEnd = Selection.End(xlToLeft).Value
                If Cells(4, nextHeaderCol).Value = "" Then
                    lastSkillRow = Cells(1000, nextHeaderCol).End(xlUp).row
                    lastSkillEnd = Cells(lastSkillRow, col).Value
                Else
                    lastSkillEnd = Cells(4, col).Value
                End If
                
                ' Check for skill within report dates
                If DateValue(skillEnd) < DateValue(reportStart) And DateValue(reportStart) - DateValue(lastSkillEnd) < 182 Then
                ' Do nothing
                ElseIf DateValue(skillStart) > DateValue(reportEnd) Then
                ' Also do nothing
                ElseIf DateValue(reportStart) - DateValue(lastSkillEnd) > 182 Then
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
                    k = Cells(1000, i).End(xlUp).row
                    l = Cells(4, i).End(xlDown).row
                    Application.ScreenUpdating = True
                    Range(Cells(l, i), Cells(k, i)).Activate
                    MCMbox.Show
                    Application.ScreenUpdating = False
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
            Next i
        End If
    Next col
    
    ' Reset report period borders to black/white/transparent
     For i = 4 To bottomDateRow
        If DateValue(reportStart) = DateValue(Cells(i, 1).Value) Then
            Rows(i).Select
            With Selection.Borders(xlEdgeTop)
                .Color = RGB(0, 0, 0)
                .LineStyle = xlContinuous
            End With
            startDateRow = i
        End If
        If DateValue(Cells(i, 1).Value) = DateValue(reportEnd) Then
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
    
    Application.ScreenUpdating = True
           
ErrorHandling:
    ErrHandling
       
End Sub

Public Sub UserForm_Initialize()

    Dim suggestStart
    Dim suggestEnd
    Dim placeHolder
    Dim reportDates As String
    
    errorTracking = "UserForm_Initialize"
    
    MCMbox.programNameBox.Value = Trim(ProgramName)
    MCMbox.skillNameBox.Value = Trim(SkillName)
    MCMbox.progMast.Value = True
    MCMbox.NextProgram.SetFocus
    If Worksheets("CI").Cells(2, 4).Value = "Final" Then
        MCMbox.progCont.Enabled = False
        MCMbox.Label4.Enabled = False
    End If
    
    
    reportDates = Worksheets("CI").Cells(2, 6).Value & " " & Worksheets("CI").Cells(2, 7).Value & _
        " - " & Worksheets("CI").Cells(2, 8).Value & " " & Worksheets("CI").Cells(2, 9).Value
    UserForm1.reportDates.Value = reportDates
    
    With UserForm1.ComboBox1
        For i = 4 To Worksheets(dataSheetName).Cells(5, 1).End(xlDown).row
            .AddItem Worksheets(dataSheetName).Cells(i, 1).Value
        Next i
    End With
    
    With UserForm1.ComboBox2
        For i = 4 To Worksheets(dataSheetName).Cells(5, 1).End(xlDown).row
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

    For i = 4 To X.Worksheets("Data").Cells(5, 1).End(xlDown).row
        If DateValue(X.Worksheets("Data").Cells(i, 1).Value) = DateValue(suggestStart) Then
            UserForm1.ComboBox1 = X.Worksheets("Data").Cells(i, 1).Value
            Exit For
        End If
        If DateValue(X.Worksheets("Data").Cells(i, 1).Value) < DateValue(suggestStart) And DateValue(X.Worksheets("Data").Cells(i + 1, 1).Value) > DateValue(suggestStart) Then
            UserForm1.ComboBox1 = X.Worksheets("Data").Cells(i + 1, 1).Value
            Exit For
        End If
    Next i
    
    For i = X.Worksheets("Data").Cells(4, 1).End(xlDown).row To 5 Step -1
        X.Worksheets("Data").Cells(i, 1).Activate
        If DateValue(X.Worksheets("Data").Cells(i, 1).Value) = DateValue(suggestEnd) Then
            UserForm1.ComboBox2 = X.Worksheets("Data").Cells(i, 1).Value
            Exit For
        End If
        If DateValue(X.Worksheets("Data").Cells(i, 1).Value) < DateValue(suggestEnd) And Not IsDate(X.Worksheets("Data").Cells(i + 1, 1).Value) Then
            UserForm1.ComboBox2 = X.Worksheets("Data").Cells(i, 1).Value
            Exit For
        End If
        If DateValue(X.Worksheets("Data").Cells(i - 1, 1).Value) < DateValue(suggestEnd) And DateValue(X.Worksheets("Data").Cells(i, 1).Value) > DateValue(suggestEnd) Then
            UserForm1.ComboBox2 = X.Worksheets("Data").Cells(i - 1, 1).Value
            Exit For
        End If
    Next i
              
End Sub


Sub CreateProgramLists()

    Dim bottomProgramRow, countMast, countCont, countMaint As Integer
    
    errorTracking = "CreateProgramLists"
    
    countMast = 1
    countCont = 1
    countMaint = 1
    
    Worksheets("Programs").Activate
    
    bottomProgramRow = Cells(1, 1).End(xlDown).row
    For row = 2 To bottomProgramRow
        If Cells(row, 3).Value = "X" Then
            If countMast = 1 Then
                Cells(1, 6).Value = Cells(1, 6) & countMast & ") " & Cells(row, 1).Value & " (" & Cells(row, 2).Value & ")"
                countMast = countMast + 1
            Else
                Cells(1, 6).Value = Cells(1, 6) & ", " & countMast & ") " & Cells(row, 1).Value & " (" & Cells(row, 2).Value & ")"
                countMast = countMast + 1
            End If
        End If
        If Cells(row, 4).Value = "X" Then
            If countCont = 1 Then
                Cells(1, 7).Value = Cells(1, 7) & countCont & ") " & Cells(row, 1).Value & " (" & Cells(row, 2).Value & ")"
                countCont = countCont + 1
            Else
                Cells(1, 7).Value = Cells(1, 7) & ", " & countCont & ") " & Cells(row, 1).Value & " (" & Cells(row, 2).Value & ")"
                countCont = countCont + 1
            End If
        End If
        If Cells(row, 5).Value = "X" Then
            If countMaint = 1 Then
                Cells(1, 8).Value = Cells(1, 8) & countMaint & ") " & Cells(row, 1).Value
                countMaint = countMaint + 1
            Else
                Cells(1, 8).Value = Cells(1, 8) & ", " & countMaint & ") " & Cells(row, 1).Value
                countMaint = countMaint + 1
            End If
        End If
    Next row
    
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

Sub SingleRestructure()
Attribute SingleRestructure.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' SingleRestructure Macro
'
' Keyboard Shortcut: Ctrl+r
'
Dim col As Integer

    errorTracking = "SingleRestructure"

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
        For row = 4 To 2000
            If Cells(row, col).Value = "" Then
                Exit For
            Else
                ' If > then cut chunk and paste one below
                If Cells(row, col).Value > Cells(row, 1) Then
                    If Cells((row + 1), col) = "" Then
                        ' Catch last row
                        Range(Cells(row, headerCol), Cells(row, (nextHeaderCol))).Select
                        Selection.Cut
                        Cells((row + 1), col).Select
                        ActiveSheet.Paste
                    Else
                        ' Ctrl+down + (next header column - 2)
                        Range(Cells(row, headerCol), Cells(Selection.End(xlDown), nextHeaderCol)).Select
                        Selection.Cut
                        Cells((row + 1), col).Select
                        ActiveSheet.Paste
                    End If
                Cells(row, col).Select
                ' If < then insert new date in master list
                ElseIf Cells(row, col).Value < Cells(row, 1) Then
                    ' Select all previous program chunks at date and below
                    Range(Cells(row, 1), Cells(Selection.End(xlDown), (headerCol - 1))).Select
                    ' Cut and paste down one row
                    Selection.Cut
                    Cells((row + 1), 1).Select
                    ActiveSheet.Paste
                    ' Copy new date to master list
                    Cells(row, col).Select
                    Selection.Copy
                    Cells(row, 1).Select
                    ActiveSheet.Paste
                ' If = then move to next row
                End If
            End If
            ' Continue until no more dates
        Next row
    End If
    
End Sub

Sub ImportSkillsPrograms()

    Dim z, w, v, Y As Workbook
    Dim k As Integer
    Dim sht As Worksheet
    
    errorTracking = "ImportSkillsPrograms"
    
    Application.DisplayAlerts = False
       
    'Check for existing worksheets, if so, delete them
    For Each Sheet In Worksheets
        If Sheet.Name = "CI" Then
            Worksheets("CI").Delete
        ElseIf Sheet.Name = "SDL" Then
            Worksheets("SDL").Delete
        ElseIf Sheet.Name = "PD" Then
            Worksheets("PD").Delete
        End If
    Next Sheet
    
    Worksheets.Add().Name = "CI"
    Worksheets.Add().Name = "SDL"
    Worksheets.Add().Name = "PD"
      
    'Import information as new worksheets
    Set z = ActiveWorkbook
    Set w = Workbooks.Open("C:\Users\jackie\Documents\Client Files\Progress Reports\FMP_DataExport\FMP_DataExport.xlsx")
    Set v = Workbooks.Open("C:\Users\jackie\Documents\Client Files\Progress Reports\FMP_DataExport\SkillDeficitList.xlsx")
    Set Y = Workbooks.Open("C:\Users\jackie\Documents\Client Files\Progress Reports\FMP_DataExport\ProgramDescriptions.xlsx")
       
    w.Sheets("CI").Range("A1:M2").Copy
    z.Sheets("CI").Range("A1:M2").PasteSpecial
    w.Close
    z.Sheets("CI").Columns("A:M").AutoFit
    
    v.Sheets("SDL").Range("A1:B112").Copy
    z.Sheets("SDL").Range("A1:B112").PasteSpecial
    v.Close
    
    k = Worksheets("PD").Cells(1000, 1).End(xlUp).row
    Y.Sheets("PD").Range("A1:C" & k).Copy
    z.Sheets("PD").Range("A1:C" & k).PasteSpecial
    Y.Close
    
    X.Worksheets("Data").Activate

    Application.DisplayAlerts = True

End Sub

Sub PopulateReport()

    Dim chunks, bxCount, currentBottomRow As Integer
    Dim chunk, bxString As String
    Dim objRange
    Dim s As Object
    Dim bx As Variant
    
    On Error Resume Next
    errorTracking = "PopulateReport"

    'Create Word object and open PRT template.
    MsgBox "Please save and close any Microsoft Word documents at this time. Any unsaved changed will be lost.", vbExclamation
    Word.Application.Quit
    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Open("C:\Users\jackie\Documents\Client Files\Progress Reports\FMP_DataExport\PRT.docx")
    objWord.Visible = True
           
    'Find/Replace sections with data.
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
    
    If Worksheets("CI").Cells(2, 4).Value = "Final" Then GoTo Final
    
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
    
Final:
    
    ProgramDescriptionsList
    ProgramMatch
    errorTracking = "PopulateReport"
    
    Worksheets("Current").Activate
    currentBottomRow = Worksheets("Current").Cells(1, 1).End(xlDown).row
    
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
        
    'Sum bx quantities over report period
    BxSetup
    'Sort bx data
    BxData
    errorTracking = "PopulateReport"
    'Write bx data to report
    bxCount = BxDict.Count
    For Each bx In BxDict
        If bxCount = BxDict.Count Then
            If BxDict(bx) = 1 Then
                bxString = ", and " & bxCount & ") " & BxDict(bx) & " count of " & bx & "."
            Else
                bxString = ", and " & bxCount & ") " & BxDict(bx) & " counts of " & bx & "."
            End If
        ElseIf bxCount = 1 Then
            If BxDict(bx) = 1 Then
                bxString = bxCount & ") " & BxDict(bx) & " count of " & bx & bxString
            Else
                bxString = bxCount & ") " & BxDict(bx) & " counts of " & bx & bxString
            End If
        Else
            If BxDict(bx) = 1 Then
                bxString = ", " & bxCount & ") " & BxDict(bx) & " count of " & bx & bxString
            Else
                bxString = ", " & bxCount & ") " & BxDict(bx) & " counts of " & bx & bxString
            End If
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
    errorTracking = "PopulateReport"
    
    'Cleanup
    Application.DisplayAlerts = False
    Worksheets("PD").Delete
    Worksheets("SDL").Delete
    Worksheets("CI").Delete
    Worksheets("Programs").Delete
    Worksheets("Current").Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    objWord.Visible = True
    objWord.Application.Activate
    objWord.Application.WindowState = wdWindowStateMaximize

ErrorHandling:
    'ErrHandling

End Sub

Sub ProgramDescriptionsList()

    errorTracking = "ProgramDescriptionsList"

    For Each Sheet In Worksheets
        If Sheet.Name = "Current" Then
            Application.DisplayAlerts = False
            Worksheets("Current").Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next Sheet
    
    'Create list of current programs
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

    errorTracking = "ProgramMatch"

    Dim found As Boolean

    For i = 1 To Worksheets("Current").Cells(1, 1).End(xlDown).row
        ProgramName = Trim(Worksheets("Current").Cells(i, 1).Value)
        found = False
        For scan = 2 To Worksheets("PD").Cells(1000, 1).End(xlUp).row
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
                For j = 2 To Worksheets("PD").Cells(2000, 1).End(xlUp).row
                    .AddItem Worksheets("PD").Cells(j, 1).Value
                 Next j
            End With

            skip = False
            DescriptionBox.CommandButton3.Enabled = True
            DescriptionBox.currentProgramName.Value = ProgramName
            DescriptionBox.programNameBox.Value = ProgramName
            DescriptionBox.programDescriptionBox = "The client will "
            DescriptionBox.programSDBox.Value = "The SD will vary depending on the instructions in the workbook, the client’s IEP, and the interventionist."
            For j = 2 To X.Worksheets("Data").Cells(2, 10000).End(xlToLeft).Column
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

    errorTracking = "BxData"

    Dim bxRow, bxColStart, bxColEnd As Integer
    
    Dim Arr() As Variant
    Dim Temp1, Temp2 As Variant
    Dim txt As String
    Dim i, j As Long
    
    Worksheets("Bx Data").Activate
    For i = 3 To Cells(4, 1).End(xlDown).row
        If DateValue(Cells(i, 1).Value) = DateValue(reportStart) Then
            bxRow = i
        End If
    Next i
    bxColStart = (Cells(2, 2).End(xlToRight).Column) + 2
    bxColEnd = Cells(2, bxColStart).End(xlToRight).Column
    
    'Store bx values in dictionary
    For i = bxColStart To bxColEnd
        If Cells(bxRow, i).Value <> 0 Then
            BxDict.Add Trim(Cells(2, i).Value), Cells(bxRow, i).Value
        End If
    Next i
    
    If BxDict Is Nothing Then
        MsgBox ("Behavior volumes null set.")
        Exit Sub
    End If
    
    ReDim Arr(0 To BxDict.Count - 1, 0 To 1)
    
    'Store bx data in temporary array
    For i = 0 To BxDict.Count - 1
        Arr(i, 0) = BxDict.Keys(i)
        Arr(i, 1) = BxDict.Items(i)
    Next i
    
    'Sort temp array
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
    
    'Remove the contents of the dictionary collection.
    BxDict.RemoveAll
    
    'Add the items back into the dictionary collection.
    For i = LBound(Arr, 1) To UBound(Arr, 1)
        BxDict.Add Key:=Arr(i, 0), Item:=Arr(i, 1)
    Next i
    
    'Build the text for behavior data.
    For i = 0 To BxDict.Count - 1
        txt = txt & BxDict.Keys(i) & vbTab & BxDict.Items(i) & vbCrLf
    Next i
    
End Sub

Sub RenamePrograms()

    errorTracking = "RenamePrograms"
   
   'Cycle through programs with option to rename
   
   renameI = 2
   
    With RenameBox.programAll
        For j = 2 To Worksheets("PD").Cells(1000, 1).End(xlUp).row
        .AddItem Worksheets("PD").Cells(j, 1).Value
        Next j
    End With
    
    With RenameBox.programCurrent
        For j = 2 To Worksheets("Data").Cells(2, 10000).End(xlToLeft).Column
            If Worksheets("Data").Cells(2, j).Value = "" Then
                'Do nothing
            Else
                .AddItem Worksheets("Data").Cells(2, j).Value
            End If
        Next j
    End With
   
    RenameBox.programChange = Worksheets("Data").Cells(2, renameI).Value
    RenameBox.programCurrent = Worksheets("Data").Cells(2, renameI).Value
    RenameBox.programExisting = Worksheets("Data").Cells(2, renameI).Value
   
    RenameBox.programChange.SetFocus
        With RenameBox.programChange
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
              
    Worksheets("Data").Cells(2, renameI).Activate
    RenameBox.btnPrev.Enabled = False
    RenameBox.Show
    
End Sub

Sub TutorHrs()

    errorTracking = "TutorHrs"

    Dim tutorHrRow, monthCount As Integer
    Dim tutorHrDate As String
    
    'Populate report with tutor hour data
    Worksheets("Tutor Hr Data").Activate
    reportStart = Format(reportStart, "MMM yyyy")
    For i = 3 To Cells(3, 1).End(xlDown).row
        If Format(Cells(i, 1).Value, "MMM yyyy") = Format(reportStart, "MMM yyyy") Then
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
        Case Is = "Final"
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

    errorTracking = "DataEntryPrograms"
       
    DataEntryBox.SessionDate.SetFocus
    DataEntryBox.buttonNextData.Default = True
    DataEntryBox.ProgramList = "Select preexisting program..."
    DataEntryBox.SkillList = "<Please select program first>"
    DataEntryBox.SkillList.Enabled = False
    DataEntryBox.AddProgram.Enabled = False
    DataEntryBox.AddSkill.Enabled = False
    DataEntryBox.buttonNextData.Enabled = False
    DataEntryBox.btnEditUp.Enabled = False
    DataEntryBox.btnEditDown.Enabled = False
    DataEntryBox.btnDelete.Enabled = False
    DataEntryBox.btnEdit.Enabled = False
    
    X.Activate
    X.Worksheets("Data").Activate
        
    DataEntryBox.ProgramList.Clear
    For col = 2 To Worksheets("Data").Cells(2, 10000).End(xlToLeft).Column
        If Worksheets("Data").Cells(2, col).Value = "" Then
        Else
            DataEntryBox.ProgramList.AddItem Cells(2, col).Value
        End If
    Next col
    
    DataEntryBox.Show

End Sub

Sub ErrHandling()

    If err.number <> 0 And err.number <> 429 Then
        Beep
        ErrorBox.Show
    End If
    
End Sub

Sub InitGuessBox()

    errorTracking = "InitGuessBox"

    Randomize
    number = Int((99 - 2 + 1) * Rnd() + 2)
    guesses = 0
    guess = 0
    bestGuesses = 100
    guessText = "Guesses: " & guesses
    lower = 0
    upper = 100
    GuessBox.lower.Caption = lower
    GuessBox.upper.Caption = upper
    
    GuessBox.btnGuess.Caption = "Guess!"
    
    'MsgBox (number)

End Sub

Sub GetSaveAsFileName()

    errorTracking = "GetSaveAsFileName"

    Dim FileName As Variant
    Dim Filt As String, Title As String, Name As String
    Dim FilterIndex As Long, Response As Long
    
    Name = X.Worksheets("Data").Cells(1, 1).Value & " - " & Format(X.Worksheets("Data").Cells(4, 1).End(xlDown).Value, "YYYY_MM_DD") & ".xlsx"
    '   Set to Specified Path\Folder
        ChDir "C:\Users\jackie\Documents\Client Files\Data\Formatted"
    '   Set File Filter
        Filt = "Excel Files (*.xlsx), *.xlsx"
    '   Set *.* to Default
        FilterIndex = 5
    '   Set Dialogue Box Caption
        Title = "Please select a file name"
    '   Get FileName
        FileName = Application.GetSaveAsFileName(InitialFileName:=Name, FileFilter:=Filt, _
            FilterIndex:=FilterIndex, Title:=Title)
    '   Exit if Dialogue box cancelled
        If FileName = False Then
            'Response = MsgBox("No File was selected", vbOKOnly & vbCritical, "Selection Error")
            Exit Sub
        End If
    '   Display Full Path & File Name
        Response = MsgBox("Saving as: " & FileName, vbInformation, "Proceed")
    '   Save & Close Workbook
        With ActiveWorkbook
            .SaveAs FileName
        End With
            
End Sub

Sub UnloadDeleteBox()

    errorTracking = "UnloadDeleteBox"

    Unload DeleteBox
    
End Sub

Sub BxSetup()

    Dim dateFlag As Boolean
    Dim bxRow, bxEnd, bxCount As Integer
    
    errorTracking = "BxSetup"
    
    X.Worksheets("Bx Data").Activate
    For i = 3 To X.Worksheets("Bx Data").Cells(3, 1).End(xlDown).row
        If DateValue(X.Worksheets("Bx Data").Cells(i, 1).Value) = DateValue(reportStart) Then
            dateFlag = True
            bxRow = i
        End If
    Next i
    If dateFlag = False Then
        For i = 3 To X.Worksheets("Bx Data").Cells(3, 1).End(xlDown).row
            If DateValue(X.Worksheets("Bx Data").Cells(i, 1).Value) < DateValue(reportStart) And DateValue(X.Worksheets("Bx Data").Cells(i + 1, 1).Value) > DateValue(reportStart) Then
                X.Worksheets("Bx Data").Rows(i + 1).EntireRow.Insert
                bxRow = i + 1
                X.Worksheets("Bx Data").Cells(bxRow, 1).Value = reportStart
                Exit For
            End If
        Next i
    End If
   
    X.Worksheets("Bx Data").Cells(X.Worksheets("Bx Data").Cells(3, 1).End(xlDown).row + 1, 1).Value = Format(Now, "MM/DD/YYYY")
    dateFlag = False
    For i = 3 To X.Worksheets("Bx Data").Cells(3, 1).End(xlDown).row
        If DateValue(X.Worksheets("Bx Data").Cells(i, 1).Value) = DateValue(reportEnd) Then
            dateFlag = True
            bxEnd = i
        End If
    Next i
    If dateFlag = False Then
        For i = 3 To X.Worksheets("Bx Data").Cells(3, 1).End(xlDown).row
            If DateValue(X.Worksheets("Bx Data").Cells(i, 1).Value) < DateValue(reportEnd) And DateValue(X.Worksheets("Bx Data").Cells(i + 1, 1).Value) > DateValue(reportEnd) Then
                X.Worksheets("Bx Data").Rows(i + 1).EntireRow.Insert
                bxEnd = i + 1
                X.Worksheets("Bx Data").Cells(bxEnd, 1).Value = reportEnd
                Exit For
            End If
        Next i
    End If
        
    For i = X.Worksheets("Bx Data").Cells(2, 2).End(xlToRight).Column + 2 To X.Worksheets("Bx Data").Cells(2, X.Worksheets("Bx Data").Cells(2, 2).End(xlToRight).Column + 2).End(xlToRight).Column
        bxCount = 0
        For j = bxRow To bxEnd
            bxCount = bxCount + X.Worksheets("Bx Data").Cells(j, i - X.Worksheets("Bx Data").Cells(2, 2).End(xlToRight).Column).Value
        Next j
        X.Worksheets("Bx Data").Cells(bxRow, i).Value = bxCount
    Next i
    
    X.Worksheets("Bx Data").Cells(3, 1).End(xlDown).Value = ""
    
    With X.Worksheets("Bx Data").Rows(bxRow - 1).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With X.Worksheets("Bx Data").Rows(bxEnd).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
End Sub
