VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataEntryBox 
   Caption         =   "Enter New Data"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8400
   OleObjectBlob   =   "DataEntryBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataEntryBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub TextBox3_Change()

End Sub

Private Sub AddProgram_Click()

    Dim col As Integer
    
    DataEntryBox.ProgramList.AddItem DataEntryBox.Program
    col = x.Worksheets("Data").Cells(3, 1000).End(xlToLeft).Column + 2
    x.Worksheets("Data").Columns(col).Select
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

    x.Worksheets("Data").Cells(2, col).Value = DataEntryBox.Program.Value
    x.Worksheets("Data").Cells(3, col).Value = " "
    DataEntryBox.ProgramList = DataEntryBox.Program.Value
    DataEntryBox.AddProgram.Enabled = False
    DataEntryBox.Skill.SetFocus
    
End Sub

Private Sub AddSkill_Click()

    Dim i As Integer
    Dim col As Integer
    
    DataEntryBox.SkillList.AddItem DataEntryBox.Skill.Value
    DataEntryBox.SkillList = DataEntryBox.Skill.Value
    For i = 2 To x.Worksheets("Data").Cells(2, 1000).End(xlToLeft).Column
        If x.Worksheets("Data").Cells(2, i).Value = DataEntryBox.ProgramList.Value Then
            col = i + 1
            If x.Worksheets("Data").Cells(3, col).Value = "" Then
                x.Worksheets("Data").Cells(3, col).Value = DataEntryBox.SkillList
            ElseIf x.Worksheets("Data").Cells(3, col + 1).Value = "" Then
                x.Worksheets("Data").Cells(3, col).Activate
                ActiveCell.EntireColumn.Offset(0, 1).Insert
                ActiveCell.Offset(0, 1).Value = DataEntryBox.SkillList
            Else
                x.Worksheets("Data").Cells(3, col).End(xlToRight).Activate
                ActiveCell.EntireColumn.Offset(0, 1).Insert
                ActiveCell.Offset(0, 1).Value = DataEntryBox.SkillList
            End If
        End If
    Next i
    
    DataEntryBox.SessionDate.SetFocus

End Sub

Private Sub btnDelete_Click()

    Dim i, j, programCol, skillCol, arraySize As Integer
    Dim row As Variant
    
    On Error GoTo ErrorHandling
    
    For i = 2 To x.Worksheets("Data").Cells(2, 1000).End(xlToLeft).Column
        If x.Worksheets("Data").Cells(2, i).Value = DataEntryBox.ProgramList.Value Then
            programCol = i
            For j = programCol To x.Worksheets("Data").Cells(3, programCol + 1).End(xlToRight).Column
                If x.Worksheets("Data").Cells(3, j).Value = DataEntryBox.SkillList.Value Then
                    skillCol = j
                End If
            Next j
        End If
    Next i
 
    x.Worksheets("Data").Cells(editRow, programCol) = ""
    x.Worksheets("Data").Cells(editRow, skillCol) = ""
    
    If DataEntryBox.SkillList <> "<Please select program first>" Then
        arraySize = 0
        For i = 4 To x.Worksheets("Data").Cells(4, 1).End(xlDown).row
            If x.Worksheets("Data").Cells(i, skillCol).Value <> "" Then arraySize = arraySize + 1
        Next i
    End If
              
    If arraySize > 0 Then arraySize = arraySize - 1
    ReDim dateRows(arraySize)
    
    j = 0
    
    If arraySize <> 0 Then
        For i = 4 To x.Worksheets("Data").Cells(4, 1).End(xlDown).row
            If x.Worksheets("Data").Cells(i, skillCol).Value <> "" Then
                dateRows(j) = x.Worksheets("Data").Cells(i, skillCol).row
                j = j + 1
            End If
        Next i
    End If
    
    'Reposition the edit position
    If rowsIndex > 0 Then rowsIndex = rowsIndex - 1
    'Keep the index from dropping below 0
    If rowsIndex < 0 Then rowsIndex = 0
    
ErrorHandling:
    ErrHandling

End Sub

Private Sub btnEdit_Click()

    Dim i, j, programCol, skillCol, arraySize As Integer
    Dim row As Variant
    
    On Error GoTo ErrorHandling
    
    'Get listing for program and skill columns
    For i = 2 To x.Worksheets("Data").Cells(2, 1000).End(xlToLeft).Column
        If x.Worksheets("Data").Cells(2, i).Value = DataEntryBox.ProgramList.Value Then
            programCol = i
            For j = programCol To x.Worksheets("Data").Cells(3, programCol + 1).End(xlToRight).Column
                If x.Worksheets("Data").Cells(3, j).Value = DataEntryBox.SkillList.Value Then
                    skillCol = j
                End If
            Next j
        End If
    Next i
    
    'Delete current entry
    x.Worksheets("Data").Cells(editRow, programCol) = ""
    x.Worksheets("Data").Cells(editRow, skillCol) = ""
    
    'Copy from edit section to new data section
    DataEntryBox.SessionDate = DataEntryBox.txtEditDate
    DataEntryBox.Score = DataEntryBox.txtEditScore
    
    'Add the data as new
    buttonNextData_Click
    
    'Set the array size
    If DataEntryBox.SkillList <> "<Please select program first>" Then
        arraySize = 0
        For i = 4 To x.Worksheets("Data").Cells(4, 1).End(xlDown).row
            If x.Worksheets("Data").Cells(i, skillCol).Value <> "" Then arraySize = arraySize + 1
        Next i
    End If
              
    'Redefine the array
    If arraySize > 0 Then arraySize = arraySize - 1
    ReDim dateRows(arraySize)
    
    j = 0
    
    If arraySize <> 0 Then
        For i = 4 To x.Worksheets("Data").Cells(4, 1).End(xlDown).row
            If x.Worksheets("Data").Cells(i, skillCol).Value <> "" Then
                dateRows(j) = x.Worksheets("Data").Cells(i, skillCol).row
                j = j + 1
            End If
        Next i
    End If
    
    'Reposition the current edit position
    If rowsIndex > 0 Then rowsIndex = rowsIndex - 1
    'Keep the index from dropping below 0
    If rowsIndex < 0 Then rowsIndex = 0
    
    'Empty edit text boxes
    DataEntryBox.txtEditDate = ""
    DataEntryBox.txtEditScore = ""
    
ErrorHandling:
    ErrHandling

End Sub

Private Sub btnEditDown_Click()

    Dim i, j, programCol, skillCol As Integer
    
    On Error GoTo ErrorHandling
    
    For i = 2 To x.Worksheets("Data").Cells(2, 1000).End(xlToLeft).Column
        If x.Worksheets("Data").Cells(2, i).Value = DataEntryBox.ProgramList.Value Then
            programCol = i
            For j = programCol To x.Worksheets("Data").Cells(3, programCol + 1).End(xlToRight).Column
                If x.Worksheets("Data").Cells(3, j).Value = DataEntryBox.SkillList.Value Then
                    skillCol = j
                End If
            Next j
        End If
    Next i
    
    If editRow = dateRows(UBound(dateRows)) Then
        rowsIndex = 0
        editRow = dateRows(rowsIndex)
    Else
        rowsIndex = rowsIndex + 1
        editRow = dateRows(rowsIndex)
    End If
    
    DataEntryBox.txtEditDate = x.Worksheets("Data").Cells(editRow, programCol).Value
    DataEntryBox.txtEditScore = x.Worksheets("Data").Cells(editRow, skillCol).Value
    
    Union(x.Worksheets("Data").Cells(editRow, programCol), x.Worksheets("Data").Cells(editRow, skillCol)).Activate
    
    current = DataEntryBox.txtEditDate.Value

ErrorHandling:
    ErrHandling

End Sub

Private Sub btnEditUp_Click()

    Dim i, j, programCol, skillCol As Integer
    Dim row As Variant
    Dim txt As String
    
    On Error GoTo ErrorHandling
    
    For i = 2 To x.Worksheets("Data").Cells(2, 1000).End(xlToLeft).Column
        If x.Worksheets("Data").Cells(2, i).Value = DataEntryBox.ProgramList.Value Then
            programCol = i
            For j = programCol To x.Worksheets("Data").Cells(3, programCol + 1).End(xlToRight).Column
                If x.Worksheets("Data").Cells(3, j).Value = DataEntryBox.SkillList.Value Then
                    skillCol = j
                End If
            Next j
        End If
    Next i
      
    If editRow = dateRows(0) Then
        rowsIndex = UBound(dateRows)
        If rowsIndex < 0 Then rowsIndex = 0
        editRow = dateRows(rowsIndex)
    Else
        rowsIndex = rowsIndex - 1
        If rowsIndex < 0 Then rowsIndex = 0
        editRow = dateRows(rowsIndex)
    End If
    
    DataEntryBox.txtEditDate = x.Worksheets("Data").Cells(editRow, programCol).Value
    DataEntryBox.txtEditScore = x.Worksheets("Data").Cells(editRow, skillCol).Value
    
    Union(x.Worksheets("Data").Cells(editRow, programCol), x.Worksheets("Data").Cells(editRow, skillCol)).Activate
    
    current = DataEntryBox.txtEditDate.Value
    
ErrorHandling:
    ErrHandling

End Sub

Private Sub buttonDoneData_Click()

    Unload Me
    DataSelect.Show

End Sub

Private Sub buttonNextData_Click()

    Dim i, j, programCol, skillCol, Score, arraySize
    Dim newDate As String
    Dim row As Variant
    
    On Error GoTo ErrorHandling
    
    newDate = DataEntryBox.SessionDate.Value
    Score = DataEntryBox.Score.Value
    
    For i = 2 To x.Worksheets("Data").Cells(2, 1000).End(xlToLeft).Column
        If x.Worksheets("Data").Cells(2, i).Value = DataEntryBox.ProgramList.Value Then
            programCol = i
            For j = programCol To x.Worksheets("Data").Cells(3, programCol + 1).End(xlToRight).Column
                If x.Worksheets("Data").Cells(3, j).Value = DataEntryBox.SkillList.Value Then
                    skillCol = j
                End If
            Next j
        End If
    Next i
    
    For i = 4 To (x.Worksheets("Data").Cells(2000, 1).End(xlUp).row + 1)
        If (x.Worksheets("Data").Cells(i - 1, 1).Value < DateValue(newDate) Or _
            x.Worksheets("Data").Cells(i - 1, 1).Value = DateValue(newDate)) _
            And (x.Worksheets("Data").Cells(i, 1).Value > DateValue(newDate) Or x.Worksheets("Data").Cells(i, 1).Value = "") Then
            If x.Worksheets("Data").Cells(i - 1, 1).Value = DateValue(newDate) And x.Worksheets("Data").Cells(i - 1, programCol).Value = "" Then
                'Do nothing
            Else
                x.Worksheets("Data").Cells(i, 1).Activate
                ActiveCell.EntireRow.Insert
                x.Worksheets("Data").Cells(i, 1).Value = DataEntryBox.SessionDate.Value
                Exit For
            End If
        End If
    Next i
    
    For i = x.Worksheets("Data").Cells(2000, 1).End(xlUp).row To 4 Step -1
        If x.Worksheets("Data").Cells(i, 1).Value = DateValue(newDate) Then
            x.Worksheets("Data").Cells(i, programCol).Activate
            x.Worksheets("Data").Cells(i, programCol).Value = DataEntryBox.SessionDate.Value
            x.Worksheets("Data").Cells(i, skillCol).Value = DataEntryBox.Score.Value
            Exit For
        End If
    Next i
    
    If DataEntryBox.SkillList <> "<Please select program first>" Then
        arraySize = 0
        For i = 4 To x.Worksheets("Data").Cells(4, 1).End(xlDown).row
            If x.Worksheets("Data").Cells(i, skillCol).Value <> "" Then arraySize = arraySize + 1
        Next i
    End If
              
    If arraySize > 0 Then arraySize = arraySize - 1
    ReDim dateRows(arraySize)
    
    j = 0
    
    If arraySize <> 0 Then
        For i = 4 To x.Worksheets("Data").Cells(4, 1).End(xlDown).row
            If x.Worksheets("Data").Cells(i, skillCol).Value <> "" Then
                dateRows(j) = x.Worksheets("Data").Cells(i, skillCol).row
                j = j + 1
            End If
        Next i
    End If
    
    DataEntryBox.SessionDate = ""
    DataEntryBox.Score = ""
    DataEntryBox.SessionDate.SetFocus
       
ErrorHandling:
    ErrHandling
                
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Program_Change()

    Dim i As Integer

    DataEntryBox.AddProgram.Enabled = True
    For i = 2 To x.Worksheets("Data").Cells(2, 1000).End(xlToLeft).Column
        If DataEntryBox.Program.Value = x.Worksheets("Data").Cells(2, i).Value Then
            DataEntryBox.AddProgram.Enabled = False
        End If
    Next i
    
End Sub

Private Sub ProgramList_Change()

    Dim programCol As Integer
    Dim i As Integer
    Dim skillCol As Integer
    
    programCol = 1
    
    DataEntryBox.SkillList.Clear
    DataEntryBox.SkillList.Enabled = True
    DataEntryBox.SkillList = "Please select skill..."

    For i = 2 To x.Worksheets("Data").Cells(2, 1000).End(xlToLeft).Column
        If x.Worksheets("Data").Cells(2, i).Value = DataEntryBox.ProgramList.Value Then
            programCol = i
            x.Worksheets("Data").Cells(2, i).Activate
            Exit For
        End If
    Next i
    
    skillCol = x.Worksheets("Data").Cells(2, programCol).End(xlToRight).Column
    
    
    For i = programCol + 1 To x.Worksheets("Data").Cells(3, skillCol).End(xlToLeft).Column
        DataEntryBox.SkillList.AddItem x.Worksheets("Data").Cells(3, i).Value
    Next i
    
    DataEntryBox.Program.Value = DataEntryBox.ProgramList.Value
    If DataEntryBox.Program.Value = "Select preexisting program..." Then
        DataEntryBox.Program = ""
    End If
    
        DataEntryBox.btnEditUp.Enabled = False
        DataEntryBox.btnEditDown.Enabled = False
        DataEntryBox.btnDelete.Enabled = False
        DataEntryBox.btnEdit.Enabled = False
          
End Sub

Private Sub ScollDown_Click()

    ActiveWindow.SmallScroll Down:=5

End Sub

Private Sub Score_Change()

    If DataEntryBox.Score.Value < 0 Or DataEntryBox.Score.Value > 100 Then
        DataEntryBox.buttonNextData.Enabled = False
    Else
        DataEntryBox.buttonNextData.Enabled = True
    End If

End Sub



Private Sub ScrollUp_Click()

    ActiveWindow.SmallScroll Down:=-5

End Sub

Private Sub SessionDate_Change()

    If IsDate(DataEntryBox.SessionDate.Value) Then DataEntryBox.buttonNextData.Enabled = True

End Sub

Private Sub SessionDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If IsDate(DataEntryBox.SessionDate.Value) Then
        DataEntryBox.buttonNextData.Enabled = True
    Else
        DataEntryBox.buttonNextData.Enabled = False
    End If
    
End Sub

Private Sub Skill_Change()

    If DataEntryBox.Skill.Value = "<Please select program first>" Or DataEntryBox.Skill.Value = "Please select skill..." Then
        DataEntryBox.Skill = ""
    End If
    If DataEntryBox.Skill.Value = "" Then
        DataEntryBox.AddSkill.Enabled = False
    Else
        DataEntryBox.AddSkill.Enabled = True
    End If
    
End Sub

Private Sub SkillList_Change()

    Dim i, j, arraySize, programCol, skillCol As Integer
    Dim row As Variant
    
    On Error GoTo ErrHandling

    DataEntryBox.Skill.Value = DataEntryBox.SkillList.Value
    
    'Check for empty programs
    If x.Worksheets("Data").Cells(2, 1).End(xlToRight).Value = "" Then Exit Sub
    
    For i = 2 To x.Worksheets("Data").Cells(2, 1000).End(xlToLeft).Column
        If x.Worksheets("Data").Cells(2, i).Value = DataEntryBox.ProgramList.Value Then
            programCol = i
            For j = programCol To x.Worksheets("Data").Cells(3, programCol + 1).End(xlToRight).Column
                If x.Worksheets("Data").Cells(3, j).Value = DataEntryBox.SkillList.Value Then
                    skillCol = j
                End If
            Next j
        End If
        If skillCol = 0 Then skillCol = 3
    Next i
    
    If DataEntryBox.SkillList <> "<Please select program first>" Then
        arraySize = 0
        For i = 4 To x.Worksheets("Data").Cells(4, 1).End(xlDown).row
            If x.Worksheets("Data").Cells(i, skillCol).Value <> "" Then arraySize = arraySize + 1
        Next i
    End If
              
    If arraySize > 0 Then arraySize = arraySize - 1
    ReDim dateRows(arraySize)
    
    j = 0
    
    If arraySize <> 0 Then
        For i = 4 To x.Worksheets("Data").Cells(4, 1).End(xlDown).row
            If x.Worksheets("Data").Cells(i, skillCol).Value <> "" Then
                dateRows(j) = x.Worksheets("Data").Cells(i, skillCol).row
                j = j + 1
            End If
        Next i
    End If
    
    editRow = dateRows(0)
    rowsIndex = 0
    
    If DataEntryBox.SkillList = "Please select skill..." Then
    Else
        DataEntryBox.btnEditUp.Enabled = True
        DataEntryBox.btnEditDown.Enabled = True
        DataEntryBox.btnDelete.Enabled = True
        DataEntryBox.btnEdit.Enabled = True
    End If
    
ErrHandling:
    ErrorHandling
                     
End Sub

Private Sub txtEditDate_Change()

    If IsDate(DataEntryBox.txtEditDate.Value) Then
        DataEntryBox.btnDelete.Enabled = True
        DataEntryBox.btnEdit.Enabled = True
    Else
        DataEntryBox.btnDelete.Enabled = False
        DataEntryBox.btnEdit.Enabled = False
    End If

End Sub

Private Sub txtEditScore_Change()

End Sub

Private Sub UserForm_Click()

End Sub
