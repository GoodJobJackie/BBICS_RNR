VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataEntryBox 
   Caption         =   "Enter New Data"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7455
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

Private Sub buttonDoneData_Click()

    Unload Me
    DataSelect.Show

End Sub

Private Sub buttonNextData_Click()

    Dim i As Integer
    Dim j As Integer
    Dim programCol As Integer
    Dim skillCol As Integer
    Dim newDate As String
    Dim Score As Integer
    
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
    
    For i = 4 To (x.Worksheets("Data").Cells(2000, 1).End(xlUp).Row + 1)
        If (x.Worksheets("Data").Cells(i - 1, 1).Value < DateValue(newDate)) And ((x.Worksheets("Data").Cells(i, 1).Value > DateValue(newDate)) Or (x.Worksheets("Data").Cells(i, 1).Value = "")) Then
            x.Worksheets("Data").Cells(i, 1).Activate
            ActiveCell.EntireRow.Insert
            x.Worksheets("Data").Cells(i, 1).Value = DataEntryBox.SessionDate.Value
            Exit For
        End If
    Next i
    
    For i = 4 To x.Worksheets("Data").Cells(2000, 1).End(xlUp).Row
        If x.Worksheets("Data").Cells(i, 1).Value = DateValue(newDate) Then
            x.Worksheets("Data").Cells(i, programCol).Activate
            x.Worksheets("Data").Cells(i, programCol).Value = DataEntryBox.SessionDate.Value
            x.Worksheets("Data").Cells(i, skillCol).Value = DataEntryBox.Score.Value
        End If
    Next i
    
    DataEntryBox.SessionDate = ""
    DataEntryBox.Score = ""
    DataEntryBox.SessionDate.SetFocus
                
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

End Sub

Private Sub Score_Change()

End Sub

Private Sub SessionDate_Change()

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

    DataEntryBox.Skill.Value = DataEntryBox.SkillList.Value
        
End Sub

Private Sub UserForm_Click()

End Sub
