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

Private Sub AddProgram_Click()

    On Error GoTo ErrorHandling
    errorTracking = "AddProgram_Click"
    
    Dim col, i, programCol As Integer
    
    'Add and format new program columns
    DataEntryBox.ProgramList.AddItem DataEntryBox.Program
    col = X.Worksheets("Data").Cells(3, 10000).End(xlToLeft).Column + 2
    X.Worksheets("Data").Columns(col).Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.NumberFormat = "mm/dd/yyyy"
  
    X.Worksheets("Data").Cells(2, col).Value = DataEntryBox.Program.Value
    X.Worksheets("Data").Cells(3, col).Value = " "
       
    DataEntryBox.ProgramList.Clear
    For col = 2 To Worksheets("Data").Cells(2, 10000).End(xlToLeft).Column
        If Worksheets("Data").Cells(2, col).Value = "" Then
        Else
            DataEntryBox.ProgramList.AddItem Cells(2, col).Value
        End If
    Next col
    
    DataEntryBox.ProgramList = DataEntryBox.Program.Value
    DataEntryBox.Skill.SetFocus
    
    'Assign program column to variable and highlight
    For i = 2 To X.Worksheets("Data").Cells(2, 10000).End(xlToLeft).Column
        If X.Worksheets("Data").Cells(2, i).Value = DataEntryBox.ProgramList.Value Then
            programCol = i
            X.Worksheets("Data").Cells(2, i).Activate
            Exit For
        End If
    Next i
    
    'Highlight the newly added program
    X.Worksheets("Data").Cells(2, 10000).End(xlToLeft).Activate
    DataEntryBox.ProgramList = X.Worksheets("Data").Cells(2, 10000).End(xlToLeft).Value
    
ErrorHandling:
    ErrHandling
    
End Sub

Private Sub AddSkill_Click()

    On Error GoTo ErrorHandling
    errorTracking = "AddSkill_Click"
    
    Dim i, j, col, programCol, skillCol, arraySize As Integer
    Dim txt As String
    
    'Add and format new skill column
    DataEntryBox.SkillList.AddItem DataEntryBox.Skill.Value
    DataEntryBox.SkillList = DataEntryBox.Skill.Value
    For i = 2 To X.Worksheets("Data").Cells(2, 10000).End(xlToLeft).Column
        If X.Worksheets("Data").Cells(2, i).Value = DataEntryBox.ProgramList.Value Then
            col = i + 1
            If X.Worksheets("Data").Cells(3, col).Value = "" Then
                X.Worksheets("Data").Cells(3, col).Value = DataEntryBox.SkillList.Value
            ElseIf X.Worksheets("Data").Cells(3, col + 1).Value = "" Then
                X.Worksheets("Data").Cells(3, col).Activate
                ActiveCell.EntireColumn.Offset(0, 1).Insert
                ActiveCell.Offset(0, 1).Value = DataEntryBox.SkillList.Value
            Else
                X.Worksheets("Data").Cells(3, col).End(xlToRight).Activate
                ActiveCell.EntireColumn.Offset(0, 1).Insert
                ActiveCell.Offset(0, 1).Value = DataEntryBox.SkillList.Value
            End If
        End If
    Next i
    
    'Find program and skill/store column values
    For i = 2 To X.Worksheets("Data").Cells(2, 10000).End(xlToLeft).Column
        If X.Worksheets("Data").Cells(2, i).Value = DataEntryBox.ProgramList.Value Then
            programCol = i
            For j = programCol To X.Worksheets("Data").Cells(3, programCol + 1).End(xlToRight).Column
                If X.Worksheets("Data").Cells(3, j).Value = DataEntryBox.SkillList.Value Then
                    skillCol = j
                End If
            Next j
        End If
    Next i
    
    'Highlight the newly added skill
    If X.Worksheets("Data").Cells(3, skillCol + 1).Value = "" Then
        X.Worksheets("Data").Cells(3, skillCol).Activate
    Else
        X.Worksheets("Data").Cells(3, skillCol).End(xlToRight).Activate
    End If
    
    DataEntryBox.SessionDate.SetFocus
    
    arraySize = 0
    
    'Populate pairings list
    If DataEntryBox.SkillList <> "<Please select program first>" And DataEntryBox.SkillList <> "Please select skill..." Then
        DataEntryBox.pairings.Clear
        For i = 4 To X.Worksheets("Data").Cells(4, 1).End(xlDown).row
            txt = ""
            If X.Worksheets("Data").Cells(i, skillCol).Value <> "" Then
                txt = X.Worksheets("Data").Cells(i, programCol).Value & "     " & X.Worksheets("Data").Cells(i, skillCol).Value
                DataEntryBox.pairings.AddItem txt
            End If
        Next i
    End If
    
ErrorHandling:
    ErrHandling

End Sub

Private Sub btnDelete_Click()

    Dim i, j, programCol, skillCol, arraySize As Integer
    Dim row As Variant
    Dim txt As String
    Dim scoreDate As Date
    Dim Score As Double
    
    On Error Resume Next
    errorTracking = "btnDelete_Click"
    
    'Find program/skill and store column values
    For i = 2 To X.Worksheets("Data").Cells(2, 10000).End(xlToLeft).Column
        If X.Worksheets("Data").Cells(2, i).Value = DataEntryBox.ProgramList.Value Then
            programCol = i
            For j = programCol To X.Worksheets("Data").Cells(3, programCol + 1).End(xlToRight).Column
                If X.Worksheets("Data").Cells(3, j).Value = DataEntryBox.SkillList.Value Then
                    skillCol = j
                End If
            Next j
        End If
    Next i
    
    'Set date/score for selected pairing
    If DataEntryBox.SkillList <> "" Or DataEntryBox.pairings <> "" Then
        scoreDate = Trim(Left(DataEntryBox.pairings.Value, 10))
        Score = CDbl(Trim(Right(DataEntryBox.pairings.Value, 3)))
    End If
 
    'Find and delete corresponding pairing
    For i = 4 To X.Worksheets("Data").Cells(4, 1).End(xlDown).row
        If X.Worksheets("Data").Cells(i, programCol).Value = scoreDate And X.Worksheets("Data").Cells(i, skillCol).Value = Score Then
            X.Worksheets("Data").Cells(i, programCol).Value = ""
            X.Worksheets("Data").Cells(i, skillCol).Value = ""
            Exit For
        End If
    Next i
    
    'Clear edit text fields
    DataEntryBox.txtEditDate = ""
    DataEntryBox.txtEditScore = ""
    DataEntryBox.SessionDate.SetFocus
    
    'Populate pairings list
    If DataEntryBox.SkillList <> "<Please select program first>" And DataEntryBox.SkillList <> "Please select skill..." Then
        DataEntryBox.pairings.Clear
        For i = 4 To X.Worksheets("Data").Cells(4, 1).End(xlDown).row
            txt = ""
            If X.Worksheets("Data").Cells(i, skillCol).Value <> "" Then
                txt = X.Worksheets("Data").Cells(i, programCol).Value & "     " & X.Worksheets("Data").Cells(i, skillCol).Value
                DataEntryBox.pairings.AddItem txt
            End If
        Next i
    End If
    DataEntryBox.pairings = ""
    
End Sub

Private Sub btnEdit_Click()

    Dim i, j, programCol, skillCol, arraySize As Integer
    Dim row As Variant
    Dim txt As String
    Dim scoreDate As Date
    Dim Score As Double
    
    On Error Resume Next
    errorTracking = "btnEdit_Click"
    
    'Get listing for program and skill columns
    For i = 2 To X.Worksheets("Data").Cells(2, 10000).End(xlToLeft).Column
        If X.Worksheets("Data").Cells(2, i).Value = DataEntryBox.ProgramList.Value Then
            programCol = i
            For j = programCol To X.Worksheets("Data").Cells(3, programCol + 1).End(xlToRight).Column
                If X.Worksheets("Data").Cells(3, j).Value = DataEntryBox.SkillList.Value Then
                    skillCol = j
                End If
            Next j
        End If
    Next i
    
    'Set date/score variables for selected pairing
    If DataEntryBox.SkillList <> "" Or DataEntryBox.pairings <> "" Then
        scoreDate = Trim(Left(DataEntryBox.pairings.Value, 10))
        Score = CDbl(Trim(Right(DataEntryBox.pairings.Value, 3)))
    End If
 
    'Find and delete corresponding pairing
    For i = 4 To X.Worksheets("Data").Cells(4, 1).End(xlDown).row
        If X.Worksheets("Data").Cells(i, programCol).Value = scoreDate And X.Worksheets("Data").Cells(i, skillCol).Value = Score Then
            X.Worksheets("Data").Cells(i, programCol).Value = ""
            X.Worksheets("Data").Cells(i, skillCol).Value = ""
            Exit For
        End If
    Next i
    
    'Copy from edit section to new data section
    DataEntryBox.SessionDate = DataEntryBox.txtEditDate
    DataEntryBox.Score = DataEntryBox.txtEditScore
    
    'Add the data as new
    buttonNextData_Click
    
    'Clear edit text fields
    DataEntryBox.txtEditDate = ""
    DataEntryBox.txtEditScore = ""
    DataEntryBox.SessionDate.SetFocus
    
    'Populate pairings list
    If DataEntryBox.SkillList <> "<Please select program first>" And DataEntryBox.SkillList <> "Please select skill..." Then
        DataEntryBox.pairings.Clear
        For i = 4 To X.Worksheets("Data").Cells(4, 1).End(xlDown).row
            txt = ""
            If X.Worksheets("Data").Cells(i, skillCol).Value <> "" Then
                txt = X.Worksheets("Data").Cells(i, programCol).Value & "     " & X.Worksheets("Data").Cells(i, skillCol).Value
                DataEntryBox.pairings.AddItem txt
            End If
        Next i
    End If
    DataEntryBox.pairings = ""
                
End Sub

Private Sub buttonDoneData_Click()

    Unload Me
    UserAction.ActionDataEntry.Enabled = False
    UserAction.version.Caption = version
    UserAction.Show

End Sub

Private Sub buttonNextData_Click()

    Dim i, j, programCol, skillCol, Score, arraySize As Integer
    Dim newDate, txt As String
    Dim row As Variant
    
    On Error Resume Next
    errorTracking = "buttonNextData_Click"
    
    newDate = DataEntryBox.SessionDate.Value
    Score = DataEntryBox.Score.Value
  
    If IsDate(DataEntryBox.SessionDate) = False Then
        MsgBox "Please enter valid date.", vbCritical
        With DataEntryBox.SessionDate
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        Exit Sub
    End If
    
    If IsDate(DataEntryBox.SessionDate.Value) Then
        If DateValue(DataEntryBox.SessionDate.Value) > Now + 30 Then
            MsgBox "Please enter valid date.", vbCritical
            With DataEntryBox.SessionDate
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
            Exit Sub
        End If
    End If
    
    'Find program and skill/store column values
    For i = 2 To X.Worksheets("Data").Cells(2, 10000).End(xlToLeft).Column
        If X.Worksheets("Data").Cells(2, i).Value = DataEntryBox.ProgramList.Value Then
            programCol = i
            For j = programCol To X.Worksheets("Data").Cells(3, programCol + 1).End(xlToRight).Column
                If X.Worksheets("Data").Cells(3, j).Value = DataEntryBox.SkillList.Value Then
                    skillCol = j
                End If
            Next j
        End If
    Next i
    
    'Find/Insert date of new data
    For i = 4 To (X.Worksheets("Data").Cells(2000, 1).End(xlUp).row + 1)
        If (X.Worksheets("Data").Cells(i - 1, 1).Value < DateValue(newDate) Or X.Worksheets("Data").Cells(i - 1, 1).Value = DateValue(newDate)) _
            And (X.Worksheets("Data").Cells(i, 1).Value > DateValue(newDate) Or X.Worksheets("Data").Cells(i, 1).Value = "") Then
            If X.Worksheets("Data").Cells(i - 1, 1).Value = DateValue(newDate) And X.Worksheets("Data").Cells(i - 1, skillCol).Value = "" Then
                'Do nothing
            Else
                X.Worksheets("Data").Cells(i, 1).Activate
                ActiveCell.EntireRow.Insert
                X.Worksheets("Data").Cells(i, 1).Value = DataEntryBox.SessionDate.Value
                Exit For
            End If
        End If
    Next i
    
    'Insert new data
    For i = X.Worksheets("Data").Cells(2000, 1).End(xlUp).row To 4 Step -1
        If X.Worksheets("Data").Cells(i, 1).Value = DateValue(newDate) Then
            X.Worksheets("Data").Cells(i, programCol).Activate
            X.Worksheets("Data").Cells(i, programCol).Value = DataEntryBox.SessionDate.Value
            X.Worksheets("Data").Cells(i, skillCol).Value = CDbl(DataEntryBox.Score.Value)
            Exit For
        End If
    Next i
    
    'Check for unselected program
    If DataEntryBox.SkillList <> "<Please select program first>" Then
        arraySize = 0
        For i = 4 To X.Worksheets("Data").Cells(4, 1).End(xlDown).row
            If X.Worksheets("Data").Cells(i, skillCol).Value <> "" Then arraySize = arraySize + 1
        Next i
    End If
              
    'Empty text boxes
    DataEntryBox.SessionDate = ""
    DataEntryBox.Score = ""
    DataEntryBox.txtEditDate = ""
    DataEntryBox.txtEditScore = ""
    DataEntryBox.SessionDate.SetFocus
    
    'Populate pairings list
    If DataEntryBox.SkillList <> "<Please select program first>" And DataEntryBox.SkillList <> "Please select skill..." Then
        DataEntryBox.pairings.Clear
        For i = 4 To X.Worksheets("Data").Cells(4, 1).End(xlDown).row
            txt = ""
            If X.Worksheets("Data").Cells(i, skillCol).Value <> "" Then
                txt = X.Worksheets("Data").Cells(i, programCol).Value & "     " & X.Worksheets("Data").Cells(i, skillCol).Value
                DataEntryBox.pairings.AddItem txt
            End If
        Next i
    End If
          
End Sub

Private Sub pairings_Change()

    Dim i, j, programCol, skillCol As Integer
    Dim scoreDate As Date
    Dim Score As Double
    
    On Error Resume Next
        
    'Find program and skill/store column values
    For i = 2 To X.Worksheets("Data").Cells(2, 10000).End(xlToLeft).Column
        If X.Worksheets("Data").Cells(2, i).Value = DataEntryBox.ProgramList.Value Then
            programCol = i
            For j = programCol To X.Worksheets("Data").Cells(3, programCol + 1).End(xlToRight).Column
                If X.Worksheets("Data").Cells(3, j).Value = DataEntryBox.SkillList.Value Then
                    skillCol = j
                End If
            Next j
        End If
    Next i
    
    'Set date/score for selected pairing
    If DataEntryBox.SkillList <> "" Or DataEntryBox.pairings <> "" Then
        scoreDate = Trim(Left(DataEntryBox.pairings.Value, 10))
        Score = CDbl(Trim(Right(DataEntryBox.pairings.Value, 3)))
    End If

    'Find and highlight selected pairing
    If DataEntryBox.SkillList <> "" Or DataEntryBox.pairings <> "" Then
        For i = 4 To X.Worksheets("Data").Cells(4, 1).End(xlDown).row
            If X.Worksheets("Data").Cells(i, skillCol).Value <> "" Then
                If DateValue(X.Worksheets("Data").Cells(i, programCol).Value) = DateValue(scoreDate) And X.Worksheets("Data").Cells(i, skillCol).Value = Score Then
                    Union(X.Worksheets("Data").Cells(i, programCol), X.Worksheets("Data").Cells(i, skillCol)).Activate
                    DataEntryBox.txtEditDate = X.Worksheets("Data").Cells(i, programCol).Value
                    DataEntryBox.txtEditScore = X.Worksheets("Data").Cells(i, skillCol).Value
                    Exit For
                End If
            End If
        Next i
    End If

End Sub

Private Sub Program_Change()

    Dim i As Integer

    DataEntryBox.AddProgram.Enabled = True
    For i = 2 To X.Worksheets("Data").Cells(2, 10000).End(xlToLeft).Column
        If DataEntryBox.Program.Value = X.Worksheets("Data").Cells(2, i).Value Then
            DataEntryBox.AddProgram.Enabled = False
        End If
    Next i
    
End Sub

Private Sub ProgramList_Change()

    Dim programCol, i, skillCol As Integer
    
    programCol = 1
    
    DataEntryBox.SkillList.Clear
    DataEntryBox.SkillList.Enabled = True
    DataEntryBox.SkillList = "Please select skill..."

    'Assign program column to variable and highlight
    For i = 2 To X.Worksheets("Data").Cells(2, 10000).End(xlToLeft).Column
        If X.Worksheets("Data").Cells(2, i).Value = DataEntryBox.ProgramList.Value Then
            programCol = i
            X.Worksheets("Data").Cells(2, i).Activate
            Exit For
        End If
    Next i
    
    'Fill skill list
    skillCol = X.Worksheets("Data").Cells(2, programCol).End(xlToRight).Column
    For i = programCol + 1 To X.Worksheets("Data").Cells(3, skillCol).End(xlToLeft).Column
        DataEntryBox.SkillList.AddItem X.Worksheets("Data").Cells(3, i).Value
    Next i
    
    'Enter selected program into program box
    DataEntryBox.Program.Value = DataEntryBox.ProgramList.Value
    If DataEntryBox.Program.Value = "Select preexisting program..." Then
        DataEntryBox.Program = ""
    End If
    
    'Reset edit panel
    DataEntryBox.btnDelete.Enabled = False
    DataEntryBox.btnEdit.Enabled = False
    
    Worksheets("Data").Cells(2, programCol).Select
          
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
    Dim txt As String
    
    On Error GoTo ErrorHandling
    errorTracking = "SkillList_Change"

    DataEntryBox.Skill.Value = DataEntryBox.SkillList.Value
    
    'Check for empty programs
    If X.Worksheets("Data").Cells(2, 1).End(xlToRight).Value = "" Then Exit Sub
    
    'Locate and assign program/skill columns to variables
    For i = 2 To X.Worksheets("Data").Cells(2, 10000).End(xlToLeft).Column
        If X.Worksheets("Data").Cells(2, i).Value = DataEntryBox.ProgramList.Value Then
            programCol = i
            For j = programCol To X.Worksheets("Data").Cells(3, programCol + 1).End(xlToRight).Column
                If X.Worksheets("Data").Cells(3, j).Value = DataEntryBox.SkillList.Value Then
                    skillCol = j
                End If
            Next j
        End If
        
        If skillCol = 0 Then skillCol = 3
    Next i
    
    'Reset edit panel
    If DataEntryBox.SkillList = "Please select skill..." Then
    Else
        DataEntryBox.btnDelete.Enabled = True
        DataEntryBox.btnEdit.Enabled = True
    End If
    
    'Clear edit panel and set focus
    DataEntryBox.txtEditDate = ""
    DataEntryBox.txtEditScore = ""
    DataEntryBox.SessionDate.SetFocus
    Worksheets("Data").Cells(3, skillCol).Select
    
    'Populate pairings list
    If DataEntryBox.SkillList <> "<Please select program first>" And DataEntryBox.SkillList <> "Please select skill..." Then
        DataEntryBox.pairings.Clear
        For i = 4 To X.Worksheets("Data").Cells(4, 1).End(xlDown).row
            txt = ""
            If X.Worksheets("Data").Cells(i, skillCol).Value <> "" Then
                txt = X.Worksheets("Data").Cells(i, programCol).Value & "     " & X.Worksheets("Data").Cells(i, skillCol).Value
                DataEntryBox.pairings.AddItem txt
            End If
        Next i
    End If
        
ErrorHandling:
    ErrHandling
                     
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
