VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MCMbox 
   Caption         =   "Select skill placement"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7260
   OleObjectBlob   =   "MCMbox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MCMbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnView_Click()

    Dim i, j, programCol, skillCol As Integer
    Dim txt As String

    'Locate and assign program/skill columns to variables
    For i = 2 To X.Worksheets("Data").Cells(2, 10000).End(xlToLeft).Column
        If X.Worksheets("Data").Cells(2, i).Value = MCMbox.programNameBox.Value Then
            programCol = i
            For j = programCol To X.Worksheets("Data").Cells(3, programCol + 1).End(xlToRight).Column
                If X.Worksheets("Data").Cells(3, j).Value = MCMbox.skillNameBox.Value Then
                    skillCol = j
                End If
            Next j
        End If
        If skillCol = 0 Then skillCol = 3
    Next i
    
    txt = ""
'    'Populate pairings list
'        For i = 4 To X.Worksheets("Data").Cells(4, 1).End(xlDown).row
'            txt = ""
'            If X.Worksheets("Data").Cells(i, skillCol).Value <> "" Then
'                txt = txt & X.Worksheets("Data").Cells(i, programCol).Value & "     " & X.Worksheets("Data").Cells(i, skillCol).Value & vbCrLf
'                If DateValue(X.Worksheets("Data").Cells(i, programCol).Value) = DateValue(reportStart) Or _
'                    DateValue(x.Worksheets("Data).Cells(i, 1).Value) < DateValue(reportStart) And
'            End If
'        Next i
'    End If

End Sub

Private Sub buttonSkip_Click()
    skipFlag = True
    MCMbox.progMast.Value = False
    MCMbox.progCont.Value = False
    MCMbox.progMaint.Value = False
    mCm = ""
    Unload Me
End Sub

Private Sub Label3_Click()
    MCMbox.progMast.Value = True
End Sub

Private Sub Label4_Click()
    MCMbox.progCont.Value = True
End Sub

Private Sub Label5_Click()
    MCMbox.progMaint.Value = True
End Sub

Private Sub NextProgram_Click()
    skipFlag = False
    ProgramName = MCMbox.programNameBox.Value
    SkillName = MCMbox.skillNameBox.Value
    If progMast = True Then
        mCm = "Mastered"
    End If
    If progCont = True Then
        mCm = "Continued"
    End If
    If progMaint = True Then
        mCm = "Maintenance"
    End If
    MCMbox.lblPairings.Caption = ""
    Unload Me
End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub progMaint_Click()

End Sub

Private Sub programNameBox_Change()

End Sub

Private Sub ScrollDown_Click()

    ActiveWindow.SmallScroll Down:=10

End Sub

Private Sub ScrollDownBig_Click()

    ActiveWindow.SmallScroll Down:=25

End Sub

Private Sub ScrollUp_Click()

    ActiveWindow.SmallScroll Down:=-10
    
End Sub

Private Sub ScrollUpBig_Click()

    ActiveWindow.SmallScroll Down:=-25

End Sub

Private Sub UserForm_Activate()
    With MCMbox
        .Top = Application.Top + 95
        .Left = Application.Left + 450
    End With
End Sub
