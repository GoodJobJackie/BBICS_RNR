VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgramNames 
   Caption         =   "Please verify program names"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ProgramNames.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgramNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()

    ProgramName = Trim(ProgramNames.NewProgramName.Value)
    Cells(2, renameI).Value = ProgramName
    Unload Me

End Sub

Private Sub CommandButton2_Click()

    renameI = prevProgramName - 1

    Unload Me
    
End Sub

Private Sub currentProgramName_Change()

End Sub

Private Sub ProgramLists_Change()

    Dim i As Integer
    
    For i = 1 To x.Worksheets("PD").Cells(3000, 1).End(xlUp).Row
        If ProgramNames.ProgramLists.Value = Worksheets("PD").Cells(i, 1).Value Then
            ProgramNames.NewProgramName.Value = Worksheets("PD").Cells(i, 1).Value
        End If
    Next i

End Sub

Private Sub UserForm_Click()

End Sub
