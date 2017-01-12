VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RenameBox 
   Caption         =   "Please verify program name:"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "RenameBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RenameBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnChange_Click()

    Dim j As Integer

    Worksheets("Data").Cells(2, renameI).Value = RenameBox.programChange
    With RenameBox.programCurrent
        For j = 2 To Worksheets("Data").Cells(2, 2000).End(xlToLeft).Column
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
    If renameI = Worksheets("Data").Cells(2, 2000).End(xlToLeft).Column Then
        '
    Else
        btnNext_Click
    End If
    
End Sub

Private Sub btnDone_Click()

    Unload Me

End Sub

Private Sub btnNext_Click()

Dim i As Integer

    For i = 2 To Worksheets("Data").Cells(2, 2000).End(xlToLeft).Column
        If Worksheets("Data").Cells(2, i).Value = RenameBox.programExisting.Value Then
            renameI = Worksheets("Data").Cells(2, i).End(xlToRight).Column
            RenameBox.programExisting = Worksheets("Data").Cells(2, renameI).Value
            RenameBox.programCurrent = Worksheets("Data").Cells(2, renameI).Value
            RenameBox.programChange = Worksheets("Data").Cells(2, renameI).Value
            Worksheets("Data").Cells(2, renameI).Activate
            Exit For
        End If
    Next i
    
    If renameI = 2 Then
        RenameBox.btnPrev.Enabled = False
    Else
        RenameBox.btnPrev.Enabled = True
    End If
    
    If renameI = Worksheets("Data").Cells(2, 2000).End(xlToLeft).Column Then
        RenameBox.btnNext.Enabled = False
    Else
        RenameBox.btnNext.Enabled = True
    End If
    
    RenameBox.programChange.SetFocus
    With RenameBox.programChange
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub btnPrev_Click()

Dim i As Integer

    For i = 2 To Worksheets("Data").Cells(2, 2000).End(xlToLeft).Column
        If Worksheets("Data").Cells(2, i).Value = RenameBox.programExisting.Value Then
            renameI = Worksheets("Data").Cells(2, i).End(xlToLeft).Column
            RenameBox.programExisting = Worksheets("Data").Cells(2, renameI).Value
            RenameBox.programCurrent = Worksheets("Data").Cells(2, renameI).Value
            RenameBox.programChange = Worksheets("Data").Cells(2, renameI).Value
            Worksheets("Data").Cells(2, renameI).Activate
            Exit For
        End If
    Next i
    
    If renameI = 2 Then
        RenameBox.btnPrev.Enabled = False
    Else
        RenameBox.btnPrev.Enabled = True
    End If
    
    If renameI = Worksheets("Data").Cells(2, 2000).End(xlToLeft).Column Then
        RenameBox.btnNext.Enabled = False
    Else
        RenameBox.btnNext.Enabled = True
    End If
    
    RenameBox.programChange.SetFocus
    With RenameBox.programChange
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub programAll_Change()

    RenameBox.programChange = RenameBox.programAll.Value
    
    RenameBox.programChange.SetFocus
    With RenameBox.programChange
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub programCurrent_Change()

Dim i As Integer

    For i = 2 To Worksheets("Data").Cells(2, 2000).End(xlToLeft).Column
        If Worksheets("Data").Cells(2, i).Value = RenameBox.programCurrent.Value Then
            renameI = i
            Worksheets("Data").Cells(2, i).Activate
        End If
    Next i
    
    RenameBox.programChange = Worksheets("Data").Cells(2, renameI).Value
    RenameBox.programExisting = Worksheets("Data").Cells(2, renameI).Value
    
    If renameI = 2 Then
        RenameBox.btnPrev.Enabled = False
    Else
        RenameBox.btnPrev.Enabled = True
    End If
    
    If renameI = Worksheets("Data").Cells(2, 2000).End(xlToLeft).Column Then
        RenameBox.btnNext.Enabled = False
    Else
        RenameBox.btnNext.Enabled = True
    End If
    
    RenameBox.programChange.SetFocus
    With RenameBox.programChange
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub programExisting_Change()

End Sub

Private Sub UserForm_Click()

End Sub
