VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ArchiveBox 
   Caption         =   "Client File Archive"
   ClientHeight    =   2115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3570
   OleObjectBlob   =   "ArchiveBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ArchiveBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAdd_Click()

    Dim i, row As Integer
    Dim found As Boolean
    
    found = False
    For i = 2 To Y.Worksheets("Client File Archive").Cells(2, 1).End(xlDown).row
        If Y.Worksheets("Client File Archive").Cells(i, 1).Value = UCase(ArchiveBox.txtInitials.Value) Then found = True
    Next i
    
    row = Y.Worksheets("Client File Archive").Cells(2, 4).End(xlDown).row + 1
    If found = False Then Y.Worksheets("Client File Archive").Cells(row, 4) = UCase(ArchiveBox.txtInitials.Value)
    
    row = Y.Worksheets("Client File Archive").Cells(2, 1).End(xlDown).row + 1
    Y.Worksheets("Client File Archive").Cells(row, 1) = UCase(ArchiveBox.txtInitials)
    Y.Worksheets("Client File Archive").Cells(row, 2) = CInt(ArchiveBox.txtBox)
    
    ArchiveBox.txtInitials = ""
    ArchiveBox.txtBox = ""
    
    Y.Save

End Sub

Private Sub btnDone_Click()

    Unload Me
    Y.Close
    UserAction.Show

End Sub

Private Sub Label2_Click()

End Sub

Private Sub selectClient_Change()

    Dim i, j As Integer
    Dim boxes As String
    
    j = 0
    For i = 2 To Y.Worksheets("Client File Archive").Cells(2, 1).End(xlDown).row
        If Y.Worksheets("Client File Archive").Cells(i, 1).Value = ArchiveBox.selectClient.Value Then
            If j = 0 Then
                boxes = Y.Worksheets("Client File Archive").Cells(i, 2).Value
                j = j + 1
            Else
                boxes = boxes & ", " & Y.Worksheets("Client File Archive").Cells(i, 2).Value
                j = j + 1
            End If
        End If
    Next i
    
    ArchiveBox.boxes = boxes

End Sub

Private Sub UserForm_Click()

End Sub
