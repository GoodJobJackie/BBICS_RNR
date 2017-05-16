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
    Dim strDataRange, keyRange As Range
    
    'Decide whether client exists already
    found = False
    For i = 2 To Y.Worksheets("Client File Archive").Cells(2, 1).End(xlDown).row
        If Y.Worksheets("Client File Archive").Cells(i, 1).Value = UCase(ArchiveBox.txtInitials.Value) Then found = True
    Next i
    
    'If new client, add to individuals list
    row = Y.Worksheets("Client File Archive").Cells(2, 4).End(xlDown).row + 1
    If found = False Then Y.Worksheets("Client File Archive").Cells(row, 4) = UCase(ArchiveBox.txtInitials.Value)
    
    'Add client/box to list
    row = Y.Worksheets("Client File Archive").Cells(2, 1).End(xlDown).row + 1
    Y.Worksheets("Client File Archive").Cells(row, 1) = UCase(ArchiveBox.txtInitials)
    Y.Worksheets("Client File Archive").Cells(row, 2) = CInt(ArchiveBox.txtBox)
    
    'Sort client list alphabetically
    Set strDataRange = Range("A2:B" & Y.Worksheets("Client File Archive").Cells(1, 1).End(xlDown).row)
    Set keyRange = Range("A2:B" & Y.Worksheets("Client File Archive").Cells(1, 1).End(xlDown).row)
    strDataRange.Sort Key1:=keyRange, Order1:=xlAscending
    Set strDataRange = Range("D2:D" & Y.Worksheets("Client File Archive").Cells(2, 4).End(xlDown).row)
    Set keyRange = Range("D2:D" & Y.Worksheets("Client File Archive").Cells(2, 4).End(xlDown).row)
    strDataRange.Sort Key1:=keyRange, Order1:=xlAscending
    
    'Save master list
    Y.Save
    
    'Reload client drop-down list
    ArchiveBox.selectClient.Clear
    For i = 2 To Y.Worksheets("Client File Archive").Cells(2, 4).End(xlDown).row
        With ArchiveBox.selectClient
            .AddItem Y.Worksheets("Client File Archive").Cells(i, 4).Value
        End With
    Next i
    
    ArchiveBox.selectClient.Value = UCase(ArchiveBox.txtInitials.Value)
    
    'Reset text boxes
    ArchiveBox.txtInitials = ""
    ArchiveBox.txtBox = ""
'    ArchiveBox.btnAdd.Enabled = False

End Sub

Private Sub btnDone_Click()

    On Error Resume Next
    
    Unload Me
    Y.Close
    X.Activate
    ActiveWindow.WindowState = xlMaximized
    UserAction.version.Caption = version
    'UserAction.Show

End Sub

Private Sub Label2_Click()

End Sub

Private Sub selectClient_Change()

    Dim i, j As Integer
    Dim boxes As String
    
    'Locate which boxes the selected client are in and store as a variable
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
    
    'Display boxes
    ArchiveBox.boxes = boxes
    
    If ArchiveBox.selectClient <> "" Then ArchiveBox.txtInitials.Value = ArchiveBox.selectClient.Value

End Sub

Private Sub txtBox_Change()

    If ArchiveBox.txtInitials.Value <> "" And ArchiveBox.txtBox.Value <> "" Then ArchiveBox.btnAdd.Enabled = True
    If ArchiveBox.txtInitials.Value = "" Or ArchiveBox.txtBox.Value = "" Then ArchiveBox.btnAdd.Enabled = False

End Sub

Private Sub txtInitials_Change()

    If ArchiveBox.txtInitials.Value <> "" And ArchiveBox.txtBox.Value <> "" Then ArchiveBox.btnAdd.Enabled = True
    If ArchiveBox.txtInitials.Value = "" Or ArchiveBox.txtBox.Value = "" Then ArchiveBox.btnAdd.Enabled = False

End Sub

Private Sub UserForm_Click()

End Sub
