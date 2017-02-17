VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Please select report dates"
   ClientHeight    =   2655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2670
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()
Dim sheetName

Dim i, bottomDateRow As Integer

    Unload Me
    sheetName = ActiveSheet.Name
    Worksheets(sheetName).Activate
    bottomDateRow = Cells(5, 1).End(xlDown).row
    reportStart = Trim(ComboBox1.Value)
    reportEnd = Trim(ComboBox2.Value)

    For i = 4 To bottomDateRow
        If Trim(ComboBox1.Value) = Trim(Cells(i, 1).Value) Then
            Rows(i).Select
            With Selection.Borders(xlEdgeTop)
                .Color = RGB(255, 0, 0)
                .LineStyle = xlContinuous
            End With
            startDateRow = i
        End If
        If Trim(Cells(i, 1).Value) = Trim(ComboBox2.Value) Then
            Rows(i).Select
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = RGB(255, 0, 0)
            End With
            endDateRow = i
        End If
    Next i
    
    Rows(startDateRow & ":" & endDateRow).Select
    Selection.Interior.Color = RGB(255, 245, 245)

End Sub

Private Sub UserForm_Click()

End Sub
