VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataSelect 
   Caption         =   "UserForm2"
   ClientHeight    =   3135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2325
   OleObjectBlob   =   "DataSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()

    Unload Me
    DataEntryPrograms

End Sub

Private Sub CommandButton2_Click()

    x.Worksheets("Bx Data").Activate
    Unload Me

End Sub

Private Sub CommandButton3_Click()

    x.Worksheets("Tutor Hr Data").Activate
    Unload Me

End Sub

Private Sub DataSelectDone_Click()

    Dim fileName As String

    Unload Me
    
    fileName = x.Worksheets("Data").Cells(1, 1).Value + " - " + Format(x.Worksheets("Data").Cells(1000, 1).End(xlUp).Value, "YYYY") _
        + "_" + Format(x.Worksheets("Data").Cells(1000, 1).End(xlUp).Value, "MM") _
        + "_" + Format(x.Worksheets("Data").Cells(1000, 1).End(xlUp).Value, "DD")
    
    'Application.GetSaveAsFilename ("C:\Users\jackie\Documents\Client Files\Data\Formatted\" + fileName)
    
    UserAction.Show

End Sub

Private Sub Label2_Click()

End Sub

Private Sub UserForm_Click()

End Sub
