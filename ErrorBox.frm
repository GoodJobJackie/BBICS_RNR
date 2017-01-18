VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ErrorBox 
   Caption         =   "UserForm2"
   ClientHeight    =   3375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6960
   OleObjectBlob   =   "ErrorBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ErrorBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()

    Unload Me

End Sub

Private Sub btnSubmit_Click()

    Dim filepath As String
    Dim errorText As String
    Dim currentFile As String
    
    filepath = "C:\Users\jackie\Documents\Client Files\Progress Reports\FMP_DataExport\Err\" & Format(Now, "yyyy_mm_dd-hh_mm_ss") & ".txt"
    
    If ThisWorkbook.Name = "PERSONAL.XLSB" Then
        currentFile = x.FullName
    Else
        currentFile = ThisWorkbook.FullName
    End If
    
    errorText = Now & vbCrLf & "Error Number: " & err.Number & vbCrLf & "Error Description: " & err.Description _
        & vbCrLf & currentFile & vbCrLf & "User Description: " & ErrorBox.txtUserDescription.Value
    
    Open filepath For Output As #1
    
    Write #1, errorText
    
    Close #1
    
    Unload Me

End Sub

Private Sub txtUserDescription_Change()

End Sub

Private Sub UserForm_Click()

End Sub
