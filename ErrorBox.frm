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

    Dim filepath, errorText, currentFile As String
    
    'Build the filepath/name
    filepath = "C:\Users\jackie\Documents\Client Files\Progress Reports\FMP_DataExport\Err\" & Format(Now, "yyyy_mm_dd-hh_mm_ss") & ".txt"
    
    'Check if the error occurred with a workbook open
    If ThisWorkbook.Name = "PERSONAL.XLSB" Then
        currentFile = X.FullName
    Else
        currentFile = ThisWorkbook.FullName
    End If
    
    'Build the text output for the error report
    errorText = Now & vbCrLf & "Error Number: " & err.number & vbCrLf & "Error Description: " & err.Description & vbCrLf & "Proceedure: " & errorTracking _
        & vbCrLf & currentFile & vbCrLf & "   Program: " & DataEntryBox.ProgramList.Value & vbCrLf & "   Skill: " & DataEntryBox.SkillList.Value _
        & vbCrLf & "   SessionDate: " & DataEntryBox.SessionDate.Value & vbCrLf _
        & "   SessionScore: " & DataEntryBox.Score.Value & vbCrLf & "User Description: " & ErrorBox.txtUserDescription.Value
    
    'Open/write/close error report text file
    Open filepath For Output As #1
        
    Write #1, errorText
    
    Close #1
    
    Unload Me

End Sub

Private Sub txtUserDescription_Change()

End Sub

Private Sub UserForm_Click()

End Sub
