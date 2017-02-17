VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AboutBox 
   Caption         =   "About BBICS-DMS"
   ClientHeight    =   4035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2985
   OleObjectBlob   =   "AboutBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnReadme_Click()

    Dim strFile

    strFile = "C:\Users\jackie\Documents\BBICS_DMS\Readme.txt"

    Shell "cmd /c """ & strFile & """", 0

End Sub

Private Sub CommandButton1_Click()

    Unload Me

End Sub

Private Sub Label3_Click()

InitGuessBox
    GuessBox.guessText.Caption = guessText
    GuessBox.Show
    
    With GuessBox.txtGuess
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With

End Sub

Private Sub UserForm_Click()

End Sub
