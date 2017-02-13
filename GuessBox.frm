VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GuessBox 
   Caption         =   "Guess The Number!"
   ClientHeight    =   2220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3360
   OleObjectBlob   =   "GuessBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GuessBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnGuess_Click()

    guess = Int(GuessBox.txtGuess.Value)
    
    If GuessBox.btnGuess.Caption = "Try Again" Then
        Unload Me
        InitGuessBox
        GuessBox.Show
    ElseIf GuessBox.btnGuess.Caption = "Guess!" Then
        If guess = number Then
            'MsgBox ("=")
            guessText = "Hooray! You got it right in " & guesses & " tries!"
            GuessBox.guessText.Caption = guessText
            GuessBox.btnGuess.Caption = "Try Again"
        ElseIf guess < number Then
            'MsgBox ("<")
            lower = guess
            GuessBox.lower.Caption = lower
            guesses = guesses + 1
            GuessBox.txtGuess = ""
            GuessBox.txtGuess.SetFocus
            guessText = "Guesses: " & guesses
            GuessBox.guessText.Caption = guessText
            GuessBox.btnGuess.Caption = "Guess!"
        ElseIf guess > number Then
            'MsgBox (">")
            upper = guess
            GuessBox.upper.Caption = upper
            guesses = guesses + 1
            GuessBox.txtGuess = ""
            GuessBox.txtGuess.SetFocus
            guessText = "Guesses: " & guesses
            GuessBox.guessText.Caption = guessText
            GuessBox.btnGuess.Caption = "Guess!"
        End If
    End If

End Sub

Private Sub txtGuess_Change()
   
End Sub

Private Sub UserForm_Click()

End Sub
