VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ClientSelect 
   Caption         =   "Select a workbook:"
   ClientHeight    =   1575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6915
   OleObjectBlob   =   "ClientSelect.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ClientSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FileSelect_Click()

    Dim filePath As String

    filePath = ClientSelect.FileList.Value
    
    Set x = Workbooks.Open(filePath)
    ActiveWindow.WindowState = xlMaximized
    x.Worksheets("Data").Activate
    
    Unload Me
    
    MsgBox ("Remember to save after making any changes.")
   
    DataSelect.Show

End Sub

Private Sub UserForm_Click()

End Sub
