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

Private Sub FileList_Change()

    ClientSelect.FileSelect.Enabled = True

End Sub

Private Sub FileSelect_Click()

    Dim filepath As String
    
    On Error Resume Next

    'Open selected client's workbook
    filepath = ClientSelect.FileList.Value
       
    Set X = Workbooks.Open(filepath)
    
    X.Activate                              '
    X.Worksheets("Data").Activate           ' Apparently, do nothing in Excel 2016...
    ActiveWindow.WindowState = xlMaximized  '
    
    'Close dialog box and update version number
    Unload Me
    UserAction.version.Caption = version
    
    'Check if opened workbook exists and change main menu buttons accordingly
    If X Is Nothing Then
        UserAction.ActionDataEntry.Enabled = True
        UserAction.actionSaveWorkbook.Enabled = False
        UserAction.actionCloseWorkbook.Enabled = False
        UserAction.actionIPG.Enabled = False
        UserAction.btnDataEntry.Enabled = False
        UserAction.VerifyProgramNames.Enabled = False
    Else
        UserAction.ActionDataEntry.Enabled = False
        UserAction.actionSaveWorkbook.Enabled = True
        UserAction.actionCloseWorkbook.Enabled = True
        UserAction.actionIPG.Enabled = True
        UserAction.btnDataEntry.Enabled = True
        UserAction.VerifyProgramNames.Enabled = True
    End If

    'Open the data select dialog box
    UserAction.Show

End Sub

Private Sub UserForm_Click()

End Sub
