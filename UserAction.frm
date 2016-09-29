VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserAction 
   Caption         =   "Please select an action."
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2640
   OleObjectBlob   =   "UserAction.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActionFullService_Click()

    Unload Me
    NewRestructuring
    UserAction.ActionFullService.Enabled = False
    UserAction.ActionRestructureSingle.Enabled = False

End Sub

Private Sub ActionProgramList_Click()

    Unload Me
    PopulatePrograms
    Cells(1, 1).Select
    ActiveWindow.Zoom = 90
    CreateProgramLists

End Sub

Private Sub ActionReformat_Click()

    Unload Me
    ActiveWindow.Zoom = 90
    CreateHeader
    EmptyBCheck
    MasterListFormat
    FormatProgramDates
    FindLastDate
    Cells(4, 2).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    UserAction.ActionFullService.Enabled = False

End Sub

Private Sub ActionRestructureSingle_Click()

    Unload Me
    SingleRestructure
    UserAction.ActionFullService.Enabled = False
    

End Sub

Private Sub ActionRestuctureFull_Click()

    Unload Me
    MoveData
    UserAction.ActionRestuctureFull.Enabled = False
    UserAction.ActionFullService.Enabled = False
    UserAction.ActionRestructureSingle.Enabled = False
    
End Sub

Private Sub UserForm_Click()

End Sub
