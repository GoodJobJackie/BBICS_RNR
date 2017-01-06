VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserAction 
   Caption         =   "Please select an action."
   ClientHeight    =   8475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5040
   OleObjectBlob   =   "UserAction.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub actionDatabase_Click()

    Dim strFile

    strFile = "C:\Users\jackie\Desktop\BBICS Employee Database.fmp12"

    Shell "cmd /c """ & strFile & """", 0

End Sub

Private Sub ActionDataEntry_Click()
      
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder("C:\Users\jackie\Documents\Client Files\Data\Formatted")
    
    ClientSelect.FileList = "Select File..."
    
    For Each objFile In objFolder.Files
        ClientSelect.FileList.AddItem CStr(objFile)
    Next objFile

    Unload Me
    
    ClientSelect.Show
    
End Sub

Private Sub actionDebug_Click()

    Application.DisplayAlerts = False
    For Each Worksheet In ActiveWorkbook.Worksheets
        Select Case Worksheet.Name
            Case Is = "PD"
                Worksheets("PD").Delete
            Case Is = "CI"
                Worksheets("CI").Delete
            Case Is = "SDL"
                Worksheets("SDL").Delete
            Case Is = "Current"
                Worksheets("Current").Delete
            Case Is = "Programs"
                Worksheets("Programs").Delete
        End Select
    Next
    Application.DisplayAlerts = True

End Sub

Private Sub actionDocuments_Click()

    Dim strFile

    strFile = "C:\Users\jackie\Desktop\Admin Documents.jar"

    Shell "cmd /c """ & strFile & """", 0

End Sub

Private Sub actionFixNames_Click()
    
    Unload Me
    ProgramDescriptionsList
    
End Sub

Private Sub ActionFullService_Click()

    Unload Me
    NewRestructuring
    UserAction.ActionFullService.Enabled = False
    UserAction.ActionRestructureSingle.Enabled = False

End Sub

Private Sub ActionImportSP_Click()
    
    Unload Me
    ImportSkillsPrograms
    
End Sub

Private Sub actionIPG_Click()

    Unload Me
 
    ImportSkillsPrograms
    PopulatePrograms
    CreateProgramLists
    PopulateReport

End Sub

Private Sub ActionPopulate_Click()
    
    Unload Me
    PopulatePrograms
    CreateProgramLists
    
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
    
    For i = 2 To Cells(2, 2000).End(xlToLeft).Column
        If Cells(2, i).Value = "Worksheets" Then
            With Cells(2, i)
                .Value = ""
                .Interior.Color = -4142
            End With
            With Cells(1, i)
                .Value = "Worksheets"
                .Interior.Color = RGB(255, 255, 0)
                .Font.Bold = True
            End With
        End If
    Next i
    
    Cells(4, 2).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    'UserAction.ActionFullService.Enabled = False
    Cells(2000, 1).End(xlUp).Offset(1, 0).Value = Date

End Sub

Private Sub ActionReport_Click()

    Unload Me
    PopulateReport
    
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
    'UserAction.ActionFullService.Enabled = False
    UserAction.ActionRestructureSingle.Enabled = False
    
End Sub

Private Sub CommandButton1_Click()
       
End Sub

Private Sub CommandButton2_Click()

    Unload Me

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub VerifyProgramNames_Click()

    Unload Me
    ImportSkillsPrograms
    RenamePrograms
    Application.DisplayAlerts = False
    For Each Worksheet In ActiveWorkbook.Worksheets
        Select Case Worksheet.Name
            Case Is = "PD"
                Worksheets("PD").Delete
            Case Is = "CI"
                Worksheets("CI").Delete
            Case Is = "SDL"
                Worksheets("SDL").Delete
            Case Is = "Current"
                Worksheets("Current").Delete
            Case Is = "Programs"
                Worksheets("Programs").Delete
        End Select
    Next
    Application.DisplayAlerts = True
    
End Sub
