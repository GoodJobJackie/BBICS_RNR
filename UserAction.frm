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

Private Sub actionNewClient_Click()

    Dim client As String
    Dim fileName As String
    
    client = InputBox("Please enter new client initials:", "New Client")
    fileName = "C:\Users\jackie\Documents\Client Files\Data\Formatted\" & UCase(client) & " - 0000_00_00.xlsx"
    
    Workbooks.Add
    ActiveSheet.Name = "Tutor Hr Data"
    Worksheets.Add().Name = "Bx Data"
    Worksheets.Add().Name = "Data"
    
    Worksheets("Data").Activate
    CreateHeader
    MasterListFormat
    ActiveWindow.Zoom = 90
    Worksheets("Data").Cells(1, 1) = UCase(client)
    Worksheets("Data").Cells(4, 1) = DateValue("01/01/2016")
    Worksheets("Data").Cells(5, 1) = Format(Now, "mm/dd/yyyy")
    Cells(4, 2).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    
    Worksheets("Bx Data").Activate
    Cells.Select
    Range("BQ21").Activate
    Selection.ColumnWidth = 11
    Range("A1:A2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Worksheets("Bx Data").Cells(1, 1) = UCase(client)
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    MasterListFormat
    Cells(3, 2).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    
    Worksheets("Tutor Hr Data").Activate
    Cells.Select
    Range("BQ21").Activate
    Selection.ColumnWidth = 11
    Range("A1:A2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Worksheets("Bx Data").Cells(1, 1) = UCase(client)
    Selection.Font.Size = 18
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    Columns("A").Select
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.NumberFormat = "MMM yyyy"
    Worksheets("Tutor Hr Data").Cells(3, 1) = "Jan 2017"
    Worksheets("Tutor Hr Data").Cells(4, 1) = "Feb 2017"
    Worksheets("Tutor Hr Data").Cells(5, 1) = "Mar 2017"
    Worksheets("Tutor Hr Data").Cells(6, 1) = "Apr 2017"
    Worksheets("Tutor Hr Data").Cells(7, 1) = "May 2017"
    Worksheets("Tutor Hr Data").Cells(8, 1) = "Jun 2017"
    Worksheets("Tutor Hr Data").Cells(9, 1) = "Jul 2017"
    Worksheets("Tutor Hr Data").Cells(10, 1) = "Aug 2017"
    Worksheets("Tutor Hr Data").Cells(11, 1) = "Sep 2017"
    Worksheets("Tutor Hr Data").Cells(12, 1) = "Oct 2017"
    Worksheets("Tutor Hr Data").Cells(13, 1) = "Nov 2017"
    Worksheets("Tutor Hr Data").Cells(14, 1) = "Dec 2017"
    Cells(3, 1).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    Worksheets("Tutor Hr Data").Cells(1, 1) = UCase(client)
    
    Worksheets("Data").Activate
        
    ActiveWorkbook.SaveAs (fileName)
    ActiveWorkbook.Close

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
    Cells(2000, 1).End(xlUp).Offset(1, 0).Value = Date

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
    UserAction.ActionRestructureSingle.Enabled = False
    
End Sub

Private Sub CommandButton1_Click()

      On Error GoTo ErrorHandling
      
      Call err.Raise(1342, "UserAction button", "User submitted message")
      
ErrorHandling:
    ErrHandling
          
End Sub

Private Sub CommandButton2_Click()

    Unload Me

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
