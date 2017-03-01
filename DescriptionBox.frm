VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DescriptionBox 
   Caption         =   "Please verify program description:"
   ClientHeight    =   3315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   OleObjectBlob   =   "DescriptionBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DescriptionBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    ProgramName = DescriptionBox.programNameBox.Value
    ProgramDescription = DescriptionBox.programDescriptionBox.Value
    ProgramSD = DescriptionBox.programSDBox.Value
    
    For i = 2 To X.Worksheets("Data").Cells(2, 10000).End(xlToLeft).Column
        If Trim(X.Worksheets("Data").Cells(2, i).Value) = DescriptionBox.currentProgramName.Value Then
            X.Worksheets("Data").Cells(2, i).Value = DescriptionBox.programNameBox.Value
        End If
    Next i
    
    Unload Me
    
End Sub

Private Sub CommandButton2_Click()

    Unload Me
    skip = True

End Sub

Private Sub CommandButton3_Click()
        
    Dim X As Workbook
        
    Set X = Workbooks.Open("C:\Users\jackie\Documents\Client Files\Progress Reports\FMP_DataExport\ProgramDescriptions.xlsx")
    
    X.Worksheets("PD").Rows(3).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    X.Worksheets("PD").Cells(3, 1).Value = DescriptionBox.programNameBox.Value
    X.Worksheets("PD").Cells(3, 2).Value = DescriptionBox.programDescriptionBox.Value
    X.Worksheets("PD").Cells(3, 3).Value = DescriptionBox.programSDBox.Value
    
    DescriptionBox.CommandButton3.Enabled = False
    
    Application.DisplayAlerts = False
    X.Save
    X.Close
    Application.DisplayAlerts = True
            
End Sub

Private Sub programSuggestion_Change()

    For i = 1 To X.Worksheets("PD").Cells(3000, 1).End(xlUp).row
        If DescriptionBox.programSuggestion.Value = Worksheets("PD").Cells(i, 1).Value Then
            DescriptionBox.programNameBox.Value = Worksheets("PD").Cells(i, 1).Value
            DescriptionBox.programDescriptionBox.Value = Worksheets("PD").Cells(i, 2).Value
            DescriptionBox.programSDBox.Value = Worksheets("PD").Cells(i, 3).Value
        End If
    Next i

End Sub

Private Sub UserForm_Click()

End Sub
