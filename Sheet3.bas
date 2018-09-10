VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub CommandButton1_Click()

    Dim sDate As String
    
    '* Create a copy of the sheet at the end of excel but don't copy the button
    Application.CopyObjectsWithCells = False
    ActiveSheet.Copy after:=Sheets(Sheets.Count)
    Application.CopyObjectsWithCells = True
    '* rename the copied sheet
    sDate = Format(Now(), "mm-dd-yy hhmm")
    ActiveSheet.Name = "Daily " + sDate
    
    '* Clear the Daily sheet.
    Cells(1, 14) = ""
    Cells(2, 6) = ""
    Cells(2, 10) = ""
    Cells(2, 12) = ""
    Cells(2, 16) = ""
    Cells(2, 18) = ""
    Cells(2, 22) = ""
    Sheets("Daily").Range("B5:Y6").ClearContents
    Sheets("Daily").Range("A7:Y67").ClearContents
    Sheets("Daily").Range("A69:Y69").ClearContents
    

End Sub


