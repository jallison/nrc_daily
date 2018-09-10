VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click()

    Dim i As Integer
    Dim x As Integer
    
    Sheet1.Cells(7, 3) = Cells(1, 2).Value
    Sheet1.Cells(9, 3) = Cells(2, 2).Value
    Sheet1.Cells(5, 3) = Cells(1, 5).Value
    Sheet1.Cells(5, 6) = Cells(1, 14).Value
    
    x = 14
    
    For i = 2 To 32
    
        If Cells(98, i) > 0 Then
        
            Sheet1.Cells(x, 1) = Cells(99, i)
            Sheet1.Cells(x, 6) = Cells(98, i)
            x = x + 1
        
        End If
    
    Next i
    
    '* sort the invoice sheet
    Sheet1.Range(Sheet1.Cells(14, 1), _
        Sheet1.Cells(69, 6)).Sort _
        Key1:=Sheet1.Range("A1"), _
        Order1:=xlAscending, Header:=xlNo, _
        Key2:=Sheet1.Range("B1"), _
        Order1:=xlAscending, _
        Header:=xlNo

End Sub

