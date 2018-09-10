VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'* Sub used to clear cell in Ticket sheet
Private Sub CommandButton1_Click()

    Dim i As Integer '* generic counter
    
    Cells(2, 3) = ""
    
    For i = TicketStartRow To TicketStopRow
    
        Cells(i, 1) = ""
        
    Next i

End Sub

'* Sub used to move production to the Daily and the Weekly sheet
Private Sub CommandButton2_Click()

    If Cells(2, 3) = "" Then
    
        MsgBox "Ticket# must be populated."
        
    Else
    
        ProdMove Sheet3
        ProdMove Sheet4
        
    End If

End Sub

'* Sub used to move production to the Daily sheet
Private Sub CommandButton3_Click()

    If Cells(2, 3) = "" Then
    
        MsgBox "Ticket# must be populated."
        
    Else
    
        ProdMove Sheet3
        
    End If

End Sub

'* Sub used to move production to the Weekly sheet
Private Sub CommandButton4_Click()

    If Cells(2, 3) = "" Then
    
        MsgBox "Ticket# must be populated."
        
    Else
    
        ProdMove Sheet4
        
    End If

End Sub

'* This fills in the desc of the task inputed by user
Private Sub Worksheet_Change(ByVal Changes As Range)

    Dim Change As Range
    
    For Each Change In Changes

        '* Populate job category and description based on job code entered in column 1
        If Change.Column = "1" And _
           Len(Change.Text) > 0 And _
           Change.Row >= TicketStartRow And _
           Change.Row <= TicketStopRow Then
           
            Application.EnableEvents = False
                Change.Value = UCase(Change.Value)
            Application.EnableEvents = True
    
            Dim Job As CJob '* Object to store job code and description
            Set Job = New CJob
        
            Job.JobCode = Change.Value
            
            Call job_lookup(Job)
            
            '* Populate job category and description
            ActiveSheet.Unprotect
            Cells(Change.Row, 2) = Job.JobCat
            Cells(Change.Row, 3) = Job.JobDesc
            ActiveSheet.Protect
            
            Set Job = Nothing
            
        '* If job code is deleted then remove the category, description, quantity for that line
        ElseIf Change.Column = "1" And _
               Change.Text = "" And _
               Change.Row >= TicketStartRow And _
               Change.Row <= TicketStopRow Then
        
            
            ActiveSheet.Unprotect
            Cells(Change.Row, 2) = ""
            Cells(Change.Row, 3) = ""
            Cells(Change.Row, 4) = ""
            ActiveSheet.Protect
            
        End If
        
    Next Change
    
End Sub


