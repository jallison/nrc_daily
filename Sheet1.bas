VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Changes As Range)

    Dim Change As Range
    
    For Each Change In Changes

        '* Populates Job Category and Description based on the code populated in column 1
        If Change.Column = "1" And _
           Len(Change.Text) > 0 And _
           Change.Row >= InvStartRow And _
           Change.Row <= InvStopRow Then
    
            Dim Job As CJob '* Object to store job category and description
            Set Job = New CJob
        
            Job.JobCode = Change.Value
            
            Call job_lookup(Job)
            
            '* Populate job category and description
            '* Not allowed to use IN for Carl
            Cells(Change.Row, 2) = Job.JobCat
            Cells(Change.Row, 3) = Job.JobDesc
            
            '* Capitalize code that was entered
            Application.EnableEvents = False
                Cells(Change.Row, 1) = UCase(Cells(Change.Row, 1))
            Application.EnableEvents = True
            
            Set Job = Nothing
            
        ElseIf Change.Column = "1" And _
               Change.Text = "" And _
               Change.Row >= InvStartRow And _
               Change.Row <= InvStopRow Then
            
            '* Deletes category and description if job code is removed
            Cells(Change.Row, 2) = ""
            Cells(Change.Row, 3) = ""
            Cells(Change.Row, 6) = ""
            
        End If
        
    Next Change
    
End Sub


