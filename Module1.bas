Attribute VB_Name = "Module1"
Option Explicit

Public Const TicketStartRow = 6
Public Const TicketStopRow = 30
Public Const InvStartRow = 14
Public Const InvStopRow = 69


'* Imports ticket into the Daily/Weekly worksheet
Public Sub ProdMove(ByVal SheetName As Worksheet)

    Dim i As Integer '* Generic loop counter
    Dim x As Integer '* Generic loop counter
    Dim y As Integer '* Generic loop counter
    Dim bBlankFound As Boolean '* Used to exit loop when blank row if found
    Dim strFind As String '* Value being searched for
    Dim rSearch As Range '* used to store found matches
    Dim Job As CJob '* Object to store job code and description
    
    Dim iRowStart As Integer
    Dim iRowStop As Integer
    Dim iColStart As Integer
    Dim iColStop As Integer
    
    '* Initialize bBlankFound
    bBlankFound = False
    
    With SheetName
    
        '* Adjust column and row stop/starts based on sheetname
        If .Name = "Daily" Then
        
            iRowStart = 7
            iRowStop = 67
            iColStart = 2
            iColStop = 25
        
        Else
        
            iRowStart = 7
            iRowStop = 97
            iColStart = 2
            iColStop = 32
        
        End If '*.Name = "Daily" Then
    
        '* Loop thru each row until a blank row
        For x = iRowStart To iRowStop
            
            '* See if the row is blank
            If .Cells(x, 1) = "" Then
            
                'Blank row was found
                bBlankFound = True
                        
                '* Bring ticket number in
                .Cells(x, 1) = Sheet2.Cells(2, 3)
    
                '* Move Ticket prdouction to the daiy sheet
                For i = TicketStartRow To TicketStopRow
                
                    '* Retrieve the value we are searching for
                    Set Job = New CJob
                    Job.JobCode = Sheet2.Cells(i, 1).Value
            
                    Call job_lookup(Job)
                    
                    '* Try to find the value of strFind in the row
                    Set rSearch = .Rows(5).Find(What:=Job.JobCode, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
                    
                    '* Check to see if someting was found
                    If Not (rSearch Is Nothing) Then
                    
                        '* Bring the quanity over
                        .Cells(x, rSearch.Column) = Sheet2.Cells(i, 4)
                        
                    Else
                        
                        '* Fill in column information
                        For y = iColStart To iColStop
                        
                            If .Cells(5, y) = "" Then
                            
                                .Cells(5, y) = Job.JobCode
                                .Cells(6, y) = Job.ShortJobDesc
                                .Cells(iRowStop + 2, y) = Job.JobCode
                                .Cells(x, y) = Sheet2.Cells(i, 4)
                                
                                y = iColStop
                            
                            End If '.Cells(5, y) = "" Then
                            
                        Next y ' = iColStart To iColStop
                    
                    End If 'Not (rSearch Is Nothing) Then
    
                Next i '= TicketStartRow To TicketStopRow
                            
            
            End If '.Cells(x, 1) = "" Then
            
            '* When blank row is found exit the loop
            If bBlankFound = True Then
            
                x = iRowStop
            
            End If 'bBlankFound = True Then
        
        Next x '= iRowStart To iRowStop
    
    End With 'SheetName
    
End Sub

'* Function looks up the category and description of a job code
Public Function job_lookup(ByVal Job As CJob)
     
    Dim rSearch As Range
    Set rSearch = Nothing
        
    Set rSearch = Sheet5.Columns(1).Find(Job.JobCode, LookIn:=xlValues, LookAt:=xlWhole, after:=Range("A65536"))
         
    If Not (rSearch Is Nothing) Then
    
        Job.JobCat = Sheet5.Cells(rSearch.Row, 2)
        Job.JobDesc = Sheet5.Cells(rSearch.Row, 3)
        Job.ShortJobDesc = Sheet1.Cells(rSearch.Row, 4)
            
    Else
        
        MsgBox "Code " & Job.JobCode & " is not defined on the Code tab."
            
    End If
    
End Function


