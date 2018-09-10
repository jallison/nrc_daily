VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pJobCode As String
Private pJobCat As String
Private pJobDesc As String
Private pShortJobDesc As String

'************************************************
Public Property Get JobCode() As String
    JobCode = pJobCode
End Property
Public Property Let JobCode(Value As String)
    pJobCode = Value
End Property
'************************************************
Public Property Get JobCat() As String
    JobCat = pJobCat
End Property
Public Property Let JobCat(Value As String)
    pJobCat = Value
End Property
'************************************************
Public Property Get JobDesc() As String
    JobDesc = pJobDesc
End Property
Public Property Let JobDesc(Value As String)
    pJobDesc = Value
End Property
'************************************************
Public Property Get ShortJobDesc() As String
    ShortJobDesc = pShortJobDesc
End Property
Public Property Let ShortJobDesc(Value As String)
    pShortJobDesc = Value
End Property
'************************************************

