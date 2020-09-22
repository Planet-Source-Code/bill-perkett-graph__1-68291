VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private cInfoName As String
Private cYValue As String
Private cInfoOther As String
Public Property Get InfoName() As String
    InfoName = cInfoName
End Property

Public Property Let InfoName(cInfoNameA As String)
    cInfoName = cInfoNameA
End Property

Public Property Get YValue() As String
    YValue = cYValue
End Property

Public Property Let YValue(cYValueA As String)
    cYValue = cYValueA
End Property
Public Property Get InfoOther() As String
   InfoOther = cInfoOther
End Property

Public Property Let InfoOther(cInfoOtherA As String)
    cInfoOther = cInfoOtherA
End Property

