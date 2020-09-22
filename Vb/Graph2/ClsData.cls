VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private cYValue As String
Private cXValue As String
Private cPname As String
Private cPDate As String
Private cPError As String
Private cPErrorType As String
Private cPColor As String

Public Property Get YValue() As String
    YValue = cYValue
End Property

Public Property Let YValue(cYValueA As String)
    cYValue = cYValueA
End Property
Public Property Get XValue() As String
    XValue = cXValue
End Property

Public Property Let XValue(cXValueA As String)
    cXValue = cXValueA
End Property
Public Property Get PName() As String
    PName = cPname
End Property

Public Property Let PName(cPNameA As String)
    cPname = cPNameA
End Property
Public Property Get PDate() As String
   PDate = cPDate
End Property

Public Property Let PDate(cPDateA As String)
    cPDate = cPDateA
End Property
Public Property Get PError() As String
   PError = cPError
End Property

Public Property Let PError(cPErrorA As String)
    cPError = cPErrorA
End Property
Public Property Get PColor() As String
   PColor = cPColor
End Property

Public Property Let PColor(cPColorA As String)
    cPColor = cPColorA
End Property
Public Property Get PErrorType() As String
   PErrorType = cPErrorType
End Property

Public Property Let PErrorType(cPErrorTypeA As String)
    cPErrorType = cPErrorTypeA
End Property
