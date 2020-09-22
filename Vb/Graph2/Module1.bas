Attribute VB_Name = "Module1"
'**********
'
' Global graph settings
' (used so graph can be generated in any form)
'
'**********
Public Ydel2 As Single
Public Ymin As Single
Public Yold As String
Public Jold As Integer

Public nData As Collection
Public MYdata As New ClsData
Public nInfo As Collection
Public MYinfo As New ClsInfo
'
' For Summary
'
Public dAve As Double
Public dSD As Double
Public sCount As Single
Public Ylow As Single
Public Yhi As Single
Public dSum As Double
Public dSum2 As Double
Public dRange As Double
Public iWarn As Integer
Public iErr As Integer
