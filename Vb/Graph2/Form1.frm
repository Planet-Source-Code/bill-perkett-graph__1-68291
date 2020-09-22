VERSION 5.00
Begin VB.Form FrmGraph 
   Caption         =   "FrmGraph (by Bill Perkett)"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   FillStyle       =   0  'Solid
   ForeColor       =   &H00404040&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Blocks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   44
      Top             =   6480
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Points"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   43
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox TxtInfo 
      Height          =   285
      Index           =   8
      Left            =   9600
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   5400
      Width           =   585
   End
   Begin VB.CommandButton CmdCopy 
      Caption         =   "CopyData"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   40
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox TxtInfo 
      Height          =   285
      Index           =   7
      Left            =   9600
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   5040
      Width           =   585
   End
   Begin VB.TextBox TxtInfo 
      Height          =   285
      Index           =   6
      Left            =   9600
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   4680
      Width           =   585
   End
   Begin VB.TextBox TxtInfo 
      Height          =   285
      Index           =   5
      Left            =   4440
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   0
      Width           =   1785
   End
   Begin VB.TextBox TxtInfo 
      Height          =   285
      Index           =   4
      Left            =   9600
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   4320
      Width           =   585
   End
   Begin VB.TextBox TxtInfo 
      Height          =   285
      Index           =   3
      Left            =   9600
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   3960
      Width           =   585
   End
   Begin VB.TextBox TxtInfo 
      Height          =   285
      Index           =   2
      Left            =   9600
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   3600
      Width           =   585
   End
   Begin VB.TextBox TxtInfo 
      Height          =   285
      Index           =   1
      Left            =   9600
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   3240
      Width           =   585
   End
   Begin VB.TextBox TxtInfo 
      Height          =   285
      Index           =   0
      Left            =   9600
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   2880
      Width           =   585
   End
   Begin VB.TextBox Txtdata 
      Height          =   285
      Index           =   7
      Left            =   9240
      TabIndex        =   23
      Text            =   "Txtdata"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Txtdata 
      Height          =   285
      Index           =   6
      Left            =   9240
      TabIndex        =   22
      Text            =   "Txtdata"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Txtdata 
      Height          =   285
      Index           =   5
      Left            =   9240
      TabIndex        =   21
      Text            =   "Txtdata"
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Txtdata 
      Height          =   285
      Index           =   4
      Left            =   9240
      TabIndex        =   20
      Text            =   "Txtdata"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Txtdata 
      Height          =   285
      Index           =   3
      Left            =   9240
      TabIndex        =   19
      Text            =   "Txtdata"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Txtdata 
      Height          =   285
      Index           =   2
      Left            =   9240
      TabIndex        =   18
      Text            =   "Txtdata"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Txtdata 
      Height          =   285
      Index           =   1
      Left            =   9240
      TabIndex        =   17
      Text            =   "Txtdata"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Txtdata 
      Height          =   285
      Index           =   0
      Left            =   9240
      TabIndex        =   16
      Text            =   "Txtdata"
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox TxtPoint 
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Text            =   "TxtPoint"
      Top             =   6480
      Width           =   615
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&Next point >>"
      Height          =   300
      Index           =   2
      Left            =   7680
      TabIndex        =   7
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "<< &Previous point"
      Height          =   300
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   6480
      Width           =   1455
   End
   Begin VB.TextBox TxtY 
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Text            =   "TxtY"
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   7680
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6360
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton CmdGraph 
      Caption         =   "Create Graph"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   8760
      TabIndex        =   0
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label LblInfo 
      Caption         =   "Inc/ Above"
      Height          =   435
      Index           =   8
      Left            =   9120
      TabIndex        =   42
      Top             =   5280
      Width           =   645
      WordWrap        =   -1  'True
   End
   Begin VB.Label LblInfo 
      AutoSize        =   -1  'True
      Caption         =   "MAX"
      Height          =   195
      Index           =   7
      Left            =   9240
      TabIndex        =   39
      Top             =   5040
      Width           =   345
   End
   Begin VB.Label LblInfo 
      AutoSize        =   -1  'True
      Caption         =   "MIN"
      Height          =   195
      Index           =   6
      Left            =   9240
      TabIndex        =   38
      Top             =   4680
      Width           =   300
   End
   Begin VB.Label LblInfo 
      AutoSize        =   -1  'True
      Caption         =   "TITLE"
      Height          =   195
      Index           =   5
      Left            =   3720
      TabIndex        =   37
      Top             =   0
      Width           =   450
   End
   Begin VB.Label LblInfo 
      AutoSize        =   -1  'True
      Caption         =   "TAR"
      Height          =   195
      Index           =   4
      Left            =   9240
      TabIndex        =   36
      Top             =   4320
      Width           =   330
   End
   Begin VB.Label LblInfo 
      AutoSize        =   -1  'True
      Caption         =   "LCL"
      Height          =   195
      Index           =   3
      Left            =   9240
      TabIndex        =   35
      Top             =   3960
      Width           =   285
   End
   Begin VB.Label LblInfo 
      AutoSize        =   -1  'True
      Caption         =   "LWL"
      Height          =   195
      Index           =   2
      Left            =   9240
      TabIndex        =   34
      Top             =   3600
      Width           =   345
   End
   Begin VB.Label LblInfo 
      AutoSize        =   -1  'True
      Caption         =   "UWL"
      Height          =   195
      Index           =   1
      Left            =   9240
      TabIndex        =   33
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label LblInfo 
      AutoSize        =   -1  'True
      Caption         =   "UCL"
      Height          =   195
      Index           =   0
      Left            =   9240
      TabIndex        =   32
      Top             =   2880
      Width           =   315
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8640
      TabIndex        =   15
      Top             =   0
      Width           =   420
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9360
      TabIndex        =   14
      Top             =   5760
      Width           =   765
   End
   Begin VB.Label LblGreen 
      BackColor       =   &H00C0FFC0&
      Caption         =   "LblGreen"
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   8160
      TabIndex        =   13
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label LblYellow 
      BackColor       =   &H00C0FFFF&
      Caption         =   "LblYellow"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   8400
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label LblRed 
      BackColor       =   &H00C0C0FF&
      Caption         =   "LblRed"
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   8400
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      Height          =   255
      Left            =   7440
      TabIndex        =   10
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Point"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   8
      Top             =   6480
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Yvalue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2760
      TabIndex        =   5
      Top             =   6480
      Width           =   600
   End
End
Attribute VB_Name = "FrmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOK_Click(Index As Integer)
   '
   ' Used to go from point to point
   '
    Dim iX As Integer
    Dim iX2 As Integer
    Dim iAdd As Integer
    Dim j As Integer
    Dim y2 As Single
    Dim Ycal As Single
    Dim bFound As Boolean
    '
    ' Find Previous or Next point
    '
    bFound = False
    iAdd = 1
    If Index = 1 Then iAdd = -1
    iX = Val(TxtPoint.Text)
    For Each MYdata In nData
      iX2 = Val(MYdata.XValue / 8)
      If iX2 = iX + iAdd Then
         TxtY.Text = MYdata.YValue
         TxtPoint.Text = MYdata.XValue / 8
'         TxtName.Text = MYdata.PName
'         TxtError.Text = MYdata.PError
'         TxtError.BackColor = MYdata.PColor
         j = MYdata.XValue
         y2 = MYdata.YValue
          bFound = True
       End If
    Next
    If bFound = False Then Exit Sub
    '
    ' Redraw point as white with a black border
    '
    FillStyle = 0
    Ycal = (y2 - Ymin) * Ydel2 + 5
    FillColor = QBColor(15)
    If Option1(0) Then
           FillColor = QBColor(15)
            Circle (j, Ycal), 1, QBColor(0)
      Else
            'Line (j, 0)-(j, Ycal)
           ' Line (j, 0)-(j + 3, Ycal)
           Line (j, 0)-(j + 5, Ycal), QBColor(0), BF
           
      End If
    'Circle (j, Ycal), 1, QBColor(0)
    '
    ' Redraw old point as black
    '
    FillColor = QBColor(0)
    If Jold > 0 Then
           If Jold <> j Then 'Circle (Jold, Yold), 1, QBColor(0)
             If Option1(0) Then
                Circle (Jold, Yold), 1, QBColor(0)
              Else
                'Line (j, 0)-(j, Ycal)
                  ' Line (j, 0)-(j + 3, Ycal)
                Line (Jold, 0)-(Jold + 5, Yold), RGB(0, 0, 255), BF
           End If
      End If
    End If
    Yold = Ycal
    Jold = j
 End Sub

Private Sub CmdCopy_Click()
    '  '
    '  ' Copy data to clipboard
    '  ' can paste information into Excel
    '  '
    '   Dim txt As String
    '   Dim i As Integer
    '   Dim j As Integer
    '   Clipboard.Clear
    '   '
    '   ' Title
    '   '
    '   txt = TxtInfo(5).Text & vbCrLf
    '   txt = txt & "X" & vbTab & "Y" & vbTab & "Date" & vbTab & "DataName " & vbTab & "Error " & vbTab & "ErrorType " & vbTab
    '   txt = txt & vbCrLf
    '   i = 1
    '   For Each MYdata In nData
    '      txt = txt & i & vbTab & MYdata.YValue & vbTab & MYdata.PDate & vbTab & MYdata.PName & vbTab & MYdata.PError & vbTab & MYdata.PErrorType
    ''      j = 1
    ''      For Each MYinfo In nInfo
    ''        If i = j Then txt = txt & vbTab & " " & vbTab & MYinfo.InfoName & vbTab & MYinfo.YValue
    ''        j = j + 1
    ''      Next
    '        txt = txt & vbCrLf
    '      i = i + 1
    '    Next
    '    '
    '    '
    '    '
    '    txt = txt & vbCrLf
    '    txt = txt & "InfoName" & vbTab & "InfoValue" & vbCrLf
    '    For Each MYinfo In nInfo
    '        txt = txt & MYinfo.InfoName & vbTab & MYinfo.YValue & vbCrLf
    '    Next
    '    '
    '    ' Add Summary data
    '    '
    '    txt = txt & vbCrLf
    '    txt = txt & "Points   " & vbTab & sCount & vbCrLf
    '    txt = txt & "Hi       " & vbTab & Yhi & vbCrLf
    '    txt = txt & "Low      " & vbTab & Ylow & vbCrLf
    '    txt = txt & "Ave      " & vbTab & dAve & vbCrLf
    '    txt = txt & "Range    " & vbTab & dRange & vbCrLf
    '    txt = txt & "SD       " & vbTab & dSD & vbCrLf
    '    txt = txt & "Warnings " & vbTab & iWarn & vbCrLf
    '    txt = txt & "Errors   " & vbTab & iErr & vbCrLf
    '    Clipboard.SetText txt, vbCFText
    Dim exlApp As New Excel.Application
    Dim exlSheet As Excel.Worksheet
    Dim j As Integer
    Dim i As Integer
    '
    ' Open Excel
    '
    With exlApp
      .DisplayAlerts = False
      .Workbooks.Add
      '
      ' Add Title
      '
      Set exlSheet = .Worksheets(1)
      exlSheet.Cells(1, 1).Value = TxtInfo(5).Text
      exlSheet.Cells(2, 1).Value = "X"
      exlSheet.Cells(2, 2).Value = "Y"
      exlSheet.Cells(2, 3).Value = "Date"
      exlSheet.Cells(2, 4).Value = "DataName "
      exlSheet.Cells(2, 5).Value = "Error "
      exlSheet.Cells(2, 6).Value = "ErrorType "
      '
      ' Add Graph Data
      '
      i = 1
      j = 3
      For Each MYdata In nData
        exlSheet.Cells(j, 1).Value = i
        exlSheet.Cells(j, 2).Value = MYdata.YValue
        exlSheet.Cells(j, 3).Value = MYdata.PDate
        exlSheet.Cells(j, 4).Value = MYdata.PName
        exlSheet.Cells(j, 5).Value = MYdata.PError
        exlSheet.Cells(j, 6).Value = MYdata.PErrorType
        i = i + 1
        j = j + 1
      Next
      '
      ' Add Limit Information
      '
      exlSheet.Cells(1, 8).Value = "Graph Limits"
      exlSheet.Cells(2, 7).Value = " "
      exlSheet.Cells(2, 8).Value = "InfoName"
      exlSheet.Cells(2, 9).Value = "InfoValue"
      j = 3
      For Each MYinfo In nInfo
         exlSheet.Cells(j, 8).Value = MYinfo.InfoName
         exlSheet.Cells(j, 9).Value = MYinfo.YValue
          j = j + 1
      Next
      '
      ' Add Summary Information
      '
      exlSheet.Cells(1, 11).Value = "Summary"
      exlSheet.Cells(2, 10).Value = " "
      exlSheet.Cells(2, 11).Value = "Name "
      exlSheet.Cells(2, 12).Value = "Value"
      exlSheet.Cells(3, 11).Value = "Points "
      exlSheet.Cells(3, 12).Value = sCount
      exlSheet.Cells(4, 11).Value = "Hi"
      exlSheet.Cells(4, 12).Value = Yhi
      exlSheet.Cells(5, 11).Value = "Low "
      exlSheet.Cells(5, 12).Value = Ylow
      exlSheet.Cells(6, 11).Value = "Average"
      exlSheet.Cells(6, 12).Value = dAve
      exlSheet.Cells(7, 11).Value = "Range"
      exlSheet.Cells(7, 12).Value = dRange
      exlSheet.Cells(8, 11).Value = "SD"
      exlSheet.Cells(8, 12).Value = dSD
      exlSheet.Cells(9, 11).Value = "Warnings"
      exlSheet.Cells(9, 12).Value = iWarn
      exlSheet.Cells(10, 11).Value = "Errors"
      exlSheet.Cells(10, 12).Value = iErr
      '
      ' Format Information
      '
      exlSheet.Activate
      exlSheet.Cells.Select
      exlSheet.Cells.EntireColumn.AutoFit
      exlSheet.Rows("1:1").Select
      Selection.Font.Bold = True
      exlSheet.Rows("1:1").Select
      Selection.Font.Bold = True
      Rows("2:2").Select
      Selection.Font.Bold = True
      Range("A1").Select
   
     ' Save the sheet
     'exlSheet.SaveAs "C:\Data\ TEST.xls"

 End With
 exlApp.Visible = True
 'exlApp.Quit
 
End Sub

Private Sub CmdGraph_Click()
    '
    ' Use Form.Paint to Draw the graph
    '
    Me.Refresh
End Sub



Private Sub CmdPrint_Click()
 'Me.p
End Sub

Private Sub Form_Load()
    '
    ' Dummy data used to Create Graph
    '
    Dim Y(50) As Single 'Data
    Dim i As Integer
    Y(1) = 3
    Y(2) = 180
    Y(3) = 15
    Y(4) = 25
    Y(5) = 26
    Y(6) = 27
    Y(7) = 0
    Y(8) = 28
    For i = 0 To 7
     Txtdata(i).Text = Y(i + 1)
    Next
    Option1(0).Value = True
    '
    ' To draw Limits
    ' (Always set MYinfo.YValue to a value)
    '
    Set nInfo = New Collection
    
    Set MYinfo = New ClsInfo
    MYinfo.InfoName = "TAR"
    MYinfo.YValue = ""
    TxtInfo(4).Text = MYinfo.YValue
    nInfo.Add MYinfo
    
    Set MYinfo = New ClsInfo
    MYinfo.InfoName = "UCL"
    MYinfo.YValue = ""
    TxtInfo(0).Text = MYinfo.YValue
    nInfo.Add MYinfo
    Set MYinfo = New ClsInfo
    MYinfo.InfoName = "LCL"
    MYinfo.YValue = ""
     TxtInfo(3).Text = MYinfo.YValue
    nInfo.Add MYinfo
    Set MYinfo = New ClsInfo
    MYinfo.InfoName = "UWL"
    MYinfo.YValue = ""
    TxtInfo(1).Text = MYinfo.YValue
    nInfo.Add MYinfo
    Set MYinfo = New ClsInfo
    MYinfo.InfoName = "LWL"
    MYinfo.YValue = ""
    TxtInfo(2).Text = MYinfo.YValue
    nInfo.Add MYinfo
    Set MYinfo = New ClsInfo
     MYinfo.InfoName = "MIN"
    MYinfo.YValue = 0
    TxtInfo(6).Text = MYinfo.YValue
    nInfo.Add MYinfo
    Set MYinfo = New ClsInfo
    MYinfo.InfoName = "MAX"
    MYinfo.YValue = ""
    TxtInfo(7).Text = MYinfo.YValue
    nInfo.Add MYinfo
    Set MYinfo = New ClsInfo
    MYinfo.InfoName = "Title"
    MYinfo.YValue = "This is the title"
    TxtInfo(5).Text = MYinfo.YValue
    nInfo.Add MYinfo
    TxtInfo(8).Text = "0"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 '
 ' if cursor inside graph
 '
 Dim j As Integer
 Dim x2 As Single
 Dim y2 As Single
 Dim delta As Single
 Dim delta2 As Single
 Dim Ycal As Single
 delta = 999
 Screen.MousePointer = vbArrow
 If X >= 0 And X <= 123 Then
   If Y >= 5 And Y <= 95 Then
    Screen.MousePointer = vbCrosshair
     Text2.Text = X
     Text3.Text = Y
      dCurrentX = X
      dCurrentY = Y
      '
      ' Determine closest point
      '
     For Each MYdata In nData
       delta2 = Abs(dCurrentX - MYdata.XValue)
       If delta2 < delta Then
           TxtY.Text = MYdata.YValue
           y2 = MYdata.YValue
           delta = Abs(dCurrentX - MYdata.XValue)
           j = MYdata.XValue
           TxtPoint.Text = j / 8
'           TxtName.Text = MYdata.PName
'           TxtError.Text = MYdata.PError
'           TxtError.BackColor = MYdata.PColor
       End If
     Next
        '
        ' Redraw point as white with a black border
        '
        Ycal = (y2 - Ymin) * Ydel2 + 5
        If Option1(0) Then
           FillColor = QBColor(15)
            Circle (j, Ycal), 1, QBColor(0)
      Else
            'Line (j, 0)-(j, Ycal)
           ' Line (j, 0)-(j + 3, Ycal)
           Line (j, 0)-(j + 5, Ycal), QBColor(0), BF
           
      End If
    'Circle (j, Ycal), 1, QBColor(0)
    '
    ' Redraw old point as black
    '
    FillColor = QBColor(0)
    If Jold > 0 Then
           If Jold <> j Then 'Circle (Jold, Yold), 1, QBColor(0)
             If Option1(0) Then
                Circle (Jold, Yold), 1, QBColor(0)
              Else
                'Line (j, 0)-(j, Ycal)
                  ' Line (j, 0)-(j + 3, Ycal)
                Line (Jold, 0)-(Jold + 5, Yold), RGB(0, 0, 255), BF
           End If
      End If
    End If
        Yold = Ycal
        Jold = j
  End If
  End If
End Sub

Private Sub Form_Paint()
    '
    ' Draw the graph
    '
    Dim i As Integer
    Dim j As Integer
    Dim Ymax As Single
    Dim YOldPnt As Single
    Dim Ydel As Single
    Dim Ydel1 As Single
    Dim Ycal As Single
    Dim Ycal2 As Single
    Dim Ycal3 As Single
    Dim bLimits As Boolean
    Dim cLcl As String
    Dim cUcl As String
    Dim cLwl As String
    Dim cUwl As String
    Dim bYMax As Boolean
    Dim bYMin As Boolean
    Dim iTick As Integer
    '
    ' Check for Inc/above points
    '
    Dim iInc As Integer
    Dim iAbove As Integer
    Dim iIncCnt As Integer
    Dim iAboveCnt As Integer
    Dim iDecCnt As Integer
    Dim iBelowCnt As Integer
    '
    ' ClsData - a class used to store all your data
    '
    Set nData = New Collection
    For i = 1 To 8
      Set MYdata = New ClsData
      MYdata.YValue = Txtdata(i - 1).Text
      MYdata.XValue = i * 8
      MYdata.PDate = Now
      MYdata.PName = "Test" & i
      MYdata.PError = ""
      MYdata.PErrorType = ""
      MYdata.PColor = vbWhite
      nData.Add MYdata
    Next
    '
    ' Old point information
    '
    Yold = ""
    Jold = -1
    '
    ' ClsInfo - a class used to draw Limits
    '
    ' InfoName = TAR    (the chart target)
    ' InfoName = UCL    (the chart UCL)
    ' InfoName = LCL    (the chart LCL)
    ' InfoName = UWL    (the chart UWL)
    ' InfoName = LWL    (the chart LWL)
    ' InfoName = MAX    (the chart Y MAX)
    ' InfoName = MIN    (the chart Y MIN)
    ' InfoName = Title  (the chart TITLE)
    ' InfoName = INC    (Check for Increasing and Decreasing points )
    ' InfoName = ABOVE  (Check for Above and Below points)
    '
    Set nInfo = New Collection
    For i = 0 To 7
     If Trim(TxtInfo(i).Text) <> "" Then
       Set MYinfo = New ClsInfo
        MYinfo.InfoName = LblInfo(i).Caption
        MYinfo.YValue = TxtInfo(i).Text
        nInfo.Add MYinfo
     End If
    Next
    '
    ' Add Inc and Above check
    '
    Set MYinfo = New ClsInfo
    MYinfo.InfoName = "INC"
    MYinfo.YValue = TxtInfo(8).Text
    nInfo.Add MYinfo
    Set MYinfo = New ClsInfo
    MYinfo.InfoName = "ABOVE"
    MYinfo.YValue = TxtInfo(8).Text
    nInfo.Add MYinfo
        '
    '********* Draw Graph *********
    '
    ' Draw Scale
    '
    Scale (-10, 110)-(155, -10) ' Set custom coordinate system.
    '
    ' Draw rectangle for the graph area color gray
    '
    Line (0, 0)-(123, 100), Label5.BackColor, BF
    '
    ' Find Limits from ClsInfo
    '
    cLcl = ""
    cUcl = ""
    cLwl = ""
    cUwl = ""
    iInc = 0
    iAbove = 0
    bYMax = False
    bYMin = False
    bLimits = False
    For Each MYinfo In nInfo
       If UCase(MYinfo.InfoName) = "TITLE" Then
           TxtInfo(5).Text = MYinfo.YValue
       End If
       If UCase(MYinfo.InfoName) = "MAX" Then
         Ymax = MYinfo.YValue
          bYMax = True
       End If
       If UCase(MYinfo.InfoName) = "MIN" Then
           Ymin = MYinfo.YValue
           bYMmin = True
       End If
    Next
    '
    ' Determin min/max from your data
    '
    i = 1
    For Each MYdata In nData
      If i = 1 Then
          If bYMax = False Then Ymax = MYdata.YValue
          If bYMmin = False Then Ymin = MYdata.YValue
      End If
      If MYdata.YValue > Ymax Then Ymax = MYdata.YValue
      If MYdata.YValue < Ymin Then Ymin = MYdata.YValue
      i = 5
    Next
    '
    ' Expand min/max to include limits
    '
    For Each MYinfo In nInfo
       If UCase(MYinfo.InfoName) = "UCL" Then
          cUcl = MYinfo.YValue
           bLimits = True
           If MYinfo.YValue > Ymax Then Ymax = MYinfo.YValue
           If MYinfo.YValue < Ymin Then Ymin = MYinfo.YValue
        End If
       If UCase(MYinfo.InfoName) = "LCL" Then
          cLcl = MYinfo.YValue
          bLimits = True
          If MYinfo.YValue > Ymax Then Ymax = MYinfo.YValue
          If MYinfo.YValue < Ymin Then Ymin = MYinfo.YValue
       End If
       If UCase(MYinfo.InfoName) = "UWL" Then
          cUwl = MYinfo.YValue
           bLimits = True
           If MYinfo.YValue > Ymax Then Ymax = MYinfo.YValue
           If MYinfo.YValue < Ymin Then Ymin = MYinfo.YValue
       End If
       If UCase(MYinfo.InfoName) = "LWL" Then
           cLwl = MYinfo.YValue
           bLimits = True
           If MYinfo.YValue > Ymax Then Ymax = MYinfo.YValue
           If MYinfo.YValue < Ymin Then Ymin = MYinfo.YValue
       End If
       If UCase(MYinfo.InfoName) = "TAR" Then
           If MYinfo.YValue > Ymax Then Ymax = MYinfo.YValue
           If MYinfo.YValue < Ymin Then Ymin = MYinfo.YValue
       End If
       If UCase(MYinfo.InfoName) = "INC" Then
           iInc = MYinfo.YValue
       End If
       If UCase(MYinfo.InfoName) = "ABOVE" Then
            iAbove = MYinfo.YValue
       End If
    Next
    '
    ' 90 = The graph goes from 5 to 95
    ' Ydel2 = the offset to use to have points fit
    '
    Ydel = Ymax - Ymin
    Ydel2 = 90# / Ydel
    '
    ' Color graph (only if control or warning limits are present)
    '
    If bLimits Then
       '
       ' Draw everything red
       '
       Line (0, 0)-(123, 100), LblRed.BackColor, BF
       '
       ' Color Ucl to UWl Yellow
       '
       If cUwl <> "" And cUcl <> "" Then
          Ycal = cUcl
          Ycal = (Ycal - Ymin) * Ydel2 + 5
          Ycal2 = cUwl
          Ycal2 = (Ycal2 - Ymin) * Ydel2 + 5
          Line (0, Ycal)-(123, Ycal2), LblYellow.BackColor, BF
       End If
       If cLwl <> "" And cLcl <> "" Then
          Ycal = cLcl
          Ycal = (Ycal - Ymin) * Ydel2 + 5
          Ycal2 = cLwl
          Ycal2 = (Ycal2 - Ymin) * Ydel2 + 5
          Line (0, Ycal)-(123, Ycal2), LblYellow.BackColor, BF
       End If
       '
       ' Color Uwl to LWl green
       '
       If cUwl <> "" Then
        If cLwl <> "" Then
          Ycal = cUwl
          Ycal = (Ycal - Ymin) * Ydel2 + 5
          Ycal2 = cLwl
          Ycal2 = (Ycal2 - Ymin) * Ydel2 + 5
          Line (0, Ycal)-(123, Ycal2), LblGreen.BackColor, BF
        Else
          Ycal = cUwl
          Ycal = (Ycal - Ymin) * Ydel2 + 5
          Line (0, Ycal)-(123, 0), LblGreen.BackColor, BF
        End If
       End If
    End If
    '
    ' Draw lines around the graph
    '
    Line (0, 0)-(123, 0)
    Line (0, 0)-(0, 100)
    
    '''Line (0, 100)-(123, 100)
    '''Line (123, 0)-(123, 100)
    '
    ' Draw Limits lines and print limit value on form
    '
    For Each MYinfo In nInfo
       If UCase(MYinfo.InfoName) = "TAR" Then
           Ycal = (MYinfo.YValue - Ymin) * Ydel2 + 5
           Line (0, Ycal)-(123, Ycal)
           CurrentY = Ycal + 2
           CurrentX = CurrentX + 1
           Print "Tar " & MYinfo.YValue
       End If
       If UCase(MYinfo.InfoName) = "UCL" Then
          Ycal = (MYinfo.YValue - Ymin) * Ydel2 + 5
           Line (0, Ycal)-(123, Ycal)
           CurrentY = Ycal + 2
           CurrentX = CurrentX + 1
           Print "UCL " & MYinfo.YValue
       End If
       If UCase(MYinfo.InfoName) = "LCL" Then
           Ycal = (MYinfo.YValue - Ymin) * Ydel2 + 5
           Line (0, Ycal)-(123, Ycal)
           CurrentY = Ycal + 2
           CurrentX = CurrentX + 1
           Print "LCL " & MYinfo.YValue
       End If
       If UCase(MYinfo.InfoName) = "UWL" Then
           Ycal = (MYinfo.YValue - Ymin) * Ydel2 + 5
           Line (0, Ycal)-(123, Ycal)
           CurrentY = Ycal + 2
           CurrentX = CurrentX + 1
           Print "UWL " & MYinfo.YValue
       End If
       If UCase(MYinfo.InfoName) = "LWL" Then
           Ycal = (MYinfo.YValue - Ymin) * Ydel2 + 5
           Line (0, Ycal)-(123, Ycal)
           CurrentY = Ycal + 2
           CurrentX = CurrentX + 1
           Print "LWL " & MYinfo.YValue
       End If
    Next
    '
    ' Draw tick 5 marks and print value on form
    '
    If Ydel > 2 Then
        iTick = Int(Ydel / 10)
        If Ydel < 6 Then iTick = 1
        If Ydel > 49 Then
          iTick = Int(Ydel / 10)
          'iTick = iTick * 10
        End If
        j = Ymin / iTick
        Ycal = iTick
        If j > 0 Then Ycal = iTick * j
        Do While Ycal < Ymax
        Ycal2 = (Ycal - Ymin) * Ydel2 + 5
        If bLimits Then
           Line (-2, Ycal2)-(0, Ycal2)
           CurrentY = Ycal2 + 2
           CurrentX = CurrentX - 10
            Print Ycal
            
         Else
            Line (-2, Ycal2)-(0, Ycal2)
           CurrentY = Ycal2 + 2
           CurrentX = CurrentX - 10
            Print Ycal
            Line (0, Ycal2)-(123, Ycal2)
           CurrentY = Ycal2 + 2
           CurrentX = CurrentX + 1
           Print Ycal
         End If
         Ycal = Ycal + iTick
        Loop
    End If
    '
    '********* Plot Points *********
    '
    ' Draw actual points
    ' The graph will go from 5 to 95
    ' QBcolor(0)=black
    '
    FillColor = QBColor(0)
    FillStyle = 0
    sCount = 0
    iErr = 0
    iWarn = 0
    dSum = 0
    dSum2 = 0
    iIncCnt = 1
    iAboveCnt = 0
    iDecCnt = 1
    iBelowCnt = 0
    For Each MYdata In nData
      sCount = sCount + 1
      Ycal = (MYdata.YValue - Ymin) * Ydel2 + 5
      dSum = dSum + MYdata.YValue
      dSum2 = dSum2 + MYdata.YValue * MYdata.YValue
      dAve = dAve + MYdata.YValue
      j = MYdata.XValue '* 8
      '
      ' Type of point
      '
      If Option1(0) Then
           Circle (j, Ycal), 1, QBColor(0)
      Else
            'Line (j, 0)-(j, Ycal)
           ' Line (j, 0)-(j + 3, Ycal)
            Line (j, 0)-(j + 5, Ycal), RGB(0, 0, 255), BF
      End If
      '
      ' Check for errors
      '
      If cUwl <> "" Then
        Ycal2 = cUwl
        If MYdata.YValue > Ycal2 Then
          MYdata.PError = "Above UWL"
          MYdata.PColor = LblYellow.BackColor
        End If
      End If
      If cUcl <> "" Then
        Ycal2 = cUcl
        If MYdata.YValue > Ycal2 Then
          MYdata.PError = "Above UCL"
          MYdata.PColor = LblRed.BackColor
        End If
      End If
      If cLwl <> "" Then
        Ycal2 = cLwl
        If MYdata.YValue < Ycal2 Then
          MYdata.PError = "Below LWL"
          MYdata.PColor = LblYellow.BackColor
        End If
      End If
      If cLcl <> "" Then
        Ycal2 = cLcl
        If MYdata.YValue < Ycal2 Then
          MYdata.PError = "Below LCL"
          MYdata.PColor = LblRed.BackColor
        End If
      End If
      '
      ' Display 1st point info
      '
      If sCount = 1 Then
          TxtY.Text = MYdata.YValue
          TxtPoint.Text = 1
'          TxtName.Text = MYdata.PName
'          TxtError.Text = MYdata.PError
'          TxtError.BackColor = MYdata.PColor
          Yhi = MYdata.YValue
          Ylow = MYdata.YValue
          YOldPnt = MYdata.YValue
       End If
       
       If MYdata.YValue > Yhi Then Yhi = MYdata.YValue
       If MYdata.YValue < Ylow Then Ylow = MYdata.YValue
       '
       ' Check for Above/Below
       '
       If cUwl <> "" And cUcl <> "" Then
          Ycal2 = cUwl
          Ycal3 = cUcl
          If MYdata.YValue > Ycal2 And MYdata.YValue < Ycal3 Then
             iAboveCnt = iAboveCnt + 1
          Else
             iAboveCnt = 0
          End If
       Else
            iAboveCnt = 0
       End If
       If iAbove > 0 And iAboveCnt > iAbove Then
          If MYdata.PColor <> LblRed.BackColor Then
            MYdata.PError = "Too many Above"
            MYdata.PColor = LblRed.BackColor
          End If
       End If
       If cLwl <> "" And cLcl <> "" Then
          Ycal2 = cLwl
          Ycal3 = cLcl
          If MYdata.YValue < Ycal2 And MYdata.YValue > Ycal3 Then
             iBelowCnt = iBelowCnt + 1
          Else
             iBelowCnt = 0
          End If
        Else
            iBelowCnt = 0
        End If
        If iAbove > 0 And iBelowCnt > iAbove Then
          If MYdata.PColor <> LblRed.BackColor Then
            MYdata.PError = "Too many Below"
            MYdata.PColor = LblRed.BackColor
          End If
       End If
       '
       ' Check for Inc/dec
       '
       If MYdata.YValue > YOldPnt Then
          iIncCnt = iIncCnt + 1
       Else
            iIncCnt = 1
       End If
       If iInc > 0 And iIncCnt > iInc Then
          If MYdata.PColor <> LblRed.BackColor Then
            MYdata.PError = "Too many Inc"
            MYdata.PColor = LblRed.BackColor
          End If
       End If
       If MYdata.YValue < YOldPnt Then
           iDecCnt = iDecCnt + 1
       Else
            iDecCnt = 1
       End If
       If iInc > 0 And iDecCnt > iInc Then
          If MYdata.PColor <> LblRed.BackColor Then
             MYdata.PError = "Too many iDec"
             MYdata.PColor = LblRed.BackColor
          End If
       End If
       '
       ' Count Errors
       '
       If MYdata.PColor = LblYellow.BackColor Then iWarn = iWarn + 1
       If MYdata.PColor = LblRed.BackColor Then iErr = iErr + 1
       If MYdata.PColor = LblYellow.BackColor Then MYdata.PErrorType = "Warning"
       If MYdata.PColor = LblRed.BackColor Then MYdata.PErrorType = "Error"
       YOldPnt = MYdata.YValue
    Next
    
    ' Compute Average/Sd
    '
    dSD = 0
    dRange = Yhi - Ylow
    If sCount > 0 Then
       dAve = dAve / (sCount)
       If sCount > 1 Then
         dSD = dSum2 / (sCount - 1#) - (dSum * dSum / (sCount * (sCount - 1#)))
         dSD = Sqr(dSD)
       End If
    End If
    '
    '  Create a Summary
    '
    List1.Clear
    List1.AddItem "Points   " & sCount
    List1.AddItem "Hi       " & Yhi
    List1.AddItem "Low      " & Ylow
    List1.AddItem "Ave      " & dAve
    List1.AddItem "Range    " & dRange
    List1.AddItem "SD       " & dSD
    List1.AddItem "Warnings " & iWarn
    List1.AddItem "Errors   " & iErr
    '
End Sub


Private Sub Option1_Click(Index As Integer)
  Me.Refresh
End Sub
