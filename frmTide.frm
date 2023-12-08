VERSION 5.00
Begin VB.Form frmTide 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tide Graph"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   Icon            =   "frmTide.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleMode       =   0  'User
   ScaleWidth      =   3377.83
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   910
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   2280
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2500
      Left            =   240
      ScaleHeight     =   2475
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   360
      Width           =   2500
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tide Station"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   900
   End
End
Attribute VB_Name = "frmTide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xInit As Integer, yInit As Integer, yOffset As Integer, timeWidth As Integer, defaultIndex As Integer, yScale As Double
Dim StationName(8) As String

Private Sub Combo1_Click()
  tmpCheckChange
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
  tmpCheckChange
End Sub

Private Sub Combo1_LostFocus()
  tmpCheckChange
End Sub

Private Sub Form_Load()
  Dim i As Integer
  Dim strTideHeight As String
  
  SetWindowPos hWnd, -1, 0, 0, 0, 0, 3
  With Picture1
    .Top = 2 + Combo1.Height
    .Left = 2
    .Width = frmTide.ScaleWidth - 4
    .Height = frmTide.ScaleHeight - 4 - Combo1.Height
    Me.Picture1.Scale (0, .Height)-(.Width, 0)
    xInit = 300
    yOffset = 300
    timeWidth = 112
    yScale = (.Height - 2 * yOffset) / 5
    yInit = yScale + yOffset
  End With
  rst.Open "select * from tblPlaceName order by fldTableNameS", conn, adOpenDynamic
  i = 0
  Do While Not rst.EOF
    Me.Combo1.AddItem rst.Fields(1)
    StationName(i) = rst.Fields(0)
    rst.MoveNext
    i = i + 1
  Loop
  rst.Close
  defaultIndex = 3
  Me.Combo1.ListIndex = defaultIndex
  Me.Picture1.BackColor = RGB(221, 221, 221)
  RefreshGraph
End Sub

Public Function RefreshGraph()
  Dim TableName As String
  Dim tmpDateStart As Date
  Dim tmpDateEnd As Date
  Dim tmpX(25) As Double
  Dim tmpY(25) As Double
  Dim fx() As Double
  Dim fy() As Double
  Dim i As Integer
  Dim strSQL As String
  Dim corX1 As Integer, corY1 As Integer
  Dim corX2 As Integer, corY2 As Integer
  Dim currentHour As Integer
  Dim currentMin As Double
  Dim strCurrentTide As String
  Dim strTideHeight As String
  
  If Year(DateCal) = 2003 Then
    TableName = StationName(defaultIndex)
    currentHour = Hour(DateCal)
    If currentHour > 12 Then
      tmpDateStart = DateAdd("h", 12, CDate(Format(DateCal, "dd mmmm, yyyy")))
      tmpDateEnd = DateAdd("h", 37, CDate(Format(DateCal, "dd mmmm, yyyy")))
    Else
      tmpDateStart = CDate(Format(DateCal, "dd mmmm, yyyy"))
      tmpDateEnd = DateAdd("h", 25, CDate(Format(DateCal, "dd mmmm, yyyy")))
    End If
    strSQL = "select tHeight from " & TableName & " where tdate >= #" & Format(tmpDateStart, "yyyy-mmm-dd hh:mm") & "# and  tdate<#" & Format(tmpDateEnd, "yyyy-mmm-dd hh:mm") & "# order by tDate"
    rst.Open strSQL, conn, adOpenKeyset
    i = 1
      
    Do While Not rst.EOF
      tmpX(i) = xInit + (i - 1) * 112
      tmpY(i) = rst.Fields(0)
      rst.MoveNext
      i = i + 1
    Loop
    rst.Close
    catint tmpX, tmpY, i - 1, 60, fx, fy
    
    'Draw Border and legend
    Me.Picture1.Cls
    Picture1.Line (xInit, Picture1.Height - yOffset)-(xInit, yOffset), RGB(0, 0, 0)
    Picture1.Line (xInit, yInit)-(timeWidth * 24 + xInit, yInit), RGB(0, 0, 0)
    With Me.Picture1
      .ForeColor = RGB(0, 0, 0)
      .FontBold = False
      .FontSize = 8
      .FontName = "Arial"
      .CurrentX = (.Width - TextWidth(Me.Combo1.Text)) / 2
      .CurrentY = .Height
    End With
    Me.Picture1.Print Me.Combo1.Text
    
    'XY Legend
    With Me.Picture1
      .CurrentX = xInit - TextWidth("(m)") / 2
      .CurrentY = .Height
    End With
    Me.Picture1.Print "(m)"
    
    With Me.Picture1
      .CurrentX = xInit + timeWidth * 24 - TextWidth("(hrs)")
      .CurrentY = yInit - 200
    End With
    Me.Picture1.Print "(hrs)"
    
    'Draw Legend
    For i = 0 To 5
      Picture1.Line (xInit - 30, yOffset + i * yScale)-(xInit + 30, yOffset + i * yScale)
      With Picture1
        .CurrentX = xInit - 30 - TextWidth(CStr(i - 1))
        .CurrentY = yOffset + i * yScale + 100
      End With
      Picture1.Print CStr((i - 1))
    Next
    
    For i = 0 To 23
      Picture1.Line (xInit + timeWidth * i, yInit - 30)-(xInit + timeWidth * i, yInit + 30)
    Next
    If currentHour > 12 Then
      With Picture1
        .FontSize = 8
        .FontName = "Arial"
        .CurrentX = xInit + timeWidth * 12 - TextWidth("24") / 2
        .CurrentY = yInit - 30
      End With
      Picture1.Print "24"
      With Picture1
        .FontSize = 8
        .FontName = "Arial"
        .CurrentX = xInit + timeWidth * 24 - TextWidth("12") / 2
        .CurrentY = yInit - 30
      End With
      Picture1.Print "12"
    Else
      With Picture1
        .FontSize = 8
        .FontName = "Arial"
        .CurrentX = xInit + timeWidth * 12 - TextWidth("12") / 2
        .CurrentY = yInit - 30
      End With
      Picture1.Print "12"
      With Picture1
        .FontSize = 8
        .FontName = "Arial"
        .CurrentX = xInit + timeWidth * 24 - TextWidth("24") / 2
        .CurrentY = yInit - 30
      End With
      Picture1.Print "24"
    End If
    
    'Draw Current Tide
    currentMin = currentHour * 60 + Minute(DateCal)
    If currentHour > 12 Then
      strTideHeight = "   " & Round(fy(currentMin - 12 * 60), 2) & "m"
    Else
      strTideHeight = "   " & Round(fy(currentMin), 2) & "m"
    End If
    strCurrentTide = Format(DateCal, "yyyy-mmm-dd hh:mm")
    
    With Picture1
      .CurrentX = (.Width - TextWidth(strCurrentTide) - TextWidth(strTideHeight)) / 2
      .CurrentY = 200
    End With
    Picture1.Print strCurrentTide
    With Picture1
      .FontBold = True
      .ForeColor = RGB(255, 0, 0)
      .CurrentX = (.Width + TextWidth(strCurrentTide) - TextWidth(strTideHeight)) / 2
      .CurrentY = 200
    End With
    Picture1.Print strTideHeight
    
    If currentHour > 12 Then
      currentHour = currentHour - 12
      currentMin = currentMin - 12 * 60
    End If
    'currentMin = currentHour + Minute(DateCal) / 60
    currentMin = currentMin / 60
    Picture1.Line (xInit + currentMin * timeWidth, Picture1.Height - yOffset)-(xInit + currentMin * timeWidth, yOffset), RGB(255, 0, 0)
    
    For i = 1 To UBound(fx) - 61
      corX1 = Int(fx(i))
      corX2 = Int(fx(i + 1))
      corY1 = Int(fy(i) * yScale) + yInit
      corY2 = Int(fy(i + 1) * yScale) + yInit
      Me.Picture1.Line (corX1, corY1)-(corX2, corY2), RGB(0, 0, 255)
    Next
  End If
End Function

Private Sub tmpCheckChange()
  If CheckChange Then
    defaultIndex = Me.Combo1.ListIndex
    Me.RefreshGraph
  End If
End Sub

Private Function CheckChange() As Boolean
  CheckChange = Me.Combo1.ListIndex <> defaultIndex
End Function

Private Sub Form_Unload(Cancel As Integer)
  frmMain.mnuShowTide.Checked = False
  frmMain.setPressButton
End Sub
