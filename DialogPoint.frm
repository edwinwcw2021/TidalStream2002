VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form DialogPoint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Control Point Conditions"
   ClientHeight    =   6795
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5010
   Icon            =   "DialogPoint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFunc 
      Height          =   380
      Index           =   1
      Left            =   2535
      Picture         =   "DialogPoint.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Save Point Setting"
      Top             =   645
      Width           =   380
   End
   Begin VB.CommandButton cmdFunc 
      Height          =   380
      Index           =   0
      Left            =   2160
      Picture         =   "DialogPoint.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Load Point Setting"
      Top             =   645
      Width           =   380
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "Control Point"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3000
      TabIndex        =   16
      Top             =   0
      Width           =   1935
      Begin VB.Label Label7 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   1485
         Width           =   1455
      End
      Begin VB.Label Label7 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   1110
         Width           =   1455
      End
      Begin VB.Label Label7 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   735
         Width           =   1455
      End
      Begin VB.Label Label7 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Export"
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   2535
      TabIndex        =   15
      ToolTipText     =   "Export"
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   495
      Index           =   0
      Left            =   1320
      TabIndex        =   12
      ToolTipText     =   "Generate"
      Top             =   6240
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   98
      TabIndex        =   11
      Top             =   3045
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5530
      View            =   3
      Arrange         =   2
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdDate 
      Height          =   380
      Left            =   2160
      Picture         =   "DialogPoint.frx":0C5E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Calendar"
      Top             =   120
      Width           =   380
   End
   Begin VB.ComboBox cboTide 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.ComboBox cboCurrent 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2640
      Width           =   2415
   End
   Begin MSMask.MaskEdBox txtStartDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "m/d/yy h:nn"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3076
         SubFormatType   =   4
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd-mmm-yy"
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtStartDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "m/d/yy h:nn"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3076
         SubFormatType   =   4
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   4
      Top             =   645
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   5
      Format          =   "hh:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDuration 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "m/d/yy h:nn"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3076
         SubFormatType   =   4
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   3
      Mask            =   "999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtTimeInterval 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "m/d/yy h:nn"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3076
         SubFormatType   =   4
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   3
      Mask            =   "999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label6 
      Caption         =   "Interval (minutes)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   14
      Top             =   2200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   10
      Top             =   150
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   90
      TabIndex        =   9
      Top             =   660
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Tide"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   8
      Top             =   1210
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Duration(hours)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   7
      Top             =   1720
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Current"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   6
      Top             =   2650
      Width           =   975
   End
End
Attribute VB_Name = "DialogPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim listCnt As Integer
Dim tmpCurrentIndex As Integer
Dim tmpTideIndex As Integer
Dim NauticalUnit As String
Dim resumeIndex As Integer

Private Sub chkCurrentChange()
  If tmpCurrentIndex <> Me.cboCurrent.ListIndex Then
    tmpCurrentIndex = Me.cboCurrent.ListIndex
    ClearArray
  End If
End Sub

Private Sub chkTideChange()
  If tmpTideIndex <> Me.cboTide.ListIndex Then
    tmpTideIndex = Me.cboTide.ListIndex
    ClearArray
  End If
End Sub

Private Sub cboCurrent_Click()
  chkCurrentChange
End Sub

Private Sub cboCurrent_KeyDown(KeyCode As Integer, Shift As Integer)
  chkCurrentChange
End Sub

Private Sub cboCurrent_LostFocus()
  chkCurrentChange
End Sub

Private Sub cboTide_Click()
  chkTideChange
End Sub

Private Sub cboTide_KeyDown(KeyCode As Integer, Shift As Integer)
  chkTideChange
End Sub

Private Sub cboTide_LostFocus()
  chkTideChange
End Sub

Private Sub cmdDate_Click()
  diaDate.Show
End Sub

Private Sub cmdFunc_Click(Index As Integer)
  Dim fs As New FileSystemObject
  Dim f As TextStream
  Dim expFileName As String
  If Index = 0 Then
    On Error GoTo exitFunc
    Me.CommonDialog1.DialogTitle = "Load Point Setting"
    Me.CommonDialog1.Filter = "*.set|*.set"
    Me.CommonDialog1.ShowOpen
    If Len(Me.CommonDialog1.filename) = 0 Then
      Exit Sub
    End If
    
    expFileName = Me.CommonDialog1.filename
    Set f = fs.OpenTextFile(expFileName, ForReading)
    Dim tmpPDate As String, tmpPTime As String, tmpPTide As Integer
    Dim tmpPDuration As Integer, tmpPInterval As Integer
    Dim tmpPCurrent As Integer
    Dim tmpPstrClickPoint As String
    Dim strLat As Long, strLong As Long
    
    On Error Resume Next
      tmpPDate = f.ReadLine
      tmpPTime = f.ReadLine
      tmpPTide = CInt(f.ReadLine)
      tmpPDuration = CInt(f.ReadLine)
      tmpPInterval = CInt(f.ReadLine)
      tmpPCurrent = CInt(f.ReadLine)
      tmpPstrClickPoint = f.ReadLine
      If Err <> 0 Then
        resumeIndex = 5
        errorFormCall = "DialogPoint"
        ErrorMsg "Invalid Point Setting File"
        On Error GoTo 0
        Exit Sub
      End If
    On Error GoTo 0
    f.Close
    
    If Not IsDate(tmpPDate) Then
      resumeIndex = 5
      errorFormCall = "DialogPoint"
      ErrorMsg "Invalid Point Setting File"
      Exit Sub
    End If
    On Error Resume Next
      tmpPDate = CDate(tmpPDate & " " & tmpPTime)
      If Err <> 0 Then
        resumeIndex = 5
        errorFormCall = "DialogPoint"
        ErrorMsg "Invalid Point Setting File"
        On Error GoTo 0
        Exit Sub
      End If
    On Error GoTo 0
    
    If Not IsDate(tmpPDate) Then
      resumeIndex = 5
      errorFormCall = "DialogPoint"
      ErrorMsg "Invalid Point Setting File"
      Exit Sub
    End If
'    If Year(tmpPDate) <> LimitedYear Then
'      resumeIndex = 5
'      errorFormCall = "DialogPoint"
'      ErrorMsg "Invalid Point Setting File"
'      Exit Sub
'    End If
                 
    If tmpPTide <> 0 And tmpPTide <> 1 Then
      resumeIndex = 5
      errorFormCall = "DialogPoint"
      ErrorMsg "Invalid Point Setting File"
      Exit Sub
    End If
    
    If tmpPCurrent <> 0 And tmpPCurrent <> 1 Then
      resumeIndex = 5
      errorFormCall = "DialogPoint"
      ErrorMsg "Invalid Point Setting File"
      Exit Sub
    End If
    
    If tmpPInterval < 15 Or tmpPInterval > 120 Then
      resumeIndex = 5
      errorFormCall = "DialogPoint"
      ErrorMsg "Invalid Point Setting File"
      Exit Sub
    End If
    
    If tmpPDuration < 1 Or tmpPDuration > 140 Then
      resumeIndex = 5
      errorFormCall = "DialogPoint"
      ErrorMsg "Invalid Point Setting File"
      Exit Sub
    End If
        
   
    rst.Open "select x, y from MNXY where mn='" & tmpPstrClickPoint & "'", conn, adOpenDynamic, adLockReadOnly
    If (rst.BOF And rst.EOF) Then
      rst.Close
      resumeIndex = 5
      errorFormCall = "DialogPoint"
      ErrorMsg "Invalid Point Setting File"
      Exit Sub
    End If
    
    'Start of no Error
    
    strClickPoint = tmpPstrClickPoint
    XY_To_Map rst.Fields(1).Value, rst.Fields(0).Value, strLat, strLong
    ClickX = rst.Fields(0).Value
    ClickY = rst.Fields(1).Value
    rst.Close
    frmMain.showPointClick strLat, strLong
    strClickY = LToDMSString2(strLat, 1)
    strClickX = LToDMSString2(strLong, 0)
    
    If tmpPTide = 0 Then
      strPointWeatherString = "W"
    Else
      strPointWeatherString = "D"
    End If
    
    If tmpPCurrent = 0 Then
      strPointSDString = "D"
    Else
      strPointSDString = "S"
    End If
    
    PointDateCal = tmpPDate
    Me.Label7(2) = ClickY & "N"
    Me.Label7(3) = ClickX & "E"
    Me.Label7(0) = strClickY
    Me.Label7(1) = strClickX
    
    Me.txtStartDate(0).Text = Format(PointDateCal, "yyyy-mm-dd")
    Me.txtStartDate(1).Text = Format(PointDateCal, "hh:mm")
    Me.txtTimeInterval.Text = tmpPInterval
    Me.txtDuration.Text = tmpPDuration
    
    If strPointSDString = "D" Then
      Me.cboCurrent.ListIndex = 0
    Else
      Me.cboCurrent.ListIndex = 1
    End If
    
    'Check Point Date Weather
    PointDateWeatherDefault
    If pointTranPeriod Then
      If tmpPTide = 0 Then
        strPointWeatherString = "W"
      Else
        strPointWeatherString = "D"
      End If
      If strPointWeatherString = "W" Then
        Me.cboTide.ListIndex = 0
      Else
        Me.cboTide.ListIndex = 1
      End If
    End If
    cmdGenerate_Click (0)
  Else
    On Error GoTo exitFunc
    Me.CommonDialog1.filename = ""
    Me.CommonDialog1.DialogTitle = "Save Point Setting"
    Me.CommonDialog1.Filter = "*.set|*.set"
    Me.CommonDialog1.ShowSave
  
    If Len(Me.CommonDialog1.filename) = 0 Then
      Exit Sub
    End If
    If Len(Me.CommonDialog1.filename) > 4 Then
      If Format(Right(Me.CommonDialog1.filename, 4), ">") = ".SET" Then
        expFileName = Me.CommonDialog1.filename
      Else
        expFileName = Me.CommonDialog1.filename & ".SET"
      End If
    End If
    If fs.FileExists(expFileName) Then
    DisableAllForm
      Hide
      If MsgBox("File exists Overwrite it", vbYesNo) = vbNo Then
        EnableAllForm
        Show
        Exit Sub
      End If
      EnableAllForm
      Show
    End If
    Set f = fs.CreateTextFile(expFileName, True)
    f.WriteLine Format(CDate(Me.txtStartDate(0).Text), "mmm dd, yyyy")
    f.WriteLine Me.txtStartDate(1).Text
    f.WriteLine Me.cboTide.ListIndex
    f.WriteLine Me.txtDuration.Text
    f.WriteLine Me.txtTimeInterval.Text
    f.WriteLine Me.cboCurrent.ListIndex
    f.WriteLine strClickPoint
    f.Close
  End If
exitFunc:
End Sub

Private Sub Form_Load()
  SetWindowPos hWnd, -1, 0, 0, 0, 0, 3
  Me.cboTide.AddItem "Wet Season", 0
  Me.cboTide.AddItem "Dry Season", 1
  If strPointWeatherString = "W" Then
    Me.cboTide.ListIndex = 0
  Else
    Me.cboTide.ListIndex = 1
  End If
  
  PointDateWeatherDefault
  FixCombo

  Me.ListView1.ColumnHeaders.Add , , "Date", (Me.ListView1.Width - 150) * 4 / 8 - 150
  Me.ListView1.ColumnHeaders.Add , , "Direction", (Me.ListView1.Width - 150) * 2 / 8
  Me.ListView1.ColumnHeaders.Add , , "Magnitude", (Me.ListView1.Width - 150) * 2 / 8

  Me.cboCurrent.AddItem "Depth Average", 0
  Me.cboCurrent.AddItem "Surface Layer", 1
  If strPointSDString = "D" Then
    Me.cboCurrent.ListIndex = 0
  Else
    Me.cboCurrent.ListIndex = 1
  End If
  Me.Label7(2) = ClickY & "N"
  Me.Label7(3) = ClickX & "E"
  Me.Label7(0) = strClickY
  Me.Label7(1) = strClickX
  Me.txtDuration.Text = duration
  listCnt = 0
  Me.txtTimeInterval.Text = CStr(Round(timePointInterVal * 60))
  Me.txtStartDate(0).Text = Format(PointDateCal, "yyyy-mm-dd")
  Me.txtStartDate(1).Text = Format(PointDateCal, "hh:mm")
End Sub

Private Sub cmdGenerate_Click(Index As Integer)
  Dim i As Integer
  If Index = 0 Then
    Dim chkMin As Double
    Dim strSQL As String
    Dim itemX As ListItem
    Dim runHourCount
    Dim runDate As Date
    
    PointDateWeatherDefault
    If Not pointTranPeriod Then
      FixCombo
    End If
    If Me.cboTide.ListIndex = 0 Then
      strPointWeatherString = "W"
    Else
      strPointWeatherString = "D"
    End If
    If Me.cboCurrent.ListIndex = 0 Then
      strPointSDString = "D"
    Else
      strPointSDString = "S"
    End If
    
    If Not checkValid Then
      Exit Sub
    End If
    
    chkMin = CDbl(Replace(Me.txtTimeInterval.Text, "_", ""))
    
    'Start Generate
    Screen.MousePointer = vbHourglass
    
    PointDateCal = CDate(Me.txtStartDate(0) & " " & Me.txtStartDate(1).Text)
    
    timePointInterVal = chkMin / 60
    strSQL = "delete from edwinTempCal2"
    conn.Execute strSQL
    
    InitPointMonthCal
    
    runHourCount = PointHourCount
    Do While runHourCount < PointHourCount + duration
      strSQL = "insert into edwinTempCal2 values ('" & strClickPoint & "','" & InitPointMonthDate & "'," & runHourCount & ")"
      conn.Execute strSQL
      runHourCount = runHourCount + timePointInterVal
    Loop
    'Removing List View
    ClearArray
    Me.cmdGenerate(1).Enabled = True
    strSQL = "select * from qryDir2_" & strPointSDString & strPointWeatherString & " order by time"
    rst.Open strSQL, conn, adOpenKeyset
    Do While Not rst.EOF
      runDate = DateAdd("n", rst.Fields(1) * 60, InitPointMonthDate)
      Set itemX = Me.ListView1.ListItems.Add(, , Format(runDate, "yyyy-mmm-dd hh:mm"))
      itemX.SubItems(1) = Round(edwinDir(rst.Fields(4), rst.Fields(5)), 1)
      If frmMain.cboKnotMS.ListIndex = 0 Then
        itemX.SubItems(2) = Round(rst.Fields(3) / KnotToMS, 2)
      Else
        itemX.SubItems(2) = Round(rst.Fields(3), 2)
      End If
      rst.MoveNext
      listCnt = listCnt + 1
    Loop
    rst.Close
    NauticalUnit = frmMain.cboKnotMS.Text
    Me.cmdGenerate(1).SetFocus
    Screen.MousePointer = vbDefault
  Else
    'Export File
    
    Dim fs As New FileSystemObject
    Dim f As TextStream
    Dim expFileName As String
    On Error GoTo exitExportFile
    Me.CommonDialog1.filename = ""
    Me.CommonDialog1.DialogTitle = "Export File"
    Me.CommonDialog1.Filter = "*.txt|*.txt"
    Me.CommonDialog1.ShowSave
    If Len(Me.CommonDialog1.filename) = 0 Then
      Exit Sub
    End If
    If Len(Me.CommonDialog1.filename) > 4 Then
      If Format(Right(Me.CommonDialog1.filename, 4), ">") = ".TXT" Then
        expFileName = Me.CommonDialog1.filename
      Else
        expFileName = Me.CommonDialog1.filename & ".TXT"
      End If
    End If
    If fs.FileExists(expFileName) Then
      DisableAllForm
      Hide
      If MsgBox("File exists Overwrite it", vbYesNo) = vbNo Then
        EnableAllForm
        Show
        Exit Sub
      End If
      EnableAllForm
      Show
    End If
    Set f = fs.CreateTextFile(expFileName, True)
    f.WriteLine "* Digital Tidal Stream Atlas 2003"
    f.WriteLine "* Coordinates:"
    f.WriteLine "* (WGS84) Lat = " & Me.Label7(0) & Space(15 - Len(Me.Label7(1))) & ", Long = " & Me.Label7(1)
    f.WriteLine "* (HK80 Grid)   " & Me.Label7(2) & Space(14 - Len(Me.Label7(3))) & ",    " & Me.Label7(3)
    f.WriteLine
    f.WriteLine "* " & Me.cboTide.Text & ", " & Me.cboCurrent.Text & ", Time Interval: " & Me.txtTimeInterval.Text & " min, Period: " & Me.txtDuration.Text & " hours"
    f.WriteLine
    f.WriteLine "* Current magnitude in " & NauticalUnit
    f.WriteLine
    f.WriteLine "Date/Time                Direction        Speed"
    f.WriteLine "--------------------------------------------------------"
    For i = 1 To listCnt
      f.WriteLine Me.ListView1.ListItems.Item(i).Text & Space(8) & Me.ListView1.ListItems.Item(i).SubItems(1) & Space(17 - Len(Me.ListView1.ListItems.Item(i).SubItems(1))) & Format(Me.ListView1.ListItems.Item(i).SubItems(2), "0.00")
    Next
    f.Close
    Set fs = Nothing
  End If
exitExportFile:
End Sub

Private Sub ImageCombo1_Change()

End Sub

Private Sub txtDuration_Change()
  ClearArray
End Sub

Private Sub txtDuration_LostFocus()
  If Not checkValid Then
    Exit Sub
  End If
End Sub

Private Sub txtStartDate_Change(Index As Integer)
  PointDateWeatherDefault
  FixCombo
  ClearArray
End Sub

Private Sub txtStartDate_LostFocus(Index As Integer)
  Select Case Index
    Case 0
      If Not IsDate(Me.txtStartDate(0).Text) Then
        resumeIndex = 0
        errorFormCall = "DialogPoint"
        ErrorMsg "Invalid date !!"
        Me.txtStartDate(0).Text = Format(PointDateCal, "yyyy-mm-dd")
        Exit Sub
      End If
    Case 1
      If Not IsDate(Me.txtStartDate(0).Text & " " & Me.txtStartDate(1).Text) Then
        resumeIndex = 1
        errorFormCall = "DialogPoint"
        ErrorMsg "Invalid Time!!"
        Me.txtStartDate(1).Text = Format(PointDateCal, "hh:mm")
        Exit Sub
      End If
  End Select
  If Not CheckDate(Me.txtStartDate(0).Text) Then
    Me.txtStartDate(0).Text = Format(PointDateCal, "yyyy-mm-dd")
    resumeIndex = 0
    errorFormCall = "DialogPoint"
    ErrorMsg "Please select date range from 1/1/1995 to 31/12/2015"
  End If
  DateWeatherDefault
  FixCombo
End Sub

Private Sub FixCombo()
  If strPointWeatherString = "W" Then
    Me.cboTide.ListIndex = 0
  Else
    Me.cboTide.ListIndex = 1
  End If
  If pointTranPeriod = False Then
    Me.cboTide.Enabled = False
  Else
    Me.cboTide.Enabled = True
  End If
End Sub

Public Sub ClearArray()
  Dim i As Integer
  listCnt = Me.ListView1.ListItems.Count
  For i = 0 To listCnt - 1
    Me.ListView1.ListItems.Remove 1
  Next
  listCnt = 0
  Me.cmdGenerate(1).Enabled = False
End Sub

Private Sub txtTimeInterval_Change()
  ClearArray
End Sub

Private Sub txtTimeInterval_LostFocus()
  checkValid
End Sub

Private Function checkValid() As Boolean
  Dim chkMin As Double
  Dim chkDuration As Integer
  
  chkDuration = CInt(Me.txtDuration.Text)
  If chkDuration < 1 Or chkDuration > 140 Then
    resumeIndex = 2
    errorFormCall = "DialogPoint"
    Me.txtDuration.Text = duration
    Me.txtTimeInterval = Round(timePointInterVal * 60)
    ErrorMsg "Duration should be between 1 and 140"
    checkValid = False
    Exit Function
  End If
  duration = chkDuration
    
  chkMin = CDbl(Replace(Me.txtTimeInterval.Text, "_", ""))
  If chkMin > duration * 60 Then
    resumeIndex = 4
    errorFormCall = "DialogPoint"
    ErrorMsg "Time Interval cannot greater than duration"
    checkValid = False
    Exit Function
  End If
  
  If chkMin < 15 Or chkMin > 120 Then
    resumeIndex = 4
    errorFormCall = "DialogPoint"
    Me.txtTimeInterval = Round(timePointInterVal * 60)
    ErrorMsg "Time Interval should be between 15 and 120"
    checkValid = False
    Exit Function
  End If
  checkValid = True
End Function

Public Sub resumeError()
  Select Case resumeIndex
    Case 0:
      Me.txtStartDate(0).SetFocus
    Case 1:
      Me.txtStartDate(1).SetFocus
    Case 3:
      Me.txtDuration.SetFocus
    Case 4:
      Me.txtTimeInterval.SetFocus
    Case 5:
      Me.cmdFunc(0).SetFocus
  End Select
End Sub

