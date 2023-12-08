VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form DialogDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Animation Conditions"
   ClientHeight    =   4740
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4275
   ClipControls    =   0   'False
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Slider Slider1 
      Height          =   615
      Left            =   210
      TabIndex        =   13
      Top             =   2760
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1085
      _Version        =   393216
      Min             =   -10
      TextPosition    =   1
   End
   Begin VB.ComboBox cboCurrent 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ComboBox cboTide 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdDate 
      Height          =   380
      Left            =   3840
      Picture         =   "Dialog.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Date"
      Top             =   120
      Width           =   380
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
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd-mmm-yy"
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cl&ose"
      Height          =   375
      Left            =   2130
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Confirm"
      Height          =   375
      Left            =   1050
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Confirm"
      Top             =   4320
      Width           =   1095
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
      Left            =   2040
      TabIndex        =   5
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   5
      Format          =   "hh:mm"
      Mask            =   "##:##"
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
      Left            =   2040
      TabIndex        =   11
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   3
      Mask            =   "999"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   615
      Left            =   210
      TabIndex        =   14
      Top             =   3600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1085
      _Version        =   393216
      Min             =   -10
      TextPosition    =   1
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "-       Vector Size        +"
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
      Left            =   210
      TabIndex        =   16
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "-     Vector Density      +"
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
      Left            =   210
      TabIndex        =   15
      Top             =   3360
      Width           =   3855
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
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Time Intervals (mins)"
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
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1815
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
      Left            =   120
      TabIndex        =   7
      Top             =   1080
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
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   735
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "DialogDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim resumeIndex As Integer

Private Sub CancelButton_Click()
  Unload Me
End Sub

Private Sub cmdDate_Click()
  diaDate.Show
End Sub

Private Sub Form_Load()
  SetWindowPos hWnd, -1, 0, 0, 0, 0, 3
  Me.cboTide.AddItem "Wet Season", 0
  Me.cboTide.AddItem "Dry Season", 1
  If strWeatherString = "W" Then
    Me.cboTide.ListIndex = 0
  Else
    Me.cboTide.ListIndex = 1
  End If
  
  DateWeatherDefault
  FixCombo

  Me.cboCurrent.AddItem "Depth Average", 0
  Me.cboCurrent.AddItem "Surface Layer", 1
  If strSDString = "D" Then
    Me.cboCurrent.ListIndex = 0
  Else
    Me.cboCurrent.ListIndex = 1
  End If
  
  Me.txtTimeInterval.Text = CStr(Round(timeInterVal * 60))
  
  Me.txtStartDate(0).Text = Format(DateCal, "yyyy-mm-dd")
  Me.txtStartDate(1).Text = Format(DateCal, "hh:mm")
  
  Me.Slider1.Value = (LegendWidth - 300) / 20
  Me.Slider2.Value = SliderBarValue
End Sub

Private Sub OKButton_Click()
  Dim chkMin As Double
  DateWeatherDefault
  If Not pointTranPeriod Then
    FixCombo
  End If
  If Me.cboTide.ListIndex = 0 Then
    strWeatherString = "W"
  Else
    strWeatherString = "D"
  End If
  If Me.cboCurrent.ListIndex = 1 Then
    strSDString = "S"
  Else
    strSDString = "D"
  End If
  DateCal = CDate(Me.txtStartDate(0) & " " & Me.txtStartDate(1).Text)
  frmMain.StatusBar1.Panels(5).Text = Format(DateCal, "yyyy-mmm-dd hh:mm")
  chkMin = CDbl(Replace(Me.txtTimeInterval.Text, "_", ""))
  If chkMin < 15 Or chkMin > 120 Then
    resumeIndex = 2
    errorFormCall = "DialogDate"
    ErrorMsg "Time Interval should be between 15 and 120"
    Me.txtTimeInterval = Round(timeInterVal * 60)
    Exit Sub
  End If
    
  LegendWidth = Me.Slider1.Value * 20 + 300
  SliderBarValue = Me.Slider2.Value
  
  timeInterVal = chkMin / 60
  If isFormLoad("frmTide") Then
    frmTide.RefreshGraph
  End If
  Unload Me
End Sub

Private Sub txtStartDate_Change(Index As Integer)
  DateWeatherDefault
  FixCombo
End Sub

Private Sub txtStartDate_LostFocus(Index As Integer)
  Select Case Index
    Case 0
      If Not IsDate(Me.txtStartDate(0).Text) Then
        resumeIndex = 0
        errorFormCall = "DialogDate"
        ErrorMsg "Invalid date !!"
        Me.txtStartDate(0).Text = Format(DateCal, "yyyy-mm-dd")
        Exit Sub
      End If
    Case 1
      If Not IsDate(Me.txtStartDate(0).Text & " " & Me.txtStartDate(1).Text) Then
        resumeIndex = 1
        errorFormCall = "DialogDate"
        ErrorMsg "Invalid Time!!"
        Me.txtStartDate(1).Text = Format(DateCal, "hh:mm")
        Exit Sub
      End If
  End Select
  If Not CheckDate(Me.txtStartDate(0).Text) Then
    Me.txtStartDate(0).Text = Format(DateCal, "yyyy-mm-dd")
    resumeIndex = 0
    errorFormCall = "DialogDate"
    ErrorMsg "Please select date range from 1/1/1995 to 31/12/2015"
  End If
  DateWeatherDefault
  FixCombo
End Sub

Private Sub FixCombo()
  If strWeatherString = "W" Then
    Me.cboTide.ListIndex = 0
  Else
    Me.cboTide.ListIndex = 1
  End If
  If TranPeriod = False Then
    Me.cboTide.Enabled = False
  Else
    Me.cboTide.Enabled = True
  End If
End Sub

Private Sub txtTimeInterval_LostFocus()
  Dim chkMin As Double
  chkMin = CDbl(Replace(Me.txtTimeInterval.Text, "_", ""))
  If chkMin < 15 Or chkMin > 120 Then
    resumeIndex = 2
    errorFormCall = "DialogDate"
    ErrorMsg "Time Interval should be between 15 and 120"
    Me.txtTimeInterval = Round(timeInterVal * 60)
    Exit Sub
  End If
End Sub

Public Sub resumeError()
  Select Case resumeIndex
    Case 0:
      Me.txtStartDate(0).SetFocus
    Case 1:
      Me.txtStartDate(1).SetFocus
    Case 2:
      Me.txtTimeInterval.SetFocus
  End Select
End Sub
