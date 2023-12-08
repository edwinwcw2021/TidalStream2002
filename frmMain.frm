VERSION 5.00
Object = "{EFAB76C0-9F63-11CF-A48A-A0AC34F4689F}#2.0#0"; "Pwstrv2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Digital Tidal Stream Atlas"
   ClientHeight    =   6120
   ClientLeft      =   2235
   ClientTop       =   2310
   ClientWidth     =   10260
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   10260
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   8400
      Top             =   720
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   975
      Left            =   9360
      ScaleHeight     =   915
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   9000
      Top             =   480
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5865
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2505
            MinWidth        =   2505
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2963
            MinWidth        =   2963
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2275
            MinWidth        =   2275
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2275
            MinWidth        =   2275
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3087
            MinWidth        =   3087
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3263
            MinWidth        =   3263
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3969
            MinWidth        =   3969
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PWSTREETLib.Pwstreet Pwstreet1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5655
      _Version        =   131072
      _ExtentX        =   9975
      _ExtentY        =   9340
      _StockProps     =   32
      RightMouseMenu  =   0   'False
      ScrollBars      =   0   'False
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "zoom"
            Object.ToolTipText     =   "Zoom/Point"
            ImageIndex      =   1
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pan"
            Object.ToolTipText     =   "Pan"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "play"
            Object.ToolTipText     =   "Start Animation"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "pause"
            Object.ToolTipText     =   "Pause"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "forward"
            Object.ToolTipText     =   "Step Forward"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Load View"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save View"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "measure"
            Object.ToolTipText     =   "Measure Distance"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "setting"
            Object.ToolTipText     =   "Set Control Point Conditions"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SettingAn"
            Object.ToolTipText     =   "Set Animation Conditions"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ShowTide"
            Object.ToolTipText     =   "Show Tide Graph"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "overview"
            Object.ToolTipText     =   "View Overview"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "zoomin"
            Object.ToolTipText     =   "Zoom In"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "zoomout"
            Object.ToolTipText     =   "Zoom Out"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "globe"
            Object.ToolTipText     =   "Full Extent"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   6170
         TabIndex        =   7
         Top             =   0
         Width           =   1695
         Begin VB.ComboBox cboKnotMS 
            Height          =   315
            Left            =   550
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   40
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Speed"
            Height          =   255
            Left            =   30
            TabIndex        =   9
            Top             =   100
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   7860
         TabIndex        =   4
         Top             =   0
         Width           =   3250
         Begin VB.Label Label1 
            Caption         =   "500m"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   6
            Top             =   100
            Width           =   400
         End
         Begin VB.Label Label1 
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   5
            Top             =   100
            Width           =   100
         End
         Begin VB.Shape Shape1 
            Height          =   135
            Index           =   1
            Left            =   1200
            Top             =   130
            Width           =   375
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   120
            Top             =   130
            Width           =   375
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10860
         Top             =   -120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   21
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":08A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0E3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0F98
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1532
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1ACC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1DE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2380
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":291A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2EB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":344E
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":39E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3D02
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":429C
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4B76
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5110
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":526A
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5804
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":595E
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5AB8
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6052
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoadZoom 
         Caption         =   "&Load View"
      End
      Begin VB.Menu mnuSaveZoom 
         Caption         =   "&Save View"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuExportMapToFile 
         Caption         =   "&Export Map to File"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuDateStart 
         Caption         =   "Set Animation Conditions"
      End
      Begin VB.Menu mnuSinglePoint 
         Caption         =   "Set Control Point Conditions"
      End
      Begin VB.Menu mnuSept2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStartAnimation 
         Caption         =   "Start Animation"
      End
      Begin VB.Menu mnuStopAnimation 
         Caption         =   "Stop Animation"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuStepForward 
         Caption         =   "Step Forward"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSept 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMeasureDistance 
         Caption         =   "Measure Distance"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuZoomIn 
         Caption         =   "Zoom &In (+)"
      End
      Begin VB.Menu mnuZoomOut 
         Caption         =   "Zoom &Out (-)"
      End
      Begin VB.Menu mnuFullExtent 
         Caption         =   "Full Extent (F11)"
      End
      Begin VB.Menu mnuSept3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowTide 
         Caption         =   "View Tide Graph"
      End
      Begin VB.Menu mnuViewOverView 
         Caption         =   "View Overview"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Layer"
      Begin VB.Menu mnuShowPlaceLabel 
         Caption         =   "Show Place Label"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShowGridlines 
         Caption         =   "Show Gridlines"
      End
      Begin VB.Menu mnuFairway 
         Caption         =   "Show Fairway"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShowAnchorage 
         Caption         =   "Show Anchorage"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuTemp 
         Caption         =   "Temp"
      End
      Begin VB.Menu mnuAboutTidalStreamAtlas 
         Caption         =   "&About Digital Tidal Stream Atlas"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n As Long
Dim nrow As Integer
Dim filename As String
Dim Layer As Long
Dim zLayer As Long
Dim xtop As Long
Dim xleft As Long
Dim xbottom As Long
Dim xright As Long
Dim resumeIndex As Integer

Dim ScaleWidihZoom As Long
Dim ScaleHeightZoom As Long
Dim MapCenterLat As Long
Dim MapCenterLong As Long
Dim tmpScaleFactor As Double
Dim ScaleFactorLX As Double
Dim ScaleFactorLY As Double

Dim xCurrentLat As Long
Dim yCurrentLong As Long

'Flag
Dim MeasureCheck As Boolean
Dim NoControlPointFound As Boolean
Dim CachedStateLabelWhenConfigurationSetting As Long
Dim clickToolButton As Integer
Dim noPoint As Boolean

'User Define Line
Dim ln As UserDefinedLine               'Arrow
Dim lnLegend As UserDefinedLine         'Legend
Dim lnGrid As UserDefinedLine           'Grid
Dim lnMeasure As UserDefinedLine        'Measure
Dim lnRegion As UserDefinedLine        'Region

Dim ll(1 To 2, 1 To 9) As Long          'Arrow Line Array
Dim llg(1 To 2, 1 To 5) As Long         'Legend Line Array
Dim llGrid(1 To 2, 1 To 3) As Long      'Grid Line Array
Dim llMeasure(1 To 2, 1 To 2) As Long   'Measure Line Array
Dim llRegion(1 To 2, 1 To 5) As Long    'Region Line Array

Private Sub cboKnotMS_Click()
  clearNautical
End Sub

Private Sub cboKnotMS_KeyDown(KeyCode As Integer, Shift As Integer)
  clearNautical
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyAdd Then
    mnuZoomIn_Click
  End If
  If KeyCode = vbKeySubtract Then
    mnuZoomOut_Click
  End If
End Sub

Private Sub Form_Load()
  'Location Map Center
  centerMap
  'Load Layer
  
  Me.MousePointer = vbHourglass
  
  DefaultPath = App.Path
  
  filename = DefaultPath & "\map\fairway_polygon.pwc"
  lyrVFair = AttachData(filename)
  DrawLayer lyrVFair, 5
  
  filename = DefaultPath & "\map\fwy_l_wgs84.pwc"
  lyrFair = AttachData(filename)
  DrawLayer lyrFair, 4

  filename = DefaultPath & "\map\anch_a_wgs84.pwc"
  lyrAnor = AttachData(filename)
  DrawLayer lyrAnor, 2
  
  filename = DefaultPath & "\map\land_a.pwc"
  Layer = AttachData(filename)
  DrawLayer Layer, 3
  
  filename = DefaultPath & "\map\land_l.pwc"
  Layer = AttachData(filename)
  DrawLayer Layer, 0
  
  renderLabel Layer, "LEGEND"
  filename = DefaultPath & "\map\sar_l_wgs84.pwc"
  zLayer = AttachData(filename)
  DrawLayer zLayer, 1
  
  ZoomLayer zLayer
    
  'Init Legend Length
  LegendWidth = 300
  
  'Init Database Connection
  InitDB
  
  'Show Place Name
  InitPlaceName
  showPlaceName
   
  'Init Flag
  MeasureCheck = False
  
  'Set Background Color and left mouse pan
  With Me.Pwstreet1
    .Feature = pwVoidPolygon
    .Value = RGB(255, 255, 255)
    .Attribute = pwFillColor
    .Action = pwSetConfig
    'Disable Measure
    .Feature = pwLeftMouseButton
    .Attribute = pwPlusShiftKey
    .MouseBehavior = pwMouseDisabled
    
    'Show Anch
    .Value = lyrAnor
    .Picture = "OBJNAM"
    .Action = pwSetAttachedEnumAttribute
    .Value = lyrVFair
    .Picture = "NAME"
    .Action = pwSetAttachedEnumAttribute
  End With
    
  'Set properties for layer easier to see
  SetConfig pwState, pwLabelWhen, pwNever
  SetConfig pwState, pwFillWhen, pwNever
  SetConfig pwMapCursor, pwDrawWhen, pwNever
  SetConfig pwSearchHighlight, pwDrawWhen, pwNever  'don't highlight Nipp St
  SetConfig pwStructureNumber, pwLabelWhen, pwNever ' and don't show street numbers either

  'Set Default Zoon Margin
  xOrgtop = xtop
  xOrgleft = xleft
  xOrgbottom = xbottom
  xOrgright = xright
  
  'Initialize Meterics Nautical ComboBox
  defaultNautical = 0
  Me.cboKnotMS.AddItem "Knots", 0
  Me.cboKnotMS.AddItem "m/s", 1
  Me.cboKnotMS.ListIndex = defaultNautical
  MagScaleFactor = 1
  
  Dim i As Integer
          
  'Draw Grid
  'DrawGrid
            
  With Pwstreet1
    .Value = 249989
    .Action = pwSetMaxScaleFactor
    .Value = 12000
    .Action = pwSetMinScaleFactor
  End With
  
  'Temp Init Date
'  If Year(Date) <> LimitedYear Then
'    DateCal = CDate("2003-Jan-01")
'    PointDateCal = CDate("2003-Jan-01")
'  Else
'    DateCal = CDate(LimitedYear & " " & Format(Date, "mmm dd"))
'    PointDateCal = CDate(LimitedYear & " " & Format(Date, "mmm dd"))
'  End If
  startLimitYear = CDate("Jan 1, 1995 00:00:00")
  endLimitYear = CDate("Dec 31, 2015 12:59:59")
  
  DateCal = Now
  PointDateCal = Now
  
  DateWeatherDefault
  timeInterVal = 1
  timePointInterVal = 1
  duration = 48
  
  'Init Record Count
  recordCnt = 0
  strSDString = "D"
  strPointSDString = "D"
  SliderBarValue = 0
  varAddFactor = 5
    
  Me.WindowState = 2
  Me.MousePointer = vbDefault
  clickToolButton = 0
End Sub

Private Sub DrawLayer(Layer As Long, c1 As Integer)
  With Pwstreet1
    .Select Layer
    .EndSelect
    n = Pwstreet1.Value
    Select Case c1
      Case 0:
        .SetSelectedObjectsRenderingAttr pwBlack, 0, 1, pwBlack, pwSolidFill
      Case 1:
        .SetSelectedObjectsRenderingAttr pwRed, 0, 1, pwRed, pwSolidFill
      Case 2:
        .SetSelectedObjectsRenderingAttr pwBlue, 0, 1, RGB(239, 255, 255), pwSolidFill
      Case 3:
        .SetSelectedObjectsRenderingAttr RGB(255, 251, 165), 0, 1, pwLightYellow, pwSolidFill
      Case 4:
        .SetSelectedObjectsRenderingAttr pwBlue, 0, 1, pwBlue, pwSolidFill
      Case 5:
        .SetSelectedObjectsRenderingAttr RGB(255, 255, 255), 0, 1, RGB(255, 255, 255), pwSolidFill
      Case 99:
        .SetSelectedObjectsRenderingAttr RGB(255, 255, 255), 0, 1, RGB(255, 255, 255), pwSolidFill
    End Select
    .Action = pwDraw
    'ZoomLayer Layer
  End With
End Sub

Private Sub Form_Resize()
  centerMap
End Sub

Private Sub Form_Unload(Cancel As Integer)
  endApp
End Sub

Private Sub mnuAboutTidalStreamAtlas_Click()
  If mnuStartAnimation.Checked Then
    mnuStopAnimation_Click
  End If
  diaAbout.Show
End Sub

Private Sub mnuDateStart_Click()
  If mnuStartAnimation.Checked Or mnuPause.Checked Then
    mnuStopAnimation_Click
  End If
  DialogDate.Show
End Sub

Public Sub mnuExit_Click()
  endApp
End Sub

Private Sub mnuExportMapToFile_Click()
  If mnuStartAnimation.Checked Then
    mnuStopAnimation_Click
  End If
  'unloadAllForm
  Dim frm As Form
  For Each frm In Forms
    If frm.Name <> "frmMain" Then
      Unload frm
    End If
  Next
  Timer2.Enabled = True
End Sub

Private Sub mnuFullExtent_Click()
  If mnuStartAnimation.Checked Then
    mnuStopAnimation_Click
  End If
  With Pwstreet1
    .Latitude = xOrgtop
    .Longitude = xOrgleft
    .ID = xOrgbottom
    .Value = xOrgright
    .Action = 997
  End With
End Sub

Private Sub mnuLoadZoom_Click()
  If mnuStartAnimation.Checked Then
    mnuStopAnimation_Click
  End If
  Dim expFileName As String
  CommonDialog1.CancelError = True
  Me.CommonDialog1.DialogTitle = "Load View"
  Me.CommonDialog1.Filter = "Viw File (*.viw)|*.viw"
  On Error GoTo exitLoadZoom
  Me.CommonDialog1.ShowOpen
  If Len(Me.CommonDialog1.filename) = 0 Then
    Exit Sub
  End If
  
  expFileName = Me.CommonDialog1.filename
  
  Dim fs As New FileSystemObject
  If Not fs.FileExists(expFileName) Then
    ErrorMsg "File does not exists !!"
    Exit Sub
  End If
  Dim f As TextStream
  Set f = fs.OpenTextFile(expFileName, ForReading)
  Dim tmpLat1 As Long, tmpLong1 As Long, tmpScaleFactor1 As Long
  On Error Resume Next
    tmpLat1 = CLng(f.ReadLine)
    tmpLong1 = CLng(f.ReadLine)
    tmpScaleFactor1 = CLng(f.ReadLine)
    If Err <> 0 Then
      ErrorMsg "Invalid View File"
      Exit Sub
    End If
  On Error GoTo 0
  'Add Check Data
  With Pwstreet1
    .Latitude = tmpLat1
    .Longitude = tmpLong1
    .ScaleFactor = tmpScaleFactor1
  End With
exitLoadZoom:
End Sub

Private Sub mnuMeasureDistance_Click()
  If mnuStartAnimation.Checked Then
    mnuStopAnimation_Click
  End If
  With Me.mnuMeasureDistance
    If .Checked Then
      .Checked = False
      clearPoint
      strClickPoint = ""
      clickToolButton = 0
      setPressButton
    Else
      With Pwstreet1
        .Feature = pwLeftMouseButton
        .Attribute = pwPlusNoKey
        .MouseBehavior = pwMouseV1LeftButton
      End With
      MeasureCheck = False
      .Checked = True
      clickToolButton = 2
      setPressButton
      unloadAllForm
    End If
  End With
End Sub

Private Sub mnuPrint_Click()
  DisableAllForm
  If mnuStartAnimation.Checked Then
    mnuStopAnimation_Click
  End If
  frmPrint.Show
End Sub

Private Sub mnuFairway_Click()
  If Me.mnuFairway.Checked Then
'    With Me.Pwstreet1
'      .Value = lyrFair
'      .Action = pwDetachFile
'      .Action = pwDraw
'      .Value = lyrVFair
'      .Action = pwDetachFile
'      .Action = pwDraw
'    End With
    DrawLayer lyrFair, 99
    'DrawLayer lyrFair, 99
    With Me.Pwstreet1
      .Value = lyrVFair
      .Picture = ""
      .Action = pwSetAttachedEnumAttribute
    End With
    Me.mnuFairway.Checked = False
  Else
'    filename = DefaultPath & "\map\fairway_polygon.pwc"
'    lyrFair = AttachData(filename)
'    DrawLayer lyrFair, 5
'    filename = DefaultPath & "\map\fwy_l_wgs84.pwc"
'    lyrFair = AttachData(filename)
    DrawLayer lyrFair, 4
    With Me.Pwstreet1
      .Value = lyrVFair
      .Picture = "NAME"
      .Action = pwSetAttachedEnumAttribute
    End With
    Me.mnuFairway.Checked = True
  End If
End Sub

Private Sub mnuSaveZoom_Click()
  If mnuStartAnimation.Checked Then
    mnuStopAnimation_Click
  End If
  unloadAllForm
  Dim expFileName As String
  Me.CommonDialog1.filename = ""
  Me.CommonDialog1.DialogTitle = "Save View"
  Me.CommonDialog1.Filter = "View File (*.viw)|*.viw"
  On Error GoTo exitSaveZoom
  Me.CommonDialog1.ShowSave
  If Len(Me.CommonDialog1.filename) = 0 Then
    Exit Sub
  End If
  If Len(Me.CommonDialog1.filename) > 4 Then
    If Format(Right(Me.CommonDialog1.filename, 4), ">") = ".VIW" Then
      expFileName = Me.CommonDialog1.filename
    Else
      expFileName = Me.CommonDialog1.filename & ".VIW"
    End If
  End If
  Dim fs As New FileSystemObject
  Dim f As TextStream
    
  If fs.FileExists(expFileName) Then
    DisableAllForm
    If MsgBox("File exists Overwrite it", vbYesNo) = vbNo Then
      EnableAllForm
      Exit Sub
    End If
    EnableAllForm
  End If
  Set f = fs.CreateTextFile(expFileName, True)
  f.WriteLine MapCenterLat
  f.WriteLine MapCenterLong
  f.WriteLine tmpScaleFactor
  f.Close
  Set fs = Nothing
exitSaveZoom:
End Sub

Private Sub mnuShowAnchorage_Click()
  If Me.mnuShowAnchorage.Checked Then
'    With Me.Pwstreet1
'      .Value = lyrAnor
'      .Picture = ""
'      .Action = pwSetAttachedEnumAttribute
'    End With
'    DrawLayer lyrAnor, 99
    With Me.Pwstreet1
      .Value = lyrAnor
      .Action = pwDetachFile
      .Action = pwDraw
    End With
    mnuShowAnchorage.Checked = False
  Else
    filename = DefaultPath & "\map\anch_a_wgs84.pwc"
    lyrAnor = AttachData(filename)
    DrawLayer lyrAnor, 2
    With Me.Pwstreet1
      .Value = lyrAnor
      .Picture = "OBJNAM"
      .Action = pwSetAttachedEnumAttribute
    End With
    mnuShowAnchorage.Checked = True
  End If
End Sub

Private Sub mnuShowGridlines_Click()
  If Me.mnuShowGridlines.Checked Then
    Me.mnuShowGridlines.Checked = False
    clearGrid
  Else
    Me.mnuShowGridlines.Checked = True
    DrawGrid
  End If
End Sub

Private Sub mnuShowPlaceLabel_Click()
  If mnuShowPlaceLabel.Checked Then
    mnuShowPlaceLabel.Checked = False
    clearPlaceName
  Else
    mnuShowPlaceLabel.Checked = True
    showPlaceName
  End If
End Sub

Private Sub mnuShowTide_Click()
  If Me.mnuShowTide.Checked Then
    Me.mnuShowTide.Checked = False
    setPressButton
    Unload frmTide
  Else
    Me.mnuShowTide.Checked = True
    setPressButton
    frmTide.Show
  End If
End Sub

Private Sub mnuSinglePoint_Click()
  If mnuStartAnimation.Checked Then
    mnuStopAnimation_Click
  End If
  If Me.mnuMeasureDistance.Checked Or strClickPoint = "" Then
    Me.mnuMeasureDistance.Checked = False
    clearPoint
    resumeIndex = 0
    errorFormCall = "frmMain"
    ErrorMsg "Please Select One Control Point"
    strClickPoint = ""
    clearPoint
    With Pwstreet1
      .Feature = pwLeftMouseButton
      .Attribute = pwPlusNoKey
      .MouseBehavior = pwMouseV1LeftButton
    End With
    clearPoint
    clickToolButton = 0
    setPressButton
    Exit Sub
  End If
  DialogPoint.Show
End Sub

Private Sub mnuStartAnimation_Click()
  If Not IsNull(DateCal) Then
    If Not mnuPause.Checked Then
      InitCorner
      InitHour
      If NoControlPointFound Then
        mnuStopAnimation_Click
        Exit Sub
      End If
      Me.StatusBar1.Panels(5).Text = Format(DateCal, "yyyy-mmm-dd hh:mm")
      UpdateStatus
    End If
    clearPoint
    DrawLengendArrow
    If isFormLoad("diaAbout") Then
      Unload diaAbout
    End If
    If isFormLoad("diaDate") Then
      Unload diaDate
    End If
    If isFormLoad("DialogDate") Then
      Unload DialogDate
    End If
    If isFormLoad("DialogPoint") Then
      Unload DialogPoint
    End If
    If isFormLoad("frmErrorMsg") Then
      Unload frmErrorMsg
    End If
    If isFormLoad("frmPrint") Then
      Unload frmPrint
    End If
    Me.mnuStartAnimation.Enabled = False
    Me.mnuStopAnimation.Enabled = True
    Me.mnuPause.Enabled = True
    Me.mnuStepForward.Enabled = False
    Me.mnuStartAnimation.Checked = True
    Me.mnuPause.Checked = False
    setPressButton
    Me.Timer1.Enabled = True
  End If
End Sub

Private Sub mnuStepForward_Click()
  updateTide
End Sub

Private Sub mnuStopAnimation_Click()
  Me.mnuStartAnimation.Checked = False
  Me.mnuPause.Checked = False
  Me.Timer1.Enabled = False
  Me.mnuStartAnimation.Enabled = True
  Me.mnuStopAnimation.Enabled = False
  Me.mnuPause.Enabled = False
  Me.mnuStepForward.Enabled = False
  setPressButton
  clearUpArrow
  
  If strClickPoint <> "" Then
    rst.Open "select x, y from MNXY where mn='" & strClickPoint & "'", conn, adOpenKeyset
    If Not (rst.BOF And rst.EOF) Then
      Dim strLat As Long, strLong As Long
      XY_To_Map rst.Fields(1).Value, rst.Fields(0).Value, strLat, strLong
      frmMain.showPointClick strLat, strLong
    End If
    rst.Close
  End If
End Sub
  
Private Sub mnuPause_Click()
  Timer1.Enabled = False
  Me.mnuStartAnimation.Enabled = True
  Me.mnuStopAnimation.Enabled = True
  Me.mnuPause.Enabled = False
  Me.mnuStepForward.Enabled = True
  Me.mnuStartAnimation.Checked = False
  Me.mnuPause.Checked = True
  setPressButton
End Sub
  
Private Sub mnuTemp_Click()
  frmTemp.Show
End Sub

Private Sub mnuViewOverView_Click()
  If mnuViewOverView.Checked Then
    mnuViewOverView.Checked = False
    setPressButton
    Unload frmOverView
  Else
    mnuViewOverView.Checked = True
    setPressButton
    frmOverView.Show
  End If
End Sub

Private Sub mnuZoomIn_Click()
  If mnuStartAnimation.Checked Then
    mnuStopAnimation_Click
  End If
  With Pwstreet1
    .Latitude = MapCenterLat
    .Longitude = MapCenterLong
    .ScaleFactor = tmpScaleFactor / 2
  End With
End Sub

Private Sub mnuZoomOut_Click()
  If mnuStartAnimation.Checked Then
    mnuStopAnimation_Click
  End If
  With Pwstreet1
    .Latitude = MapCenterLat
    .Longitude = MapCenterLong
    .ScaleFactor = tmpScaleFactor * 2
  End With
End Sub

Private Sub Pwstreet1_AfterZoom()
  InitCorner
  If mnuStartAnimation.Checked Or Me.mnuPause.Checked Then
    mnuStopAnimation_Click
  End If
  DrawScaleBar
End Sub

Private Sub Pwstreet1_EnumerateString(ByVal Element As String, Num As Long, SendMore As Long)
  If mnuStartAnimation.Checked Or mnuPause.Checked Then
    noPoint = False
    Dim i As Integer
    i = Num - 2001
    If Me.cboKnotMS.ListIndex = 0 Then
      MagScaleFactor = 1 / KnotToMS
    Else
      MagScaleFactor = 1
    End If
    On Error Resume Next
      Me.StatusBar1.Panels(6).Text = Round(PointMN(i).Dir) & " deg  " & Round(PointMN(i).Mag * MagScaleFactor, 2) & " " & Me.cboKnotMS.Text
      Me.Pwstreet1.ToolTipText = Round(PointMN(i).Dir) & " deg  " & Round(PointMN(i).Mag * MagScaleFactor, 2) & " " & Me.cboKnotMS.Text
    On Error GoTo 0
  Else
    On Error Resume Next
      'Me.StatusBar1.Panels(6).Text = Round(PointMN(i).Dir) & " deg  " & Round(PointMN(i).Mag * MagScaleFactor, 2) & " " & Me.cboKnotMS.Text
      Me.Pwstreet1.ToolTipText = Element
    On Error GoTo 0
  End If
End Sub

Private Sub Pwstreet1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyAdd Then
    mnuZoomIn_Click
  End If
  If KeyCode = vbKeySubtract Then
    mnuZoomOut_Click
  End If
  If KeyCode = vbKeyF11 Then
    mnuFullExtent_Click
  End If
End Sub

Private Sub Pwstreet1_MapClick(Lat As Long, Lng As Long)
  If Me.mnuMeasureDistance.Checked Then
    With Pwstreet1
      'MsgBox "lat: " & .Latitude & Chr(13) & "lng: " & .Longitude
      If Not MeasureCheck Then
        clearPoint
        .ID = 1500
        MeasureCheck = True
      Else
        .ID = 1501
        MeasureCheck = False
      End If
      .Latitude = Lat
      .Longitude = Lng
      .Text = ""
      .Picture = App.Path & "\pingreen.bmp"
      .Flags = pwVisible Or pwAnimate Or pwBitmap Or pwTransparentTopLeft Or pwFocusBottomLeft Or &H8000
      .Action = pwPointLoad
            

      If .ID = 1500 Then
        lnMeasure.ID = 1502
        lnMeasure.DrawStyle = 0
        lnMeasure.DrawWeight = 2
        lnMeasure.DrawColor = RGB(39, 177, 5)
        lnMeasure.FillStyle = pwSolidFill
        lnMeasure.FillColor = RGB(39, 177, 5)
        lnMeasure.Layer = pwTopLayer
        lnMeasure.Flags = pwPolygon Or pwVisible 'Or pwDrawWeightIsFeet
        lnMeasure.NumPoints = 2
        Me.Pwstreet1.CreateLine lnMeasure.ID, llMeasure(1, 1)
        llMeasure(2, 1) = .Longitude
        llMeasure(1, 1) = .Latitude
        .Action = pwSetFromPoint
      Else
        .Action = pwGetDistance
        lnMeasure.ID = 1502
        llMeasure(2, 2) = .Longitude
        llMeasure(1, 2) = .Latitude
        .SetLine lnMeasure.ID, llMeasure(1, 1)
        unloadAllForm
        MsgBox "Distance = " & Round((.Value) / feetToMeter, 2) & " m"
      End If
    End With
  Else
    If Not NearestPoint(Lng, Lat) Then
      ErrorMsg ("No Nearest Control Point")
      Exit Sub
    End If
    strClickY = LToDMSString2(Lat, 1)
    strClickX = LToDMSString2(Lng, 0)
    showPointClick Lat, Lng
    If isFormLoad("DialogPoint") Then
      DialogPoint.Label7(2) = ClickY & "N"
      DialogPoint.Label7(3) = ClickX & "E"
      DialogPoint.Label7(0) = strClickY
      DialogPoint.Label7(1) = strClickX
      DialogPoint.ClearArray
    End If
  End If
End Sub

Private Sub Pwstreet1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  With Pwstreet1
    .Action = pwMouseToLatLng
    xCurrentLat = .Latitude
    yCurrentLong = .Longitude
'    Me.StatusBar1.Panels(6).Text = ""
    noPoint = True
    .Action = pwEnumPolygonsEnclosingLatLng
    If noPoint Then
      Me.Pwstreet1.ToolTipText = ""
      Me.StatusBar1.Panels(6).Text = ""
'    Else
'      Me.Pwstreet1.ToolTipText = Me.StatusBar1.Panels(6).Text
    End If
    Dim D As Double
    Dim tmpText As String
    Dim m As Double, X1 As Double, Y1 As Double
    Map_To_XY .Latitude, .Longitude, Y1, X1
    
    Me.StatusBar1.Panels(1).Text = "Lat:" & LToDMSString2(.Latitude, 1)
    Me.StatusBar1.Panels(2).Text = "Long:" & LToDMSString2(.Longitude, 0)
'    Me.StatusBar1.Panels(1).Text = .Latitude
'    Me.StatusBar1.Panels(2).Text = .Longitude
    
    X1 = Round(X1, 2)
    Y1 = Round(Y1, 2)
    Me.StatusBar1.Panels(3).Text = "N: " & Y1
    Me.StatusBar1.Panels(4).Text = "E: " & X1
    
    If mnuStartAnimation.Checked = False And mnuPause.Checked = False Then
      .Value = -1 ' the layer to be checked (or -1 to check all layers)
      .Action = pwEnumAttachedPolyEnclosingLatLng
    End If
  End With
End Sub

Private Sub ZoomLayer(Layer As Long)
'  xtop = -196000000
'  xright = xtop
'  xbottom = 196000000
'  xleft = xbottom

  xtop = 22568334
  xleft = 113817109
  xbottom = 22136721
  xright = 114502450

  With Pwstreet1
'    .Select Layer
'    .EndSelect
'    n = Pwstreet1.Value
'    For nrow = 1 To n
'      .Value = Layer
'      .ID = nrow - 1
'      .Action = 998
'      If (.Latitude > xtop) Then xtop = .Latitude
'      If (.Longitude < xleft) Then xleft = .Longitude
'      .Value = Layer
'      .ID = nrow - 1
'      .Action = 999
'      If (.Latitude < xbottom) Then xbottom = .Latitude
'      If (.Longitude > xright) Then xright = .Longitude
'    Next
    .Latitude = xtop
    .Longitude = xleft
    .ID = xbottom
    .Value = xright
    .Action = 997
  End With
End Sub

Private Function AttachData(filename) As Long
  Dim AttachedFileHandle As Long
  With Pwstreet1
    .Value = pwBehindRoads
    .Picture = filename
    .Action = pwAttachFile
    AttachedFileHandle = .Value
  End With
  AttachData = AttachedFileHandle
End Function

Private Sub DrawArrowDB()
'  Me.MousePointer = vbHourglass
  Dim X1 As Double
  Dim Y1 As Double, tLat As Long, tLong As Long
  Dim strSQL As String
  Dim LineID As Integer
  Dim i As Integer
    
  strSQL = "select * from qryDir_" & strSDString & strWeatherString
  'clearUpArrow
  
  Erase llRegion
  With lnRegion
    .ID = 2001
    .DrawStyle = 0
    .DrawWeight = 1
    .DrawColor = pwYellow
    .FillStyle = pwSolidFill
    .FillColor = pwYellow
    .Layer = pwTopLayer
    .Flags = pwPolygon 'Or pwVisible 'Or pwDrawWeightIsFeet
    .NumPoints = 5
  End With
  Me.Pwstreet1.CreateLine lnRegion.ID, llRegion(1, 1)
  
  Erase ll
  With ln
    .ID = 1
    .DrawStyle = 0
    .DrawWeight = 2
    .DrawColor = pwRed
    .FillStyle = pwSolidFill
    .FillColor = pwRed
    .Layer = pwTopLayer
    .Flags = pwPolygon Or pwVisible 'Or pwDrawWeightIsFeet
    .NumPoints = 6
  End With
  Me.Pwstreet1.CreateLine ln.ID, ll(1, 1)
  
  i = 0
  rst.Open strSQL, conn, adOpenKeyset
  LineID = 1
  Do While Not rst.EOF
    X1 = rst.Fields(6).Value
    Y1 = rst.Fields(7).Value
    PointMN(i).MN = rst.Fields(0).Value
    PointMN(i).X = rst.Fields(6).Value
    PointMN(i).Y = rst.Fields(7).Value
    PointMN(i).Mag = rst.Fields(3)
    PointMN(i).Dir = edwinDir(rst.Fields(4), rst.Fields(5))
    'HKGEO 2, y1, x1, tLat, tLong
    XY_To_Map Y1, X1, tLat, tLong
    DrawArrow LineID, tLong, tLat, PointMN(i).Mag, PointMN(i).Dir
    i = i + 1
    LineID = LineID + 1
    rst.MoveNext
  Loop
  
  numOfArrow = LineID - 1
  recordCnt = i
  rst.Close
' DrawLengendArrow
' Me.MousePointer = vbDefault
End Sub

Private Sub DrawArrow(ArrowID As Integer, X1 As Long, Y1 As Long, MagA As Double, DirA As Double)
  Dim i As Integer, j As Integer
  Dim X2 As Long, Y2 As Long, X3 As Long, Y3 As Long
  Dim X4 As Long, Y4 As Long
  Dim XX1 As Long, YY1 As Long, XX2 As Long, YY2 As Long
  Dim XX3 As Long, YY3 As Long, XX4 As Long, YY4 As Long
  Dim IncFactor As Double
  
  ln.ID = ArrowID
  lnRegion.ID = ArrowID + 2000
  Dim zero As Long
  zero = 0
  With Pwstreet1
    i = .GetLine(ln.ID, zero, zero)
    j = .GetLine(lnRegion.ID, zero, zero)

    With ln
      .NumPoints = 6
    End With
    With lnRegion
      .NumPoints = 5
    End With
    
    
    ll(1, 1) = Y1
    ll(2, 1) = X1
        
'LegendWidth * ScaleFactorLX / 30
        
    If MagA < 0.5 Then
      IncFactor = 0.5
    Else
      IncFactor = 0.2
    End If
    
    XX1 = X1 - (MagA + IncFactor) * LegendWidth * Sin(DirA * piValue / 180) * ScaleFactorLX
    YY1 = Y1 - (MagA + IncFactor) * LegendWidth * Cos(DirA * piValue / 180) * ScaleFactorLX
        
    X2 = X1 + MagA * LegendWidth * Sin(DirA * piValue / 180) * ScaleFactorLX
    Y2 = Y1 + MagA * LegendWidth * Cos(DirA * piValue / 180) * ScaleFactorLX
    
    XX2 = X1 + (MagA + IncFactor) * LegendWidth * Sin(DirA * piValue / 180) * ScaleFactorLX
    YY2 = Y1 + (MagA + IncFactor) * LegendWidth * Cos(DirA * piValue / 180) * ScaleFactorLX
    
    ll(1, 2) = Y2
    ll(2, 2) = X2
        
        
    X3 = X2 - MagA * LegendWidth / 4 * Cos(-(-135 - 270 + DirA) * piValue / 180) * ScaleFactorLX
    Y3 = Y2 - MagA * LegendWidth / 4 * Sin(-(-135 - 270 + DirA) * piValue / 180) * ScaleFactorLX
    
    XX3 = X2 - (MagA + IncFactor) * LegendWidth / 4 * Cos(-(-135 - 270 + DirA) * piValue / 180) * ScaleFactorLX
    YY3 = Y2 - (MagA + IncFactor) * LegendWidth / 4 * Sin(-(-135 - 270 + DirA) * piValue / 180) * ScaleFactorLX
    
    ll(1, 3) = Y3
    ll(2, 3) = X3
        
    ll(1, 4) = Y2
    ll(2, 4) = X2
    
    X4 = X2 - MagA * LegendWidth / 4 * Cos(-(135 - 270 + DirA) * piValue / 180) * ScaleFactorLX
    Y4 = Y2 - MagA * LegendWidth / 4 * Sin(-(135 - 270 + DirA) * piValue / 180) * ScaleFactorLX
    
    XX4 = X2 - (MagA + IncFactor) * LegendWidth / 4 * Cos(-(135 - 270 + DirA) * piValue / 180) * ScaleFactorLX
    YY4 = Y2 - (MagA + IncFactor) * LegendWidth / 4 * Sin(-(135 - 270 + DirA) * piValue / 180) * ScaleFactorLX
    
    ll(1, 5) = Y4
    ll(2, 5) = X4
       
    ll(1, 6) = Y2
    ll(2, 6) = X2
        
    .CreateLine ln.ID, ll(1, 1)
    .SetLine ln.ID, ll(1, 1)
    
    X1 = X1 - MagA * LegendWidth * ScaleFactorLX / 30 * Sin(DirA * piValue / 180) * ScaleFactorLX
    Y1 = Y1 - MagA * LegendWidth * ScaleFactorLX / 30 * Cos(DirA * piValue / 180) * ScaleFactorLX
    
    llRegion(1, 1) = YY1
    llRegion(2, 1) = XX1
        
    llRegion(1, 2) = YY3
    llRegion(2, 2) = XX3
                
    llRegion(1, 3) = YY2
    llRegion(2, 3) = XX2
    
    llRegion(1, 4) = YY4
    llRegion(2, 4) = XX4
    
    llRegion(1, 5) = YY1
    llRegion(2, 5) = XX1
    
    .CreateLine lnRegion.ID, llRegion(1, 1)
    .SetLine lnRegion.ID, llRegion(1, 1)
  End With
End Sub

Private Sub centerMap()
  With Pwstreet1
    .Top = 10 + Me.Toolbar1.Height
    .Left = 10
    Dim tmpWidth As Long
    If Me.ScaleWidth > 20 Then
      .Width = Me.ScaleWidth - 20
    Else
      .Width = 0
    End If
    If Me.ScaleHeight > (20 + Me.Toolbar1.Height + Me.StatusBar1.Height) Then
      .Height = Me.ScaleHeight - (20 + Me.Toolbar1.Height + Me.StatusBar1.Height)
    Else
      .Height = 0
    End If
    Me.Picture1.Top = .Top
    Me.Picture1.Left = .Left
    Me.Picture1.Width = .Width
    Me.Picture1.Height = .Height
  End With
End Sub

Private Sub DrawLengendArrow()
  Dim i As Integer
  Dim zero As Long
  zero = 0
  
  lnLegend.DrawStyle = 0
  lnLegend.DrawWeight = 2
  lnLegend.DrawColor = pwRed
  lnLegend.FillStyle = pwSolidFill
  lnLegend.FillColor = pwRed
  lnLegend.Layer = pwTopLayer
  lnLegend.Flags = pwPolygon Or pwVisible 'Or pwDrawWeightIsFeet
  lnLegend.NumPoints = 5
  
  Dim tmpScaleMagFact As Double
  Dim tmpLargeFactor As Double
  Dim txtLegendText As String
  
  If Me.cboKnotMS.ListIndex = 0 Then
    If LegendWidth < 300 Then
      txtLegendText = "2 Knot  " & Chr(13)
      tmpLargeFactor = LegendWidth
      LegendWidth = LegendWidth * 2
    Else
      txtLegendText = "1 Knot  " & Chr(13)
      tmpLargeFactor = 1
      tmpLargeFactor = LegendWidth
    End If
    tmpScaleMagFact = KnotToMS
  Else
    If LegendWidth < 300 Then
      txtLegendText = "2 m/s" & Chr(13)
      tmpLargeFactor = LegendWidth
      LegendWidth = LegendWidth * 2
    Else
      txtLegendText = "1 m/s" & Chr(13)
      tmpLargeFactor = LegendWidth
    End If
    tmpScaleMagFact = 1
  End If
  
  If Me.Pwstreet1.Height > 2000 And Me.Pwstreet1.Width > 3000 Then
    lnLegend.ID = 1600
    i = Me.Pwstreet1.GetLine(lnLegend.ID, zero, zero)
    With lnLegend
      .NumPoints = 3
      .FillStyle = 0
      .FillColor = pwRed
    End With
        
    With Me.Pwstreet1
      'Arrow Tail
      llg(2, 1) = (.Width - LegendLeft) * ScaleFactorLX + BoundLeft
      llg(1, 1) = LegendHeight * ScaleFactorLY + BoundBottom
      llg(2, 2) = (.Width - LegendLeft + LegendWidth * tmpScaleMagFact) * ScaleFactorLX + BoundLeft
      llg(1, 2) = LegendHeight * ScaleFactorLY + BoundBottom
      llg(2, 3) = (.Width - LegendLeft) * ScaleFactorLX + BoundLeft
      llg(1, 3) = LegendHeight * ScaleFactorLY + BoundBottom
      .CreateLine lnLegend.ID, llg(1, 1)
      .SetLine lnLegend.ID, llg(1, 1)
          
      'Arrow Head
      lnLegend.ID = 1601
      With lnLegend
        .NumPoints = 5
        .FillStyle = 0
        .FillColor = pwRed
      End With
           
      llg(2, 1) = (.Width - LegendLeft + LegendWidth * tmpScaleMagFact - LegendWidth * tmpScaleMagFact * 0.4 * Cos(45 * CDbl(piValue) / 180)) * ScaleFactorLX + BoundLeft
      llg(1, 1) = (LegendHeight - LegendWidth * tmpScaleMagFact * 0.4 * Sin(45 * CDbl(piValue) / 180)) * ScaleFactorLY + BoundBottom
      
      llg(2, 2) = (.Width - LegendLeft + LegendWidth * tmpScaleMagFact) * ScaleFactorLX + BoundLeft
      llg(1, 2) = LegendHeight * ScaleFactorLY + BoundBottom
      
      llg(2, 3) = (.Width - LegendLeft + LegendWidth * tmpScaleMagFact - LegendWidth * tmpScaleMagFact * 0.4 * Cos(45 * CDbl(piValue) / 180)) * ScaleFactorLX + BoundLeft
      llg(1, 3) = (LegendHeight + LegendWidth * tmpScaleMagFact * 0.4 * Sin(45 * CDbl(piValue) / 180)) * ScaleFactorLY + BoundBottom
      
      llg(2, 4) = (.Width - LegendLeft + LegendWidth * tmpScaleMagFact) * ScaleFactorLX + BoundLeft
      llg(1, 4) = LegendHeight * ScaleFactorLY + BoundBottom
      
      llg(2, 5) = (.Width - LegendLeft + LegendWidth * tmpScaleMagFact - LegendWidth * tmpScaleMagFact * 0.4 * Cos(45 * CDbl(piValue) / 180)) * ScaleFactorLX + BoundLeft
      llg(1, 5) = (LegendHeight - LegendWidth * tmpScaleMagFact * 0.4 * Sin(45 * CDbl(piValue) / 180)) * ScaleFactorLY + BoundBottom
      
      .CreateLine lnLegend.ID, llg(1, 1)
      .SetLine lnLegend.ID, llg(1, 1)
      .Refresh
      .Action = pwDraw
    End With
    LegendWidth = tmpLargeFactor
    
    With Me.Pwstreet1
      .Feature = pwUserPoint
      .Attribute = pwLabelMinFont
      .Value = 10
      .Action = pwSetConfig
      .Feature = pwUserPoint
      .Attribute = pwLabelMaxFont
      .Value = 12
      .Action = pwSetConfig
      .Picture = AppPath & "\spacer.bmp"
      .Flags = pwVisible Or pwAnimate Or pwBitmap Or pwTransparentTopLeft Or pwFocusBottomLeft Or pwNoAutoWrap
      .ID = 1101
      .Latitude = (LegendHeight + 500) * ScaleFactorLY + BoundBottom
      .Longitude = (.Width - LegendLeft) * ScaleFactorLX + BoundLeft
      .Text = txtLegendText
      .Action = pwPointLoad
    End With
  End If
End Sub

Private Sub InitHour()
  Dim strSQL As String
  Dim strSQL2 As String
  strSQL = "delete from edwinTempCal"
  conn.Execute strSQL
  Dim X As Double, Y As Double, m As Integer, n As Integer
  Dim MinWidthX As Double, MinWidthY As Double
  
  Screen.MousePointer = vbHourglass
  
  'Calculate MinWidth First
  MinWidthX = ScaleFactorX * LegendWidth * 2
  MinWidthY = ScaleFactorY * LegendWidth * 2
      
  'Calculate Hour count
  InitMonthCal
  
'***** Third Method Start ********
  Dim strTempWhere As String
  Dim tmpTopRightFlag As Boolean
  Dim loopCnt As Integer
  Dim TmpFactorVal As Integer
  Dim strTempPoint As String
  Dim TopLeftTop_m As Integer
  
  strTempWhere = "where (x between " & TopX & " and " & BottomX & ") and (y between " & BottomY & " and " & TopY & ")"
  strSQL = "select count(*) from MNXY " & strTempWhere
  rst.Open strSQL, conn, adOpenKeyset
  NoControlPointFound = False
  
  If rst.Fields(0).Value = 0 Then
    NoControlPointFound = True
    ErrorMsg "Control Point Not found in this range"
    rst.Close
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  rst.Close
    
  TmpFactorVal = 1 + Round((Me.Pwstreet1.ScaleFactor - 12000) / ((249989 - 12000) / 5))
  varAddFactor = TmpFactorVal + Round((6 - TmpFactorVal) / 20 * (-SliderBarValue + 10))
  
  strSQL = "select count(*) from MNXY " & strTempWhere & " and (m between 49 and 61) and (n>=90)"
  rst.Open strSQL, conn, adOpenKeyset
  tmpTopRightFlag = False
    
  If rst.Fields(0).Value <> 0 Then
  'If Not (rst.BOF And rst.EOF) Then
    tmpTopRightFlag = True
    rst.Close
    strSQL = "select top 1 * from MNXY " & strTempWhere & " and (m between 49 and 61) order by n desc"
    rst.Open strSQL, conn, adOpenKeyset
    X = rst.Fields(2).Value
    Y = rst.Fields(3).Value
    m = rst.Fields(0).Value
    n = rst.Fields(1).Value
    TopLeftTop_m = m
    rst.Close
    Do
      strSQL = "select count(*) from MNXY " & strTempWhere & " and m>=" & m
      rst.Open strSQL, conn, adOpenKeyset
      loopCnt = rst.Fields(0).Value
      rst.Close
      If loopCnt <> 0 Then
        'Loop for n or m
        strSQL = "select mn, n from MNXY " & strTempWhere & " and m=" & m & " order by n desc"
        rst.Open strSQL, conn, adOpenKeyset
        If Not (rst.BOF And rst.EOF) Then
          Do While Not rst.EOF
            If (Abs(rst.Fields(1).Value - n)) Mod varAddFactor = 0 Then
              strTempPoint = rst.Fields(0).Value
              strSQL2 = "insert into edwinTempCal values ('" & strTempPoint & "','" & InitMonthDate & "'," & HourCount & ")"
              conn.Execute strSQL2
            End If
            rst.MoveNext
          Loop
        End If
        rst.Close
      End If
      m = m - varAddFactor
    Loop While m >= 49
    m = TopLeftTop_m + varAddFactor
  Else
    rst.Close
    strSQL = "select min(m) from MNXY " & strTempWhere
    rst.Open strSQL, conn, adOpenKeyset
    m = rst.Fields(0).Value
    rst.Close
    strSQL = "select top 1 * from MNXY " & strTempWhere & " and m>=" & m & " and m<=" & (m + 2) & " order by y desc"
    rst.Open strSQL, conn, adOpenKeyset
    X = rst.Fields(2).Value
    Y = rst.Fields(3).Value
    m = rst.Fields(0).Value
    n = rst.Fields(1).Value
    rst.Close
  End If
        
  Do
    strSQL = "select count(*) from MNXY " & strTempWhere & " and m>=" & m
    rst.Open strSQL, conn, adOpenKeyset
    loopCnt = rst.Fields(0).Value
    rst.Close
    If loopCnt <> 0 Then
      'Loop for n or m
      strSQL = "select mn, n from MNXY " & strTempWhere & " and m=" & m & " order by n desc"
      rst.Open strSQL, conn, adOpenKeyset
      If Not (rst.BOF And rst.EOF) Then
        Do While Not rst.EOF
          If (Abs(rst.Fields(1).Value - n)) Mod varAddFactor = 0 Then
            strTempPoint = rst.Fields(0).Value
            strSQL2 = "insert into edwinTempCal values ('" & strTempPoint & "','" & InitMonthDate & "'," & HourCount & ")"
            conn.Execute strSQL2
          End If
          rst.MoveNext
        Loop
      End If
      rst.Close
    End If
    m = m + varAddFactor
  Loop While loopCnt <> 0
'***** Third Method End ********
  Screen.MousePointer = vbDefault
  DrawArrowDB
End Sub

Private Sub UpdateDateCal()
  currentMonth = Month(DateCal)
  DateCal = DateAdd("n", Round(timeInterVal * 60), DateCal)
'  If CheckDate(DateCal) Then
'    MsgBox "Date out of range from " & startLimitYear & " to " & endLimitYear
'    Me.Timer1.Enabled = False
'    DateCal = CDate(Year(Date) & "-" & Month(Date) & "-" & Day(Date))
'    Exit Sub
'  End If
  If Month(DateCal) <> currentMonth Then
    InitMonthCal
    Dim strSQL As String
    strSQL = "update edwinTempCal Set [DateCal]='" & InitMonthDate & "'"
    conn.Execute strSQL
  End If
  Me.StatusBar1.Panels(5).Text = Format(DateCal, "yyyy-mmm-dd hh:mm")
End Sub

Private Sub Timer1_Timer()
  updateTide
End Sub

Private Sub updateTide()
  Dim strSQL As String
  strSQL = "update edwinTempCal Set [Time] = [Time] + " & timeInterVal
  conn.Execute strSQL
  UpdateDateCal
  DateWeatherDefault
  DrawArrowDB
  UpdateStatus
  If isFormLoad("frmTide") Then
    frmTide.RefreshGraph
  End If
End Sub
Private Sub ShowInPoint(X As Double, Y As Double)
  Dim i As Integer
  Dim noFlag As Boolean
  noFlag = True
  Me.StatusBar1.Panels(6).Text = ""
  For i = 0 To recordCnt - 1
    If X > PointMN(i).X - LegendWidth * ScaleFactorLX / 30 And _
    X < PointMN(i).X + LegendWidth * ScaleFactorLX / 30 And _
    Y > PointMN(i).Y - LegendWidth * ScaleFactorLX / 30 And _
    Y < PointMN(i).Y + LegendWidth * ScaleFactorLX / 30 Then
      'Me.StatusBar1.Panels(6).Text = PointMN(i).MN & " " & Round(PointMN(i).Mag, 2) & " " & Round(PointMN(i).Dir)
      If Me.cboKnotMS.ListIndex = 0 Then
        MagScaleFactor = 1 / KnotToMS
      Else
        MagScaleFactor = 1
      End If
      Me.StatusBar1.Panels(6).Text = Round(PointMN(i).Dir) & " deg  " & Round(PointMN(i).Mag * MagScaleFactor, 2) & " " & Me.cboKnotMS.Text
      Me.Pwstreet1.ToolTipText = Round(PointMN(i).Dir) & " deg  " & Round(PointMN(i).Mag * MagScaleFactor, 2) & " " & Me.cboKnotMS.Text
      noFlag = False
      Exit For
    End If
  Next
  If noFlag Then
    Me.Pwstreet1.ToolTipText = ""
  End If
End Sub

Public Sub Map_To_XY(Y1 As Long, X1 As Long, Y As Double, X As Double)
  GEOHK 2, Y1 / 1000000, X1 / 1000000, Y, X
  Y = Round(Y, 2)
  X = Round(X, 2)
End Sub



Public Sub clearUpArrow()
  Dim i As Integer
  With Pwstreet1
    For i = 1 To numOfArrow
      ln.ID = i
      .ID = ln.ID
      .Action = pwDestroyLine
      ln.NumPoints = 0
      lnRegion.ID = i + 2000
      .ID = lnRegion.ID
      .Action = pwDestroyLine
      lnRegion.NumPoints = 0
    Next
    lnLegend.ID = 1600
    .ID = lnLegend.ID
    .Action = pwDestroyLine
    lnLegend.NumPoints = 0
    lnLegend.ID = 1601
    .ID = lnLegend.ID
    .Action = pwDestroyLine
    lnLegend.NumPoints = 0
    .ID = 1101
    .Action = pwPointDelete
    .Action = pwDraw
    .Refresh
  End With
End Sub

Private Sub SetConfig(f As Long, a As Long, v As Long)
    Pwstreet1.Feature = f
    Pwstreet1.Attribute = a
    Pwstreet1.Value = v
    Pwstreet1.Action = pwSetConfig
End Sub

Private Sub ShowLayer(lyr As Long, Dval As Long, showHide As Boolean)
  With Me.Pwstreet1
    .Value = lyr
    .DValue = Dval
    If showHide Then
      .Action = pwSetLayerDrawOn
    Else
      .Action = pwSetLayerDrawOff
    End If
    .Action = pwDraw
    .Refresh
  End With
End Sub

Private Sub DrawGrid()
  Dim runLat As Long, runLong As Long, runMin As Long
  Dim tmpY1 As Long, tmpY2 As Long, tmpX1 As Long, tmpX2 As Long, i As Long
  Dim zero As Long
  
  zero = 0
  tmpY1 = DMSToL("22", "5", "0")
  tmpY2 = DMSToL("22", "30", "0")
    
  Erase llGrid
  lnGrid.ID = 1001
  lnGrid.DrawStyle = 0
  lnGrid.DrawWeight = 1
  lnGrid.DrawColor = RGB(185, 159, 253)
  lnGrid.FillStyle = pwSolidFill
  lnGrid.FillColor = RGB(185, 159, 253)
  lnGrid.Layer = pwTopLayer
  lnGrid.Flags = pwPolygon Or pwVisible 'Or pwDrawWeightIsFeet
  lnGrid.NumPoints = 3
  Me.Pwstreet1.CreateLine lnGrid.ID, llGrid(1, 1)
  
  'With lnGrid
    '.ID = 1
  'End With
  'Me.Pwstreet1.SetLine lnGrid.ID, llGrid(1, 1)

  i = 1001
  runLat = 22
  runMin = 50
  runLong = 113
  Do While (runLong < 114) Or (runMin <= 30)
    If runMin = 60 Then
      runMin = 0
      runLong = runLong + 1
    End If
    tmpX1 = DMSToL(CStr(runLong), CStr(runMin), "0")
    lnGrid.ID = i
    Pwstreet1.GetLine lnGrid.ID, zero, zero
    lnGrid.Caption = Str(runLong) & "-" & Format(runMin, "00") & "'E"
    llGrid(1, 1) = tmpY1
    llGrid(2, 1) = tmpX1
    llGrid(1, 2) = tmpY2
    llGrid(2, 2) = tmpX1
    llGrid(1, 3) = tmpY1
    llGrid(2, 3) = tmpX1
    Me.Pwstreet1.CreateLine lnGrid.ID, llGrid(1, 1)
    Me.Pwstreet1.SetLine lnGrid.ID, llGrid(1, 1)
    i = i + 1
    runMin = runMin + 5
  Loop
  
  tmpY1 = DMSToL("22", "35", "0")
  tmpY2 = DMSToL("22", "30", "0")
  runLat = 22
  runMin = 50
  runLong = 113
  Do While (runLong < 114) Or (runMin <= 30)
    If runMin = 60 Then
      runMin = 0
      runLong = runLong + 1
    End If
    tmpX1 = DMSToL(CStr(runLong), CStr(runMin), "0")
    lnGrid.ID = i
    Pwstreet1.GetLine lnGrid.ID, zero, zero
    lnGrid.Caption = ""
    llGrid(1, 1) = tmpY1
    llGrid(2, 1) = tmpX1
    llGrid(1, 2) = tmpY2
    llGrid(2, 2) = tmpX1
    llGrid(1, 3) = tmpY1
    llGrid(2, 3) = tmpX1
    Me.Pwstreet1.CreateLine lnGrid.ID, llGrid(1, 1)
    Me.Pwstreet1.SetLine lnGrid.ID, llGrid(1, 1)
    i = i + 1
    runMin = runMin + 5
  Loop
  
  tmpX1 = DMSToL("113", "50", "0")
  tmpX2 = DMSToL("114", "25", "0")
  runMin = 5
  Do While (runMin <= 35)
    tmpY1 = DMSToL(CStr(runLat), CStr(runMin), "0")
    lnGrid.ID = i
    lnGrid.Caption = Str(runLat) & "-" & Format(runMin, "00") & "'N"
    llGrid(1, 1) = tmpY1
    llGrid(2, 1) = tmpX1
    llGrid(1, 2) = tmpY1
    llGrid(2, 2) = tmpX2
    llGrid(1, 3) = tmpY1
    llGrid(2, 3) = tmpX1
    Me.Pwstreet1.CreateLine lnGrid.ID, llGrid(1, 1)
    Me.Pwstreet1.SetLine lnGrid.ID, llGrid(1, 1)
    runMin = runMin + 5
    i = i + 1
  Loop
  
  tmpX1 = DMSToL("114", "30", "0")
  tmpX2 = DMSToL("114", "25", "0")
  runMin = 5
  Do While (runMin <= 35)
    tmpY1 = DMSToL(CStr(runLat), CStr(runMin), "0")
    lnGrid.ID = i
    lnGrid.Caption = ""
    llGrid(1, 1) = tmpY1
    llGrid(2, 1) = tmpX1
    llGrid(1, 2) = tmpY1
    llGrid(2, 2) = tmpX2
    llGrid(1, 3) = tmpY1
    llGrid(2, 3) = tmpX1
    Me.Pwstreet1.CreateLine lnGrid.ID, llGrid(1, 1)
    Me.Pwstreet1.SetLine lnGrid.ID, llGrid(1, 1)
    runMin = runMin + 5
    i = i + 1
  Loop
  numOfGrid = i - 1
  Me.Pwstreet1.Refresh
  Me.Pwstreet1.Action = pwDraw
End Sub

Private Sub clearGrid()
  Dim i As Long
  With Pwstreet1
    For i = 1 To numOfGrid
      lnGrid.ID = i + 1000
      .ID = lnGrid.ID
      .Action = pwDestroyLine
      lnGrid.NumPoints = 0
    Next
    .Action = pwDraw
  End With
End Sub

Private Sub renderLabel(lyr As Long, attr As String)
  With Me.Pwstreet1
    .Value = lyr
    .Picture = attr
    .Action = pwSetAttachedLabelAttribute
  End With
End Sub

Private Sub InitCorner()
  With Pwstreet1
    .Action = pwGetTopLeft
    BoundTop = .Latitude
    BoundLeft = .Longitude
    Map_To_XY .Latitude, .Longitude, TopY, TopX
    
    .Action = pwGetBottomRight
    BoundBottom = .Latitude
    BoundRight = .Longitude
    Map_To_XY .Latitude, .Longitude, BottomY, BottomX
    
    ScaleWidihZoom = -BoundLeft + BoundRight
    ScaleHeightZoom = BoundTop - BoundBottom
    .Action = pwGetMapCenter
    MapCenterLat = .Latitude
    MapCenterLong = .Longitude
    tmpScaleFactor = .ScaleFactor
    
    ScaleFactorX = (BottomX - TopX) / Me.Pwstreet1.Width
    ScaleFactorY = (TopY - BottomY) / Me.Pwstreet1.Height
    
    ScaleFactorLX = (BoundRight - BoundLeft) / Me.Pwstreet1.Width
    ScaleFactorLY = (BoundTop - BoundBottom) / Me.Pwstreet1.Width
    
    If isFormLoad("frmOverView") Then
      frmOverView.DrawRect
    End If
    'Me.StatusBar1.Panels(8).Text = .ScaleFactor
  End With
End Sub

Private Sub UpdateStatus()
  Dim tmpStr As String
  If strWeatherString = "D" Then
    tmpStr = "Dry "
  Else
    tmpStr = "Wet "
  End If
  If strSDString = "S" Then
    tmpStr = tmpStr & "/ Surface"
  Else
    tmpStr = tmpStr & "/ Depth"
  End If
  tmpStr = tmpStr & " / " & Round(timeInterVal * 60, 0) & "min"
  Me.StatusBar1.Panels(7).Text = tmpStr
  With Me.Pwstreet1
    .Latitude = xCurrentLat
    .Longitude = yCurrentLong
    Me.StatusBar1.Panels(6).Text = ""
    Me.Pwstreet1.ToolTipText = ""
    .Action = pwEnumPolygonsEnclosingLatLng
  End With
End Sub

Private Sub clearPoint()
  With Me.Pwstreet1
    .ID = 1500
    .Action = pwPointDelete
    .ID = 1501
    .Action = pwPointDelete
    
    Erase llMeasure
    .ID = 1502
    .Action = pwDestroyLine
    lnMeasure.ID = 1502
    lnMeasure.NumPoints = 0
    .Action = pwDraw
  End With
End Sub

Public Function NearestPoint(strLong As Long, strLat As Long) As Boolean
  Dim tmpX1 As Double, tmpY1 As Double, strSQL As String
  Dim TempInitWidth As Long, rctCnt As Integer
  Map_To_XY strLat, strLong, tmpY1, tmpX1
  TempInitWidth = 0
  NearestPoint = False
  Do
    TempInitWidth = TempInitWidth + 1000
    strSQL = "select x, y, mn from MNXY where x between " & (tmpX1 - TempInitWidth) & " and " & (tmpX1 + TempInitWidth) & " and y between " & (tmpY1 - TempInitWidth) & " and " & (tmpY1 + TempInitWidth) & " ORDER by (x-" & tmpX1 & ")^2+(y-" & tmpY1 & ")^2  "
    rst.Open strSQL, conn, adOpenKeyset
    rctCnt = 0
    If Not (rst.BOF And rst.EOF) Then
      rctCnt = rst.RecordCount
      XY_To_Map rst.Fields(1).Value, rst.Fields(0).Value, strLat, strLong
      ClickX = rst.Fields(0).Value
      ClickY = rst.Fields(1).Value
      strClickPoint = rst.Fields(2).Value
      rst.Close
      NearestPoint = True
      Exit Do
    End If
    rst.Close
  Loop While (rctCnt = 0 And TempInitWidth < 10000)
End Function

Private Sub DrawScaleBar()
  Dim tmpWidth As Double
  Dim initWidth As Integer
  Dim newWidth As Integer
  
  Me.Label1(0).Left = 0
  Me.Label1(0).Top = 100
  Me.Label1(1).Top = 100
  Me.Shape1(0).Top = 130
  Me.Shape1(1).Top = 130
  Me.Shape1(0).Left = 130
  Me.Shape1(0).Height = 135
  Me.Shape1(1).Height = 135
  newWidth = 2500
  Do
    initWidth = newWidth
    tmpWidth = Round(Pwstreet1.Width * initWidth * feetToMeter / tmpScaleFactor)
    If initWidth = 2500 Then
      newWidth = 1000
    Else
      newWidth = initWidth / 2
    End If
  Loop While tmpWidth > 1300
  If tmpWidth Mod 2 = 1 Then
    tmpWidth = tmpWidth + 1
  End If
  Me.Shape1(0).Width = tmpWidth
  Me.Shape1(1).Left = 128 + Me.Shape1(0).Width
  Me.Shape1(1).Width = tmpWidth
  Me.Label1(1).Left = 160 + tmpWidth * 2
  If initWidth >= 500 Then
    Me.Label1(1).Caption = initWidth / 500 & "km"
  Else
    Me.Label1(1).Caption = initWidth * 2 & "m"
  End If
End Sub

Private Sub Timer2_Timer()
  Me.Pwstreet1.Action = pwCopyToClipboard
  Picture1 = Clipboard.GetData()
  Dim expFileName As String
  CommonDialog1.CancelError = True
  Me.CommonDialog1.filename = ""
  Me.CommonDialog1.DialogTitle = "Export File"
  Me.CommonDialog1.Filter = "Image File (*.bmp)|*.bmp"
  On Error GoTo exitExportMap
  Me.CommonDialog1.ShowSave
  If Len(Me.CommonDialog1.filename) = 0 Then
    Exit Sub
  End If
  If Len(Me.CommonDialog1.filename) > 4 Then
    If Format(Right(Me.CommonDialog1.filename, 4), ">") = ".BMP" Then
      expFileName = Me.CommonDialog1.filename
    Else
      expFileName = Me.CommonDialog1.filename & ".BMP"
    End If
  End If
  expFileName = Me.CommonDialog1.filename


  Dim fs As New FileSystemObject
  If fs.FileExists(expFileName) Then
    DisableAllForm
    If MsgBox("File exists Overwrite it", vbYesNo) = vbNo Then
      EnableAllForm
      Exit Sub
    End If
    EnableAllForm
  End If
  Set fs = Nothing
  SavePicture Picture1.Image, expFileName
  MsgBox "haha : " & TopY & ", " & TopX & ", " & BottomY & ", " & BottomX
exitExportMap:
  Timer2.Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "zoom":
      With Pwstreet1
        .Feature = pwLeftMouseButton
        .Attribute = pwPlusNoKey
        .MouseBehavior = pwMouseV1LeftButton
      End With
      clearPoint
      mnuMeasureDistance.Checked = False
      clickToolButton = 0
      setPressButton
      mnuStopAnimation_Click
    Case "pan":
      With Pwstreet1
        .Feature = pwLeftMouseButton
        .Attribute = pwPlusNoKey
        .MouseBehavior = pwMousePan
      End With
      clearPoint
      mnuMeasureDistance.Checked = False
      clickToolButton = 1
      setPressButton
      mnuStopAnimation_Click
    Case "stop":
      mnuStopAnimation_Click
    Case "play":
      mnuStartAnimation_Click
    Case "print":
      mnuPrint_Click
    Case "open":
      mnuLoadZoom_Click
    Case "save":
      mnuSaveZoom_Click
    Case "measure":
      mnuMeasureDistance_Click
    Case "setting":
      mnuSinglePoint_Click
    Case "zoomin":
      mnuZoomIn_Click
    Case "zoomout":
      mnuZoomOut_Click
    Case "globe":
      mnuFullExtent_Click
    Case "pause":
      mnuPause_Click
    Case "SettingAn":
      mnuDateStart_Click
    Case "forward":
      mnuStepForward_Click
    Case "ShowTide":
      mnuShowTide_Click
    Case "overview":
      mnuViewOverView_Click
  End Select
End Sub

Private Sub showPlaceName()
  Dim i As Integer
  With Me.Pwstreet1
    .Feature = pwUserPoint
    .Attribute = pwLabelMinFont
    .Value = 10
    .Action = pwSetConfig
    .Feature = pwUserPoint
    .Attribute = pwLabelMaxFont
    .Value = 12
    .Action = pwSetConfig
    .Picture = AppPath & "\spacer.bmp"
    .Flags = pwVisible Or pwAnimate Or pwBitmap Or pwTransparentTopLeft Or pwFocusBottomLeft Or pwNoAutoWrap
    For i = 0 To varPlaceNameCnt
      .ID = 1800 + i
      .Latitude = PlaceNameArr(i).Y
      .Longitude = PlaceNameArr(i).X
      .Text = PlaceNameArr(i).Name
      .Action = pwPointLoad
    Next
  End With
End Sub

Private Sub clearPlaceName()
  Dim i As Integer
  With Pwstreet1
    For i = 0 To varPlaceNameCnt
      .ID = 1800 + i
      .Action = pwPointDelete
    Next
  End With
End Sub

Private Sub endApp()
  unloadAllForm
  End
End Sub

Private Sub unloadAllForm()
  If isFormLoad("diaAbout") Then
    Unload diaAbout
  End If
  If isFormLoad("diaDate") Then
    Unload diaDate
  End If
  If isFormLoad("DialogDate") Then
    Unload DialogDate
  End If
  If isFormLoad("DialogPoint") Then
    Unload DialogPoint
  End If
  If isFormLoad("frmTide") Then
    Unload frmTide
  End If
  If isFormLoad("frmOpening") Then
    Unload frmOpening
  End If
  If isFormLoad("frmErrorMsg") Then
    Unload frmErrorMsg
  End If
  If isFormLoad("frmOverView") Then
    Unload frmOverView
  End If
  If isFormLoad("frmPrint") Then
    Unload frmPrint
  End If
  If isFormLoad("frmErrorMsg") Then
    Unload frmErrorMsg
  End If
End Sub

Public Sub setPressButton()
  Select Case clickToolButton
    Case 0: 'Zoom
      Me.Toolbar1.Buttons(1).Value = tbrPressed
      Me.Toolbar1.Buttons(2).Value = tbrUnpressed
      Me.Toolbar1.Buttons(11).Value = tbrUnpressed
    Case 1: 'Pan
      Me.Toolbar1.Buttons(1).Value = tbrUnpressed
      Me.Toolbar1.Buttons(2).Value = tbrPressed
      Me.Toolbar1.Buttons(11).Value = tbrUnpressed
    Case 2: 'Measure
      Me.Toolbar1.Buttons(1).Value = tbrUnpressed
      Me.Toolbar1.Buttons(2).Value = tbrUnpressed
      Me.Toolbar1.Buttons(11).Value = tbrPressed
  End Select
  
  If mnuShowTide.Checked Then
    Me.Toolbar1.Buttons(14).Value = tbrPressed
  Else
    Me.Toolbar1.Buttons(14).Value = tbrUnpressed
  End If
  If Me.mnuViewOverView.Checked Then
    Me.Toolbar1.Buttons(15).Value = tbrPressed
  Else
    Me.Toolbar1.Buttons(15).Value = tbrUnpressed
  End If
  
  If mnuStartAnimation.Checked Then
    'start Animation
    Me.Toolbar1.Buttons(3).Enabled = False
    Me.Toolbar1.Buttons(4).Enabled = True
    Me.Toolbar1.Buttons(5).Enabled = True
    Me.Toolbar1.Buttons(6).Enabled = False
    Me.Toolbar1.Buttons(3).Value = tbrPressed
    Me.Toolbar1.Buttons(4).Value = tbrUnpressed
  Else
    If Me.mnuPause.Checked Then
      'pause case
      Me.Toolbar1.Buttons(3).Enabled = True
      Me.Toolbar1.Buttons(4).Enabled = False
      Me.Toolbar1.Buttons(5).Enabled = True
      Me.Toolbar1.Buttons(6).Enabled = True
      Me.Toolbar1.Buttons(3).Value = tbrUnpressed
      Me.Toolbar1.Buttons(4).Value = tbrPressed
    Else
      'stop case
      Me.Toolbar1.Buttons(3).Enabled = True
      Me.Toolbar1.Buttons(4).Enabled = False
      Me.Toolbar1.Buttons(5).Enabled = False
      Me.Toolbar1.Buttons(6).Enabled = False
      Me.Toolbar1.Buttons(3).Value = tbrUnpressed
      Me.Toolbar1.Buttons(4).Value = tbrUnpressed
    End If
  End If
  Me.Toolbar1.Refresh
End Sub

Private Sub clearNautical()
  If defaultNautical <> cboKnotMS.ListIndex Then
    mnuStopAnimation_Click
    defaultNautical = cboKnotMS.ListIndex
    If isFormLoad("DialogPoint") Then
      DialogPoint.ClearArray
    End If
  End If
End Sub

Public Sub resumeError()
  Select Case resumeIndex
    Case 0:
      'No focus
  End Select
End Sub

Public Sub showPointClick(Lat As Long, Lng As Long)
  With Pwstreet1
    .ID = 1500
    .Action = pwPointDelete
    .Latitude = Lat
    .Longitude = Lng
    .Text = ""
    .Picture = App.Path & "\pingreen.bmp"
    .Flags = pwVisible Or pwAnimate Or pwBitmap Or pwTransparentTopLeft Or pwFocusBottomLeft Or &H8000
    .Action = pwPointLoad
  End With
End Sub

Private Function checkRegion(X As Double, Y As Double) As Integer
  If X < 823352 And Y > 832000 Then
    checkRegion = 1
    Exit Function
  End If
  If X < 822723 And Y < 832000 Then
    checkRegion = 0
    Exit Function
  End If
  If (X > 822723 And Y > 822917) Or (X > 830894) Then
    checkRegion = 1
    Exit Function
  End If
  checkRegion = 2
End Function


