VERSION 5.00
Object = "{EFAB76C0-9F63-11CF-A48A-A0AC34F4689F}#2.0#0"; "Pwstrv2.ocx"
Begin VB.Form frmOverView 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Overview"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   Icon            =   "frmOverView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin PWSTREETLib.Pwstreet pwOverView 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      _Version        =   131072
      _ExtentX        =   8916
      _ExtentY        =   6588
      _StockProps     =   32
      RightMouseMenu  =   0   'False
      ScrollBars      =   0   'False
   End
End
Attribute VB_Name = "frmOverView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnRect As UserDefinedLine
Dim llRect(1 To 2, 1 To 6) As Long

Public Sub DrawRect()
  Dim i As Long
  Dim zero As Long
  zero = 0
  lnRect.ID = 10
  With pwOverView
    .ID = lnRect.ID
    .Action = pwDestroyLine
    lnRect.NumPoints = 0
    .Action = pwDraw
  End With
  lnRect.DrawStyle = 0
  lnRect.DrawWeight = 1
  lnRect.DrawColor = RGB(255, 0, 0)
  lnRect.FillStyle = pwNoFill
  lnRect.Layer = pwTopLayer
  lnRect.Flags = pwPolygon Or pwVisible
  lnRect.NumPoints = 6
  llRect(1, 1) = BoundTop
  llRect(2, 1) = BoundLeft
  llRect(1, 2) = BoundTop
  llRect(2, 2) = BoundRight
  llRect(1, 3) = BoundBottom
  llRect(2, 3) = BoundRight
  llRect(1, 4) = BoundBottom
  llRect(2, 4) = BoundLeft
  llRect(1, 5) = BoundTop
  llRect(2, 5) = BoundLeft
  llRect(1, 6) = BoundBottom
  llRect(2, 6) = BoundLeft
  pwOverView.CreateLine lnRect.ID, llRect(1, 1)
  pwOverView.SetLine lnRect.ID, llRect(1, 1)
  With pwOverView
    .Action = pwDraw
  End With
End Sub
Private Function AttachData(filename) As Long
  Dim AttachedFileHandle As Long
  With pwOverView
    .Value = pwBehindRoads
    .Picture = filename
    .Action = pwAttachFile
    AttachedFileHandle = .Value
  End With
  AttachData = AttachedFileHandle
End Function

Private Sub DrawLayer(Layer As Long, c1 As Integer)
  Dim n As Long
  With pwOverView
    .Select Layer
    .EndSelect
    n = pwOverView.Value
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
    End Select
    .Action = pwDraw
  End With
End Sub

Private Sub Form_Load()
  SetWindowPos hWnd, -1, 0, 0, 0, 0, 3
  Dim filename As String
  Dim Layer As Long
  
  filename = DefaultPath & "\map\land_a.pwc"
  Layer = AttachData(filename)
  DrawLayer Layer, 3

  filename = DefaultPath & "\map\land_l.pwc"
  Layer = AttachData(filename)
  DrawLayer Layer, 0
  
  centerMap
  With Me.pwOverView
    .Feature = pwRightMouseButton
    .Attribute = pwPlusShiftKey
    .MouseBehavior = pwMouseDisabled
    .Feature = pwMiddleMouseButton
    .Attribute = pwPlusShiftKey
    .MouseBehavior = pwMouseDisabled
    .Feature = pwLeftMouseButton
    .Attribute = pwPlusShiftKey
    .MouseBehavior = pwMouseDisabled
    
    .Feature = pwRightMouseButton
    .Attribute = pwPlusCtrlKey
    .MouseBehavior = pwMouseDisabled
    .Feature = pwMiddleMouseButton
    .Attribute = pwPlusCtrlKey
    .MouseBehavior = pwMouseDisabled
    .Feature = pwLeftMouseButton
    .Attribute = pwPlusCtrlKey
    .MouseBehavior = pwMouseDisabled
  
    .Feature = pwRightMouseButton
    .Attribute = pwPlusNoKey
    .MouseBehavior = pwMouseDisabled
    .Feature = pwMiddleMouseButton
    .Attribute = pwPlusNoKey
    .MouseBehavior = pwMouseDisabled
    .Feature = pwLeftMouseButton
    .Attribute = pwPlusNoKey
    .MouseBehavior = pwMouseDisabled
  
  
    .Feature = pwVoidPolygon
    .Value = RGB(255, 255, 255)
    .Attribute = pwFillColor
    .Action = pwSetConfig
    .Latitude = xOrgtop
    .Longitude = xOrgleft
    .ID = xOrgbottom
    .Value = xOrgright
    .Action = 997
  End With
  DrawRect
End Sub

Private Sub centerMap()
  With pwOverView
    .Top = 10
    .Left = 10
    .Width = Me.ScaleWidth - 20
    .Height = Me.ScaleHeight - 20
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMain.mnuViewOverView.Checked = False
  frmMain.setPressButton
End Sub
