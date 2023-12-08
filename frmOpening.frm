VERSION 5.00
Begin VB.Form frmOpening 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5400
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOpening.frx":0000
   ScaleHeight     =   5400
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cl&ose"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   4320
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   2280
      Top             =   1920
   End
End
Attribute VB_Name = "frmOpening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Form_Load()
'  Me.Width = Me.Picture.Width
'  Me.Height = Me.Picture.Height
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2
  frmMain.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
'  frmMain.Show
End Sub

Private Sub Timer1_Timer()
  Unload Me
End Sub
