VERSION 5.00
Begin VB.Form frmErrorMsg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Error Message"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   380
      Left            =   2040
      TabIndex        =   1
      Top             =   850
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      Picture         =   "frmErrorMsg.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmErrorMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Dim FormErr As Form
  For Each FormErr In Forms
    FormErr.Enabled = True
  Next
  For Each FormErr In Forms
    If FormErr.Name = errorFormCall Then
      FormErr.SetFocus
      FormErr.resumeError
      Exit For
    End If
  Next
  Unload Me
End Sub

Private Sub Form_Activate()
  Me.Command1.SetFocus
End Sub

Private Sub Form_Load()
  SetWindowPos hWnd, -1, 0, 0, 0, 0, 3
  Me.Top = (Screen.Height - Me.Height) / 2
  Me.Left = (Screen.Width - Me.Width) / 2
  'Me.Command1.SetFocus
End Sub
