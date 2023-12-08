VERSION 5.00
Begin VB.Form diaAbout 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Digital Tidal Stream Atlas 2003"
   ClientHeight    =   3930
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5985
   Icon            =   "diaAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   495
      Left            =   2212
      TabIndex        =   6
      Top             =   3400
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3375
      Left            =   105
      TabIndex        =   2
      Top             =   0
      Width           =   5775
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   3720
         Picture         =   "diaAbout.frx":08CA
         ScaleHeight     =   1500
         ScaleWidth      =   1860
         TabIndex        =   4
         Top             =   720
         Width           =   1860
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         Caption         =   "Digital Tidal Stream Atlas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   280
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Height          =   2535
         Left            =   240
         TabIndex        =   3
         Top             =   705
         Width           =   3375
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&System Information"
      Height          =   495
      Left            =   3780
      TabIndex        =   1
      Top             =   3400
      Width           =   1575
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   495
      Left            =   630
      TabIndex        =   0
      Top             =   3400
      Width           =   1575
   End
End
Attribute VB_Name = "diaAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
  Dim SysInfoPath As String
  If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
  ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
    If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
      SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
    Else
      GoTo SysInfoErr
    End If
  Else
      GoTo SysInfoErr
  End If

  Call Shell(SysInfoPath, vbNormalFocus)
  Exit Sub

SysInfoErr:
  MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Private Sub cmdHelp_Click()
  Dim hwndHelp As Long
  hwndHelp = HtmlHelp(hWnd, DefaultPath & "\Tide\DTSA2003.chm", HH_DISPLAY_TOPIC, 0)
  If hwndHelp = 0 Then
    errorFormCall = "diaAbout"
    ErrorMsg "Help file not found !!!"
  End If
End Sub

Private Sub Form_Load()
  Me.Label2 = "Hydrographic Office, " & Chr(13) & Chr(10) & "Marine Department, " & Chr(13) & Chr(10) & "Hong Kong Government SAR, " & Chr(13) & Chr(10) & " Version 1.0" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Copyright(c) Hong Kong Government SAR" & Chr(13) & Chr(10) & "Copyright Reserved" & Chr(13) & Chr(10) & "Project Team" & Chr(13) & Chr(10) & "Project Manager: Michael C.M. Chau" & Chr(13) & Chr(10) & "Software Developer: Edwin Wong"
End Sub

Private Sub OKButton_Click()
  Unload Me
End Sub

Public Sub resumeError()

End Sub
