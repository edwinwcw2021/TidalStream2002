VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printing Dialog"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   615
      Left            =   1995
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   615
      Left            =   435
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Orientation"
      Height          =   855
      Left            =   255
      TabIndex        =   2
      Top             =   480
      Width           =   3495
      Begin VB.OptionButton Option1 
         Caption         =   "Portrait"
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Landscape"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Printer"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   135
      Width           =   615
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo errorPrint
  Dim prn As Printer
  Dim strHeader As String
  Set prn = Printers(Me.cboPrinter.ListIndex)
  Set Printer = prn
  If Me.Option1(0).Value Then
    Printer.Orientation = 2
  Else
    Printer.Orientation = 1
  End If
  strHeader = "Digital Tidal Stream Atlas 2003  Printed on:" & Format(Now, "mmm dd,yyyy hh:mm:ss")
  Printer.FontSize = 14
  Printer.FontBold = True
  Printer.Print
  Printer.CurrentX = (Printer.Width - TextWidth(strHeader) * 14 / 8) / 2
  Printer.Print strHeader
  With frmMain.Pwstreet1
    If .ScaleFactor < 32000 Then
      .ScaleFactor = 32000
    End If
    .Action = pwDraw
    .FriendHandle = Printer.hDC
    .Flags = pwSuppressPageErase
    .PrintTop = 0
    .PrintLeft = 0
    .PrintWidth = .Width / .ScaleFactor
    .PrintHeight = .Height / .ScaleFactor
    .Action = pwPrint
  End With
  Printer.EndDoc
  GoTo exitPrint

errorPrint:
  If Err <> 0 Then
    errorFormCall = "frmPrint"
    ErrorMsg "Printing with errors"
  End If

exitPrint:
  Unload Me
End Sub

Private Sub Form_Load()
  SetWindowPos hWnd, -1, 0, 0, 0, 0, 3
  Dim i As Integer
  Dim iDefault As Integer

  On Error GoTo errorPrint
  For i = 0 To Printers.Count - 1
    If Printers(i).TrackDefault Then
      iDefault = i
    End If
    Me.cboPrinter.AddItem Printers(i).DeviceName
  Next
  Me.cboPrinter.ListIndex = iDefault
  Me.Option1(0).Value = True
  
  GoTo exitPrint
errorPrint:
  If Err <> 0 Then
    errorFormCall = "frmPrint"
    ErrorMsg "No Printer found or Printer Error !!"
  End If
exitPrint:
End Sub

Private Sub Form_Unload(Cancel As Integer)
  EnableAllForm
End Sub

Private Sub Option1_Click(Index As Integer)
  Select Case Index
    Case 0:
      Option1(1).Value = False
      Option1(0).Value = True
    Case 1:
      Option1(1).Value = True
      Option1(0).Value = False
  End Select
End Sub

Public Sub resumeError()
  Unload Me
End Sub
