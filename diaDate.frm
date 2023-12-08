VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form diaDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Date"
   ClientHeight    =   2520
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2580
   Icon            =   "diaDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   2580
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4445
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   22806529
      CurrentDate     =   37622
      MaxDate         =   42369
      MinDate         =   34700
   End
End
Attribute VB_Name = "diaDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  SetWindowPos hWnd, -1, 0, 0, 0, 0, 3
  With Me.MonthView1
    If isFormLoad("DialogPoint") Then
      .Month = Month(PointDateCal)
      .Day = Day(PointDateCal)
      .Year = Year(PointDateCal)
    End If
    If isFormLoad("DialogDate") Then
      .Month = Month(DateCal)
      .Day = Day(DateCal)
      .Year = Year(DateCal)
    End If
    Me.Width = .Width + 75
    Me.Height = .Height + 400
  End With
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
  If CheckDate(Me.MonthView1.Value) Then
    If isFormLoad("DialogPoint") Then
      PointDateCal = Format(Me.MonthView1.Value, "yyyy-mm-dd")
      DialogPoint.txtStartDate(0).Text = Format(Me.MonthView1.Value, "yyyy-mm-dd")
    End If
    If isFormLoad("DialogDate") Then
      DateCal = Format(Me.MonthView1.Value, "yyyy-mm-dd")
      DialogDate.txtStartDate(0).Text = Format(Me.MonthView1.Value, "yyyy-mm-dd")
    End If
    Unload Me
  Else
    errorFormCall = "DialogPoint"
    ErrorMsg "Please select date range from 1/1/1995 to 31/12/2015"
    Me.MonthView1.Value = DateCal
  End If
End Sub

Public Sub resumeError()

End Sub
