VERSION 5.00
Begin VB.Form frmTemp 
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmTemp.frx":0000
      Top             =   240
      Width           =   8775
   End
End
Attribute VB_Name = "frmTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Me.Text1.Text = edwinEncryption(Me.Text1.Text)
End Sub

Private Sub Form_Load()
  Dim strSQL As String
  Dim tmpString As String
  Dim i As Integer
  
  strSQL = "select tHeight from Tide_WAG order by tdate"
  rst.Open strSQL, conn, adOpenKeyset
  Dim fs As New FileSystemObject
  Dim f As TextStream
  Set f = fs.CreateTextFile("wag.txt", True)
  i = 0
  tmpString = ""
  Do While Not rst.EOF
    tmpString = tmpString & rst.Fields(0)
    i = i + 1
    If i = 24 Then
      f.WriteLine edwinEncryption(tmpString)
      'f.WriteLine tmpString
      tmpString = ""
      i = 0
    Else
      tmpString = tmpString & ","
    End If
    rst.MoveNext
  Loop
  f.WriteLine edwinEncryption(Left(tmpString, Len(tmpString) - 1))
  'f.WriteLine Left(tmpString, Len(tmpString) - 1)
  f.Close
  Set fs = Nothing
  rst.Close
End Sub

