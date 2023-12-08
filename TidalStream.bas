Attribute VB_Name = "TidalStream"
Option Explicit

Public BoundTop As Long
Public BoundLeft As Long
Public BoundBottom As Long
Public BoundRight As Long

Public defaultNautical As Integer
Public DefaultPath As String
Public Const feetToMeter = 3.28083989501
'Public Const LimitedYear = 2003
Public Const KnotToMS = 0.51444

Public conn As New ADODB.Connection
Public rst As New ADODB.Recordset
Public startLimitYear As Date
Public endLimitYear As Date
Public DateCal As Date
Public PointDateCal As Date
Public HourCount As Double
Public PointHourCount As Double
Public timeInterVal As Double
Public timePointInterVal As Double
Public duration As Integer
Public strClickPoint As String
Public ClickX As String
Public ClickY As String
Public strClickX As String
Public strClickY As String
Public currentMonth As Integer
Public InitMonthDate As Date
Public InitPointMonthDate As Date
Public PlaceNameArr(0 To 100) As PlaceName
Public varPlaceNameCnt As Integer
Public errorFormCall As String

Public varAddFactor As Integer
Public varAddFactorC As Integer
Public SliderBarValue As Integer

Public xOrgtop As Long
Public xOrgleft As Long
Public xOrgbottom As Long
Public xOrgright As Long

Public MagScaleFactor As Double
Public Type PlaceName
  Name As String
  X As Long
  Y As Long
  group As String
End Type
Public Type CurrentRecord
  MN As String
  X As Double
  Y As Double
  Mag As Double
  Dir As Double
End Type
Public PointMN(1000) As CurrentRecord
Public recordCnt As Integer
Type UserDefinedLine
  ID As Long
  DrawColor As Long
  LabelColor As Long
  FillColor As Long
  NumPoints As Integer
  Flags As Integer
  Layer As Integer
  DrawStyle As Integer
  DrawWeight As Integer
  FillStyle As Integer
  Caption As String * 32
End Type
Global Const pwBlack = &H0&
Global Const pwBlue = &HFF0000
Global Const pwGreen = &H7F00&
Global Const pwCyan = &HFFFF00
Global Const pwRed = &HFF&
Global Const pwMagenta = &HFF00FF
Global Const pwBrown = &H4080&
Global Const pwYellow = &HFFFF&
Global Const pwLightYellow = &H9FFFFF
Global Const pwWhite = &HFFFFFF
Global Const pwVeryLightGray = &HC0C0C0
Global Const pwLightGray = &H7F7F7F
Global Const pwGray = &H3F3F3F

'For System Info 32
Const KEY_ALL_ACCESS = &H2003F
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const ERROR_SUCCESS = 0
Global Const REG_SZ = 1
Global Const REG_DWORD = 4
Global Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Global Const gREGVALSYSINFOLOC = "MSINFO"
Global Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Global Const gREGVALSYSINFO = "PATH"

Public Const strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=5;Jet OLEDB:Database Password=T1d1lSt31m6594;Data Source="

'API
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

'Legend Location Constance
Public LegendWidth As Double
Public Const LegendHeight = 600
Public Const LegendLeft = 700

'Scale Bar Location Constance
Public Const BarHeight = 600

Public lyrAnor As Long
Public lyrFair As Long
Public lyrVFair As Long

Public strWeatherString As String
Public strSDString As String

Public strPointWeatherString As String
Public strPointSDString As String

Public TopX As Double
Public TopY As Double
Public BottomX As Double
Public BottomY As Double
Public ScaleFactorX As Double
Public ScaleFactorY As Double

Public numOfArrow As Long
Public numOfGrid As Long

Public TranPeriod As Boolean
Public pointTranPeriod As Boolean

Public Const HH_DISPLAY_TOPIC = &H0
Public Const HH_SET_WIN_TYPE = &H4
Public Const HH_GET_WIN_TYPE = &H5
Public Const HH_GET_WIN_HANDLE = &H6
Public Const HH_DISPLAY_TEXT_POPUP = &HE   ' Display string resource ID or
                                          ' text in a pop-up window.
Public Const HH_HELP_CONTEXT = &HF         ' Display mapped numeric value in
                                          ' dwData.
Public Const HH_TP_HELP_CONTEXTMENU = &H10 ' Text pop-up help, similar to
                                          ' WinHelp's HELP_CONTEXTMENU.
Public Const HH_TP_HELP_WM_HELP = &H11     ' text pop-up help, similar to
                                          ' WinHelp's HELP_WM_HELP.

Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long

Public Const piValue = 3.14159265358979

Public Sub InitDB()
  Dim frm As Form
  If Len(conn.ConnectionString) = 0 Then
    On Error Resume Next
    conn.Open strConnString & App.Path & "\Tide\AtlasFullXP.mdb"
    If Err <> 0 Then
      MsgBox "Open Database with errors !! ", vbCritical
      For Each frm In Forms
        Unload frm
      Next
    End If
    On Error GoTo 0
  End If
End Sub

Public Function AppPath() As String
  AppPath = App.Path
End Function

Public Function edwinDir(strN As Double, strE As Double) As Double
  If strE = 0 Then
    edwinDir = 90
  End If
  If strN = 0 Then
    edwinDir = 0
  End If
  If strE < 0 And strN < 0 Then
    edwinDir = 270 - Atn(strN / strE) * 180 / piValue
  End If
  If strE < 0 And strN > 0 Then
    edwinDir = 270 + Atn(-strN / strE) * 180 / piValue
  End If
  If strE > 0 And strN > 0 Then
    edwinDir = 90 - Atn(strN / strE) * 180 / piValue
  End If
  If strE > 0 And strN < 0 Then
    edwinDir = 90 + Atn(-strN / strE) * 180 / piValue
  End If
End Function

Public Sub ErrorMsg(strMsg As String)
  'MsgBox strMsg, vbCritical, "Error Founds"
  If isFormLoad("frmErrorMsg") Then
    Unload frmErrorMsg
  End If
  Dim FormErr As Form
  For Each FormErr In Forms
    FormErr.Enabled = False
  Next
  frmErrorMsg.Show
  frmErrorMsg.Label1 = strMsg
End Sub

Public Sub DateWeatherDefault()
  Dim strYear As String
  Dim dateWetStart As Date
  Dim dateWetEnd As Date
  
  Dim dateTran1Start As Date
  Dim dateTran1End As Date
  Dim dateTran2Start As Date
  Dim dateTran2End As Date
     
  strYear = Year(DateCal)
  dateWetStart = DateAdd("d", 152, CDate("Jan 1, " & Year(DateCal)))
  'dateWetEnd = CDate("July 30, " & LimitedYear)
  dateWetEnd = DateAdd("d", 210, CDate("Jan 1, " & Year(DateCal)))
   
  'dateTran1Start = CDate("May 2," & LimitedYear)
  dateTran1Start = DateAdd("d", 121, CDate("Jan 1, " & Year(DateCal)))
  'dateTran1End = CDate("June 1," & LimitedYear)
  dateTran1End = DateAdd("d", 151, CDate("Jan 1, " & Year(DateCal)))
  
  'dateTran2Start = CDate("July 31," & LimitedYear)
  dateTran2Start = DateAdd("d", 211, CDate("Jan 1, " & Year(DateCal)))
  'dateTran2End = CDate("Aug 30," & LimitedYear)
  dateTran2End = DateAdd("d", 241, CDate("Jan 1, " & Year(DateCal)))
    
  TranPeriod = False
  If DateCal >= dateTran1Start And DateCal <= dateTran2End Then
    strWeatherString = "W"
    If DateCal < dateWetStart Or DateCal > dateWetEnd Then
      TranPeriod = True
    End If
  Else
    strWeatherString = "D"
  End If
End Sub

Public Sub PointDateWeatherDefault()
  Dim strYear As String
  Dim dateWetStart As Date
  Dim dateWetEnd As Date
  
  Dim dateTran1Start As Date
  Dim dateTran1End As Date
  Dim dateTran2Start As Date
  Dim dateTran2End As Date
     
  strYear = Year(PointDateCal)
  dateWetStart = DateAdd("d", 152, CDate("Jan 1, " & Year(DateCal)))
  'dateWetEnd = CDate("July 30, " & LimitedYear)
  dateWetEnd = DateAdd("d", 210, CDate("Jan 1, " & Year(DateCal)))
   
  'dateTran1Start = CDate("May 2," & LimitedYear)
  dateTran1Start = DateAdd("d", 121, CDate("Jan 1, " & Year(DateCal)))
  'dateTran1End = CDate("June 1," & LimitedYear)
  dateTran1End = DateAdd("d", 151, CDate("Jan 1, " & Year(DateCal)))
  
  'dateTran2Start = CDate("July 31," & LimitedYear)
  dateTran2Start = DateAdd("d", 211, CDate("Jan 1, " & Year(DateCal)))
  'dateTran2End = CDate("Aug 30," & LimitedYear)
  dateTran2End = DateAdd("d", 241, CDate("Jan 1, " & Year(DateCal)))
    
  pointTranPeriod = False
  If PointDateCal >= dateTran1Start And PointDateCal <= dateTran2End Then
    strPointWeatherString = "W"
    If PointDateCal < dateWetStart Or PointDateCal > dateWetEnd Then
      pointTranPeriod = True
    End If
  Else
    strPointWeatherString = "D"
  End If
End Sub

Function DMSToL(D As String, m As String, s As String) As Long
  Dim deg As Double
  Dim min As Double
  Dim sec As Double
  deg = Val(D)
  min = Val(m)
  sec = Val(s)
  deg = deg * 1000000
  min = min * 16666.66666667
  sec = sec * 277.77777778
  deg = deg + min + sec
  DMSToL = CLng(deg)
End Function

Function LToDMSString(l)
  Dim deg As Double
  Dim min As Double
  Dim sec As Double
  Dim v As Double
  If (l < 0) Then l = -l
  v = l / 1000000
  deg = Int(v)
  v = v - deg
  min = Int(v * 60)
  v = v - min / 60
  sec = v * 3600
  LToDMSString = Str(deg) & " " & Str(min) & "'" & Left$(Str(sec), 5) & Chr$(34)
End Function

Function CheckDate(chkDate As Date) As Boolean
  If chkDate < startLimitYear Or chkDate > endLimitYear Then
    CheckDate = False
  Else
    CheckDate = True
  End If
End Function

Public Sub InitMonthCal()
  'InitMonthDate = CDate(Year(DateCal) & "-" & Month(DateCal) & "-01")
  InitMonthDate = CDate("01 " & Format(DateCal, "mmm, yyyy"))
  HourCount = DateDiff("n", InitMonthDate, DateCal) / 60
End Sub

Public Sub InitPointMonthCal()
  'InitPointMonthDate = CDate(Year(PointDateCal) & "-" & Month(PointDateCal) & "-01")
  InitPointMonthDate = CDate("01 " & Format(PointDateCal, "mmm, yyyy"))
  PointHourCount = DateDiff("n", InitPointMonthDate, PointDateCal) / 60
End Sub

Public Function isFormLoad(strFormName As String) As Boolean
  Dim i As Integer
  isFormLoad = False
  For i = 0 To Forms.Count - 1
    If Forms(i).Name = strFormName Then
      isFormLoad = True
      Exit Function
    End If
  Next
End Function

Public Sub InitPlaceName()
  Dim i As Integer, X1 As Long, Y1 As Long
  rst.Open "select x, y, Label, group from Places order by group, label", conn, adOpenKeyset
  i = 0
  Do While Not rst.EOF
    XY_To_Map rst.Fields(1).Value, rst.Fields(0).Value, Y1, X1
    PlaceNameArr(i).X = X1
    PlaceNameArr(i).Y = Y1
    PlaceNameArr(i).Name = rst.Fields(2).Value
    PlaceNameArr(i).group = rst.Fields(3).Value
    i = i + 1
    rst.MoveNext
  Loop
  varPlaceNameCnt = i - 1
  rst.Close
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
  Dim i As Long ' Loop Counter
  Dim rc As Long ' Return Code
  Dim hKey As Long ' Handle To An Open Registry Key
  Dim hDepth As Long
  Dim KeyValType As Long ' Data Type Of A Registry Key
  Dim tmpVal As String ' Tempory Storage For A Registry Key Value
  Dim KeyValSize As Long ' Size Of Registry Key Variable
    
  rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    
  If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    
  tmpVal = String$(1024, 0)
  KeyValSize = 1024
  rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)
  If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
  If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then
    tmpVal = Left(tmpVal, KeyValSize - 1)
  Else
    tmpVal = Left(tmpVal, KeyValSize)
  End If

  Select Case KeyValType
    Case REG_SZ
      KeyVal = tmpVal
    Case REG_DWORD
      For i = Len(tmpVal) To 1 Step -1 ' Convert Each Bit
        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1))) ' Build Value Char. By
      Next
      KeyVal = Format$("&h" + KeyVal) ' Convert Double Word To String
  End Select
    
  GetKeyValue = True ' Return Success
  rc = RegCloseKey(hKey) ' Close Registry Key
  Exit Function ' Exit
  
GetKeyError:
  KeyVal = ""
  GetKeyValue = False
  rc = RegCloseKey(hKey)
End Function

Public Sub DisableAllForm()
  Dim formAll As Form
  For Each formAll In Forms
    DoEvents
    formAll.Enabled = False
  Next
End Sub

Public Sub EnableAllForm()
  Dim formAll As Form
  For Each formAll In Forms
    DoEvents
    formAll.Enabled = True
  Next
End Sub

Public Function LToDMSString2(ldeg As Long, x0y1 As Integer)
  Dim l As Double
  Dim v As Double
  Dim deg As Double
  Dim min As Double
  Dim sec As Double
  Dim directional As String
On Error Resume Next
  l = ldeg
  If (l < 0) Then
    l = -l
    If (x0y1 = 0) Then
      directional = "W"
    Else
      directional = "S"
    End If
  Else
    If (x0y1 = 0) Then
      directional = "E"
    Else
      directional = "N"
    End If
  End If
  v = l
  v = v / 1000000
  deg = Int(v)
  v = v - deg
  min = Round(v * 60, 3)
  LToDMSString2 = CStr(deg) & "-" & CStr(min) & "'" & directional
On Error GoTo 0
End Function

Public Sub XY_To_Map(Y1 As Double, X1 As Double, Y As Long, X As Long) '
  Dim Y2 As Double, X2 As Double
  Y2 = CDbl(Y)
  X2 = CDbl(X)
  HKGEO 2, Y1, X1, Y2, X2
  Y = CLng(Y2 * 1000000)
  X = CLng(X2 * 1000000)
End Sub
