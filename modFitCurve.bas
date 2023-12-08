Attribute VB_Name = "modFitCurve"
Option Explicit

Public Sub catint(ax() As Double, ay() As Double, npt As Integer, nit As Integer, ffx() As Double, ffy() As Double)
  Dim px(4) As Double, py(4) As Double
  ReDim ffx(0 To npt * nit) As Double, ffy(0 To npt * nit) As Double
  ReDim fxx(0 To nit) As Double, fyy(0 To nit) As Double
  Dim ii As Integer, i As Integer, k As Integer, m As Integer, j As Integer
  Dim nintvl As Integer
  ii = 0
  ffx(0) = ax(1)
  ffy(0) = ay(1)
  For j = 1 To npt - 1
    k = j - 1
    For i = 1 To 4
      If k = 0 Then
        k = 1
        px(i) = ax(k)
        py(i) = ay(k)
        k = 0
      End If
      If (k <> 0 And k < npt) Then
        px(i) = ax(k)
        py(i) = ay(k)
      End If
      If (k > npt) Then
        k = k - 1
        px(i) = ax(k)
        py(i) = ay(k)
        k = k + 1
      End If
      k = k + 1
    Next
    nintvl = nit
    catmullsb px, py, nintvl, fxx, fyy
    For m = 1 To nintvl
      ii = ii + 1
      ffx(ii) = fxx(m)
      ffy(ii) = fyy(m)
    Next
  Next
End Sub

Public Sub catmullsb(px() As Double, py() As Double, n As Integer, fx() As Double, fy() As Double)
  Dim i As Integer, j As Integer, m As Integer, t As Double
  Dim ma As Variant, cx(4) As Double, cy(4) As Double
  ReDim fx(0 To n) As Double
  ReDim fy(0 To n) As Double
  Erase cx
  Erase cy
  ma = Array(-1, 3, -3, 1, 2, -5, 4, -1, -1, 0, 1, 0, 0, 2, 0, 0)
  For i = 1 To 4
    For j = 0 To 3
      cx(i) = cx(i) + ma(j + (i - 1) * 4) * px(j + 1)
      cy(i) = cy(i) + ma(j + (i - 1) * 4) * py(j + 1)
    Next
  Next
  For m = 0 To n
    t = m / n
    fx(m) = 0.5 * (cx(1) * t ^ 3 + cx(2) * t ^ 2 + cx(3) * t + cx(4))
    fy(m) = 0.5 * (cy(1) * t ^ 3 + cy(2) * t ^ 2 + cy(3) * t + cy(4))
  Next
End Sub

