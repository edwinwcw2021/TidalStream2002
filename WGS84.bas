Attribute VB_Name = "WGS84"
Option Explicit

Public Const pi = 3.14159265359

Sub HKGEO(ByVal IG As Double, ByVal X As Double, ByVal Y As Double, PHI As Double, FLAM As Double)
'C**** CONVERT HK METRIC GRID COORDINATES TO GEODETIC COORDINTES
'C ENTRY:
'C         IG   ;  1 FOR HAYFORD SPHEROID, 2 FOR WGS84 SPHEROID
'C         X    :  NORTHING  (HK 1980 DATAM)
'C         Y    :  EASTING
'C RETURN :
'C         PHI  : LATITUDE
'C         FLAM : LONGITUDE
    Dim RAD As Double, PHI0 As Double, FLAM0 As Double
    Dim a As Double, b As Double, C As Double, D As Double
    Dim WX As Double, WY As Double, AA As Double, BB As Double, DPHI As Double
    Dim PHIF As Double, DPH As Double, SM As Double, CR As Double
    Dim TPHI As Double, TPHI2 As Double, TPHI4 As Double, TPHI6 As Double
    Dim TT As Double, TT2 As Double, TT3 As Double, TT4 As Double
    Dim DUE As Double, DX As Double, CPHI1 As Double, CPHI2 As Double
    Dim CPHI3 As Double, CPHI4 As Double, CLAM1 As Double, CLAM2 As Double
    Dim CLAM3 As Double, CLAM4 As Double, DY As Double, DX2 As Double
    Dim RHO As Double, RMU As Double

      RAD = pi / 180#
'C---- CONVERT PROJECTION ORIGIN TO RADIANS
      If IG = 1 Then
        PHI0 = (22# + 18# / 60# + 43.68 / 3600#) * RAD
        FLAM0 = (114# + 10# / 60# + 42.8 / 3600#) * RAD
      Else
        PHI0 = (22# + 18# / 60# + 38.17 / 3600#) * RAD
        FLAM0 = (114# + 10# / 60# + 51.65 / 3600#) * RAD
'C----   TRANSFORM HK 1980 GRID TO WGS84 GRID
        a = 1.0000001619
        b = 0.000027858
        C = 23.098979
        D = -23.149125
        WX = a * X - b * Y + C
        WY = b * X + a * Y + D
        X = WX
        Y = WY
      End If
'C---- REMOVE FALSE GRID ORIGIN COORDINATES
      DX = X - 819069.8
      DY = Y - 836694.05
'C---- COMPUTE PROVISIONAL PHIF (APPROXIMATE)
      AA = 6.853561524
      BB = 110736.3925
      DPHI = ((Sqr(DX * AA * 4# + BB ^ 2) - BB) * 0.5 / AA) * RAD
      PHIF = PHI0 + DPHI
      DPH = 0#
'C---- EVALUATE PHIF, ITERATE UNTIL CR IS NEAR ZERO
Do
      PHIF = PHIF + DPH
      SM = SMER(IG, PHI0, PHIF)
      CR = DX - SM
      RADIUS IG, PHIF, RHO, RMU
      DPH = CR / RHO
Loop Until Abs(CR) < 0.00001

' C---- COMPUTE RADII
      RADIUS IG, PHIF, RHO, RMU
      TPHI = Tan(PHIF)
      TPHI2 = TPHI * TPHI
      TPHI4 = TPHI2 * TPHI2
      TPHI6 = TPHI2 * TPHI4
      TT = RMU / RHO
      TT2 = TT ^ 2
      TT3 = TT ^ 3
      TT4 = TT ^ 4
'C---- COMPUTE LATITUDE
      DUE = DY
      DX = DUE / RMU
      DX2 = DX * DX
      CPHI1 = DUE / RHO * DX * TPHI / 2#
      CPHI2 = CPHI1 / 12# * DX2 * (9# * TT * (1# - TPHI2) - 4# * TT2 + 12# * TPHI2)
      CPHI3 = CPHI1 / 360# * DX2 * DX2 * (8# * TT4 * (11# - 24# * TPHI2) - 12# * TT3 * (21# - 71# * TPHI2) + 15# * TT2 * (15# - 98# * TPHI2 + 15# * TPHI4) + 180# * TT * (5# * TPHI2 - 3# * TPHI4) + 360# * TPHI4)
      CPHI4 = CPHI1 / 20160# * DX2 * DX2 * DX2 * (1385# + 3633# * TPHI2 + 4095# * TPHI4 + 1575# * TPHI2 * TPHI4)
      PHI = PHIF - CPHI1 + CPHI2 - CPHI3 + CPHI4
'C---- COMPUTE LONGITUDE
      CLAM1 = DX / Cos(PHIF)
      CLAM2 = CLAM1 * DX2 / 6# * (TT + 2# * TPHI2)
      CLAM3 = CLAM1 * DX2 * DX2 / 120# * (TT2 * (9# - 68# * TPHI2) - 4# * TT3 * (1# - 6# * TPHI2) + 72# * TT * TPHI2 + 24# * TPHI4)
      CLAM4 = CLAM1 * DX2 * DX2 * DX2 / 5040# * (61# + 662# * TPHI2 + 1320# * TPHI4 + 720# * TPHI2 * TPHI4)
      FLAM = FLAM0 + CLAM1 - CLAM2 + CLAM3 - CLAM4
'C---- CONVERT TO DECIMAL DEGREES
      PHI = PHI / RAD
      FLAM = FLAM / RAD
 
End Sub

Sub GEOHK(ByVal IG As Double, ByVal PHI As Double, ByVal FLAM As Double, X As Double, Y As Double)
'C**** CONVERT GEODETIC COORDINTES TO HK METRIC GRID COORDINATES
'C ENTRY:
'C         IG   ; 1 FOR HAYFORD SPHEROID, 2 FOR WG282 SPHEROID
'C         PHI  : LATITUDE   IN DECIMAL DEGREES
'C         FLAM : LONGITUDE  IN DECIMAL DEGREES
'C RETURN :
'C         X    :  NORTHING  (HK 1980 METRIC DATAM)
'C         Y    :  EASTING

    Dim RAD As Double, PHI0 As Double, FLAM0 As Double, RPHI As Double
    Dim RLAM As Double, SM0 As Double, SM1 As Double, CJ  As Double
    Dim TPHI  As Double, TPHI2 As Double, TPHI4 As Double, TPHI6 As Double
    Dim TT As Double, TT2 As Double, TT3  As Double, TT4 As Double
    Dim XF  As Double, X1 As Double, X2 As Double, X3 As Double, X4 As Double
    Dim YF As Double, Y1 As Double, Y2 As Double, Y3 As Double
    Dim WX As Double, WY As Double, a As Double, b As Double, C As Double, D As Double
    Dim RHO As Double, RMU  As Double
    
      RAD = pi / 180
 'C---- CONVERT PROJECTION ORIGIN TO RADIANS
      If IG = 1 Then
        PHI0 = (22# + 18# / 60# + 43.68 / 3600#) * RAD
        FLAM0 = (114# + 10# / 60# + 42.8 / 3600#) * RAD
      Else
        PHI0 = (22# + 18# / 60# + 38.17 / 3600#) * RAD
        FLAM0 = (114# + 10# / 60# + 51.65 / 3600#) * RAD
      End If
'C---- CONVERT LATITUDE AND LONGITUDE TO RADIANS
      RPHI = PHI * RAD
      RLAM = FLAM * RAD
'C---- COMPUTE MERIDIAN ARCS
      SM0 = SMER(IG, 0, PHI0)
      SM1 = SMER(IG, 0, RPHI)
'C---- COMPUTE RADII
      RADIUS IG, RPHI, RHO, RMU
'C---- COMPUTE CJ (IN RADIANS)
      CJ = (RLAM - FLAM0) * Cos(RPHI)
      TPHI = Tan(RPHI)
      TPHI2 = TPHI * TPHI
      TPHI4 = TPHI2 * TPHI2
      TPHI6 = TPHI2 * TPHI4
      TT = RMU / RHO
      TT2 = TT ^ 2
      TT3 = TT ^ 3
      TT4 = TT ^ 4
'C---- COMPUTE  NORTHING
      
      XF = SM1 - SM0
      X1 = RMU / 2# * CJ ^ 2 * TPHI
      X2 = X1 / 12# * CJ ^ 2 * (4# * TT2 + TT - TPHI2)
      X3 = X2 / 30# * CJ ^ 2 * (8# * TT4 * (11# - 24# * TPHI2) - 28# * TT3 * (1# - 6# * TPHI2) + TT2 * (1# - 32# * TPHI2) - 2# * TT * TPHI2 + TPHI4)
      X4 = X3 / 56# * CJ ^ 2 * (1385# - 3111# * TPHI2 + 543# * TPHI4 - TPHI6)
      X = XF + X1 + X2 + X3 + X4 + 819069.8
'C---- COMPUTE  EASTING
      YF = RMU * CJ
      Y1 = YF / 6# * CJ ^ 2
      Y2 = Y1 / 20# * CJ ^ 2
      Y3 = Y2 / 42# * CJ ^ 2
      Y1 = Y1 * (TT - TPHI2)
      Y2 = Y2 * (4# * TT3 * (1# - 6# * TPHI2) + TT2 * (1# + 8# * TPHI2) - TT * 2# * TPHI2 + TPHI4)
      Y3 = Y3 * (61# - 479# * TPHI2 + 179# * TPHI4 - TPHI6)
      Y = YF + Y1 + Y2 + Y3 + 836694.05
      If IG = 2 Then
        WX = X
        WY = Y
'C----   TRANSFROM WGS84 GRID TO HK 1980 GRID
        a = 0.9999998373
        b = -0.000027858
        C = -23.098331
        D = 23.149765
        X = a * WX - b * WY + C
        Y = b * WX + a * WY + D
      End If
      
End Sub

Function SMER(ByVal IG As Double, ByVal PHI0 As Double, ByVal PHIF As Double) As Double
'C**** COMPUTE MERIDIAN ARC
'C ENTRY:
'C         PHI0 : LATITUDE OF ORIGIN
'C         PHIF : LATITUDE OF PROJECTION TO CENTRAL MERIDIAN
'C RETURN :
'C         SMER : MERIDIAN ARC
    Dim AXISM As Double, FLAT  As Double, ECC  As Double
    Dim a As Double, b As Double, C As Double, D As Double, DP0 As Double
    Dim DPO  As Double, DP2 As Double
    Dim DP4 As Double, DP6 As Double

      If IG = 1 Then
        AXISM = 6378388#
        FLAT = 1# / 297#
      Else
        AXISM = 6378137#
        FLAT = 1# / 298.2572235634
      End If
      
      ECC = 2# * FLAT - FLAT ^ 2
      ECC = Sqr(ECC)
      a = 1# + 3# / 4# * ECC ^ 2 + 45# / 64# * ECC ^ 4 + 175# / 256# * ECC ^ 6
      b = 3# / 4# * ECC ^ 2 + 15# / 16# * ECC ^ 4 + 525# / 512# * ECC ^ 6
      C = 15# / 64# * ECC ^ 4 + 105# / 256# * ECC ^ 6
      D = 35# / 512# * ECC ^ 6
      DP0 = PHIF - PHI0
      DP2 = Sin(2# * PHIF) - Sin(2# * PHI0)
      DP4 = Sin(4# * PHIF) - Sin(4# * PHI0)
      DP6 = Sin(6# * PHIF) - Sin(6# * PHI0)
      SMER = AXISM * (1# - ECC ^ 2)
      SMER = SMER * (a * DP0 - b * DP2 / 2# + C * DP4 / 4# - D * DP6 / 6#)
End Function


Sub RADIUS(ByVal IG As Double, ByVal PHI As Double, RHO As Double, RMU As Double)
'C**** COMPUTE RADII OF CURVATURE OF A GIVEN LATITUDE
'C ENTRY:
'C         PHI  : LATITUDE
'C RETURN:
'C         RHO  : RADIUS OF MERIDIAN
'C         PMU  : RADIUS OF PRIME VERTICAL
    Dim AXISM As Double, FLAT  As Double, ECC As Double
    Dim FAC As Double
      If IG = 1 Then
        AXISM = 6378388#
        FLAT = 1# / 297#
      Else
        AXISM = 6378137#
        FLAT = 1# / 298.2572235634
      End If
      ECC = 2# * FLAT - FLAT ^ 2
      FAC = 1# - ECC * (Sin(PHI) ^ 2)
      RHO = AXISM * (1# - ECC) / FAC ^ 1.5
      RMU = AXISM / Sqr(FAC)

End Sub

Sub To_dms(ByVal n, D, m, s)
    'change the latitude or longitude from a number into degree, minute and second
    D = Int(n)
    m = Int((n - D) * 60)
    s = ((n - D) * 60 - m) * 60
    If Abs(60 - s) < 0.000000000001 Then
        m = m + 1
        s = 0
    End If
    If Abs(60 - m) < 0.000000000001 Then
        D = D + 1
        m = 0
    End If

End Sub

Sub To_dms2(ByVal n, D, m)
    'change the latitude or longitude from a number into degree, minute and second
    D = Int(n)
    m = (n - D) * 60
    If Abs(60 - m) < 0.000000000001 Then
        D = D + 1
        m = 0
    End If
    m = Format(m, "00.000000000")
End Sub



