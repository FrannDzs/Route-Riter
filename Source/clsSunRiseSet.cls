VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSunRiseSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------ Class clsSunRiseSet
' Version 2.0

Option Explicit

' -- The following properties are exposed:
'Sunrise (r) - Sunrise time
'Sunset (r) - Sunset time
'SolarNoon (r) - Solar noon
'
'CityCount (r) - Number of cities
'CityName (r) - Name of city, by index
'City (w) - Sets the longitude/latitude/timezone based off a city
'           name or city index
'
'TimeZone (r/w) - Current Timezone - Positive for W, negative for E
'DaySavings (r/w) - Daylight savings time in effect
'Longitude (r/w) - Longitude to calculate for - Positive for W, negative for E
'Latitude (r/w) - Latitude to calculate for - Positive for N, negative for S
'
'DateDay (r/w) - Date to calculate for
'
'
' -- The following method is exposed
'CalculateSun - Calculate sunrise, sunset and solar noon
'
'
' Scott Seligman <scott@scottandmichelle.net>
' Based off of
'   http://www.srrb.noaa.gov/highlights/sunrise/gen.html

Private Type typeMonth
    Name As String
    NumDays As Long
End Type

Private Type typeCity
    Name As String
    longitude As Double
    latitude As Double
    TimeZone As Long
End Type

Private m_cNumberCities As Long
Private m_Cities() As typeCity
    
Private m_monthList(0 To 11) As typeMonth
Private m_monthLeap(0 To 11) As typeMonth

Private m_nTimeZone As Long
Private m_bDaySavings As Boolean
Private m_nLongitude As Double
Private m_nLatitude As Double
Private m_dateSel As Date

Private m_dateSunrise As Date
Private m_dateSunset As Date
Private m_dateNoon As Date

Public Property Get Sunrise() As Date
    Sunrise = m_dateSunrise
End Property

Public Property Get Sunset() As Date
    Sunset = m_dateSunset
End Property

Public Property Get SolarNoon() As Date
    SolarNoon = m_dateNoon
End Property

Public Property Get CityCount() As Long
    CityCount = m_cNumberCities + 1
End Property

Public Property Get CityName(nCity As Long) As String
    If nCity < 0 Or nCity > m_cNumberCities Then
        CityName = "(Error)"
    Else
        CityName = m_Cities(nCity).Name
    End If
End Property

Public Property Let City(City)
    Dim nCity As Long
    Dim bFound As Boolean

    If VarType(City) = vbString Then
        For nCity = 0 To m_cNumberCities
            If Trim$(LCase(City)) = _
                Trim$(LCase(m_Cities(nCity).Name)) Then
                bFound = True
                Exit For
            End If
        Next
        If Not bFound Then
            nCity = -1
        End If
    Else
        If IsNumeric(City) Then
            nCity = City
        Else
            nCity = -1
        End If
    End If

    If nCity < 0 Or nCity > m_cNumberCities Then
        m_nTimeZone = 0
        m_bDaySavings = False
        m_nLongitude = 0
        m_nLatitude = 0
    Else
        m_nTimeZone = m_Cities(nCity).TimeZone
        m_bDaySavings = False
        m_nLongitude = m_Cities(nCity).longitude
        m_nLatitude = m_Cities(nCity).latitude
    End If

End Property

Public Property Let TimeZone(nNew As Long)
    m_nTimeZone = nNew
End Property

Public Property Get TimeZone() As Long
    TimeZone = m_nTimeZone
End Property

Public Property Let DaySavings(bNew As Boolean)
    m_bDaySavings = bNew
End Property

Public Property Get DaySavings() As Boolean
    DaySavings = m_bDaySavings
End Property

Public Property Let longitude(nNew As Double)
    m_nLongitude = nNew
End Property

Public Property Get longitude() As Double
    longitude = m_nLongitude
End Property

Public Property Let latitude(nNew As Double)
    m_nLatitude = nNew
End Property

Public Property Get latitude() As Double
    latitude = m_nLatitude
End Property
    
Public Property Let DateDay(dateNew As Date)
    m_dateSel = dateNew
End Property
    
Public Property Get DateDay() As Date
    DateDay = m_dateSel
End Property


Private Function IsLeapYear(nYear As Long) As Boolean
    If (nYear Mod 4 = 0 And nYear Mod _
        100 <> 0) Or nYear Mod 400 = 0 Then
        IsLeapYear = True
    Else
        IsLeapYear = False
    End If
End Function

Private Function RadToDeg(angleRad As Double) As Double
    RadToDeg = 180 * angleRad / 3.1415926535
End Function

Private Function DegToRad(angleDeg As Double) As Double
    DegToRad = 3.1415926535 * angleDeg / 180
End Function

Private Function acos(x As Double) As Double
    On Error Resume Next
    acos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Function

Private Sub InitMonths()

    m_monthList(0).Name = "January": m_monthList(0).NumDays = 31
    m_monthList(1).Name = "February": m_monthList(1).NumDays = 28
    m_monthList(2).Name = "March": m_monthList(2).NumDays = 31
    m_monthList(3).Name = "April": m_monthList(3).NumDays = 30
    m_monthList(4).Name = "May": m_monthList(4).NumDays = 31
    m_monthList(5).Name = "June": m_monthList(5).NumDays = 30
    m_monthList(6).Name = "July": m_monthList(6).NumDays = 31
    m_monthList(7).Name = "August": m_monthList(7).NumDays = 31
    m_monthList(8).Name = "September": m_monthList(8).NumDays = 30
    m_monthList(9).Name = "October": m_monthList(9).NumDays = 31
    m_monthList(10).Name = "November": m_monthList(10).NumDays = 30
    m_monthList(11).Name = "DEcember": m_monthList(11).NumDays = 31
    
    m_monthLeap(0).Name = "January": m_monthLeap(0).NumDays = 31
    m_monthLeap(1).Name = "February": m_monthLeap(1).NumDays = 29
    m_monthLeap(2).Name = "March": m_monthLeap(2).NumDays = 31
    m_monthLeap(3).Name = "April": m_monthLeap(3).NumDays = 30
    m_monthLeap(4).Name = "May": m_monthLeap(4).NumDays = 31
    m_monthLeap(5).Name = "June": m_monthLeap(5).NumDays = 30
    m_monthLeap(6).Name = "July": m_monthLeap(6).NumDays = 31
    m_monthLeap(7).Name = "August": m_monthLeap(7).NumDays = 31
    m_monthLeap(8).Name = "September": m_monthLeap(8).NumDays = 30
    m_monthLeap(9).Name = "October": m_monthLeap(9).NumDays = 31
    m_monthLeap(10).Name = "November": m_monthLeap(10).NumDays = 30
    m_monthLeap(11).Name = "DEcember": m_monthLeap(11).NumDays = 31

End Sub

Private Sub Class_Initialize()

    InitMonths
    
End Sub

Private Function calcDayOfYear(ByVal nmn As Long, ByVal ndy As Long, _
    bLeapYear As Boolean) As Long

    calcDayOfYear = Int((275 * nmn) / 9) - IIf(bLeapYear, 1, 2) _
        * Int((nmn + 9) / 12) + ndy - 30

End Function

Private Function calcJD(ByVal nYear As Long, ByVal nMonth As Long, _
    ByVal nDay As Long) As Double

    If nMonth <= 2 Then
        nYear = nYear - 1
        nMonth = nMonth + 12
    End If
    Dim A As Long
    Dim B As Long
    
    A = Int(nYear / 100)
    B = 2 - A + Int(A / 4)
    
    calcJD = Int(365.25 * (nYear + 4716)) + Int(30.6001 * _
        (nMonth + 1)) + nDay + B - 1524.5

End Function

Private Function calcTimeJulianCent(ByVal njd As Double) As Double

    calcTimeJulianCent = (njd - 2451545#) / 36525#
    
End Function

Private Function calcJDFromJulianCent(ByVal nt As Double) As Double

    calcJDFromJulianCent = nt * 36525# + 2451545#

End Function

Private Function calcGeomMeanLongSun(ByVal nt As Double) As Double

    Dim nLO As Double
    nLO = 280.46646 + nt * (36000.76983 + 0.0003032 * nt)
    Do While nLO > 360#
        nLO = nLO - 360#
    Loop
    Do While nLO < 0
        nLO = nLO + 360#
    Loop
    calcGeomMeanLongSun = nLO

End Function

Private Function calcGeomMeanAnomalySun(ByVal nt As Double) As Double

    calcGeomMeanAnomalySun = 357.52911 + nt * (35999.05029 - _
        0.0001537 * nt)

End Function

Private Function calcEccentricityEarthOrbit(ByVal nt As Double) _
    As Double

    calcEccentricityEarthOrbit = 0.016708634 - nt * _
        (0.000042037 + 0.0000001267 * nt)

End Function

Private Function calcSunEqOfCenter(ByVal nt As Double) As Double

    Dim nm As Double
    Dim nmrad As Double
    Dim nsinm As Double
    Dim nsin2m As Double
    Dim nsin3m As Double

    nm = calcGeomMeanAnomalySun(nt)
    
    nmrad = DegToRad(nm)
    nsinm = Sin(nmrad)
    nsin2m = Sin(nmrad + nmrad)
    nsin3m = Sin(nmrad + nmrad + nmrad)
    
    calcSunEqOfCenter = nsinm * (1.914602 - nt * _
        (0.004817 + 0.000014 * nt)) + nsin2m * _
        (0.019993 - 0.000101 * nt) + nsin3m * 0.000289

End Function

Private Function calcSunTrueLong(ByVal nt As Double) As Double

    Dim n10 As Double
    Dim nc As Double
    
    n10 = calcGeomMeanLongSun(nt)
    nc = calcSunEqOfCenter(nt)
    
    calcSunTrueLong = n10 + nc

End Function

Private Function calcSunApparentLong(ByVal nt As Double) As Double
    Dim no As Double
    Dim nomega As Double
    
    no = calcSunTrueLong(nt)
    nomega = 125.04 - 1934.136 * nt
    calcSunApparentLong = no - 0.00569 - 0.00478 * _
        Sin(DegToRad(nomega))
    
End Function

Private Function calcMeanObliquityOfEcliptic(ByVal nt As Double) _
    As Double

    Dim nseconds As Double
    
    nseconds = 21.448 - nt * (46.815 + nt * _
        (0.00059 - nt * (0.001813)))
    calcMeanObliquityOfEcliptic = 23# + (26# + _
        (nseconds / 60#)) / 60#

End Function

Private Function calcObliquityCorrection(ByVal nt As Double) As Double

    Dim ne0 As Double
    ne0 = calcMeanObliquityOfEcliptic(nt)
    
    Dim nomega As Double
    nomega = 125.04 - 1934.136 * nt
    calcObliquityCorrection = ne0 + 0.00256 * Cos(DegToRad(nomega))

End Function

Private Function calcSunRtAscension(nt As Double) As Double

    Dim ne As Double
    Dim nlambda As Double
    Dim ntananum As Double
    Dim ntanadenom As Double
    
    ne = calcObliquityCorrection(nt)
    nlambda = calcSunApparentLong(nt)

    ntananum = (Cos(DegToRad(ne)) * Sin(DegToRad(nlambda)))
    ntanadenom = (Cos(DegToRad(nlambda)))
    
    calcSunRtAscension = RadToDeg(atan2(ntananum, ntanadenom))
    
End Function

Private Function atan2(ByVal Y As Double, ByVal x As Double) As Double

    If x > 0 Then
        atan2 = Atn(Y / x)
    ElseIf x < 0 Then
        atan2 = Atn(Y / x) + 3.1415926535
    Else
        atan2 = 3.1415926535 / 2 * Sgn(Y)
    End If

End Function

Private Function Asin(ByVal x As Double) As Double
    Asin = Atn(x / Sqr(-x * x + 1))
End Function

Private Function calcSunDeclination(ByVal nt As Double) As Double

    Dim ne As Double
    Dim nlambda As Double
    Dim nsint As Double
    
    ne = calcObliquityCorrection(nt)
    nlambda = calcSunApparentLong(nt)
    
    nsint = Sin(DegToRad(ne)) * Sin(DegToRad(nlambda))
    calcSunDeclination = RadToDeg(Asin(nsint))
    
End Function

Private Function calcEquationOfTime(ByVal nt As Double) As Double

    Dim nepsilon As Double
    Dim nl0 As Double
    Dim ne As Double
    Dim nm As Double
    Dim ny As Double
    Dim nsin2l0 As Double
    Dim nsinm As Double
    Dim ncos2l0 As Double
    Dim nsin4l0 As Double
    Dim nsin2m As Double
    Dim nEtime As Double
    
    nepsilon = calcObliquityCorrection(nt)
    nl0 = calcGeomMeanLongSun(nt)
    ne = calcEccentricityEarthOrbit(nt)
    nm = calcGeomMeanAnomalySun(nt)
    
    ny = Math.Tan(DegToRad(nepsilon) / 2#)
    ny = ny * ny
    
    nsin2l0 = Sin(2# * DegToRad(nl0))
    nsinm = Sin(DegToRad(nm))
    ncos2l0 = Cos(2# * DegToRad(nl0))
    nsin4l0 = Sin(4# * DegToRad(nl0))
    nsin2m = Sin(2# * DegToRad(nm))
    
    nEtime = ny * nsin2l0 - 2# * ne * nsinm + 4# * ne * _
        ny * nsinm * ncos2l0 - 0.5 * ny * ny * nsin4l0 - _
        1.25 * ne * ne * nsin2m
    
    calcEquationOfTime = RadToDeg(nEtime) * 4#
End Function

Private Function calcHourAngleSunrise(ByVal nlat As Double, _
    ByVal nsolarDec As Double) As Double

    Dim nlatRad As Double
    Dim nsdRad As Double
    Dim nHAarg As Double
    Dim nHA As Double

    nlatRad = DegToRad(nlat)
    nsdRad = DegToRad(nsolarDec)

    nHAarg = (Cos(DegToRad(90.833)) / (Cos(nlatRad) * _
        Cos(nsdRad)) - Tan(nlatRad) * Tan(nsdRad))

    Dim nTemp As Double
    nTemp = Cos(DegToRad(90.833)) / (Cos(nlatRad) * _
        Cos(nsdRad)) - Tan(nlatRad) * Tan(nsdRad)
    If Abs(nTemp) > 1 Then
        nHA = -999
    Else
        nHA = (acos(nTemp))
    End If

    calcHourAngleSunrise = nHA

End Function

Private Function calcHourAngleSunset(ByVal nlat As Double, _
    ByVal nsolarDec As Double) As Double

    Dim nlatRad As Double
    Dim nsdRad As Double
    Dim nHAarg As Double
    Dim nHA As Double

    nlatRad = DegToRad(nlat)
    nsdRad = DegToRad(nsolarDec)
    
    nHAarg = (Cos(DegToRad(90.833)) / (Cos(nlatRad) * _
        Cos(nsdRad)) - Tan(nlatRad) * Tan(nsdRad))
    
    Dim nTemp As Double
    nTemp = Cos(DegToRad(90.833)) / (Cos(nlatRad) * _
        Cos(nsdRad)) - Tan(nlatRad) * Tan(nsdRad)
    If Abs(nTemp) > 1 Then
        nHA = 999
    Else
        nHA = (acos(nTemp))
    End If
    
    calcHourAngleSunset = -nHA

End Function

Private Function calcSunriseUTC(ByVal njd As Double, _
    ByVal nLatitude As Double, ByVal nLongitude As Double) As Double

    Dim nt As Double
    Dim neqTime As Double
    Dim nsolarDec As Double
    Dim nhourAngle As Double
    
    Dim ndelta As Double
    Dim ntimeDiff As Double
    Dim ntimeUTC As Double
    
    nt = calcTimeJulianCent(njd)

    neqTime = calcEquationOfTime(nt)
    nsolarDec = calcSunDeclination(nt)
    nhourAngle = calcHourAngleSunrise(nLatitude, nsolarDec)
    If nhourAngle = -999 Then
        calcSunriseUTC = -999
        Exit Function
    End If

    ndelta = nLongitude - RadToDeg(nhourAngle)
    ntimeDiff = 4 * ndelta
    ntimeUTC = 720 + ntimeDiff - neqTime

    Dim nnewt As Double
    nnewt = calcTimeJulianCent(calcJDFromJulianCent(nt) + _
        ntimeUTC / 1440#)
    neqTime = calcEquationOfTime(nnewt)
    nsolarDec = calcSunDeclination(nnewt)
    nhourAngle = calcHourAngleSunrise(nLatitude, nsolarDec)
    If nhourAngle = -999 Then
        calcSunriseUTC = -999
        Exit Function
    End If
    ndelta = nLongitude - RadToDeg(nhourAngle)
    ntimeDiff = 4 * ndelta
    ntimeUTC = 720 + ntimeDiff - neqTime

    calcSunriseUTC = ntimeUTC
    
End Function

Private Function calcSolNoonUTC(ByVal nt As Double, _
    ByVal nLongitude As Double) As Double

    Dim nnewt As Double
    Dim neqTime As Double
    Dim nsolarNoonDec As Double
    Dim nsolNoonUTC As Double
    
    nnewt = calcTimeJulianCent(calcJDFromJulianCent(nt) + _
        0.5 + nLongitude / 360#)
    neqTime = calcEquationOfTime(nt)
    nsolarNoonDec = calcSunDeclination(nt)
    nsolNoonUTC = 720 + (nLongitude * 4) - neqTime
    
    calcSolNoonUTC = nsolNoonUTC

End Function

Private Function calcSunsetUTC(ByVal njd As Double, _
    ByVal nLatitude As Double, ByVal nLongitude As _
    Double) As Double

    Dim neqTime As Double
    Dim nsolarDec As Double
    Dim nhourAngle As Double
    
    Dim ndelta As Double
    Dim ntimeDiff As Double
    Dim ntimeUTC As Double
    Dim nnewt As Double
    Dim nt As Double

    nt = calcTimeJulianCent(njd)

    neqTime = calcEquationOfTime(nt)
    nsolarDec = calcSunDeclination(nt)
    nhourAngle = calcHourAngleSunset(nLatitude, nsolarDec)
    If nhourAngle = -999 Then
        calcSunsetUTC = -999
        Exit Function
    End If

    ndelta = nLongitude - RadToDeg(nhourAngle)
    ntimeDiff = 4 * ndelta
    ntimeUTC = 720 + ntimeDiff - neqTime
    
    nnewt = calcTimeJulianCent(calcJDFromJulianCent(nt) + _
        ntimeUTC / 1440#)
    neqTime = calcEquationOfTime(nnewt)
    nsolarDec = calcSunDeclination(nnewt)
    nhourAngle = calcHourAngleSunset(nLatitude, nsolarDec)
    If nhourAngle = -999 Then
        calcSunsetUTC = -999
        Exit Function
    End If
    
    ndelta = nLongitude - RadToDeg(nhourAngle)
    ntimeDiff = 4 * ndelta
    ntimeUTC = 720 + ntimeDiff - neqTime
    
    calcSunsetUTC = ntimeUTC

End Function

Private Function findRecentSunrise(ByVal njd As Double, _
    ByVal nLatitude As Double, ByVal nLongitude _
    As Double) As Double

    Dim njulianday As Double
    njulianday = njd
    Dim nBail As Long
    
    Dim ntime As Double
    ntime = calcSunriseUTC(njulianday, nLatitude, nLongitude)
    Do While ntime = -999 And nBail < 367
        nBail = nBail + 1
        njulianday = njulianday - 1
        ntime = calcSunriseUTC(njulianday, nLatitude, nLongitude)
    Loop
    
    findRecentSunrise = njulianday

End Function

Private Function findRecentSunset(ByVal njd As Double, _
    ByVal nLatitude As Double, ByVal nLongitude _
    As Double) As Double

    Dim njulianday As Double
    Dim ntime As Double
    Dim nBail As Long
    
    njulianday = njd

    ntime = calcSunsetUTC(njulianday, nLatitude, nLongitude)
    Do While ntime = -999 And nBail < 367
        nBail = nBail + 1
        njulianday = njulianday - 1
        ntime = calcSunsetUTC(njulianday, nLatitude, nLongitude)
    Loop
    
    findRecentSunset = njulianday

End Function

Private Function findNextSunrise(ByVal njd As Double, ByVal _
    nLatitude As Double, ByVal nLongitude As Double) As Double
    Dim njulianday As Double
    Dim ntime As Double
    Dim nBail As Long

    njulianday = njd

    ntime = calcSunriseUTC(njulianday, nLatitude, nLongitude)
    Do While ntime = -999 And nBail < 367
        nBail = nBail + 1
        njulianday = njulianday + 1
        ntime = calcSunriseUTC(njulianday, nLatitude, nLongitude)
    Loop
    
    findNextSunrise = njulianday

End Function

Private Function findNextSunset(ByVal njd As Double, ByVal _
    nLatitude As Double, ByVal nLongitude As Double) As Double

    Dim njulianday As Double
    Dim ntime As Double
    Dim nBail As Long
    njulianday = njd
    ntime = calcSunsetUTC(njulianday, nLatitude, nLongitude)
    Do While ntime = -999 And nBail < 367
        nBail = nBail + 1
        njulianday = njulianday + 1
        ntime = calcSunsetUTC(njulianday, nLatitude, nLongitude)
    Loop
    
    findNextSunset = njulianday

End Function

Public Function CalculateSun()
Rem ************* Get from form
m_dateSel = frmGetSun.txtDate      '"21/6/1999"
    
        If frmGetSun.optTZEast.Value = True Then
        m_nTimeZone = "-" & frmGetSun.txtTZ
        Else
        m_nTimeZone = frmGetSun.txtTZ
        End If
        If frmGetSun.DaySave.Value = 0 Then
        m_bDaySavings = False
        ElseIf frmGetSun.optNorth.Value = True And (m_dateSel = "1/4/2002" Or m_dateSel = "1/7/2002") Then
        m_bDaySavings = True
        ElseIf frmGetSun.optSouth.Value = True And (m_dateSel = "1/10/2002" Or m_dateSel = "1/1/2002") Then
        m_bDaySavings = True
        Else
        m_bDaySavings = False
        End If
        If frmGetSun.optEast.Value = True Then
        m_nLongitude = "-" & frmGetSun.txtLon
        Else
        m_nLongitude = frmGetSun.txtLon
        End If
        If frmGetSun.optSouth.Value = True Then
        m_nLatitude = "-" & frmGetSun.txtLat
        Else
        m_nLatitude = frmGetSun.txtLat
        End If



Rem ******************************

    Dim nLatitude As Double
    Dim nLongitude As Double
    nLatitude = m_nLatitude
    nLongitude = m_nLongitude

    If nLatitude >= -90 And nLatitude < -89.5 Then
        nLatitude = -89.5
    End If
    If nLatitude <= 90 And nLatitude > 89.8 Then
        nLatitude = 89.8
    End If

    Dim njd As Double
    Dim ndoy As Double
    Dim nt As Double
    
    njd = calcJD(Year(m_dateSel), Month(m_dateSel), Day(m_dateSel))
    ndoy = calcDayOfYear(Month(m_dateSel), Day(m_dateSel), _
        IsLeapYear(Year(m_dateSel)))
    nt = calcTimeJulianCent(njd)
    
    Dim nAlpha As Double
    Dim nTheta As Double
    Dim nEtime As Double
    
    nAlpha = calcSunRtAscension(nt)
    nTheta = calcSunDeclination(nt)
    nEtime = calcEquationOfTime(nt)
    
    Dim neqTime As Double
    Dim nsolarDec As Double
    
    neqTime = nEtime
    nsolarDec = nTheta
    
    'Calculate sunrise
    Dim bNoSunrise As Boolean
    bNoSunrise = False
    
    Dim nRiseTimeGMT As Double
    nRiseTimeGMT = calcSunriseUTC(njd, nLatitude, nLongitude)
    
    If nRiseTimeGMT = -999 Then
        bNoSunrise = True
    End If

    'Calculate sunset
    Dim bNoSunset As Boolean
    bNoSunset = False
    Dim nSetTimeGMT As Double
    nSetTimeGMT = calcSunsetUTC(njd, nLatitude, nLongitude)
    If nSetTimeGMT = -999 Then
        bNoSunset = True
    End If
    
    Dim ndaySavings As Double
    Dim nZone As Double
    
    ndaySavings = IIf(m_bDaySavings, 60, 0)
    nZone = m_nTimeZone
    If nZone > 12 Or nZone < -12.5 Then
        nZone = 0
    End If
    
    If Not bNoSunrise Then
        Dim nriseTimeLST As Double
        nriseTimeLST = nRiseTimeGMT - (60 * nZone) + ndaySavings
        
        m_dateSunrise = DateAdd("s", nriseTimeLST * 60, _
            Int(m_dateSel))
    End If

    If Not bNoSunset Then
        Dim nsetTimeLST As Double
        nsetTimeLST = nSetTimeGMT - (60 * nZone) + ndaySavings
        
        m_dateSunset = DateAdd("s", nsetTimeLST * 60, Int(m_dateSel))
    End If
    
    'Calculate solar noon for this date
    Dim nsolNoonGMT As Double
    Dim nsolNoonLST As Double
    nsolNoonGMT = calcSolNoonUTC(nt, nLongitude)
    nsolNoonLST = nsolNoonGMT - (60 * nZone) + ndaySavings
    
    m_dateNoon = DateAdd("s", nsolNoonLST * 60, Int(m_dateSel))

    Dim nnewjd As Double
    Dim nnewtime As Double

    If bNoSunrise Then
        If ((nLatitude > 66.4) And (ndoy > 79) And _
            (ndoy < 267)) Or ((nLatitude < -66.4) And _
            ((ndoy < 83) Or (ndoy > 236))) Then

            nnewjd = findRecentSunrise(njd, nLatitude, nLongitude)
            nnewtime = calcSunriseUTC(nnewjd, nLatitude, _
                nLongitude) - (60 * nZone) + ndaySavings

            If nnewtime > 1440 Then
                nnewtime = nnewtime - 1440
                nnewjd = nnewjd + 1
            End If
            If nnewtime < 0 Then
                nnewtime = nnewtime + 1440
                nnewjd = nnewjd - 1
            End If
            
            m_dateSunrise = DateAdd("s", nnewtime * 60, _
                Int(m_dateSel))
            m_dateSunrise = DateAdd("d", nnewjd - njd, m_dateSunrise)
        
        ElseIf ((nLatitude > 66.4) And ((ndoy < 83) Or _
            (ndoy > 263))) Or ((nLatitude < -66.4) And _
            (ndoy > 79) And (ndoy < 267)) Then

            nnewjd = findNextSunrise(njd, nLatitude, nLongitude)
            nnewtime = calcSunriseUTC(nnewjd, nLatitude, _
                nLongitude) - (60 * nZone) + ndaySavings
            If nnewtime > 1440 Then
                nnewtime = nnewtime - 1440
                nnewjd = nnewjd + 1
            End If
            If nnewtime < 0 Then
                nnewtime = nnewtime + 1440
                nnewjd = nnewjd - 1
            End If
            
            m_dateSunrise = DateAdd("s", nnewtime * 60, _
                Int(m_dateSel))
            m_dateSunrise = DateAdd("d", nnewjd - njd, m_dateSunrise)
        
        End If
    End If
    
    If bNoSunset Then
        If (((nLatitude > 66.4) And (ndoy > 79) And _
            (ndoy < 267)) Or ((nLatitude < -66.4) And _
            ((ndoy < 83) Or (ndoy > 263)))) Then

            nnewjd = findNextSunset(njd, nLatitude, nLongitude)
            nnewtime = calcSunsetUTC(nnewjd, nLatitude, _
                nLongitude) - (60 * nZone) + ndaySavings
            If nnewtime > 1440 Then
                nnewtime = nnewtime - 1440
                nnewjd = nnewjd + 1
            End If
            If nnewtime < 0 Then
                nnewtime = nnewtime + 1440
                nnewjd = nnewjd - 1
            End If

            m_dateSunset = DateAdd("s", nnewtime * 60, Int(m_dateSel))
            m_dateSunset = DateAdd("d", nnewjd - njd, m_dateSunset)

        ElseIf (((nLatitude > 66.4) And ((ndoy < 83) Or _
            (ndoy > 263))) Or ((nLatitude < -66.4) And _
            (ndoy > 79) And (ndoy < 267))) Then

            nnewjd = findRecentSunset(njd, nLatitude, nLongitude)
            nnewtime = calcSunsetUTC(nnewjd, nLatitude, _
                nLongitude) - (60 * nZone) + ndaySavings
            If nnewtime > 1440 Then
                nnewtime = nnewtime - 1440
                nnewjd = nnewjd + 1
            End If
            If nnewtime < 0 Then
                nnewtime = nnewtime + 1440
                nnewjd = nnewjd - 1
            End If

            m_dateSunset = DateAdd("s", nnewtime * 60, Int(m_dateSel))
            m_dateSunset = DateAdd("d", nnewjd - njd, m_dateSunset)

        End If
    End If
    
End Function

' ------ End of class clsSunRiseSet


Private Sub Class_Terminate()

End Sub


