VERSION 5.00
Begin VB.Form frmGetSun 
   Caption         =   "Provide Details for your Route"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmTimeZone 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   2415
      Width           =   5600
      Begin VB.CheckBox DaySave 
         Caption         =   "Daylight Saving Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.OptionButton optTZEast 
         Caption         =   "East of Greenwich"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3200
         TabIndex        =   17
         Top             =   550
         Width           =   2370
      End
      Begin VB.OptionButton optTZWest 
         Caption         =   "West of Greenwich"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3200
         TabIndex        =   16
         Top             =   100
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.TextBox txtTZ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   15
         Top             =   100
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Time Zone (hrs)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   100
         Width           =   1815
      End
   End
   Begin VB.Frame frmLon 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   600
      Left            =   120
      TabIndex        =   8
      Top             =   1965
      Width           =   5600
      Begin VB.OptionButton optEast 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4320
         TabIndex        =   12
         Top             =   100
         Width           =   615
      End
      Begin VB.OptionButton optWest 
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   120
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.TextBox txtLon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1965
         TabIndex        =   10
         Top             =   100
         Width           =   1020
      End
      Begin VB.Label Label3 
         Caption         =   "Longitude (deg)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   100
         Width           =   1935
      End
   End
   Begin VB.Frame frmLat 
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   5600
      Begin VB.OptionButton optNorth 
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   7
         Top             =   100
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox txtLat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   6
         Top             =   100
         Width           =   1095
      End
      Begin VB.OptionButton optSouth 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   4
         Top             =   100
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Latitude (deg)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   100
         Width           =   1575
      End
   End
   Begin VB.TextBox txtDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1965
      TabIndex        =   2
      Top             =   1020
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   5535
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   5760
      Y1              =   915
      Y2              =   915
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   5760
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5760
      Y1              =   1395
      Y2              =   1395
   End
End
Attribute VB_Name = "frmGetSun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim col As Double
Dim row As Double
Dim x As Double
Dim z As Double
Dim feast(0 To 11) As Double
Dim lon_centre(0 To 11) As Double
Dim retval(0 To 1) As Double
Dim wt_ew_offset As Double
Dim wt_ns_offset As Double
Dim pixsiz As Double
Dim ul_x As Double
Dim ul_y As Double
Dim R2D As Double
Dim OK As Double
Dim E_radius As Double
Dim HALFPI As Double
Dim EPSLN As Double
Dim PI As Double
Private Sub ParseTime(strTime As String, strNewTime As String, strMoon As String, booTime As Boolean)
Dim hr As String, mi As String, se As String, moon As String, mn As Integer
'Hour(srs.Sunrise) & ":" & Minute(srs.Sunrise) & ":" & Second(srs.Sunrise)
hr = Hour(strTime)
mn = Val(hr)
If booTime = False Then
mn = mn + 1
Else
mn = mn - 1
End If
moon = str(mn)
moon = Trim$(moon)
If Len(Trim$(moon)) = 1 Then
moon = "0" & moon
End If
If Len(Trim$(hr)) = 1 Then
hr = "0" & hr
End If
mi = Minute(strTime)
If Len(Trim$(mi)) = 1 Then
mi = "0" & mi
End If
se = Hour(strTime)
If Len(Trim$(se)) = 1 Then
se = "0" & se
End If

strNewTime = hr & ":" & mi & ":" & se
strMoon = moon & ":" & mi & ":" & se
End Sub

Private Sub cmdCalc_Click()
Dim srs As New clsSunRiseSet, NewTime As String, MoonTime As String
Dim MyYear As String

MyYear = Year(Now)
If optSouth.value = True Then
booSouth = True
Else
booSouth = False
End If
   ' srs.City = "London, England"
    txtDate.Text = "1/1/" & MyYear
    srs.CalculateSun
    Call ParseTime(srs.Sunrise, NewTime, MoonTime, False)
    RiseSet(7) = NewTime
    MoonSet(8) = MoonTime
    Call ParseTime(srs.Sunset, NewTime, MoonTime, True)
    RiseSet(8) = NewTime
    MoonSet(7) = MoonTime
    DoEvents
    txtDate.Text = "1/4/" & MyYear
    
    srs.CalculateSun
    Call ParseTime(srs.Sunrise, NewTime, MoonTime, False)
    RiseSet(1) = NewTime
    MoonSet(2) = MoonTime
    Call ParseTime(srs.Sunset, NewTime, MoonTime, True)
    RiseSet(2) = NewTime
    MoonSet(1) = MoonTime
    DoEvents
    txtDate.Text = "1/7/" & MyYear
    srs.CalculateSun
    Call ParseTime(srs.Sunrise, NewTime, MoonTime, False)
    RiseSet(3) = NewTime
    MoonSet(4) = MoonTime
    Call ParseTime(srs.Sunset, NewTime, MoonTime, True)
    RiseSet(4) = NewTime
    MoonSet(3) = MoonTime
    DoEvents
    txtDate.Text = "1/10/" & MyYear
    srs.CalculateSun
    Call ParseTime(srs.Sunrise, NewTime, MoonTime, False)
    RiseSet(5) = NewTime
    MoonSet(6) = MoonTime
    Call ParseTime(srs.Sunset, NewTime, MoonTime, True)
    RiseSet(6) = NewTime
    MoonSet(5) = MoonTime
    Unload Me
    
    End Sub
Private Sub cmdExit_Click()
booCancel = True

Unload Me
End Sub

Private Sub Form_Load()
Dim TimeDiff As Integer
Me.Caption = Lang(228)
Label2.Caption = Lang(229)
Label3.Caption = Lang(230)
Label4.Caption = Lang(231)
DaySave.Caption = Lang(232)
optTZWest.Caption = Lang(233)
optTZEast.Caption = Lang(234)
cmdCalc.Caption = Lang(235)
cmdExit.Caption = Lang(203)

Label1.Caption = Lang(366)

col = RStart(1)
row = RStart(2)
x = RStart(3)
z = RStart(4)
wt_ew_offset = -16385
 wt_ns_offset = 16385
 pixsiz = 2048
 ul_x = -20015000
 ul_y = 8673000
 R2D = 57.2957795130823
 OK = 1
 E_radius = 6370997#
 HALFPI = 1.5707963267949
 EPSLN = 0.0000000001
 PI = 3.14159265358979
        feast(0) = E_radius * -1.74532925199
        feast(1) = E_radius * -1.74532925199
        feast(2) = E_radius * 0.523598775598
        feast(3) = E_radius * 0.523598775598
        feast(4) = E_radius * -2.79252680319
        feast(5) = E_radius * -1.0471975512
        feast(6) = E_radius * -2.79252680319
        feast(7) = E_radius * -1.0471975512
        feast(8) = E_radius * 0.349065850399
        feast(9) = E_radius * 2.44346095279
        feast(10) = E_radius * 0.349065850399
        feast(11) = E_radius * 2.44306095279
        
        lon_centre(0) = -1.74532925199
        lon_centre(1) = -1.74532925199
        lon_centre(2) = 0.523598775598
        lon_centre(3) = 0.523598775598
        lon_centre(4) = -2.79252680319
        lon_centre(5) = -1.0471975512
        lon_centre(6) = -2.79252680319
        lon_centre(7) = -1.0471975512
        lon_centre(8) = 0.349065850399
        lon_centre(9) = 2.44346095279
        lon_centre(10) = 0.349065850399
        lon_centre(11) = 2.44346095279
 
Call main(col, row, x, z)
If Left$(retval(0), 1) = "-" Then
optSouth.value = True
txtLat = Mid$(retval(0), 2)
Else
optNorth.value = True
txtLat = retval(0)
End If
If Left$(retval(1), 1) = "-" Then
optWest.value = True
txtLon = Mid$(retval(1), 2)
Else
optEast.value = True
txtLon = retval(1)
End If
If Len(txtLon) > 6 Then
txtLon = Left$(txtLon, 6)
End If
If Len(txtLat) > 6 Then
txtLat = Left$(txtLat, 6)
End If

TimeDiff = Int(retval(1) / 15)
txtTZ.Text = Abs(TimeDiff)

End Sub

Public Sub main(col As Double, row As Double, x As Double, z As Double)
Dim tileX As Double, tileZ As Double

        tileX = ((col * 2048) + 1024) + x ' //absolute x value (in meters)
        tileZ = ((row * 2048) + 1024) + z ' //absolute z value (in meters)
                
        Call getLLfromMSTileM(tileX, tileZ)
End Sub

Public Function adjust_lon(xx As Double) As Double

    'adjust_lon=(fabs(xl)< PI ) ? xl : (xl-(sign(xl)*TWO_PI))
    If fabs(xx) < PI Then
    adjust_lon = xx
    Else
    adjust_lon = xx - (sign(xx) * (PI * 2))
    
End If
End Function

Public Function sign(x1 As Double) As Integer

        If (x1 < 0#) Then
               sign = (-1)
        Else
        sign = 1
End If
End Function
Public Function fabs(x1 As Double) As Double

        If (x1 < 0) Then
                fabs = (x1 * -1)
        Else
         fabs = x1
End If
End Function
        
Public Sub goodeInverse(gx As Double, gy As Double)
Dim Temp As Double, arg As Double, Theta As Double, Region As Integer

        If (gy >= E_radius * 0.710987989993) Then
        
                If (gx <= E_radius * -0.698131700798) Then
                Region = 0
                Else
                Region = 2
                End If
        
        ElseIf (gy >= 0#) Then
        
                If (gx <= E_radius * -0.698131700798) Then
                 Region = 1
                Else
                Region = 3
                End If
        
        ElseIf (gy >= E_radius * -0.710987989993) Then
        
                If (gx <= E_radius * -1.74532925199) Then
                Region = 4
                ElseIf (gx <= E_radius * -0.349065850399) Then
                Region = 5
                ElseIf (gx <= E_radius * 1.3962634016) Then
                Region = 8
                Else
                 Region = 9
                 End If
        
        Else
        
                If (gx <= E_radius * -1.74532925199) Then
                Region = 6
                ElseIf (gx <= E_radius * -0.349065850399) Then
                Region = 7
                ElseIf (gx <= E_radius * 1.3962634016) Then
                 Region = 10
                Else
                Region = 11
                End If
        End If
        
                
        gx = gx - feast(Region)
        
        If Region = 1 Or Region = 3 Or Region = 4 Or Region = 5 Or Region = 8 Or Region = 9 Then
        
        
                retval(0) = gy / E_radius
                If (fabs(retval(0)) > HALFPI) Then
                
                        retval(0) = 9999
                        retval(1) = 9999
                        Exit Sub
                End If
                
                 Temp = fabs(retval(0)) - HALFPI
                If (fabs(Temp) > EPSLN) Then
               ' If (fabs(Temp) < EPSLN) Then
                        Temp = lon_centre(Region) + gx / (E_radius * Cos(retval(0)))
                       retval(1) = adjust_lon(Temp)
                
                Else
                retval(1) = lon_centre(Region)
                End If
        
        Else
        
                arg = (gy + 0.0528035274542 * E_radius * sign(gy)) / (1.4142135623731 * E_radius)
                If (fabs(arg) > 1#) Then
                
                        retval(0) = 9999
                        retval(1) = 9999
                        Exit Sub
                End If
                Theta = Atn(arg / Sqr(-arg * arg + 1))
                retval(1) = lon_centre(Region) + (gx / (0.900316316158 * E_radius * Cos(Theta)))
                If (retval(1) < (-PI)) Then
                
                        retval(0) = 9999
                        retval(1) = 9999
                        Exit Sub
                End If
                arg = (2# * Theta + Sin(2# * Theta)) / PI
                If (fabs(arg) > 1#) Then
                
                        retval(0) = 9999
                        retval(1) = 9999
                        Exit Sub
                End If
                retval(0) = Atn(arg / Sqr(-arg * arg + 1)) ' asin(arg)
        End If
        Select Case Region
                Case 0
        If retval(1) < -PI Or retval(1) > -0.698131700798 Then
                retval(0) = 9999
                retval(1) = 9999
        End If
                Case 1
                If retval(1) < -PI Or retval(1) > -0.698131700798 Then
                        retval(0) = 9999
                retval(1) = 9999
        End If
                Case 2
                If retval(1) < -0.698131700798 Or retval(1) > PI Then
                        retval(0) = 9999
                retval(1) = 9999
        End If
                Case 3
                If retval(1) < -0.698131700798 Or retval(1) > PI Then
                        retval(0) = 9999
                retval(1) = 9999
        End If
                Case 4
                If retval(1) < -PI Or retval(1) > -1.74532925199 Then
                        retval(0) = 9999
                retval(1) = 9999
        End If
                Case 5
                If retval(1) < -1.74532925199 Or retval(1) > -0.349065850399 Then
                        retval(0) = 9999
                retval(1) = 9999
        End If
                Case 6
                If retval(1) < -PI Or retval(1) > -1.74532925199 Then
                        retval(0) = 9999
                retval(1) = 9999
        End If
                Case 7
                If retval(1) < -1.74532925199 Or retval(1) > -0.349065850399 Then
                        retval(0) = 9999
                retval(1) = 9999
        End If
                Case 8
                If retval(1) < -0.349065850399 Or retval(1) > 1.3962634016 Then
                        retval(0) = 9999
                retval(1) = 9999
        End If
                Case 9
                If retval(1) < 1.3962634016 Or retval(1) > PI Then
                        retval(0) = 9999
                retval(1) = 9999
        End If
                Case 10
                If retval(1) < -0.349065850399 Or retval(1) > 1.3962634016 Then
                        retval(0) = 9999
                retval(1) = 9999
        End If
                Case 11
                If retval(1) < 1.3962634016 Or retval(1) > PI Then
        
                retval(0) = 9999
                retval(1) = 9999
                
                End If
        End Select
        'End If
        If retval(0) = 9999 Then Exit Sub
        '//convert from radians to degrees
        retval(0) = R2D * retval(0)
        retval(1) = R2D * retval(1)
        
End Sub
        
Private Sub getLLfromMSTileM(tileX As Double, tileZ As Double)
Dim samp As Double, line As Double, gx As Double, gy As Double
Rem convert the abs MSTS coords into Goode grid coords

        tileX = tileX / 2048
        tileZ = tileZ / 2048
         samp = (tileX - wt_ew_offset)
         line = (wt_ns_offset - tileZ)
         gx = ul_x + ((samp - 1) * pixsiz)
         gy = ul_y - ((line - 1) * pixsiz)
                

        '//perform the Inverse Goode Projection
        Call goodeInverse(gx, gy)
        '//return values of 9999,9999 indicates that the values failed to converge
End Sub


Private Sub optEast_Click()
optTZEast = True
optTZEast.ForeColor = &H80000012
End Sub

Private Sub optWest_Click()
optTZWest = True
optTZWest.ForeColor = &H80000012
End Sub
Private Sub optTZWest_Click()
If optWest = True Then optTZWest.ForeColor = &H80000012
If optEast = True Then optTZWest.ForeColor = &HFF&
End Sub
Private Sub optTZEast_Click()
If optEast = True Then optTZWest.ForeColor = &H80000012
If optWest = True Then optTZEast.ForeColor = &HFF&
End Sub
Public Function Julianday(Day, Month, Year As Integer, dayfraction As Double) As Double

Dim a%, b%
    If Month < 3 Then Month = Month + 12: Year = Year - 1
    a = Int(Year / 100)
    b = 2 - a + Int(a / 4)
Julianday = Int(365.25 * (Year + 4716#)) + Int(30.6001 * (Month + 1)) + Day + dayfraction + b - 1524.5
End Function
Public Function JD2CalendarDate(JD As Double) As String

Dim z&, Alpha&, a&, b&, c&, d&, E&, Day%, Month%, Year%

z = Int(JD + 0.5)
Alpha = Int((z - 1867216.25) / 36524.25)
a = z + 1 + Alpha - Int(Alpha / 4)
b = a + 1524
c = Int((b - 122.1) / 365.25)
d = Int(365.25 * c)
E = Int((b - d) / 30.6001)
Day = b - d - Int(30.6001 * E)

If E < 14 Then Month = E - 1 Else: Month = E - 13
If Month < 3 Then Year = c - 4715 Else: Year = c - 4716

JD2CalendarDate = Format(Day, "00") & "-" & Format(Month, "00") & "-" & Year

End Function
Public Function JDSidGreenwich(Julianday As Double) As Double

Dim T#, Sid#

T = (Julianday - 2451545#) / 36525#
Sid = 280.46061837 + 360.98564736629 * (Julianday - 2451545#) + T * T * (0.000387933 - (1 / 38710000#) * T)
JDSidGreenwich = (Sid - 360# * Int(Sid / 360#))
End Function
Public Function JDRightAscSun(Julianday As Double) As Double

Dim Deg2Rad#, T#, Obl#
Dim MASun#, Center#, MLSun#, TLSun#, RASun#

Deg2Rad = Atn(1#) / 45#
Obl = 23.439

T = (Julianday - 2451545#) / 36525#

MASun = 357.52911 + T * (35999.05029 - 0.0001537 * T)
    Do Until MASun > 0#
    MASun = MASun + 360#
    Loop
    
    Do Until MASun < 360#
    MASun = MASun - 360#
    Loop

Center = (1.914602 - T * (0.004817 - 0.000014 * T)) * Sin(MASun * Deg2Rad) + (0.019993 - 0.000101 * T) * Sin(2# * MASun * Deg2Rad) + 0.000289 * Sin(3# * MASun * Deg2Rad)

MLSun = 280.4664567 + T * (36000.76983 + 0.0003032 * T)
    Do Until MLSun > 0#
    MLSun = MLSun + 360#
    Loop
    
    Do Until MLSun < 360#
    MLSun = MLSun - 360#
    Loop

TLSun = MLSun + Center
    If TLSun > 360# Then TLSun = TLSun - 360#

RASun = (1 / Deg2Rad) * Atn(Sin(TLSun * Deg2Rad) * Cos(Obl * Deg2Rad) / Cos(TLSun * Deg2Rad))
    If Cos(TLSun * Deg2Rad) < 0# Then RASun = RASun + 180#
    If RASun < 0# Then RASun = RASun + 360#

JDRightAscSun = RASun

End Function
Public Function JDDeclSun(Julianday As Double) As Double

Dim Deg2Rad#, Obl#, T#
Dim MASun#, Center#, MLSun#, TLSun#, xa#

Deg2Rad = Atn(1#) / 45#
Obl = 23.439

T = (Julianday - 2451545#) / 36525#

MASun = 357.52911 + T * (35999.05029 - 0.0001537 * T)
    Do Until MASun > 0#
    MASun = MASun + 360#
    Loop
    
    Do Until MASun < 360#
    MASun = MASun - 360#
    Loop

Center = (1.914602 - T * (0.004817 - 0.000014 * T)) * Sin(MASun * Deg2Rad) + (0.019993 - 0.000101 * T) * Sin(2# * MASun * Deg2Rad) + 0.000289 * Sin(3# * MASun * Deg2Rad)

MLSun = 280.46646 + T * (36000.76983 + 0.003032 * T)
    Do Until MLSun > 0#
    MLSun = MLSun + 360#
    Loop
    
    Do Until MLSun < 360#
    MLSun = MLSun - 360#
    Loop

TLSun = MLSun + Center
    If TLSun > 360# Then TLSun = TLSun - 360#

xa = Sin(Obl * Deg2Rad) * Sin(TLSun * Deg2Rad)
JDDeclSun = (1# / Deg2Rad) * Atn(xa / Sqr(-xa * xa + 1#))

End Function
Public Function Asin(x As Double) As Double
Dim Deg2Rad As Double
Deg2Rad = Atn(1#) / 45#
Asin = (1# / Deg2Rad) * Atn(x / Sqr(-x * x + 1#))
End Function

