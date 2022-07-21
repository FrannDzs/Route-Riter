Attribute VB_Name = "ConvertXY"
 Global Pie As Variant, HalfPie As Variant, TwoPie As Variant
 Global E_Radius As Variant, EPSLN As Variant, R2D As Variant
 Global D2R As Variant
 
 Private Lon_Center(12) As Double, FEast(12) As Double
 Private pixsiz As Integer, ul_x As Long, ul_y As Long
 Private nl As Integer, ns As Integer
 Private wt_ew_offset As Integer, wt_ns_offset As Integer
 Private X, Y

Function SetVariables()

E_Radius = 6370997                              'earth radius
Pie = CDec(3.14159265358979) + CDec(3.238E-15)  'pi
HalfPie = CDec(Pie * 0.5)                       '.5 * pi
TwoPie = CDec(Pie * 2)                          '2*pi
EPSLN = 0.0000000001                            'error factor
R2D = 57.2957795131                             'radian to degree
D2R = 0.0174532925199                           'degree to radian

pixsiz = 2048                                   'size of tile in meters

ul_x = -20015000 '20015000 appears to be -180 deg in Goode projection
ul_y = 8673000      'appears to be +90 deg lat in Goode projection
nl = 8471 '(Abs(ul_y) * 2) / pixsiz
ns = 19547 '(Abs(ul_x) * 2) / pixsiz

'the Upperleft corner of the Goode is ul_x,ul_y
'the bottomright corner of the goode is -ul_x,-ul_y

'offsets to convert Goode raster coord to MSTS world tile coord
wt_ew_offset = -16385
wt_ns_offset = 16385

'set scroll bars to 0
'Form1.HScroll1 = 0
'Form1.VScroll1 = 0

'intilize the goode projection
GoodeInit


End Function
Private Function GoodeInit()
'initialize central meridians for each of the 12 regions
Lon_Center(0) = -1.74532925199   '-100.0 degrees
Lon_Center(1) = -1.74532925199   '-100.0 degrees
Lon_Center(2) = 0.523598775598   '  30.0 degrees
Lon_Center(3) = 0.523598775598   '  30.0 degrees
Lon_Center(4) = -2.79252680319   '-160.0 degrees
Lon_Center(5) = -1.0471975512    ' -60.0 degrees
Lon_Center(6) = -2.79252680319   '-160.0 degrees
Lon_Center(7) = -1.0471975512    ' -60.0 degrees
Lon_Center(8) = 0.349065850399   '  20.0 degrees
Lon_Center(9) = 2.44346095279    ' 140.0 degrees
Lon_Center(10) = 0.349065850399 '  20.0 degrees
Lon_Center(11) = 2.44346095279   ' 140.0 degrees

'init false easting for each of te 12 regions
FEast(0) = E_Radius * -1.74532925199
FEast(1) = E_Radius * -1.74532925199
FEast(2) = E_Radius * 0.523598775598
FEast(3) = E_Radius * 0.523598775598
FEast(4) = E_Radius * -2.79252680319
FEast(5) = E_Radius * -1.0471975512
FEast(6) = E_Radius * -2.79252680319
FEast(7) = E_Radius * -1.0471975512
FEast(8) = E_Radius * 0.349065850399
FEast(9) = E_Radius * 2.44346095279
FEast(10) = E_Radius * 0.349065850399
FEast(11) = E_Radius * 2.44346095279
End Function
Function ConvertL(ByVal Lon, ByVal Lat)
'decimal degrees is assumed
'get goode x,y
If Goode_Forward(CDec(Lon * D2R), CDec(Lat * D2R)) Then

    'convert goode raster coord
    Gline = CDec((ul_y - Y) / pixsiz + 1)
    GSamp = CDec((X - ul_x) / pixsiz + 1)

'emit MSTS world tile coord, if calculated
    WorldTX = Int(CDec(GSamp + wt_ew_offset))
    WorldTY = Int(CDec(wt_ns_offset - Gline))
    'for debugging
    'Form1.Text1 = Round(WorldTX, 0)
    'Form1.Text2 = Round(WorldTY, 0)
    frmUtils.Text3(3) = WorldTX
    frmUtils.Text3(4) = WorldTY

Else
    frmUtils.Text3(3) = "There was an error."
End If
End Function
Private Function Goode_Forward(Lon, Lat) As Boolean
'forward equations
If Lat >= 0.710987929993 Then
        'if on or above 40 44' 11.8"
        If Lon <= -0.698131700798 Then
            'if to the left of -40
            Region = 0
        Else
            Region = 2
        End If
ElseIf Lat >= 0 Then
        'between 0.0 and 40 44' 11.8"
        If Lon <= -0.698131700798 Then
            'if to the left of -40
            Region = 1
        Else
            Region = 3
        End If
ElseIf Lat >= -0.710987989993 Then
        'between 0.0 and -40 44' 11.8"
        If Lon <= -1.74532925199 Then
                Region = 4    'if between -180 and -100
        ElseIf Lon <= -0.349065850399 Then
                Region = 5   'between -100 and -20
        ElseIf Lon <= 1.3962634016 Then
                Region = 8      'between -20 and 80
        Else
                'between 80 and 180
                Region = 9
        End If
Else
        'below -40 44'
        If Lon <= -1.74532925199 Then
                'between -180 and -100
                Region = 6
        ElseIf Lon <= -0.349065850399 Then
                'between -100 and -20
                Region = 7
        ElseIf Lon <= 1.3962634016 Then
                'between -20 and 80
                Region = 10
        Else
                'between 80 and 180
                Region = 11
        End If
End If

Select Case (Region)
    Case 1, 3, 4, 5, 8, 9: 'select case
        delta_lon = CDec(Adjust_Lon(Lon - Lon_Center(Region)))
        X = CDec(FEast(Region) + E_Radius * delta_lon * Cos(Lat))
        Y = CDec(E_Radius * Lat)
    Case Else
        delta_lon = CDec(Adjust_Lon(Lon - Lon_Center(Region)))
        Theta = Lat
        constant = CDec(Pie * Sin(Lat))
        
        'iterlate using the Newton-Raphson method to find theta
        working = False
        For i = 0 To 30
            delta_theta = CDec(-(Theta + Sin(Theta) - constant) / (1 + Cos(Theta)))
            Theta = CDec(Theta + delta_theta)
            If Fabs(delta_theta) < EPSLN Then
                working = True
                Exit For
            End If
        Next i
        If working = False Then
            Goode_Forward = False
            Exit Function
        End If
        Theta = CDec(Theta / 2) 'original
        X = CDec(FEast(Region) + 0.900316316158 * E_Radius * delta_lon * Cos(Theta))
        Y = CDec(E_Radius * (1.4142135623731 * Sin(Theta) - 0.0528035274542 * Sign(Lat)))
End Select
Goode_Forward = True

End Function
Private Function Sign(H)
'this will return the sign of a value
If H < 0 Then
    Sign = -1
Else
    Sign = 1
End If
End Function
Private Function Adjust_Lon(ByVal Z)
If Fabs(Z) > Pie Then
    Adjust_Lon = Z - (Sign(Z) * TwoPie)
Else
    Adjust_Lon = Z
End If

End Function
Private Function Fabs(X1)
If X1 < 0 Then
    Fabs = X1 * -1
Else
    Fabs = X1
End If

End Function
Private Function asin(ByVal X1)
'asin = atan2(X1 / sqrt(-X1 * X1 + 1), 1)
asin = CDec(Math.Atn((X1 / Sqr(-X1 * X1 + 1)) / 1))
End Function
Public Function Goode_Inverse(ByVal GX, ByVal GY) As Integer
'Goode's Homolosine inverse equations
'mapping x,y to lat, lon
'gx and gy must be offset in order to be in raw goode coord
'this may alter lon and lat values
Dim Region As Integer
'need to have offsets placed on values

'Inverse equations
If GY >= E_Radius * 0.710987989993 Then         'if on or above 40 44' 11.8"
    If GX <= E_Radius * -0.698131700798 Then    'if to the left of -40
        Region = 0
    Else
        Region = 2
    End If
ElseIf GY >= 0 Then                             'between 0.0 and 40 44' 11.8"
    If GX <= E_Radius * -0.698131700798 Then    'if to the left of -40
        Region = 1
    Else
        Region = 3
    End If
ElseIf GY >= E_Radius * -0.710987989993 Then    'between 0.0 and -40 44' 11.8"
    If GX <= E_Radius * -1.74532925199 Then
        Region = 4      'if between -180 and -100
    ElseIf GX <= E_Radius * -0.349065850399 Then
        Region = 5 'if between -100 and -20
    ElseIf GX <= E_Radius * 1.3962634016 Then
        Region = 8    'if between -20 and 80
    Else: Region = 9 'if between 80 and 180
    End If
Else
    If GX <= E_Radius * -1.74532925199 Then
        Region = 6                                  'if between -180 and -100
    ElseIf GX <= E_Radius * -0.349065850399 Then
        Region = 5                                  'if between -100 and -20
    ElseIf GX <= E_Radius * 1.3962634016 Then
        Region = 10                                 'if between -20 and 80
    Else: Region = 11                               'if between 80 and 180
    End If
End If
GX = GX - FEast(Region)
Select Case (Region)
    Case 1, 3, 4, 5, 8, 9:
        Lat = GY / E_Radius
        If Fabs(Lat) > HalfPie Then
            'return(error) return -2
            Goode_Inverse = -1
            Exit Function
        End If
        Temp = Fabs(Lat) - HalfPie
        If Fabs(Temp) > EPSLN Then
            Temp = Lon_Center(Region) + GX / (E_Radius * Cos(Lat))
            Lon = Adjust_Lon(Temp)
        Else
            Lon = Lon_Center(Temp)
        End If
    Case Else:
        arg = (GY + 0.0528035274542 * E_Radius * Sign(GY)) / (1.4142135623731 * E_Radius)
        If Fabs(arg) > 1 Then
            'return(In_break) return -2
            Goode_Inverse = -2
            Exit Function
        End If
        Theta = asin(arg)
        Lon = Lon_Center(Region) + (GX / (0.900316316158 * E_Radius * Cos(Theta)))
        If Lon < -Pie Then
            'return(In_break) return -2
            Goode_Inverse = -2
            Exit Function
        End If
        arg = (2 * Theta + Sin(2 * Theta)) / Pie
        If Fabs(arg) > 1 Then
            'return(In_break) return -2
            Goode_Inverse = -2
            Exit Function
        End If
        Lat = asin(arg)
End Select

'are we in a interrupted area? if so, return status code on in_break
Select Case (Region)
    Case 0:
            If Lon < -Pie Or Lon > -0.698131700798 Then
                'return(In_break) return -2
                Goode_Inverse = -2
                Exit Function
            End If
    Case 1:
            If Lon < -Pie Or Lon > -0.698131700798 Then
                'return(In_break) return -2
                Goode_Inverse = -2
                Exit Function
            End If
    Case 2:
            If Lon < -0.698131700798 Or Lon > Pie Then
                'return(In_break) return -2
                Goode_Inverse = -2
                Exit Function
            End If
    Case 3:
            If Lon < -0.698131700798 Or Lon > Pie Then
                'return(In_break) return -2
                Goode_Inverse = -2
                Exit Function
            End If
    Case 4:
            If Lon < -Pie Or Lon > -1.74532925199 Then
                'return(In_break) return -2
                Goode_Inverse = -2
                Exit Function
            End If
    Case 5:
            If Lon < -1.74532925199 Or Lon > -0.349065850399 Then
                'return(In_break) return -2
                Goode_Inverse = -2
                Exit Function
            End If
    Case 6:
            If Lon < -Pie Or Lon > -1.74532925199 Then
                'return(In_break) return -2
                Goode_Inverse = -2
                Exit Function
            End If
    Case 7:
            If Lon < -1.74532925199 Or Lon > -0.349065850399 Then
                'return(In_break) return -2
                Goode_Inverse = -2
                Exit Function
            End If
    Case 8:
            If Lon < -0.349065850399 Or Lon > 1.3962634016 Then
                'return(In_break) return -2
                Goode_Inverse = -2
                Exit Function
            End If
    Case 9:
            If Lon < 1.3962634016 Or Lon > Pie Then
                'return(In_break) return -2
                Goode_Inverse = -2
                Exit Function
            End If
    Case 10:
            If Lon < -0.349065850399 Or Lon > 1.3962634016 Then
                'return(In_break) return -2
                Goode_Inverse = -2
                Exit Function
            End If
    Case 11:
            If Lon < 1.3962634016 Or Lon > Pie Then
                'return(In_break) return -2
                Goode_Inverse = -2
                Exit Function
            End If
End Select
frmUtils.Text3(0) = Lon * R2D
frmUtils.Text3(1) = Lat * R2D
Goode_Inverse = 1
End Function
Function ConvertWTC(ByVal wt_ew_dat, ByVal wt_ns_dat)
'decimal degrees is assumed
GSamp = CDec((wt_ew_dat - wt_ew_offset))  'GSamp is Goode world tile x
Gline = CDec((wt_ns_offset - wt_ns_dat))  'GLine is Goode world tile Y
Y = CDec(ul_y - ((Gline - 1) * pixsiz))   'actual Goode X
X = CDec(ul_x + ((GSamp - 1) * pixsiz))   'actual Goode Y

'get lon, lat from x,y
success = Goode_Inverse(X, Y)
End Function

