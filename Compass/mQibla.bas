Attribute VB_Name = "mQibla"
'==========================================================
'           Copyright Information
'==========================================================
'Program Name     : Mewsoft Qibla Direction Compass
'Program Author   : Elsheshtawy, A. A.
'Home Page        : http://www.mewsoft.com
'Home Page        : http://www.islamware.com
'Copyrights © 2006 Mewsoft Corporation. All rights reserved.
'==========================================================
'==========================================================
'Muslim Qibla Direction Compass, Great Circle Distance
'and Great Circle Direction Calculator.

'Qibla is an Arabic word referring to the direction that should be
'faced when a Muslim prays. This program calculates the Qibla direction
'from any point on the Earth. It also uses and claculates
'the Great Circle Distance and the Great Circle Direction.
'==========================================================
'==========================================================

Option Explicit

Public Const pi As Double = 3.14159265358979    ' PI=22/7, Pi = Atn(1) * 4
Public Const DtoR As Double = (pi / 180#)       ' Degree to Radians
Public Const RtoD As Double = (180# / pi)       ' Radians to Degrees

'====================================================================
'*  The Coordinates for the KABAH (in Makkah, Saudi Arabia) are:         *
'*            KABAH  Latitude:   21 deg  27 minutes  North               *
'*            KABAH  Longitude:  39 deg  49 minutes  East                *
'*                                                                       *
'   Kabah Lat=21 Deg N, Long 40 Deg E
'   Makkah: Lat=21.4, Long=39.8
'   Cairo: Lat=30.1, Long=31.3
'====================================================================
'Inverse Cosine
Public Function Arccos(ByVal X As Double) As Double
    
    Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
    
End Function

'====================================================================
'Hints
'1. To convert degrees to radians, first convert the number of degrees, minutes,
' and seconds to decimal form. Divide the number of minutes by 60 and add to the
' number of degrees. So, for example, 12° 28' is 12 + 28/60 which equals 12.467°.
' Next multiply by  and divide by 180 to get the angle in radians.

'2. Conversely, to convert radians to degrees divide by  and multiply by 180.
' So, 0.47623 divided by  and multiplied by 180 gives 27.286°. You can convert
' the fractions of a degree to minutes and seconds as follows. Multiply the
' fraction by 60 to get the number of minutes. Here, 0.286 times 60 equals
' 17.16, so the angle could be written as 27° 17.16'. Then take any fraction
' of a minute that remains and multiply by 60 again to get the number of seconds.
' Here, 0.16 times 60 equals about 10, so the angle can also be written as
' 27° 17' 10".

' Converting from Degrees, Minutes and Seconds to Decimal Degrees
' Example: 30 degrees 15 minutes 22 seconds = 30 + 15/60 + 22/3600 = 30.2561
'-----------------------------------------------------------------------------
'Converting Between Decimal Degrees, Degrees, Minutes and Seconds, and Radians
'(dd + mm/60 +ss/3600) to Decimal degrees (dd.ff)
'dd = whole degrees, mm = minutes, ss = seconds
'dd.ff = dd + mm / 60 + ss / 3600
'Example: 30 degrees 15 minutes 22 seconds = 30 + 15/60 + 22/3600 = 30.2561
'----------------------------------
'Decimal degrees (dd.ff) to (dd + mm/60 +ss/3600)
'For the reverse conversion, we want to convert dd.ff to dd mm ss. Here ff = the fractional part of a decimal degree.
'mm = 60 * ff
'ss = 60*(fractional part of mm)
'Use only the whole number part of mm in the final result.
'30.2561 degrees = 30 degrees
'.2561*60 = 15.366 minutes
'.366 minutes = 22 seconds, so the final result is 30 degrees 15 minutes 22 seconds
'----------------------------------
'Decimal degrees (dd.ff) to Radians
'Radians = (dd.ff) * pi / 180
'----------------------------------
'Radians to Decimal degrees (dd.ff)
'(dd.ff) = Radians*180/pi
'----------------------------------

Public Function DegreeToDecimal(ByVal Degrees As Double, _
        ByVal Minutes As Double, ByVal Seconds As Double) As Double
    
    DegreeToDecimal = Degrees + Minutes / 60 + Seconds / 3600
    
End Function

'====================================================================
'Converting from Decimal Degrees to Degrees, Minutes and Seconds
Public Sub DecimalToDegree(ByVal DecimalDegree As Double, _
        ByRef Degrees As Integer, ByRef Minutes As Integer, _
        ByRef Seconds As Integer)
    
    Dim ff As Double
    
    Degrees = Fix(DecimalDegree)
    ff = DecimalDegree - Degrees
    Minutes = Fix(60 * ff)
    Seconds = 60 * ((60 * ff) - Minutes)
    
End Sub

'====================================================================
' The shortest distance between points 1 and 2 on the earth's surface is
' d = arccos{cos(Dlat) - [1 - cos(Dlong)]cos(lat1)cos(lat2)}
' Dlat = lab - lat2
' Dlong = 10ng• - long2
' lati, = latitude of point i
' longi, = longitude of point i

' The following are the mathematical formulas for calculating great circle distance and bearing.
' The following are the conversion factors, one nautical mile equals to:
'   6076.10 feet
'   2027 yards
'   1.852 kilometers
'   1.151 statute mile
' Trigonometric computation in BASIC employs radians,
' not degrees, so each input datum must be divided by 57.2958 (which is 180/pi,
' the number of degrees in a radian).

'Conversion of grad to degrees is as follows:
'Grad=400-degrees/0.9 or Degrees=0.9x(400-Grad)

'Latitude is determined by the earth's polar axis. Longitude is determined
'by the earth's rotation. If you can see the stars and have a sextant and
'a good clock set to Greenwich time, you can find your latitude and longitude.

' Calculates the distance between any two points on the Earth
Public Function GreatCircleDistance(ByVal OriginLatitude As Double, _
    ByVal DestinationLatitude As Double, ByVal OriginLongitude As Double, _
    ByVal DestinationLongitude As Double) As Double

    Dim D As Double
    Dim L1 As Double, L2 As Double
    Dim I1 As Double, I2 As Double
    
    L1 = OriginLatitude * DtoR
    L2 = DestinationLatitude * DtoR
    I1 = OriginLongitude * DtoR
    I2 = DestinationLongitude * DtoR
    
    D = Arccos(Cos(L1 - L2) - (1 - Cos(I1 - I2)) * Cos(L1) * Cos(L2))
    GreatCircleDistance = D * 60 * RtoD
    ' One degree of such an arc on the earth's surface is 60
    ' international nautical miles NM
    
End Function

'====================================================================
' Calculates the direction from one point to another on the Earth
' a = arccos{[sin(lat2) - cos(d + lat1 - 1.5708)]/cos(lat1)/sin(d) + 1}
' Great Circle Bearing

Public Function GreatCircleDirection(ByVal OriginLatitude As Double, _
    ByVal DestinationLatitude As Double, ByVal OriginLongitude As Double, _
    ByVal DestinationLongitude As Double, ByVal Distance As Double) As Double
    
    On Error GoTo ErrHandler
    
    Dim A As Double, B As Double
    Dim D As Double
    Dim L1 As Double, L2 As Double
    Dim I1 As Double, I2 As Double
    Dim Result As Double
    Dim Dlong As Double
    
    L1 = OriginLatitude * DtoR
    L2 = DestinationLatitude * DtoR
    D = (Distance / 60) * DtoR ' divide by 60 for nautical miles NM to degree
    
    I1 = OriginLongitude * DtoR
    I2 = DestinationLongitude * DtoR
    Dlong = I1 - I2
    
    ' Pi/2 = 1.5708
    A = Sin(L2) - Cos(D + L1 - pi / 2)
    B = Arccos(A / (Cos(L1) * Sin(D)) + 1)
    
    'If (Abs(Dlong) < pi And Dlong < 0) Or (Abs(Dlong) > pi And Dlong > 0) Then
    '        Result = (2 * pi) - B
    'Else
    '        Result = B
    'End If
    
    Result = B
    GreatCircleDirection = Result * RtoD
    Exit Function
     
ErrHandler:
    GreatCircleDirection = 0
    MsgBox ("Error calculating Bearing. " & Err.Description)
End Function

'====================================================================
'The Equivalent Earth redius is 6,378.14 Kilometers.

' Calculates the direction of the Qibla from any point on
' the Earth From North Clocklwise

Public Function QiblaDirection(ByVal OriginLatitude As Double, ByVal OriginLongitude As Double) As Double

    Dim Distance As Double, Bearing As Double
    Dim L1 As Double, L2 As Double
    Dim I1 As Double, I2 As Double
    
    L1 = OriginLatitude
    I1 = OriginLongitude
    
    ' Kabah Lat=21 Deg N, Long 40 Deg E
    L2 = 21: I2 = 40 ' Kabah
    
    Distance = GreatCircleDistance(L1, L2, I1, I2)
    
    Bearing = GreatCircleDirection(L1, L2, I1, I2, Distance)
    
    QiblaDirection = Bearing
    
End Function

'====================================================================
'====================================================================
' Calculates the distance between any two points on the Earth
Public Function GreatCircleDistance1(ByVal OriginLatitude As Double, _
    ByVal DestinationLatitude As Double, ByVal OriginLongitude As Double, _
    ByVal DestinationLongitude As Double) As Double

    Dim A As Double, B As Double
    Dim C As Double, D As Double
    Dim L1 As Double, L2 As Double
    Dim I1 As Double, I2 As Double
    
    L1 = OriginLatitude * DtoR
    I1 = OriginLongitude * DtoR
    L2 = DestinationLatitude * DtoR
    I2 = DestinationLongitude * DtoR
    
    A = Sin(L1) * Sin(L2)
    B = Cos(L1) * Cos(L2)
    C = Cos(I2 - I1)
    D = A + (B * C)
       
    GreatCircleDistance1 = 60 * Arccos(D) * RtoD
    'Distance (in Nautical Miles)
    
End Function

'====================================================================
' Calculates the direction from one point to another on the Earth
'                   NOT COMPLETED YET
Public Function GreatCircleDirection1(ByVal OriginLatitude As Double, _
    ByVal DestinationLatitude As Double, ByVal Distance As Double) As Double
    
    On Error GoTo ErrHandler
    
    Dim A As Double, B As Double
    Dim C As Double, D As Double
    Dim L1 As Double, L2 As Double
    
    L1 = OriginLatitude * DtoR
    L2 = DestinationLatitude * DtoR
    D = Distance * DtoR
    
    A = (Sin(L1) * Sin(L2)) + Cos(D / 60)
    
    B = Sin(D / 60) * Cos(L1)
    
    C = (A / B)
    GreatCircleDirection1 = Arccos(C) * RtoD
    ' Course (in degrees)
    
    Exit Function

ErrHandler:
    GreatCircleDirection1 = 0
    MsgBox ("Error calculating Bearing. " & Err.Description)
    
End Function
'====================================================================

