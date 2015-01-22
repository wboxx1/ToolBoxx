Imports ToolBoxx.Math
Imports System.Math

Public Class GPS
    
    Const earthRadius As Double = 6371000 'meters

    ''' <summary>
    ''' Calculates distance between coordinates in meters
    ''' </summary>
    ''' <param name="aLatitude1"></param>
    ''' <param name="aLongitude1"></param>
    ''' <param name="aLatitude2"></param>
    ''' <param name="aLongitude2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Shared Function DistanceBetweenCoordinatePoints(ByVal aLatitude1 As String, ByVal aLongitude1 As String, ByVal aLatitude2 As String, ByVal aLongitude2 As String) As Double

        Dim latitude1 As Double = 0
        Dim latitude2 As Double = 0
        Dim longitude1 As Double = 0
        Dim longitude2 As Double = 0

        'check what format each coordinate is: Degrees Minutes Seconds (DMS) or Decimal Degrees

        If aLatitude1 Like "*'*" Then
            latitude1 = convertDMStoRadians(aLatitude1)
        Else
            latitude1 = convertDegreesToRadians(CDbl(aLatitude1))
        End If

        If aLongitude1 Like "*'*" Then
            longitude1 = convertDMStoRadians(aLongitude1)
        Else
            longitude1 = convertDegreesToRadians(CDbl(aLongitude1))
        End If

        If aLatitude2 Like "*'*" Then
            latitude2 = convertDMStoRadians(aLatitude2)
        Else
            latitude2 = convertDegreesToRadians(CDbl(aLatitude2))
        End If

        If aLongitude2 Like "*'*" Then
            longitude2 = convertDMStoRadians(aLongitude2)
        Else
            longitude2 = convertDegreesToRadians(CDbl(aLongitude2))
        End If


        Dim latDelta As Double = latitude2 - latitude1
        Dim longDelta As Double = longitude2 - longitude1

        'Using the Haversine formula
        'a = sin^2(delLat/2) + cos(lat1)*cos(lat2)*sin^2(delLong/2)
        'c = 2*arctan(sqrt(a)/ sqrt(1-a))
        'd = R*c

        Dim A As Double = Sin(latDelta / 2) * Sin(latDelta / 2) + Cos(latitude1) * Cos(latitude2) * Sin(longDelta / 2) * Sin(longDelta / 2)
        Dim C As Double = 2 * Atan2(Sqrt(A), Sqrt(1 - A))
        Dim D As Double = earthRadius * C
        Return D

    End Function
End Class
