
Public Class Math
    Const PI = 3.1415926535897931

    Shared Function convertDMStoRadians(ByVal aCoordinate As String) As Double
        Dim tempStr() As String = aCoordinate.Split(" ")
        Dim wholeDegrees As Double = CDbl(tempStr(0))
        Dim minSec() As String = tempStr(1).Split("'")
        Dim minutes As Double = CDbl(minSec(0))
        Dim seconds As Double = 0
        If tempStr.Length <= 2 Then
            'skip
        ElseIf tempStr(2) = "" Then
            'skip
        Else
            Dim remainingStr() As String = tempStr(2).Split(Chr(34))
            seconds = CDbl(remainingStr(0))
        End If

        'convert DMS to decimal degrees
        Dim temp As Double = (wholeDegrees + (minutes * 60 + seconds) / 3600)
        temp = convertDegreesToRadians(temp)
        Return temp

    End Function

    Shared Function convertDegreesToRadians(ByVal aValue As Double)
        Return aValue * PI / 180
    End Function

    
End Class
