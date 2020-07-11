Attribute VB_Name = "Matematicas"

Option Explicit

Public Function Porcentaje(ByVal Total As Long, ByVal Porc As Long) As Long
    Porcentaje = (Total * Porc) * 0.01
End Function

Function Distancia(ByRef wp1 As WorldPos, ByRef wp2 As WorldPos) As Long
On Error GoTo hayerror
    'Encuentra la distancia entre dos WorldPos
    Distancia = Abs(wp1.x - wp2.x) + Abs(wp1.Y - wp2.Y) + (Abs(wp1.map - wp2.map) * 100)

Exit Function
hayerror:
    LogError ("Error en Distancia: " & err.description)


End Function
Function RangoVision(ByRef wp1 As WorldPos, ByRef wp2 As WorldPos, ByVal tHeading As Byte) As Boolean
    Dim SignoNS As Integer
    Dim SignoEO As Integer

    Select Case tHeading
        Case eHeading.NORTH
            SignoNS = -1: SignoEO = 0
        Case eHeading.EAST
            SignoNS = 0: SignoEO = 1
        Case eHeading.SOUTH
            SignoNS = 1: SignoEO = 0
        Case eHeading.WEST
            SignoEO = -1: SignoNS = 0
    End Select
    
    If Abs(wp1.x - wp2.x) <= RANGO_VISION_X And _
       Sgn(wp1.x - wp2.x) = SignoEO Then
        If Abs(wp1.Y - wp2.Y) <= RANGO_VISION_Y And _
           Sgn(wp1.Y - wp2.Y) = SignoNS Then
            RangoVision = True
            Exit Function
        End If
    End If
    RangoVision = False
End Function
Public Function sAbs(ByVal val As Long) As Long
    If val < 0 Then
        sAbs = Not val + 1
    Else
        sAbs = val
    End If
End Function

Function Distance(X1 As Variant, Y1 As Variant, X2 As Variant, Y2 As Variant) As Double

'Encuentra la distancia entre dos puntos

Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))

End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Generates a random number in the range given - recoded to use longs and work properly with ranges
'**************************************************************
    RandomNumber = Fix(Rnd * (UpperBound - LowerBound + 1)) + LowerBound
End Function
