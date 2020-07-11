Attribute VB_Name = "AoDefenderEncrypt2"
Option Explicit


Public s As String


Public Function Semilla(strClave As String) As String

    Dim lngSemilla1 As Long
    Dim lngSemilla2 As Long
    Dim j As Long
    Dim i As Long
    lngSemilla1 = 0
    lngSemilla2 = 0
    j = Len(strClave)

    For i = 1 To Len(strClave)

        lngSemilla1 = lngSemilla1 + Asc(mid$(strClave, i, 1)) * i
        lngSemilla2 = lngSemilla2 + Asc(mid$(strClave, i, 1)) * j
        j = j - 1

    Next

    Semilla = LTrim$(Str$(lngSemilla1)) + "," + LTrim$(Str$(lngSemilla2))

End Function

Public Function Codificar(strCadena As String, strSemilla As String) As String

    Dim lngIi1 As Long
    Dim lngIi2 As Long
    Dim i As Long
    Dim j As Long
    lngIi1 = Val(Left$(strSemilla, InStr(strSemilla, ",") - 1))
    lngIi2 = Val(mid$(strSemilla, InStr(strSemilla, ",") + 1))

    For i = 1 To Len(strCadena)

        lngIi1 = lngIi1 - i
        lngIi2 = lngIi2 + i

        If (i Mod 2) = 0 Then

            Mid(strCadena, i, 1) = Chr$((Asc(mid$(strCadena, i, 1)) - lngIi1) And &HFF)

        Else

            Mid(strCadena, i, 1) = Chr$((Asc(mid$(strCadena, i, 1)) + lngIi2) And &HFF)

        End If

    Next

    Codificar = strCadena

End Function

Public Function DeCodificar(strCadena As String, strSemilla As String) As String

    Dim lngIi1 As Long
    Dim lngIi2 As Long
    Dim i As Long
    Dim j As Long
    lngIi1 = Val(Left$(strSemilla, InStr(strSemilla, ",") - 1))
    lngIi2 = Val(mid$(strSemilla, InStr(strSemilla, ",") + 1))

    For i = 1 To Len(strCadena)

        lngIi1 = lngIi1 - i
        lngIi2 = lngIi2 + i

        If (i Mod 2) = 0 Then

            Mid(strCadena, i, 1) = Chr$((Asc(mid$(strCadena, i, 1)) + lngIi1) And &HFF)

        Else

            Mid(strCadena, i, 1) = Chr$((Asc(mid$(strCadena, i, 1)) - lngIi2) And &HFF)

        End If

    Next

    DeCodificar = strCadena

End Function


