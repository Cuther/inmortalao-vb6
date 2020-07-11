VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AoDefenderConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Compare Text

'Declaradas a nivel de módulo
Dim unidad(0 To 9) As String
Dim decena(0 To 9) As String
Dim centena(0 To 10) As String
Dim deci(0 To 9) As String
Dim otros(0 To 15) As String

Private m_Sexo1 As String
Private m_Sexo2 As String
Private m_LenSexo1 As Long

Public Enum eSexo
    Femenino '= 0
    Masculino '= 1
End Enum

Public Function Numero2Letra(ByVal strNum As String, _
                            Optional ByVal Lo As Long = 0&, _
                            Optional ByVal NumDecimales As Long = 2&, _
                            Optional ByVal sMoneda As String = "", _
                            Optional ByVal sCentimos As String = "", _
                            Optional ByVal SexoMoneda As eSexo = Femenino, _
                            Optional ByVal SexoCentimos As eSexo = Masculino) As String
   
    Dim i As Long
    Dim iHayDecimal As Long             'Posición del signo decimal
    Dim sDecimal As String              'Signo decimal a usar
    Dim sDecimalNo As String            'Signo no decimal
    Dim sEntero As String
    Dim sFraccion As String
    Dim fFraccion As Single
    Dim sNumero As String
    Static SexoAntMoneda As eSexo
    Static SexoAntCentimos As eSexo
    Dim sSexoCents As String
    '
  
    If SexoCentimos = Femenino Then
        sSexoCents = "as"
    Else
        sSexoCents = "os"
    End If
    '
   
    If SexoMoneda = Femenino Then
        m_Sexo1 = "a"
        m_Sexo2 = "as"
    Else
        m_Sexo1 = ""
        m_Sexo2 = "os"
    End If
    'por si se cambia en el trascurso el sexo de la moneda
    If SexoMoneda <> SexoAntMoneda Then
        unidad(1) = "" ' Aquí ponía: unidad(2) = ""                 (08/Sep/01)
        SexoAntMoneda = SexoMoneda
    End If
    m_LenSexo1 = Len(m_Sexo1)

    'Si se especifica, se usarán
    sMoneda = Trim$(sMoneda)
    If Len(Trim$(sMoneda)) Then
        sMoneda = " " & sMoneda & " "
    Else
        sMoneda = " "
    End If

    sCentimos = Trim$(sCentimos)
    If Len(Trim$(sCentimos)) Then
        sCentimos = " " & sCentimos & " "
    Else
        sCentimos = " "
    End If

    'Si no se especifica el ancho...
    '
    If Lo Then
        sNumero = Space$(Lo)
    Else
        sNumero = ""
    End If

    'Comprobar el signo decimal y devolver los adecuados a la config. regional
    strNum = ConvDecimal(strNum, sDecimal, sDecimalNo)

    'Comprobar si tiene decimales
    iHayDecimal = InStr(strNum, sDecimal)
    If iHayDecimal Then
        sEntero = Left$(strNum, iHayDecimal - 1)
        sFraccion = mid$(strNum, iHayDecimal + 1) & String$(NumDecimales, "0")
        
        '
        sFraccion = Left$(sFraccion, NumDecimales)
       
        fFraccion = val(sFraccion)
        ' Si no hay decimales... no agregar nada...
        If fFraccion < 1 Then
            If Len(Trim$(sMoneda)) Then
                sMoneda = Pluralizar(sNumero, sMoneda)
            End If
            strNum = RTrim$(UnNumero(sEntero, m_Sexo1) & sMoneda)
            If Lo Then
                LSet sNumero = strNum
            Else
                sNumero = strNum
            End If
            Numero2Letra = sNumero
            Exit Function
        End If

        If Len(Trim$(sMoneda)) Then
            sMoneda = Pluralizar(sEntero, sMoneda)
        End If

        sEntero = UnNumero(sEntero, m_Sexo1)

        If Len(Trim$(sCentimos)) Then
            sCentimos = Pluralizar(sFraccion, sCentimos)
        End If

        
        m_Sexo2 = sSexoCents
        If SexoCentimos = Masculino Then
            sFraccion = UnNumero(sFraccion, "")
        Else
            sFraccion = UnNumero(sFraccion, "a")
        End If
        '
        strNum = sEntero & sMoneda & "con " & sFraccion & sCentimos
        If Lo Then
            LSet sNumero = RTrim$(strNum)
        Else
            sNumero = RTrim$(strNum)
        End If
        Numero2Letra = sNumero
    Else
        If Len(Trim$(sMoneda)) Then
            sMoneda = Pluralizar(strNum, sMoneda)
        End If
        strNum = RTrim$(UnNumero(strNum, m_Sexo1) & sMoneda)
        If Lo Then
            LSet sNumero = strNum
        Else
            sNumero = strNum
        End If
        Numero2Letra = sNumero
    End If
End Function

Private Function UnNumero(ByVal strNum As String, ByVal Sexo1 As String) As String
    '----------------------------------------------------------
   
    Dim dblNumero As Double

    Dim Negativo As Boolean
    Dim L As Integer
    Dim Una As Boolean
    Dim Millon As Boolean
    Dim Millones As Boolean
    Dim vez As Long
    Dim MaxVez As Long
    Dim k As Long
    Dim strQ As String
    Dim strB As String
    Dim strU As String
    Dim strD As String
    Dim strC As String
    Dim iA As Long
    '
    Dim strN() As String
    Dim Sexo1Ant As String

    
    Const cAncho = 12
    Const cGrupos = cAncho \ 3

   
    Sexo1Ant = m_Sexo1
    m_Sexo1 = Sexo1

    m_LenSexo1 = Len(m_Sexo1)
    '
   
    If unidad(1) <> "un" & Sexo1 Then
        InicializarArrays
    End If
    '
'
    On Local Error GoTo 0

    If Len(strNum) = 0 Then
        strNum = "0"
    End If

    dblNumero = Abs(CDbl(strNum))
    Negativo = (dblNumero <> CDbl(strNum))
    strNum = LTrim$(RTrim$(str$(dblNumero)))
    L = Len(strNum)

    If dblNumero < 1 Then
        UnNumero = "cero"
        Exit Function
    End If
    '
    Una = True
    Millon = False
    Millones = False
    If L < 4 Then Una = False
    If dblNumero > 999999 Then Millon = True
    If dblNumero > 1999999 Then Millones = True
    strB = ""
    strQ = strNum
    vez = 0

    ReDim strN(1 To cGrupos)
    strQ = Right$(String$(cAncho, "0") & strNum, cAncho)
    For k = Len(strQ) To 1 Step -3
        vez = vez + 1
        strN(vez) = mid$(strQ, k - 2, 3)
    Next
    MaxVez = cGrupos
    For k = cGrupos To 1 Step -1
        If strN(k) = "000" Then
            MaxVez = MaxVez - 1
        Else
            Exit For
        End If
    Next
    For vez = 1 To MaxVez
        strU = ""
        strD = ""
        strC = ""
        strNum = strN(vez)
        L = Len(strNum)
        k = val(Right$(strNum, 2))
        If Right$(strNum, 1) = "0" Then
            k = k \ 10
            strD = decena(k)
        ElseIf k > 10 And k < 16 Then
            k = val(mid$(strNum, L - 1, 2))
            strD = otros(k)
        Else
            strU = unidad(val(Right$(strNum, 1)))
            If L - 1 > 0 Then
                k = val(mid$(strNum, L - 1, 1))
                strD = deci(k)
            End If
        End If
        '---Parche de Esteve
        If L - 2 > 0 Then
            k = val(mid$(strNum, L - 2, 1))
            'Con esto funcionará bien el 100100, por ejemplo...
            If k = 1 Then                       'Parche
                If val(strNum) = 100 Then       'Parche
                    k = 10                      'Parche
                End If                          'Parche
            End If
            strC = centena(k) & " "
        End If
        '------
        If strU = "uno" And Left$(strB, 4) = " mil" Then strU = ""
        strB = strC & strD & strU & " " & strB

        If (vez = 1 Or vez = 3) Then
            If strN(vez + 1) <> "000" Then strB = " mil " & strB
        End If
        If vez = 2 And Millon Then
            If Millones Then
                strB = " millones " & strB
            Else
                strB = "un millón " & strB
            End If
        End If
    Next
    strB = Trim$(strB)
    If Right$(strB, 3) = "uno" Then
        strB = Left$(strB, Len(strB) - 1) & m_Sexo1 '"a"
    End If
    Do                              'Quitar los espacios dobles que haya por medio
        iA = InStr(strB, "  ")
        If iA = 0 Then Exit Do
        strB = Left$(strB, iA - 1) & mid$(strB, iA + 1)
    Loop
    '
    If Left$(strB, 5 + m_LenSexo1) = "un" & m_Sexo1 & " un" Then
        strB = mid$(strB, 4 + m_LenSexo1)
    End If
    '---Nueva comparación                                     (01:16 25/Ene/99)
    If Left$(strB, 5) = "un un" Then
        strB = mid$(strB, 4)
    End If
    
    If Left$(strB, 7 + m_LenSexo1) = "un" & m_Sexo1 & " mil " Then
        strB = mid$(strB, 4 + m_LenSexo1)
    ' Puede que el importe sea sólo "un mil" o "una mil"            (19/Ago/00)
    ElseIf strB = "un" & m_Sexo1 & " mil" Then
        strB = mid$(strB, 4 + m_LenSexo1)
    End If
    '
 
    If Left$(strB, 7) = "un mil " Then
        strB = mid$(strB, 4)
    End If
    '
    If Right$(strB, 15 + m_LenSexo1) <> "millones mil un" & m_Sexo1 Then
        iA = InStr(strB, "millones mil un" & m_Sexo1)
        If iA Then strB = Left$(strB, iA + 8) & mid$(strB, iA + 13)
    End If
    '---Nueva comparación                                   (15:13 25/Ene/99)
    If Right$(strB, 15) <> "millones mil un" Then
        iA = InStr(strB, "millones mil un")
        If iA Then strB = Left$(strB, iA + 8) + mid$(strB, iA + 13)
    End If
  
    If Millones Then
       
        If m_Sexo1 = "a" Then
            'If (strB Like "*as * millones*") Then
            ' Usar un bucle Do por si hay varias coincidencias      (07/Dic/00)
            Do While (strB Like "*as * millones*")
                ' Buscar la primera terminación "as " y cambiar por "os "
                k = InStr(strB, "as ")
                If k Then
                    Mid$(strB, k) = "os "
                End If
            Loop
          
            Do While (strB Like "*as millones*")
                ' Buscar la primera terminación "as " y cambiar por "os "
                k = InStr(strB, "as millones")
                If k Then
                    Mid$(strB, k) = "os millones"
                End If
            Loop
            'End If
            '
            '
            
            k = InStr(strB, "una mill")
            If k Then
                strB = Left$(strB, k + 1) & mid$(strB, k + 3)
            End If
            '
            '
        End If
    End If
    '
    '
    '--------------------------------------------------------------------------
    ' Cambiar los veintiun por veintiún, etc por sus acentuadas     (08/Jul/00)
    Do
        k = InStr(strB, "veintiun ")
        If k Then
            Mid$(strB, k) = "veintiún "
        End If
    Loop While k
    ' El veintidos creo que nunca lo he acentuado...                (08/Jul/00)
    ' pero en la enciclopedia consultada lo acentúa
    Do
        k = InStr(strB, "veintidos ")
        If k Then
            Mid$(strB, k) = "veintidós "
        End If
    Loop While k
    Do
        k = InStr(strB, "veintitres ")
        If k Then
            Mid$(strB, k) = "veintitrés "
        End If
    Loop While k
    Do
        k = InStr(strB, "veintiseis ")
        If k Then
            Mid$(strB, k) = "veintiséis "
        End If
    Loop While k
    '--------------------------------------------------------------------------
    '
    '
    If Right$(strB, 6) = "ciento" Then
        strB = Left$(strB, Len(strB) - 2)
    End If
    If Negativo Then strB = "menos " & strB

    UnNumero = Trim$(strB)

    ' Restablecer el valor anterior
    m_Sexo1 = Sexo1Ant
    m_LenSexo1 = Len(m_Sexo1)
End Function

Private Sub InicializarArrays()
    'Asignar los valores
    unidad(1) = "un" & m_Sexo1
    unidad(2) = "dos"
    unidad(3) = "tres"
    unidad(4) = "cuatro"
    unidad(5) = "cinco"
    unidad(6) = "seis"
    unidad(7) = "siete"
    unidad(8) = "ocho"
    unidad(9) = "nueve"
    '
    decena(1) = "diez"
    decena(2) = "veinte"
    decena(3) = "treinta"
    decena(4) = "cuarenta"
    decena(5) = "cincuenta"
    decena(6) = "sesenta"
    decena(7) = "setenta"
    decena(8) = "ochenta"
    decena(9) = "noventa"
    '
    centena(1) = "ciento"
    centena(2) = "doscient" & m_Sexo2
    centena(3) = "trescient" & m_Sexo2
    centena(4) = "cuatrocient" & m_Sexo2
    centena(5) = "quinient" & m_Sexo2
    centena(6) = "seiscient" & m_Sexo2
    centena(7) = "setecient" & m_Sexo2
    centena(8) = "ochocient" & m_Sexo2
    centena(9) = "novecient" & m_Sexo2
    centena(10) = "cien"                'Parche
    '
    deci(1) = "dieci"
    deci(2) = "veinti"
    deci(3) = "treinta y "
    deci(4) = "cuarenta y "
    deci(5) = "cincuenta y "
    deci(6) = "sesenta y "
    deci(7) = "setenta y "
    deci(8) = "ochenta y "
    deci(9) = "noventa y "
    '
    otros(1) = "1"
    otros(2) = "2"
    otros(3) = "3"
    otros(4) = "4"
    otros(5) = "5"
    otros(6) = "6"
    otros(7) = "7"
    otros(8) = "8"
    otros(9) = "9"
    otros(10) = "10"
    otros(11) = "once"
    otros(12) = "doce"
    otros(13) = "trece"
    otros(14) = "catorce"
    otros(15) = "quince"
End Sub

Private Sub Class_Initialize()
    m_Sexo1 = "a"
    m_Sexo2 = "as"
End Sub

Public Function Pluralizar( _
                ByVal sNumero As String, _
                ByVal sMoneda As String, _
                Optional ByVal bCadaPalabra As Boolean = False) As String
   
    Dim dblTotal As Double
    Dim sTmp As String
    Dim i As Long

    If Len(Trim$(sMoneda)) Then
        ' He quitado el Val             (08/Jul/00)
        'dblTotal = Val(sNumero)
        '
        ' Si entra una cadena vacia, da error                       (08/Jul/00)
        If Len(sNumero) = 0 Then
            sNumero = "0"
        End If
        dblTotal = (sNumero)
        '
        If dblTotal <> 1# Then
            sMoneda = Trim$(sMoneda)
            ' Si se pluralizan todas las palabras                   (08/Jul/00)
            If bCadaPalabra Then
                sMoneda = sMoneda & " "
                sTmp = ""
                For i = 1 To Len(sMoneda)
                    If mid$(sMoneda, i, 1) = " " Then
                        ' pluralizar
                        If InStr("aeiou", Right$(sTmp, 1)) Then
                            sTmp = sTmp & "s"
                        Else
                            sTmp = sTmp & "es"
                        End If
                    End If
                    sTmp = sTmp & mid$(sMoneda, i, 1)
                Next
                sMoneda = " " & Trim$(sTmp) & " "
            Else
                If InStr("aeiou", Right$(sMoneda, 1)) Then
                    sMoneda = " " & sMoneda & "s "
                Else
                    sMoneda = " " & sMoneda & "es "
                End If
            End If
        End If
    End If
    Pluralizar = sMoneda
End Function

Public Function ConvDecimal(ByVal strNum As String, _
                            Optional ByRef sDecimal As String = ",", _
                            Optional ByRef sDecimalNo As String = ".") As String
    ' Asigna el signo decimal adecuado (o lo intenta)               (10/Ene/99)
    ' Devuelve una cadena con el signo decimal del sistema
    Dim sNumero As String
    Dim i As Long
    Dim j As Long

    On Error Resume Next        ' Si se produce un error, continuar (07/Jul/00)

    ' Averiguar el signo decimal
    sNumero = Format$(25.5, "#.#")
    If InStr(sNumero, ".") Then
        sDecimal = "."
        sDecimalNo = ","
    Else
        sDecimal = ","
        sDecimalNo = "."
    End If

    strNum = Trim$(strNum)
    If Left$(strNum, 1) = sDecimalNo Then
        Mid$(strNum, 1, 1) = sDecimal
    End If

    ' Si el número introducido contiene signos no decimales
    j = 0
    i = 1
    Do
        i = InStr(i, strNum, sDecimalNo)
        If i Then
            j = j + 1
            i = i + 1
        End If
    Loop While i

    If j = 1 Then
        ' cambiar ese símbolo por un espacio, si sólo hay uno de esos signos
        i = InStr(strNum, sDecimalNo)
        If i Then
            If InStr(strNum, sDecimal) Then
                Mid$(strNum, i, 1) = " "
            Else
                Mid$(strNum, i, 1) = sDecimal
            End If
        End If
    Else
        'En caso de que tenga más de uno de estos símbolos
      
        '
        i = 1
        Do
            i = InStr(i, strNum, sDecimalNo)
            If i Then
                j = j - 1
                If j = 0 Then
                    Mid$(strNum, i, 1) = sDecimal
                Else
                    Mid$(strNum, i, 1) = " "
                End If
                i = i + 1
            End If
        Loop While i
    End If

    j = 0
    ' Quitar los espacios que haya por medio
    Do
        i = InStr(strNum, " ")
        If i = 0 Then Exit Do
        strNum = Left$(strNum, i - 1) & mid$(strNum, i + 1)
    Loop

    ConvDecimal = strNum

    Err = 0
End Function


