Attribute VB_Name = "Extra"

Option Explicit
Public Function ObtenerSuerte(ByVal valor As Long) As Byte
    If valor <= 10 And valor >= -1 Then
        ObtenerSuerte = 35
    ElseIf valor <= 20 And valor >= 11 Then
        ObtenerSuerte = 30
    ElseIf valor <= 30 And valor >= 21 Then
        ObtenerSuerte = 28
    ElseIf valor <= 40 And valor >= 31 Then
        ObtenerSuerte = 24
    ElseIf valor <= 50 And valor >= 41 Then
        ObtenerSuerte = 22
    ElseIf valor <= 60 And valor >= 51 Then
        ObtenerSuerte = 20
    ElseIf valor <= 70 And valor >= 61 Then
        ObtenerSuerte = 18
    ElseIf valor <= 80 And valor >= 71 Then
        ObtenerSuerte = 15
    ElseIf valor <= 90 And valor >= 81 Then
        ObtenerSuerte = 12
    ElseIf valor < 100 And valor >= 91 Then
        ObtenerSuerte = 8
    ElseIf valor = 100 Then
        ObtenerSuerte = 7
    End If
End Function
Public Function ClaseToEnum(ByVal Clase As String) As eClass
Dim i As Byte
    For i = 1 To NUMCLASES
        If UCase$(ListaClases(i)) = UCase$(Clase) Then
            ClaseToEnum = i
        End If
    Next i
End Function
Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
    EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie
End Function
Public Function esArmada(ByVal UserIndex As Integer) As Boolean
    esArmada = (UserList(UserIndex).faccion.ArmadaReal = 1)
End Function
Public Function esCaos(ByVal UserIndex As Integer) As Boolean
    esCaos = (UserList(UserIndex).faccion.FuerzasCaos = 1)
End Function
Public Function esMili(ByVal UserIndex As Integer) As Boolean
    esMili = (UserList(UserIndex).faccion.Milicia = 1)
End Function
Public Function esFaccion(ByVal UserIndex As Integer) As Boolean
    esFaccion = (UserList(UserIndex).faccion.ArmadaReal = 1 Or UserList(UserIndex).faccion.FuerzasCaos = 1 Or UserList(UserIndex).faccion.Milicia = 1)
End Function
Public Function esRene(ByVal UserIndex As Integer) As Boolean
    esRene = (UserList(UserIndex).faccion.Renegado)
End Function
Public Function esCiuda(ByVal UserIndex As Integer) As Boolean
    esCiuda = (UserList(UserIndex).faccion.Ciudadano)
End Function
Public Function esRepu(ByVal UserIndex As Integer) As Boolean
    esRepu = (UserList(UserIndex).faccion.Republicano)
End Function
Public Function esMismoBando(ByVal U1 As Integer, ByVal U2 As Integer) As Boolean
    With UserList(U1)
        esMismoBando = (.faccion.Republicano = 1 And UserList(U2).faccion.Republicano = 1) Or ( _
                        .faccion.Ciudadano = 1 And UserList(U2).faccion.Ciudadano = 1)
    End With
End Function

Public Function EsUSER(ByVal UserIndex As Integer) As Boolean
    EsUSER = Not (UserList(UserIndex).flags.Privilegios And (PlayerType.Conse Or PlayerType.Semi Or PlayerType.Dios Or PlayerType.Admin Or PlayerType.VIP))
End Function

Public Function EsFacc(ByVal UserIndex As Integer) As Boolean
    EsFacc = (UserList(UserIndex).flags.Privilegios And (PlayerType.FaccImpe Or PlayerType.FaccRepu Or PlayerType.FaccCaos Or PlayerType.Conse Or PlayerType.Semi Or PlayerType.Dios Or PlayerType.Admin Or PlayerType.VIP))
End Function
Public Function EsFaccImpe(ByVal UserIndex As Integer) As Boolean
    EsFaccImpe = (UserList(UserIndex).flags.Privilegios And PlayerType.FaccImpe)
End Function
Public Function EsFaccRepu(ByVal UserIndex As Integer) As Boolean
    EsFaccRepu = (UserList(UserIndex).flags.Privilegios And PlayerType.FaccRepu)
End Function
Public Function EsFaccCaos(ByVal UserIndex As Integer) As Boolean
    EsFaccCaos = (UserList(UserIndex).flags.Privilegios And PlayerType.FaccCaos)
End Function

Public Function EsCONSE(ByVal UserIndex As Integer) As Boolean
    EsCONSE = (UserList(UserIndex).flags.Privilegios And (PlayerType.Conse Or PlayerType.Semi Or PlayerType.Dios Or PlayerType.Admin Or PlayerType.VIP))
End Function
Public Function EsSEMI(ByVal UserIndex As Integer) As Boolean
    EsSEMI = (UserList(UserIndex).flags.Privilegios And (PlayerType.Semi Or PlayerType.Dios Or PlayerType.Admin Or PlayerType.VIP))
End Function
Public Function EsDIOS(ByVal UserIndex As Integer) As Boolean
    EsDIOS = (UserList(UserIndex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.VIP))
End Function
Public Function EsADMIN(ByVal UserIndex As Integer) As Boolean
    EsADMIN = (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.VIP))
End Function
Public Function EsVIP(ByVal UserIndex As Integer) As Boolean
    EsVIP = (UserList(UserIndex).flags.Privilegios And (PlayerType.VIP))
End Function
Public Function EsGM(ByVal UserIndex As Integer) As Boolean
    EsGM = (UserList(UserIndex).flags.Privilegios And (PlayerType.Semi Or PlayerType.Dios Or PlayerType.Admin Or PlayerType.VIP))
End Function


Public Sub DoTileEvents(ByVal UserIndex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Handles the Map passage of Users. Allows the existance
'of exclusive maps for Newbies, Royal Army and Caos Legion members
'and enables GMs to enter every map without restriction.
'Uses: Mapinfo(map).Restringir = "NEWBIE" (newbies), "ARMADA", "CAOS", "FACCION" or "NO".
'***************************************************
    Dim nPos As WorldPos
    Dim FxFlag As Boolean
    
On Error GoTo Errhandler
    'Controla las salidas
    If InMapBounds(map, x, Y) Then
        With MapData(map, x, Y)
            If .ObjInfo.ObjIndex > 0 Then
                FxFlag = ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport
            End If
            
            If .TileExit.map > 0 And .TileExit.map <= NumMaps Then
                '¿Es mapa de newbies?
                If .TileExit.map = 37 Or .TileExit.map = 208 Then
                    '¿El usuario es un newbie?
                    If EsNewbie(UserIndex) Or EsGM(UserIndex) Then
                        If LegalPos(.TileExit.map, .TileExit.x, .TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, .TileExit.map, .TileExit.x, .TileExit.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(.TileExit, nPos)
                            If nPos.x <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es newbie
                        Call WriteConsoleMsg(1, UserIndex, "Mapa exclusivo para newbies.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
        
                        If nPos.x <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.Y, False)
                        End If
                    End If
                    
                'Add Nod kopfnickend Carrera
                ElseIf Carrera_estado And .TileExit.map = 49 And UserList(UserIndex).Pos.map = MapaCarrera Then

                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Eventos> El primero en llegar a la meta es " & UserList(UserIndex).Name, FontTypeNames.FONTTYPE_GUILD))
                    'Entrega de premio
                    'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, UserList(UserIndex).Name & " Ha ganado 2 creditos que han sido acreditados en su cuenta.", FontTypeNames.FONTTYPE_BROWNI))
                    'Call agregarcreditos(UserList(UserIndex).IndexAccount)
        
                    Carrera_estado = False
                    Carrera_puestos = 0
                '\Add
                                
                ElseIf .TileExit.map = 757 Or .TileExit.map = 760 Or .TileExit.map = 250 Then
                    '¿El usuario es donador
                    If UserList(UserIndex).donador = True Or EsGM(UserIndex) Then
                        If LegalPos(.TileExit.map, .TileExit.x, .TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, .TileExit.map, .TileExit.x, .TileExit.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(.TileExit, nPos)
                            If nPos.x <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es donador
                        Call WriteConsoleMsg(1, UserIndex, "Mapa exclusivo para DONADORES. Enterate mas en http://inmortalao.com.ar/", FontTypeNames.FONTTYPE_BROWNI)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
        
                        If nPos.x <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.Y, False)
                        End If
                    End If
                    
                    
                    
                ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "ARMADA" Then '¿Es mapa de Armadas?
                    '¿El usuario es Armada?
                    If esArmada(UserIndex) Or EsGM(UserIndex) Then
                        If LegalPos(.TileExit.map, .TileExit.x, .TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, .TileExit.map, .TileExit.x, .TileExit.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(.TileExit, nPos)
                            If nPos.x <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es armada
                        Call WriteConsoleMsg(1, UserIndex, "Mapa exclusivo para miembros del ejército Real", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.x <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.Y, FxFlag)
                        End If
                    End If
                ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "MILICIA" Then '¿Es mapa de la Milicia?
                    '¿El usuario es Armada?
                    If esMili(UserIndex) Or EsGM(UserIndex) Then
                        If LegalPos(.TileExit.map, .TileExit.x, .TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, .TileExit.map, .TileExit.x, .TileExit.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(.TileExit, nPos)
                            If nPos.x <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es miliciano
                        Call WriteConsoleMsg(1, UserIndex, "Mapa exclusivo para miembros de la Milicia Republicana", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.x <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.Y, FxFlag)
                        End If
                    End If
                ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "CAOS" Then '¿Es mapa de Caos?
                    '¿El usuario es Caos?
                    If esCaos(UserIndex) Or EsGM(UserIndex) Then
                        If LegalPos(.TileExit.map, .TileExit.x, .TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, .TileExit.map, .TileExit.x, .TileExit.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(.TileExit, nPos)
                            If nPos.x <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es caos
                        Call WriteConsoleMsg(1, UserIndex, "Mapa exclusivo para miembros del ejército Oscuro.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.x <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.Y, FxFlag)
                        End If
                    End If
                ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "FACCION" Then '¿Es mapa de faccionarios?
                    '¿El usuario es Armada o Caos?
                    If esArmada(UserIndex) Or esCaos(UserIndex) Or EsGM(UserIndex) Then
                        If LegalPos(.TileExit.map, .TileExit.x, .TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, .TileExit.map, .TileExit.x, .TileExit.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(.TileExit, nPos)
                            If nPos.x <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es Faccionario
                        Call WriteConsoleMsg(1, UserIndex, "Solo se permite entrar al Mapa si eres miembro de alguna Facción", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.x <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.Y, FxFlag)
                        End If
                    End If
                Else 'No es un mapa de newbies, ni Armadas, ni Caos, ni faccionario.
                    If LegalPos(.TileExit.map, .TileExit.x, .TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                        Call WarpUserChar(UserIndex, .TileExit.map, .TileExit.x, .TileExit.Y, FxFlag)
                    Else
                        Call ClosestLegalPos(.TileExit, nPos)
                        If nPos.x <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.Y, FxFlag)
                        End If
                    End If
                End If
                
                'Te fusite del mapa. La criatura ya no es más tuya ni te reconoce como que vos la atacaste.
                Dim aN As Integer
                
                aN = UserList(UserIndex).flags.AtacadoPorNpc
                If aN > 0 Then
                   Npclist(aN).Movement = Npclist(aN).flags.OldMovement
                   Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
                   Npclist(aN).flags.AttackedBy = 0
                End If
            
                aN = UserList(UserIndex).flags.NPCAtacado
                If aN > 0 Then
                    If Npclist(aN).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                        Npclist(aN).flags.AttackedFirstBy = vbNullString
                    End If
                End If
                UserList(UserIndex).flags.AtacadoPorNpc = 0
                UserList(UserIndex).flags.NPCAtacado = 0
                
            End If
        End With
    End If
    
    'Add Marius Camptura la Bandera
    If Bandera_estado And UserList(UserIndex).Pos.map = Bandera_mapa Then
        Call Bandera_verifica_banderas(UserIndex)
    End If
    '\Add
    
    
Exit Sub

Errhandler:
    Call LogError("Error en DoTileEvents. Error: " & err.Number & " - Desc: " & err.description)
End Sub

Function InRangoVision(ByVal UserIndex As Integer, ByVal x As Integer, ByVal Y As Integer) As Boolean

If x > UserList(UserIndex).Pos.x - MinXBorder And x < UserList(UserIndex).Pos.x + MinXBorder Then
    If Y > UserList(UserIndex).Pos.Y - MinYBorder And Y < UserList(UserIndex).Pos.Y + MinYBorder Then
        InRangoVision = True
        Exit Function
    End If
End If
InRangoVision = False

End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, x As Integer, Y As Integer) As Boolean

If x > Npclist(NpcIndex).Pos.x - MinXBorder And x < Npclist(NpcIndex).Pos.x + MinXBorder Then
    If Y > Npclist(NpcIndex).Pos.Y - MinYBorder And Y < Npclist(NpcIndex).Pos.Y + MinYBorder Then
        InRangoVisionNPC = True
        Exit Function
    End If
End If
InRangoVisionNPC = False

End Function


Function InMapBounds(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer) As Boolean
            
If (map <= 0 Or map > NumMaps) Or x < MinXBorder Or x > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapBounds = False
Else
    InMapBounds = True
End If

End Function

Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos, Optional PuedeAgua As Boolean = False, Optional PuedeTierra As Boolean = True)
'*****************************************************************
'Author: Unknown (original version)
'Last Modification: 24/01/2007 (ToxicWaste)
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim loopC As Integer
Dim tX As Long
Dim tY As Long

nPos.map = Pos.map

Do While Not LegalPos(Pos.map, nPos.x, nPos.Y, PuedeAgua, PuedeTierra)
    If loopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - loopC To Pos.Y + loopC
        For tX = Pos.x - loopC To Pos.x + loopC
            
            If LegalPos(nPos.map, tX, tY, PuedeAgua, PuedeTierra) Then
                nPos.x = tX
                nPos.Y = tY
                '¿Hay objeto?
                
                tX = Pos.x + loopC
                tY = Pos.Y + loopC
  
            End If
        
        Next tX
    Next tY
    
    loopC = loopC + 1
    
Loop

If Notfound = True Then
    nPos.x = 0
    nPos.Y = 0
End If

End Sub

Private Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim loopC As Integer
Dim tX As Long
Dim tY As Long

nPos.map = Pos.map

Do While Not LegalPos(Pos.map, nPos.x, nPos.Y)
    If loopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - loopC To Pos.Y + loopC
        For tX = Pos.x - loopC To Pos.x + loopC
            
            If LegalPos(nPos.map, tX, tY) And MapData(nPos.map, tX, tY).TileExit.map = 0 Then
                nPos.x = tX
                nPos.Y = tY
                '¿Hay objeto?
                
                tX = Pos.x + loopC
                tY = Pos.Y + loopC
  
            End If
        
        Next tX
    Next tY
    
    loopC = loopC + 1
    
Loop

If Notfound = True Then
    nPos.x = 0
    nPos.Y = 0
End If

End Sub

Function NameIndex(ByVal Name As String) As Integer
    Dim UserIndex As Long
    
    '¿Nombre valido?
    If LenB(Name) = 0 Then
        NameIndex = 0
        Exit Function
    End If
    
    If InStrB(Name, "+") <> 0 Then
        Name = UCase$(Replace(Name, "+", " "))
    End If
    
    UserIndex = 1
    Do Until UCase$(UserList(UserIndex).Name) = UCase$(Name)
        
        UserIndex = UserIndex + 1
        
        If UserIndex > MaxUsers Then
            NameIndex = 0
            Exit Function
        End If
    Loop
     
    NameIndex = UserIndex
End Function

Function CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Boolean
    Dim loopC As Long
    'BY CASTELLI... la extencion del for es preferible que sea
    'hasta last user...
    
    
    For loopC = 1 To LastUser
        If UserList(loopC).flags.UserLogged = True Then
            If UserList(loopC).ip = UserIP And UserIndex <> loopC Then
                CheckForSameIP = True
                Exit Function
            End If
        End If
    Next loopC
    
    CheckForSameIP = False
End Function

Function CheckForSameName(ByVal Name As String) As Boolean
'Controlo que no existan usuarios con el mismo nombre
    Dim loopC As Long
    
    For loopC = 1 To LastUser
        If UserList(loopC).flags.UserLogged Then
            
            'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
            'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
            'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
            'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
            'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
            
            If UCase$(UserList(loopC).Name) = UCase$(Name) Then
                CheckForSameName = True
                'UserList(loopC).Counters.Saliendo = True
                'UserList(LoopC).Counters.Salir = 1
                Exit Function
            End If
        End If
    Next loopC
    
    CheckForSameName = False
End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef Pos As WorldPos)
'*****************************************************************
'Toma una posicion y se mueve hacia donde esta perfilado
'*****************************************************************
    Select Case Head
        Case eHeading.NORTH
            Pos.Y = Pos.Y - 1
        
        Case eHeading.SOUTH
            Pos.Y = Pos.Y + 1
        
        Case eHeading.EAST
            Pos.x = Pos.x + 1
        
        Case eHeading.WEST
            Pos.x = Pos.x - 1
    End Select
End Sub

Function LegalPos(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Checks if the position is Legal.
'***************************************************
'¿Es un mapa valido?
If (map <= 0 Or map > NumMaps) Or _
   (x < MinXBorder Or x > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
            LegalPos = False
Else
    If PuedeAgua And PuedeTierra Then
        LegalPos = (MapData(map, x, Y).Blocked <> 1) And _
                   (MapData(map, x, Y).UserIndex = 0) And _
                   (MapData(map, x, Y).NpcIndex = 0)
    ElseIf PuedeTierra And Not PuedeAgua Then
        LegalPos = (MapData(map, x, Y).Blocked <> 1) And _
                   (MapData(map, x, Y).UserIndex = 0) And _
                   (MapData(map, x, Y).NpcIndex = 0) And _
                   (Not HayAgua(map, x, Y))
    ElseIf PuedeAgua And Not PuedeTierra Then
        LegalPos = (MapData(map, x, Y).Blocked <> 1) And _
                   (MapData(map, x, Y).UserIndex = 0) And _
                   (MapData(map, x, Y).NpcIndex = 0) And _
                   (HayAgua(map, x, Y))
    Else
        LegalPos = False
    End If
   
End If

End Function

Function MoveToLegalPos(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True) As Boolean
'***************************************************
'Autor: ZaMa
'Last Modification: 26/03/2009
'Checks if the position is Legal, but considers that if there's a casper, it's a legal movement.
'***************************************************

Dim UserIndex As Integer
Dim IsDeadChar As Boolean
Dim IsAdminInvisible As Boolean

'¿Es un mapa valido?
    If (map <= 0 Or map > NumMaps) Or _
       (x < MinXBorder Or x > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
            MoveToLegalPos = False
    Else
        With MapData(map, x, Y)
            UserIndex = .UserIndex
            If UserIndex > 0 Then
                IsDeadChar = UserList(UserIndex).flags.Muerto = 1
                IsAdminInvisible = (UserList(UserIndex).flags.AdminInvisible = 1)
            Else
                IsDeadChar = False
            End If
        
            If PuedeAgua And PuedeTierra Then
                MoveToLegalPos = (.Blocked <> 1) And _
                           (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                           (.NpcIndex = 0)
            ElseIf PuedeTierra And Not PuedeAgua Then
                MoveToLegalPos = (.Blocked <> 1) And _
                           (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                           (.NpcIndex = 0) And _
                           (Not HayAgua(map, x, Y))
            ElseIf PuedeAgua And Not PuedeTierra Then
                MoveToLegalPos = (.Blocked <> 1) And _
                           (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And _
                           (.NpcIndex = 0) And _
                           (HayAgua(map, x, Y))
            Else
                MoveToLegalPos = False
            End If
        End With
  
    End If


End Function

Public Sub FindLegalPos(ByVal UserIndex As Integer, ByVal map As Integer, ByRef x As Integer, ByRef Y As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 26/03/2009
'Search for a Legal pos for the user who is being teleported.
'***************************************************

    If MapData(map, x, Y).UserIndex <> 0 Or _
        MapData(map, x, Y).NpcIndex <> 0 Then
                    
        ' Se teletransporta a la misma pos a la que estaba
        If MapData(map, x, Y).UserIndex = UserIndex Then Exit Sub
                            
        Dim FoundPlace As Boolean
        Dim tX As Long
        Dim tY As Long
        Dim Rango As Long
        Dim OtherUserIndex As Integer
        Rango = 5
        
        For tY = Y - Rango To Y + Rango
            For tX = x - Rango To x + Rango
                'Reviso que no haya User ni NPC
                If MapData(map, tX, tY).UserIndex = 0 And _
                    MapData(map, tX, tY).NpcIndex = 0 Then
                    
                    If InMapBounds(map, tX, tY) Then
                        FoundPlace = True
                        Exit For
                    End If
                End If

            Next tX
    
            If FoundPlace Then _
                Exit For
        Next tY

    
        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            x = tX
            Y = tY
            End If
        End If
    

End Sub

Function LegalPosNPC(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal AguaValida As Byte) As Boolean
'***************************************************
'Autor: Unkwnown
'Last Modification: 27/04/2009
'Checks if it's a Legal pos for the npc to move to.
'***************************************************
Dim IsDeadChar As Boolean
Dim UserIndex As Integer

    If (map <= 0 Or map > NumMaps) Or _
        (x < MinXBorder Or x > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPosNPC = False
        Exit Function
    End If

    UserIndex = MapData(map, x, Y).UserIndex
    If UserIndex > 0 Then
        IsDeadChar = UserList(UserIndex).flags.Muerto = 1
    Else
        IsDeadChar = False
    End If
    
    If AguaValida = 0 Then
        LegalPosNPC = (MapData(map, x, Y).Blocked <> 1) And _
        (MapData(map, x, Y).UserIndex = 0 Or IsDeadChar) And _
        (MapData(map, x, Y).NpcIndex = 0) And _
        (MapData(map, x, Y).Trigger <> eTrigger.POSINVALIDA) _
        And Not HayAgua(map, x, Y)
    Else
        LegalPosNPC = (MapData(map, x, Y).Blocked <> 1) And _
        (MapData(map, x, Y).UserIndex = 0 Or IsDeadChar) And _
        (MapData(map, x, Y).NpcIndex = 0) And _
        (MapData(map, x, Y).Trigger <> eTrigger.POSINVALIDA)
    End If
End Function

Sub SendHelp(ByVal index As Integer)
Dim NumHelpLines As Integer
Dim loopC As Integer

NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

For loopC = 1 To NumHelpLines
    Call WriteConsoleMsg(1, index, GetVar(DatPath & "Help.dat", "Help", "Line" & loopC), FontTypeNames.FONTTYPE_INFO)
Next loopC

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    If Npclist(NpcIndex).NroExpresiones > 0 Then
        Dim randomi
        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Npclist(NpcIndex).Expresiones(randomi), Npclist(NpcIndex).Char.CharIndex, vbWhite))
    End If
End Sub

Sub LookatTile(ByVal UserIndex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 26/03/2009
'13/02/2009: ZaMa - EL nombre del gm que aparece por consola al clickearlo, tiene el color correspondiente a su rango
'***************************************************


'Responde al click del usuario sobre el mapa
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Stat As String
Dim ft As FontTypeNames

'¿Rango Visión? (ToxicWaste)
If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.x - x) > RANGO_VISION_X) Then
    Exit Sub
End If

'¿Posicion valida?
If InMapBounds(map, x, Y) Then
    UserList(UserIndex).flags.TargetMap = map
    UserList(UserIndex).flags.TargetX = x
    UserList(UserIndex).flags.TargetY = Y
    '¿Es un obj?
    If MapData(map, x, Y).ObjInfo.ObjIndex > 0 Then
        'Informa el nombre
        UserList(UserIndex).flags.TargetObjMap = map
        UserList(UserIndex).flags.TargetObjX = x
        UserList(UserIndex).flags.TargetObjY = Y
        FoundSomething = 1
    ElseIf MapData(map, x + 1, Y).ObjInfo.ObjIndex > 0 Then
        'Informa el nombre
        If ObjData(MapData(map, x + 1, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            UserList(UserIndex).flags.TargetObjMap = map
            UserList(UserIndex).flags.TargetObjX = x + 1
            UserList(UserIndex).flags.TargetObjY = Y
            FoundSomething = 1
        End If
    ElseIf MapData(map, x + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
        If ObjData(MapData(map, x + 1, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            'Informa el nombre
            UserList(UserIndex).flags.TargetObjMap = map
            UserList(UserIndex).flags.TargetObjX = x + 1
            UserList(UserIndex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    ElseIf MapData(map, x, Y + 1).ObjInfo.ObjIndex > 0 Then
        If ObjData(MapData(map, x, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            'Informa el nombre
            UserList(UserIndex).flags.TargetObjMap = map
            UserList(UserIndex).flags.TargetObjX = x
            UserList(UserIndex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    End If
    
    If FoundSomething = 1 Then
        UserList(UserIndex).flags.TargetObj = MapData(map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex
    End If
    
    '¿Es un personaje?
    If Y + 1 <= YMaxMapSize Then
        If MapData(map, x, Y + 1).UserIndex > 0 Then
            TempCharIndex = MapData(map, x, Y + 1).UserIndex
            FoundChar = 1
        End If
        If MapData(map, x, Y + 1).NpcIndex > 0 Then
            TempCharIndex = MapData(map, x, Y + 1).NpcIndex
            FoundChar = 2
        End If
    End If
    '¿Es un personaje?
    If FoundChar = 0 Then
        If MapData(map, x, Y).UserIndex > 0 Then
            TempCharIndex = MapData(map, x, Y).UserIndex
            FoundChar = 1
        End If
        If MapData(map, x, Y).NpcIndex > 0 Then
            TempCharIndex = MapData(map, x, Y).NpcIndex
            FoundChar = 2
        End If
    End If
    
    'Reaccion al personaje
    If FoundChar = 1 Then '  ¿Encontro un Usuario?
       If UserList(TempCharIndex).flags.AdminInvisible = 0 Or UserList(UserIndex).flags.Privilegios And PlayerType.User Then
            If UserList(TempCharIndex).showName Then
                WriteCharMsgStatus UserIndex, TempCharIndex
            End If
            
            'Add Marius Movemos a los muertos con dos clicks
            If UserIndex <> TempCharIndex And UserList(TempCharIndex).flags.Muerto = 1 And MapData(UserList(TempCharIndex).Pos.map, UserList(TempCharIndex).Pos.x, UserList(TempCharIndex).Pos.Y).ObjInfo.ObjIndex <> 0 Then
                Call Sum(TempCharIndex, UserList(TempCharIndex).Pos.map, UserList(TempCharIndex).Pos.x, UserList(TempCharIndex).Pos.Y, True)
            End If
            '\Add
            
            FoundSomething = 1
            UserList(UserIndex).flags.TargetUser = TempCharIndex
            UserList(UserIndex).flags.TargetNPC = 0
            UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
       End If
    End If
    
    If FoundChar = 2 Then '¿Encontro un NPC?
        Dim estatus As String
        
        If EsGM(UserIndex) Or UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) = 100 Then
            estatus = "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ") "
        Else
            If UserList(UserIndex).flags.Muerto = 0 Then
                If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.05) Then
                    estatus = estatus & " (Muerto)"
                ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.1) Then
                    estatus = estatus & " (Casi muerto)"
                ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                    estatus = estatus & " (Malherido)"
                ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                    estatus = estatus & " (Herido)"
                ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP) Then
                    estatus = estatus & " (Levemente Herido)"
                Else
                    estatus = estatus & " (Intacto)"
                End If
            End If
        End If
        
        If Len(Npclist(TempCharIndex).desc) > 1 Then
            Call WriteChatOverHead(UserIndex, Npclist(TempCharIndex).desc, Npclist(TempCharIndex).Char.CharIndex, vbWhite)
        Else
            If Npclist(TempCharIndex).MaestroUser > 0 Then
                If Npclist(TempCharIndex).MaestroUser = UserIndex Then
                    Call WriteConsoleMsg(1, UserIndex, "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ") " & Npclist(TempCharIndex).Name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).Name, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(1, UserIndex, estatus & Npclist(TempCharIndex).Name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).Name, FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(1, UserIndex, Npclist(TempCharIndex).Name & estatus & ".", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        FoundSomething = 1
        UserList(UserIndex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
        UserList(UserIndex).flags.TargetNPC = TempCharIndex
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
    End If
    
    If FoundChar = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
    End If
    
    '*** NO ENCOTRO NADA ***
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
    End If

Else
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
    End If
End If


End Sub

Function FindDirection(ByVal NPCI As Integer, Target As WorldPos) As eHeading

''*****************************************************************
''Devuelve la direccion en la cual el target se encuentra
''desde pos, 0 si la direc es igual
''*****************************************************************
'


Dim x As Integer
Dim Y As Integer
Dim Pos As WorldPos
Dim puedeX As Boolean
Dim puedeY As Boolean

Pos = Npclist(NPCI).Pos
x = Npclist(NPCI).Pos.x - Target.x
Y = Npclist(NPCI).Pos.Y - Target.Y
'
'misma
If Sgn(x) = 0 And Sgn(Y) = 0 Then
    FindDirection = 0
    Exit Function
End If
'
''Lo tenemos al lado
If Distancia(Pos, Target) = 1 Then
    FindDirection = 0
    Exit Function
End If
'

If Rodeado(Target) Then
    FindDirection = 0
    Exit Function
End If
'
'
''Sur
If Sgn(x) = 0 And Sgn(Y) = -1 Then
    If Not PuedeNpc(Pos.map, Pos.x, Pos.Y + 1) Then
        If RandomNumber(1, 10) > 5 Then
            If PuedeNpc(Pos.map, Pos.x - 1, Pos.Y) Then
                FindDirection = eHeading.WEST: Exit Function
            Else
                FindDirection = eHeading.EAST: Exit Function
            End If
        Else
            If PuedeNpc(Pos.map, Pos.x + 1, Pos.Y) Then
                FindDirection = eHeading.EAST: Exit Function
            Else
                FindDirection = eHeading.WEST: Exit Function
            End If
        End If
    Else
        FindDirection = eHeading.SOUTH: Exit Function
    End If
End If

''norte
If Sgn(x) = 0 And Sgn(Y) = 1 Then
    If Not PuedeNpc(Pos.map, Pos.x, Pos.Y - 1) Then
        If RandomNumber(1, 10) > 5 Then
            If PuedeNpc(Pos.map, Pos.x - 1, Pos.Y) Then
                FindDirection = eHeading.WEST: Exit Function
            Else
                FindDirection = eHeading.EAST: Exit Function
            End If
        Else
            If PuedeNpc(Pos.map, Pos.x + 1, Pos.Y) Then
                FindDirection = eHeading.EAST: Exit Function
            Else
                FindDirection = eHeading.WEST: Exit Function
            End If
        End If
    Else
        FindDirection = eHeading.NORTH: Exit Function
    End If
End If

''oeste
If Sgn(x) = 1 And Sgn(Y) = 0 Then
    If Not PuedeNpc(Pos.map, Pos.x - 1, Pos.Y) Then
        If RandomNumber(1, 10) > 5 Then
            If PuedeNpc(Pos.map, Pos.x, Pos.Y - 1) Then
                FindDirection = eHeading.NORTH: Exit Function
            Else
                FindDirection = eHeading.SOUTH: Exit Function
            End If
        Else
            If PuedeNpc(Pos.map, Pos.x, Pos.Y + 1) Then
                FindDirection = eHeading.SOUTH: Exit Function
            Else
                FindDirection = eHeading.NORTH: Exit Function
            End If
        End If
    Else
        FindDirection = eHeading.WEST: Exit Function
    End If
End If

''este
If Sgn(x) = -1 And Sgn(Y) = 0 Then
    If Not PuedeNpc(Pos.map, Pos.x + 1, Pos.Y) Then
        If RandomNumber(1, 10) > 5 Then
            If PuedeNpc(Pos.map, Pos.x, Pos.Y - 1) Then
                FindDirection = eHeading.NORTH: Exit Function
            Else
                FindDirection = eHeading.SOUTH: Exit Function
            End If
        Else
            If PuedeNpc(Pos.map, Pos.x, Pos.Y + 1) Then
                FindDirection = eHeading.SOUTH: Exit Function
            Else
                FindDirection = eHeading.NORTH: Exit Function
            End If
        End If
    Else
        FindDirection = eHeading.EAST: Exit Function
    End If
End If
'
''NW
If Sgn(x) = 1 And Sgn(Y) = 1 Then
    puedeX = PuedeNpc(Pos.map, Pos.x - 1, Pos.Y)
    puedeY = PuedeNpc(Pos.map, Pos.x, Pos.Y - 1)
    If puedeX And puedeY Then
        puedeX = Not (Npclist(NPCI).oldPos.x = Pos.x - 1)
        puedeY = Not (Npclist(NPCI).oldPos.Y = Pos.Y - 1)
        If puedeX And puedeY Then
            If RandomNumber(1, 20) < 10 Then
                FindDirection = eHeading.WEST: Exit Function
            Else
                FindDirection = eHeading.NORTH: Exit Function
            End If
        Else
            If puedeX Then
                FindDirection = eHeading.WEST: Exit Function
            ElseIf puedeY Then
                FindDirection = eHeading.NORTH: Exit Function
            End If
        End If
    ElseIf puedeX Then
        FindDirection = eHeading.WEST: Exit Function
    ElseIf puedeY Then
        FindDirection = eHeading.NORTH: Exit Function
    End If

'    'llego aca porque no pudo en nada
    puedeX = PuedeNpc(Pos.map, Pos.x - 1, Pos.Y)
    puedeY = PuedeNpc(Pos.map, Pos.x, Pos.Y + 1)
    If Not puedeY Or Npclist(NPCI).oldPos.Y = Pos.Y + 1 Then
        FindDirection = eHeading.EAST: Exit Function
    ElseIf puedeY Then
        FindDirection = eHeading.SOUTH: Exit Function
    End If
End If
'
''NE
If Sgn(x) = -1 And Sgn(Y) = 1 Then
    puedeX = PuedeNpc(Pos.map, Pos.x + 1, Pos.Y)
    puedeY = PuedeNpc(Pos.map, Pos.x, Pos.Y - 1)
    If puedeX And puedeY Then
        puedeX = Not (Npclist(NPCI).oldPos.x = Pos.x + 1)
        puedeY = Not (Npclist(NPCI).oldPos.Y = Pos.Y - 1)
        If puedeX And puedeY Then
            If RandomNumber(1, 20) < 10 Then
                FindDirection = eHeading.EAST: Exit Function
            Else
                FindDirection = eHeading.NORTH: Exit Function
            End If
        Else
            If puedeX Then
                FindDirection = eHeading.EAST: Exit Function
            ElseIf puedeY Then
                FindDirection = eHeading.NORTH: Exit Function
            End If
        End If
    ElseIf puedeX Then
        FindDirection = eHeading.EAST: Exit Function
    ElseIf puedeY Then
        FindDirection = eHeading.NORTH: Exit Function
    End If

'    'llego aca porque no pudo en nada
    puedeX = PuedeNpc(Pos.map, Pos.x - 1, Pos.Y)
    puedeY = PuedeNpc(Pos.map, Pos.x, Pos.Y + 1)
    If Not puedeY Or Npclist(NPCI).oldPos.Y = Pos.Y + 1 Then
        FindDirection = eHeading.WEST: Exit Function
    ElseIf puedeY Then
        FindDirection = eHeading.SOUTH: Exit Function
    End If
End If
'
''SW
If Sgn(x) = 1 And Sgn(Y) = -1 Then
    puedeX = PuedeNpc(Pos.map, Pos.x - 1, Pos.Y)
    puedeY = PuedeNpc(Pos.map, Pos.x, Pos.Y + 1)
    If puedeX And puedeY Then
        puedeX = Not (Npclist(NPCI).oldPos.x = Pos.x - 1)
        puedeY = Not (Npclist(NPCI).oldPos.Y = Pos.Y + 1)
        If puedeX And puedeY Then
            If RandomNumber(1, 20) < 10 Then
                FindDirection = eHeading.WEST: Exit Function
            Else
                FindDirection = eHeading.SOUTH: Exit Function
            End If
       Else
            If puedeX Then
                FindDirection = eHeading.WEST: Exit Function
            ElseIf puedeY Then
                FindDirection = eHeading.SOUTH: Exit Function
            End If
        End If
    ElseIf puedeX Then
        FindDirection = eHeading.WEST: Exit Function
    ElseIf puedeY Then
        FindDirection = eHeading.SOUTH: Exit Function
    End If

'    'llego aca porque no pudo en nada
    puedeX = PuedeNpc(Pos.map, Pos.x + 1, Pos.Y)
    puedeY = PuedeNpc(Pos.map, Pos.x, Pos.Y - 1)
    If Not puedeY Or Npclist(NPCI).oldPos.Y = Pos.Y - 1 Then
        FindDirection = eHeading.EAST: Exit Function
    ElseIf puedeY Then
        FindDirection = eHeading.NORTH: Exit Function
    End If
End If

''SE
If Sgn(x) = -1 And Sgn(Y) = -1 Then
    puedeX = PuedeNpc(Pos.map, Pos.x + 1, Pos.Y)
    puedeY = PuedeNpc(Pos.map, Pos.x, Pos.Y + 1)
    If puedeX And puedeY Then
        puedeX = Not (Npclist(NPCI).oldPos.x = Pos.x + 1)
        puedeY = Not (Npclist(NPCI).oldPos.Y = Pos.Y + 1)
        If puedeX And puedeY Then
            If RandomNumber(1, 20) < 10 Then
                FindDirection = eHeading.EAST: Exit Function
            Else
                FindDirection = eHeading.SOUTH: Exit Function
           End If
        Else
            If puedeX Then
                FindDirection = eHeading.EAST: Exit Function
            ElseIf puedeY Then
                FindDirection = eHeading.SOUTH: Exit Function
            End If
        End If
    ElseIf puedeX Then
        FindDirection = eHeading.EAST: Exit Function
    ElseIf puedeY Then
        FindDirection = eHeading.SOUTH: Exit Function
    End If

    'llego aca porque no pudo en nada
    puedeX = PuedeNpc(Pos.map, Pos.x - 1, Pos.Y)
    puedeY = PuedeNpc(Pos.map, Pos.x, Pos.Y - 1)
    If Not puedeY Or Npclist(NPCI).oldPos.Y = Pos.Y - 1 Then
        FindDirection = eHeading.WEST: Exit Function
    Else
        FindDirection = eHeading.NORTH: Exit Function
    End If
End If

End Function
Function Rodeado(ByRef Pos As WorldPos) As Boolean
   
    If Not PuedeNpc(Pos.map, Pos.x + 1, Pos.Y) Then
        If Not PuedeNpc(Pos.map, Pos.x - 1, Pos.Y) Then
            If Not PuedeNpc(Pos.map, Pos.x, Pos.Y + 1) Then
                If Not PuedeNpc(Pos.map, Pos.x, Pos.Y - 1) Then
                    Rodeado = True
                End If
            End If
        End If
    End If
End Function
Function PuedeNpc(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer)
On Error GoTo hayerror
    'Add Marius agregamos el if para la validacion asi no se va el indice a la mierda.
    If x > 0 And Y > 0 Then
        PuedeNpc = (MapData(map, x, Y).NpcIndex = 0 And _
                    MapData(map, x, Y).Blocked = 0 And _
                    MapData(map, x, Y).UserIndex = 0)
    Else
        PuedeNpc = False
    End If
                
    Exit Function
hayerror:
     LogError ("Error en PuedeNPC:" & err.Number & " Descripcion: " & err.description & " map:" & map & " x:" & x & " y:" & Y)
End Function
Function FindDonde(ByVal NPCI As Integer, Target As WorldPos) As eHeading
Dim x As Integer
Dim Y As Integer
Dim Pos As WorldPos

    Pos = Npclist(NPCI).Pos
    x = Npclist(NPCI).Pos.x - Target.x
    Y = Npclist(NPCI).Pos.Y - Target.Y
    
    If Sgn(x) = 0 And Sgn(Y) = 0 Then FindDonde = 0: Exit Function
    
    If Sgn(x) = 0 And Sgn(Y) = -1 Then FindDonde = eHeading.SOUTH
    If Sgn(x) = 0 And Sgn(Y) = 1 Then FindDonde = eHeading.NORTH
    If Sgn(x) = 1 And Sgn(Y) = 0 Then FindDonde = eHeading.WEST
    If Sgn(x) = -1 And Sgn(Y) = 0 Then FindDonde = eHeading.EAST

End Function
'[Barrin 30-11-03]
Public Function ItemNoEsDeMapa(ByVal index As Integer) As Boolean

ItemNoEsDeMapa = ObjData(index).OBJType <> eOBJType.otPuertas And _
            ObjData(index).OBJType <> eOBJType.otForos And _
            ObjData(index).OBJType <> eOBJType.otCarteles And _
            ObjData(index).OBJType <> eOBJType.otArboles And _
            ObjData(index).OBJType <> eOBJType.otYacimiento And _
            ObjData(index).OBJType <> eOBJType.otTeleport And _
            ObjData(index).OBJType <> eOBJType.otCorreo
End Function
'[/Barrin 30-11-03]

Public Function MostrarCantidad(ByVal index As Integer) As Boolean
MostrarCantidad = ObjData(index).OBJType <> eOBJType.otPuertas And _
            ObjData(index).OBJType <> eOBJType.otForos And _
            ObjData(index).OBJType <> eOBJType.otCarteles And _
            ObjData(index).OBJType <> eOBJType.otArboles And _
            ObjData(index).OBJType <> eOBJType.otYacimiento And _
            ObjData(index).OBJType <> eOBJType.otTeleport And _
            ObjData(index).OBJType <> eOBJType.otCorreo
End Function

Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean

EsObjetoFijo = OBJType = eOBJType.otForos Or _
               OBJType = eOBJType.otCarteles Or _
               OBJType = eOBJType.otArboles Or _
               OBJType = eOBJType.otYacimiento Or _
               OBJType = eOBJType.otCorreo Or _
               OBJType = eOBJType.otArboles

End Function
Public Function ParticleToLevel(ByVal UserIndex As Integer) As Integer

    If UserList(UserIndex).Stats.ELV < 13 Then
        If UserList(UserIndex).donador = False Then
            ParticleToLevel = 42
        Else
            ParticleToLevel = 121
        End If
    ElseIf UserList(UserIndex).Stats.ELV < 25 Then
        If UserList(UserIndex).donador = False Then
            ParticleToLevel = 2
        Else
            ParticleToLevel = 122
        End If
    ElseIf UserList(UserIndex).Stats.ELV < 35 Then
        If UserList(UserIndex).donador = False Then
            ParticleToLevel = 81
        Else
            ParticleToLevel = 123
        End If
    ElseIf UserList(UserIndex).Stats.ELV < 50 Then
        
        'Des Marius a pedido de la gente. Estaria para hacer las mismas meditaciones con el mismo color todo y con estrellitas
        'If UserList(UserIndex).donador = False Then
        
            If UserList(UserIndex).faccion.Renegado = 1 Then
                ParticleToLevel = 39
            ElseIf UserList(UserIndex).faccion.Ciudadano = 1 Then
                ParticleToLevel = 40
            ElseIf UserList(UserIndex).faccion.Republicano = 1 Then
                ParticleToLevel = 71
            ElseIf UserList(UserIndex).faccion.ArmadaReal = 1 Then
                ParticleToLevel = 38
            ElseIf UserList(UserIndex).faccion.FuerzasCaos = 1 Then
                ParticleToLevel = 37
            ElseIf UserList(UserIndex).faccion.Milicia = 1 Then
                ParticleToLevel = 66
            End If
            
        'Else
        '    ParticleToLevel = 120
        'End If
        
    Else
    
        If UserList(UserIndex).donador = False Then
            ParticleToLevel = 36
        Else
            ParticleToLevel = 124
        End If
        
    End If
    
    Debug.Print "particula: " & ParticleToLevel
End Function

Public Sub ReproducirSonido(ByVal Destino As SendTarget, ByVal index As Integer, ByVal SoundIndex As Integer)
    Call SendData(Destino, index, PrepareMessagePlayWave(SoundIndex, UserList(index).Pos.x, UserList(index).Pos.Y))
End Sub
Function Tilde(ByRef F As String) As String
    Tilde = Replace$(Replace$(Replace$(Replace$(Replace$(F, "í", "i"), "è", "e"), "ó", "o"), "á", "a"), "ú", "u")
End Function
