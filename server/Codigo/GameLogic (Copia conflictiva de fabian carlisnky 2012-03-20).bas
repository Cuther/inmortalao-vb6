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
Public Function EsNewbie(ByVal userindex As Integer) As Boolean
    EsNewbie = UserList(userindex).Stats.ELV <= LimiteNewbie
End Function
Public Function esArmada(ByVal userindex As Integer) As Boolean
    esArmada = (UserList(userindex).Faccion.ArmadaReal = 1)
End Function
Public Function esCaos(ByVal userindex As Integer) As Boolean
    esCaos = (UserList(userindex).Faccion.FuerzasCaos = 1)
End Function
Public Function esMili(ByVal userindex As Integer) As Boolean
    esMili = (UserList(userindex).Faccion.Milicia = 1)
End Function
Public Function esFaccion(ByVal userindex As Integer) As Boolean
    esFaccion = (UserList(userindex).Faccion.ArmadaReal = 1 Or UserList(userindex).Faccion.FuerzasCaos = 1 Or UserList(userindex).Faccion.Milicia = 1)
End Function
Public Function esRene(ByVal userindex As Integer) As Boolean
    esRene = (UserList(userindex).Faccion.Renegado)
End Function
Public Function esCiuda(ByVal userindex As Integer) As Boolean
    esCiuda = (UserList(userindex).Faccion.Ciudadano)
End Function
Public Function esRepu(ByVal userindex As Integer) As Boolean
    esRepu = (UserList(userindex).Faccion.Republicano)
End Function
Public Function esMismoBando(ByVal U1 As Integer, ByVal U2 As Integer) As Boolean
    With UserList(U1)
        esMismoBando = (.Faccion.Republicano = 1 And UserList(U2).Faccion.Republicano = 1) Or ( _
                        .Faccion.Ciudadano = 1 And UserList(U2).Faccion.Ciudadano = 1)
    End With
End Function
Public Function EsGM(ByVal userindex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************
    EsGM = (UserList(userindex).flags.Privilegios And (PlayerType.Dios Or PlayerType.VIP))
End Function

Public Sub DoTileEvents(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer)
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
                    If EsNewbie(userindex) Or EsGM(userindex) Then
                        If LegalPos(.TileExit.map, .TileExit.x, .TileExit.Y, PuedeAtravesarAgua(userindex)) Then
                            Call WarpUserChar(userindex, .TileExit.map, .TileExit.x, .TileExit.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(.TileExit, nPos)
                            If nPos.x <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(userindex, nPos.map, nPos.x, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es newbie
                        Call WriteConsoleMsg(1, userindex, "Mapa exclusivo para newbies.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(userindex).pos, nPos)
        
                        If nPos.x <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(userindex, nPos.map, nPos.x, nPos.Y, False)
                        End If
                    End If
                    
                    
                    
                ElseIf .TileExit.map = 757 Or .TileExit.map = 760 Or .TileExit.map = 250 Then
                    '¿El usuario es donador
                    If UserList(userindex).donador = 1 Or EsGM(userindex) Then
                        If LegalPos(.TileExit.map, .TileExit.x, .TileExit.Y, PuedeAtravesarAgua(userindex)) Then
                            Call WarpUserChar(userindex, .TileExit.map, .TileExit.x, .TileExit.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(.TileExit, nPos)
                            If nPos.x <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(userindex, nPos.map, nPos.x, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es donador
                        Call WriteConsoleMsg(1, userindex, "Mapa exclusivo para DONADORES. Enterate mas en http://inmortalao.com.ar/", FontTypeNames.FONTTYPE_BROWNI)
                        Call ClosestStablePos(UserList(userindex).pos, nPos)
        
                        If nPos.x <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(userindex, nPos.map, nPos.x, nPos.Y, False)
                        End If
                    End If
                    
                    
                    
                ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "ARMADA" Then '¿Es mapa de Armadas?
                    '¿El usuario es Armada?
                    If esArmada(userindex) Or EsGM(userindex) Then
                        If LegalPos(.TileExit.map, .TileExit.x, .TileExit.Y, PuedeAtravesarAgua(userindex)) Then
                            Call WarpUserChar(userindex, .TileExit.map, .TileExit.x, .TileExit.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(.TileExit, nPos)
                            If nPos.x <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(userindex, nPos.map, nPos.x, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es armada
                        Call WriteConsoleMsg(1, userindex, "Mapa exclusivo para miembros del ejército Real", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(userindex).pos, nPos)
                        
                        If nPos.x <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(userindex, nPos.map, nPos.x, nPos.Y, FxFlag)
                        End If
                    End If
                ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "CAOS" Then '¿Es mapa de Caos?
                    '¿El usuario es Caos?
                    If esCaos(userindex) Or EsGM(userindex) Then
                        If LegalPos(.TileExit.map, .TileExit.x, .TileExit.Y, PuedeAtravesarAgua(userindex)) Then
                            Call WarpUserChar(userindex, .TileExit.map, .TileExit.x, .TileExit.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(.TileExit, nPos)
                            If nPos.x <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(userindex, nPos.map, nPos.x, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es caos
                        Call WriteConsoleMsg(1, userindex, "Mapa exclusivo para miembros del ejército Oscuro.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(userindex).pos, nPos)
                        
                        If nPos.x <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(userindex, nPos.map, nPos.x, nPos.Y, FxFlag)
                        End If
                    End If
                ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "FACCION" Then '¿Es mapa de faccionarios?
                    '¿El usuario es Armada o Caos?
                    If esArmada(userindex) Or esCaos(userindex) Or EsGM(userindex) Then
                        If LegalPos(.TileExit.map, .TileExit.x, .TileExit.Y, PuedeAtravesarAgua(userindex)) Then
                            Call WarpUserChar(userindex, .TileExit.map, .TileExit.x, .TileExit.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(.TileExit, nPos)
                            If nPos.x <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(userindex, nPos.map, nPos.x, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es Faccionario
                        Call WriteConsoleMsg(1, userindex, "Solo se permite entrar al Mapa si eres miembro de alguna Facción", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(userindex).pos, nPos)
                        
                        If nPos.x <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(userindex, nPos.map, nPos.x, nPos.Y, FxFlag)
                        End If
                    End If
                Else 'No es un mapa de newbies, ni Armadas, ni Caos, ni faccionario.
                    If LegalPos(.TileExit.map, .TileExit.x, .TileExit.Y, PuedeAtravesarAgua(userindex)) Then
                        Call WarpUserChar(userindex, .TileExit.map, .TileExit.x, .TileExit.Y, FxFlag)
                    Else
                        Call ClosestLegalPos(.TileExit, nPos)
                        If nPos.x <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(userindex, nPos.map, nPos.x, nPos.Y, FxFlag)
                        End If
                    End If
                End If
                
                'Te fusite del mapa. La criatura ya no es más tuya ni te reconoce como que vos la atacaste.
                Dim aN As Integer
                
                aN = UserList(userindex).flags.AtacadoPorNpc
                If aN > 0 Then
                   Npclist(aN).Movement = Npclist(aN).flags.OldMovement
                   Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
                   Npclist(aN).flags.AttackedBy = 0
                End If
            
                aN = UserList(userindex).flags.NPCAtacado
                If aN > 0 Then
                    If Npclist(aN).flags.AttackedFirstBy = UserList(userindex).Name Then
                        Npclist(aN).flags.AttackedFirstBy = vbNullString
                    End If
                End If
                UserList(userindex).flags.AtacadoPorNpc = 0
                UserList(userindex).flags.NPCAtacado = 0
                
            End If
        End With
    End If
Exit Sub

Errhandler:
    Call LogError("Error en DoTileEvents. Error: " & err.Number & " - Desc: " & err.description)
End Sub

Function InRangoVision(ByVal userindex As Integer, ByVal x As Integer, ByVal Y As Integer) As Boolean

If x > UserList(userindex).pos.x - MinXBorder And x < UserList(userindex).pos.x + MinXBorder Then
    If Y > UserList(userindex).pos.Y - MinYBorder And Y < UserList(userindex).pos.Y + MinYBorder Then
        InRangoVision = True
        Exit Function
    End If
End If
InRangoVision = False

End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, x As Integer, Y As Integer) As Boolean

If x > Npclist(NpcIndex).pos.x - MinXBorder And x < Npclist(NpcIndex).pos.x + MinXBorder Then
    If Y > Npclist(NpcIndex).pos.Y - MinYBorder And Y < Npclist(NpcIndex).pos.Y + MinYBorder Then
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

Sub ClosestLegalPos(pos As WorldPos, ByRef nPos As WorldPos, Optional PuedeAgua As Boolean = False, Optional PuedeTierra As Boolean = True)
'*****************************************************************
'Author: Unknown (original version)
'Last Modification: 24/01/2007 (ToxicWaste)
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Long
Dim tY As Long

nPos.map = pos.map

Do While Not LegalPos(pos.map, nPos.x, nPos.Y, PuedeAgua, PuedeTierra)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = pos.Y - LoopC To pos.Y + LoopC
        For tX = pos.x - LoopC To pos.x + LoopC
            
            If LegalPos(nPos.map, tX, tY, PuedeAgua, PuedeTierra) Then
                nPos.x = tX
                nPos.Y = tY
                '¿Hay objeto?
                
                tX = pos.x + LoopC
                tY = pos.Y + LoopC
  
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.x = 0
    nPos.Y = 0
End If

End Sub

Private Sub ClosestStablePos(pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Long
Dim tY As Long

nPos.map = pos.map

Do While Not LegalPos(pos.map, nPos.x, nPos.Y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = pos.Y - LoopC To pos.Y + LoopC
        For tX = pos.x - LoopC To pos.x + LoopC
            
            If LegalPos(nPos.map, tX, tY) And MapData(nPos.map, tX, tY).TileExit.map = 0 Then
                nPos.x = tX
                nPos.Y = tY
                '¿Hay objeto?
                
                tX = pos.x + LoopC
                tY = pos.Y + LoopC
  
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.x = 0
    nPos.Y = 0
End If

End Sub

Function NameIndex(ByVal Name As String) As Integer
    Dim userindex As Long
    
    '¿Nombre valido?
    If LenB(Name) = 0 Then
        NameIndex = 0
        Exit Function
    End If
    
    If InStrB(Name, "+") <> 0 Then
        Name = UCase$(Replace(Name, "+", " "))
    End If
    
    userindex = 1
    Do Until UCase$(UserList(userindex).Name) = UCase$(Name)
        
        userindex = userindex + 1
        
        If userindex > MaxUsers Then
            NameIndex = 0
            Exit Function
        End If
    Loop
     
    NameIndex = userindex
End Function

Function CheckForSameIP(ByVal userindex As Integer, ByVal UserIP As String) As Boolean
    Dim LoopC As Long
    'BY CASTELLI... la extencion del for es preferible que sea
    'hasta last user...
    
    
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged = True Then
            If UserList(LoopC).ip = UserIP And userindex <> LoopC Then
                CheckForSameIP = True
                Exit Function
            End If
        End If
    Next LoopC
    
    CheckForSameIP = False
End Function

Function CheckForSameName(ByVal Name As String) As Boolean
'Controlo que no existan usuarios con el mismo nombre
    Dim LoopC As Long
    
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged Then
            
            'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
            'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
            'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
            'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
            'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
            
            If UCase$(UserList(LoopC).Name) = UCase$(Name) Then
                CheckForSameName = True
                'UserList(LoopC).Counters.Saliendo = True
                'UserList(LoopC).Counters.Salir = 1
                Exit Function
            End If
        End If
    Next LoopC
    
    CheckForSameName = False
End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef pos As WorldPos)
'*****************************************************************
'Toma una posicion y se mueve hacia donde esta perfilado
'*****************************************************************
    Select Case Head
        Case eHeading.NORTH
            pos.Y = pos.Y - 1
        
        Case eHeading.SOUTH
            pos.Y = pos.Y + 1
        
        Case eHeading.EAST
            pos.x = pos.x + 1
        
        Case eHeading.WEST
            pos.x = pos.x - 1
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
                   (MapData(map, x, Y).userindex = 0) And _
                   (MapData(map, x, Y).NpcIndex = 0)
    ElseIf PuedeTierra And Not PuedeAgua Then
        LegalPos = (MapData(map, x, Y).Blocked <> 1) And _
                   (MapData(map, x, Y).userindex = 0) And _
                   (MapData(map, x, Y).NpcIndex = 0) And _
                   (Not HayAgua(map, x, Y))
    ElseIf PuedeAgua And Not PuedeTierra Then
        LegalPos = (MapData(map, x, Y).Blocked <> 1) And _
                   (MapData(map, x, Y).userindex = 0) And _
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

Dim userindex As Integer
Dim IsDeadChar As Boolean


'¿Es un mapa valido?
If (map <= 0 Or map > NumMaps) Or _
   (x < MinXBorder Or x > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
            MoveToLegalPos = False
    Else
        userindex = MapData(map, x, Y).userindex
        If userindex > 0 Then
            IsDeadChar = UserList(userindex).flags.Muerto = 1
        Else
            IsDeadChar = False
        End If
    
    If PuedeAgua And PuedeTierra Then
        MoveToLegalPos = (MapData(map, x, Y).Blocked <> 1) And _
                   (userindex = 0 Or IsDeadChar) And _
                   (MapData(map, x, Y).NpcIndex = 0)
    ElseIf PuedeTierra And Not PuedeAgua Then
        MoveToLegalPos = (MapData(map, x, Y).Blocked <> 1) And _
                   (userindex = 0 Or IsDeadChar) And _
                   (MapData(map, x, Y).NpcIndex = 0) And _
                   (Not HayAgua(map, x, Y))
    ElseIf PuedeAgua And Not PuedeTierra Then
        MoveToLegalPos = (MapData(map, x, Y).Blocked <> 1) And _
                   (userindex = 0 Or IsDeadChar) And _
                   (MapData(map, x, Y).NpcIndex = 0) And _
                   (HayAgua(map, x, Y))
    Else
        MoveToLegalPos = False
    End If
  
End If


End Function

Public Sub FindLegalPos(ByVal userindex As Integer, ByVal map As Integer, ByRef x As Integer, ByRef Y As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 26/03/2009
'Search for a Legal pos for the user who is being teleported.
'***************************************************

    If MapData(map, x, Y).userindex <> 0 Or _
        MapData(map, x, Y).NpcIndex <> 0 Then
                    
        ' Se teletransporta a la misma pos a la que estaba
        If MapData(map, x, Y).userindex = userindex Then Exit Sub
                            
        Dim FoundPlace As Boolean
        Dim tX As Long
        Dim tY As Long
        Dim Rango As Long
        Dim OtherUserIndex As Integer
        Rango = 5
        
        For tY = Y - Rango To Y + Rango
            For tX = x - Rango To x + Rango
                'Reviso que no haya User ni NPC
                If MapData(map, tX, tY).userindex = 0 And _
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
Dim userindex As Integer

    If (map <= 0 Or map > NumMaps) Or _
        (x < MinXBorder Or x > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPosNPC = False
        Exit Function
    End If

    userindex = MapData(map, x, Y).userindex
    If userindex > 0 Then
        IsDeadChar = UserList(userindex).flags.Muerto = 1
    Else
        IsDeadChar = False
    End If
    
    If AguaValida = 0 Then
        LegalPosNPC = (MapData(map, x, Y).Blocked <> 1) And _
        (MapData(map, x, Y).userindex = 0 Or IsDeadChar) And _
        (MapData(map, x, Y).NpcIndex = 0) And _
        (MapData(map, x, Y).Trigger <> eTrigger.POSINVALIDA) _
        And Not HayAgua(map, x, Y)
    Else
        LegalPosNPC = (MapData(map, x, Y).Blocked <> 1) And _
        (MapData(map, x, Y).userindex = 0 Or IsDeadChar) And _
        (MapData(map, x, Y).NpcIndex = 0) And _
        (MapData(map, x, Y).Trigger <> eTrigger.POSINVALIDA)
    End If
End Function

Sub SendHelp(ByVal index As Integer)
Dim NumHelpLines As Integer
Dim LoopC As Integer

NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

For LoopC = 1 To NumHelpLines
    Call WriteConsoleMsg(1, index, GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC), FontTypeNames.FONTTYPE_INFO)
Next LoopC

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal userindex As Integer)
    If Npclist(NpcIndex).NroExpresiones > 0 Then
        Dim randomi
        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageChatOverHead(Npclist(NpcIndex).Expresiones(randomi), Npclist(NpcIndex).Char.CharIndex, vbWhite))
    End If
End Sub

Sub LookatTile(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer)
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
If (Abs(UserList(userindex).pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(userindex).pos.x - x) > RANGO_VISION_X) Then
    Exit Sub
End If

'¿Posicion valida?
If InMapBounds(map, x, Y) Then
    UserList(userindex).flags.TargetMap = map
    UserList(userindex).flags.TargetX = x
    UserList(userindex).flags.TargetY = Y
    '¿Es un obj?
    If MapData(map, x, Y).ObjInfo.ObjIndex > 0 Then
        'Informa el nombre
        UserList(userindex).flags.TargetObjMap = map
        UserList(userindex).flags.TargetObjX = x
        UserList(userindex).flags.TargetObjY = Y
        FoundSomething = 1
    ElseIf MapData(map, x + 1, Y).ObjInfo.ObjIndex > 0 Then
        'Informa el nombre
        If ObjData(MapData(map, x + 1, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            UserList(userindex).flags.TargetObjMap = map
            UserList(userindex).flags.TargetObjX = x + 1
            UserList(userindex).flags.TargetObjY = Y
            FoundSomething = 1
        End If
    ElseIf MapData(map, x + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
        If ObjData(MapData(map, x + 1, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            'Informa el nombre
            UserList(userindex).flags.TargetObjMap = map
            UserList(userindex).flags.TargetObjX = x + 1
            UserList(userindex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    ElseIf MapData(map, x, Y + 1).ObjInfo.ObjIndex > 0 Then
        If ObjData(MapData(map, x, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            'Informa el nombre
            UserList(userindex).flags.TargetObjMap = map
            UserList(userindex).flags.TargetObjX = x
            UserList(userindex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    End If
    
    If FoundSomething = 1 Then
        UserList(userindex).flags.TargetObj = MapData(map, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).ObjInfo.ObjIndex
    End If
    
    '¿Es un personaje?
    If Y + 1 <= YMaxMapSize Then
        If MapData(map, x, Y + 1).userindex > 0 Then
            TempCharIndex = MapData(map, x, Y + 1).userindex
            FoundChar = 1
        End If
        If MapData(map, x, Y + 1).NpcIndex > 0 Then
            TempCharIndex = MapData(map, x, Y + 1).NpcIndex
            FoundChar = 2
        End If
    End If
    '¿Es un personaje?
    If FoundChar = 0 Then
        If MapData(map, x, Y).userindex > 0 Then
            TempCharIndex = MapData(map, x, Y).userindex
            FoundChar = 1
        End If
        If MapData(map, x, Y).NpcIndex > 0 Then
            TempCharIndex = MapData(map, x, Y).NpcIndex
            FoundChar = 2
        End If
    End If
    
    'Reaccion al personaje
    If FoundChar = 1 Then '  ¿Encontro un Usuario?
       If UserList(TempCharIndex).flags.AdminInvisible = 0 Or UserList(userindex).flags.Privilegios And PlayerType.User Then
            If UserList(TempCharIndex).showName Then
                WriteCharMsgStatus userindex, TempCharIndex
            End If
            
            FoundSomething = 1
            UserList(userindex).flags.TargetUser = TempCharIndex
            UserList(userindex).flags.TargetNPC = 0
            UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
       End If
    End If
    
    If FoundChar = 2 Then '¿Encontro un NPC?
        Dim estatus As String
        
        If UserList(userindex).flags.Privilegios And (PlayerType.VIP Or PlayerType.Dios) Or UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) = 100 Then
            estatus = "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ") "
        Else
            If UserList(userindex).flags.Muerto = 0 Then
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
            Call WriteChatOverHead(userindex, Npclist(TempCharIndex).desc, Npclist(TempCharIndex).Char.CharIndex, vbWhite)
        Else
            If Npclist(TempCharIndex).MaestroUser > 0 Then
                If Npclist(TempCharIndex).MaestroUser = userindex Then
                    Call WriteConsoleMsg(1, userindex, "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ") " & Npclist(TempCharIndex).Name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).Name, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(1, userindex, estatus & Npclist(TempCharIndex).Name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).Name, FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(1, userindex, Npclist(TempCharIndex).Name & estatus & ".", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        FoundSomething = 1
        UserList(userindex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
        UserList(userindex).flags.TargetNPC = TempCharIndex
        UserList(userindex).flags.TargetUser = 0
        UserList(userindex).flags.TargetObj = 0
    End If
    
    If FoundChar = 0 Then
        UserList(userindex).flags.TargetNPC = 0
        UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(userindex).flags.TargetUser = 0
    End If
    
    '*** NO ENCOTRO NADA ***
    If FoundSomething = 0 Then
        UserList(userindex).flags.TargetNPC = 0
        UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(userindex).flags.TargetUser = 0
        UserList(userindex).flags.TargetObj = 0
        UserList(userindex).flags.TargetObjMap = 0
        UserList(userindex).flags.TargetObjX = 0
        UserList(userindex).flags.TargetObjY = 0
    End If

Else
    If FoundSomething = 0 Then
        UserList(userindex).flags.TargetNPC = 0
        UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(userindex).flags.TargetUser = 0
        UserList(userindex).flags.TargetObj = 0
        UserList(userindex).flags.TargetObjMap = 0
        UserList(userindex).flags.TargetObjX = 0
        UserList(userindex).flags.TargetObjY = 0
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
Dim pos As WorldPos
Dim puedeX As Boolean
Dim puedeY As Boolean

pos = Npclist(NPCI).pos
x = Npclist(NPCI).pos.x - Target.x
Y = Npclist(NPCI).pos.Y - Target.Y
'
'misma
If Sgn(x) = 0 And Sgn(Y) = 0 Then
    FindDirection = 0
    Exit Function
End If
'
''Lo tenemos al lado
If Distancia(pos, Target) = 1 Then
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
    If Not PuedeNpc(pos.map, pos.x, pos.Y + 1) Then
        If RandomNumber(1, 10) > 5 Then
            If PuedeNpc(pos.map, pos.x - 1, pos.Y) Then
                FindDirection = eHeading.WEST: Exit Function
            Else
                FindDirection = eHeading.EAST: Exit Function
            End If
        Else
            If PuedeNpc(pos.map, pos.x + 1, pos.Y) Then
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
    If Not PuedeNpc(pos.map, pos.x, pos.Y - 1) Then
        If RandomNumber(1, 10) > 5 Then
            If PuedeNpc(pos.map, pos.x - 1, pos.Y) Then
                FindDirection = eHeading.WEST: Exit Function
            Else
                FindDirection = eHeading.EAST: Exit Function
            End If
        Else
            If PuedeNpc(pos.map, pos.x + 1, pos.Y) Then
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
    If Not PuedeNpc(pos.map, pos.x - 1, pos.Y) Then
        If RandomNumber(1, 10) > 5 Then
            If PuedeNpc(pos.map, pos.x, pos.Y - 1) Then
                FindDirection = eHeading.NORTH: Exit Function
            Else
                FindDirection = eHeading.SOUTH: Exit Function
            End If
        Else
            If PuedeNpc(pos.map, pos.x, pos.Y + 1) Then
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
    If Not PuedeNpc(pos.map, pos.x + 1, pos.Y) Then
        If RandomNumber(1, 10) > 5 Then
            If PuedeNpc(pos.map, pos.x, pos.Y - 1) Then
                FindDirection = eHeading.NORTH: Exit Function
            Else
                FindDirection = eHeading.SOUTH: Exit Function
            End If
        Else
            If PuedeNpc(pos.map, pos.x, pos.Y + 1) Then
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
    puedeX = PuedeNpc(pos.map, pos.x - 1, pos.Y)
    puedeY = PuedeNpc(pos.map, pos.x, pos.Y - 1)
    If puedeX And puedeY Then
        puedeX = Not (Npclist(NPCI).oldPos.x = pos.x - 1)
        puedeY = Not (Npclist(NPCI).oldPos.Y = pos.Y - 1)
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
    puedeX = PuedeNpc(pos.map, pos.x - 1, pos.Y)
    puedeY = PuedeNpc(pos.map, pos.x, pos.Y + 1)
    If Not puedeY Or Npclist(NPCI).oldPos.Y = pos.Y + 1 Then
        FindDirection = eHeading.EAST: Exit Function
    ElseIf puedeY Then
        FindDirection = eHeading.SOUTH: Exit Function
    End If
End If
'
''NE
If Sgn(x) = -1 And Sgn(Y) = 1 Then
    puedeX = PuedeNpc(pos.map, pos.x + 1, pos.Y)
    puedeY = PuedeNpc(pos.map, pos.x, pos.Y - 1)
    If puedeX And puedeY Then
        puedeX = Not (Npclist(NPCI).oldPos.x = pos.x + 1)
        puedeY = Not (Npclist(NPCI).oldPos.Y = pos.Y - 1)
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
    puedeX = PuedeNpc(pos.map, pos.x - 1, pos.Y)
    puedeY = PuedeNpc(pos.map, pos.x, pos.Y + 1)
    If Not puedeY Or Npclist(NPCI).oldPos.Y = pos.Y + 1 Then
        FindDirection = eHeading.WEST: Exit Function
    ElseIf puedeY Then
        FindDirection = eHeading.SOUTH: Exit Function
    End If
End If
'
''SW
If Sgn(x) = 1 And Sgn(Y) = -1 Then
    puedeX = PuedeNpc(pos.map, pos.x - 1, pos.Y)
    puedeY = PuedeNpc(pos.map, pos.x, pos.Y + 1)
    If puedeX And puedeY Then
        puedeX = Not (Npclist(NPCI).oldPos.x = pos.x - 1)
        puedeY = Not (Npclist(NPCI).oldPos.Y = pos.Y + 1)
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
    puedeX = PuedeNpc(pos.map, pos.x + 1, pos.Y)
    puedeY = PuedeNpc(pos.map, pos.x, pos.Y - 1)
    If Not puedeY Or Npclist(NPCI).oldPos.Y = pos.Y - 1 Then
        FindDirection = eHeading.EAST: Exit Function
    ElseIf puedeY Then
        FindDirection = eHeading.NORTH: Exit Function
    End If
End If

''SE
If Sgn(x) = -1 And Sgn(Y) = -1 Then
    puedeX = PuedeNpc(pos.map, pos.x + 1, pos.Y)
    puedeY = PuedeNpc(pos.map, pos.x, pos.Y + 1)
    If puedeX And puedeY Then
        puedeX = Not (Npclist(NPCI).oldPos.x = pos.x + 1)
        puedeY = Not (Npclist(NPCI).oldPos.Y = pos.Y + 1)
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
    puedeX = PuedeNpc(pos.map, pos.x - 1, pos.Y)
    puedeY = PuedeNpc(pos.map, pos.x, pos.Y - 1)
    If Not puedeY Or Npclist(NPCI).oldPos.Y = pos.Y - 1 Then
        FindDirection = eHeading.WEST: Exit Function
    Else
        FindDirection = eHeading.NORTH: Exit Function
    End If
End If

End Function
Function Rodeado(ByRef pos As WorldPos) As Boolean
   
    If Not PuedeNpc(pos.map, pos.x + 1, pos.Y) Then
        If Not PuedeNpc(pos.map, pos.x - 1, pos.Y) Then
            If Not PuedeNpc(pos.map, pos.x, pos.Y + 1) Then
                If Not PuedeNpc(pos.map, pos.x, pos.Y - 1) Then
                    Rodeado = True
                End If
            End If
        End If
    End If
End Function
Function PuedeNpc(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer)
' ---Castelli--- ///_
'On Error Resume Next ' Puse esto porq cuando se mata a un bicho pierde su poss
'y queda en un out of range, entonces lo mas sencillo es esto y no es problematico
'saque los on error goto err.... del Timer_AI y NPCAI que hacian matar a las mascotas
'Esto fue todooo.. chau... xD
' ---Castelli--- ///_
On Error GoTo hayerror

    PuedeNpc = (MapData(map, x, Y).NpcIndex = 0 And _
                MapData(map, x, Y).Blocked = 0 And _
                MapData(map, x, Y).userindex = 0)
                
                
                Exit Function
                
hayerror:
     LogError ("Error en PuedeNPC:" & err.Number & " Descripcion: " & err.description)
        
                
End Function
Function FindDonde(ByVal NPCI As Integer, Target As WorldPos) As eHeading
Dim x As Integer
Dim Y As Integer
Dim pos As WorldPos

pos = Npclist(NPCI).pos
x = Npclist(NPCI).pos.x - Target.x
Y = Npclist(NPCI).pos.Y - Target.Y

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
Public Function ParticleToLevel(ByVal userindex As Integer) As Integer
If UserList(userindex).Stats.ELV < 13 Then
    ParticleToLevel = 42
ElseIf UserList(userindex).Stats.ELV < 25 Then
    ParticleToLevel = 2
ElseIf UserList(userindex).Stats.ELV < 35 Then
    ParticleToLevel = 81
ElseIf UserList(userindex).Stats.ELV < 50 Then
    If UserList(userindex).Faccion.Renegado = 1 Then
        ParticleToLevel = 39
    ElseIf UserList(userindex).Faccion.Ciudadano = 1 Then
        ParticleToLevel = 40
    ElseIf UserList(userindex).Faccion.Republicano = 1 Then
        ParticleToLevel = 71
    ElseIf UserList(userindex).Faccion.ArmadaReal = 1 Then
        ParticleToLevel = 38
    ElseIf UserList(userindex).Faccion.FuerzasCaos = 1 Then
        ParticleToLevel = 37
    ElseIf UserList(userindex).Faccion.Milicia = 1 Then
        ParticleToLevel = 66
    End If
Else
    ParticleToLevel = 36
End If
End Function

Public Sub ReproducirSonido(ByVal Destino As SendTarget, ByVal index As Integer, ByVal SoundIndex As Integer)
    Call SendData(Destino, index, PrepareMessagePlayWave(SoundIndex, UserList(index).pos.x, UserList(index).pos.Y))
End Sub
Function Tilde(ByRef F As String) As String
    Tilde = Replace$(Replace$(Replace$(Replace$(Replace$(F, "í", "i"), "è", "e"), "ó", "o"), "á", "a"), "ú", "u")
End Function
