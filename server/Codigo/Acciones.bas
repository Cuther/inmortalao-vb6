Attribute VB_Name = "Acciones"

Option Explicit

''
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal UserIndex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer)
    Dim tempIndex As Integer
    Dim DummyInt As Integer
    
On Error GoTo hayerror


    '¿Rango Visión? (ToxicWaste)
    If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.x - x) > RANGO_VISION_X) Then
        Exit Sub
    End If

    If UserList(UserIndex).flags.Trabajando = True Then
        UserList(UserIndex).flags.Trabajando = False
        
        Call WriteConsoleMsg(1, UserIndex, "Dejas de trabajar.", FontTypeNames.FONTTYPE_BROWNI)
    End If

    '¿Posicion valida?
    If InMapBounds(map, x, Y) Then
        With UserList(UserIndex)
            'Trabajo
            If .Invent.AnilloEqpSlot <> 0 Then
                Select Case .Invent.AnilloEqpObjIndex
                    Case RED_PESCA, CAÑA_PESCA
                        If MapData(.Pos.map, .Pos.x, .Pos.Y).Trigger = 1 Then
                            Call WriteConsoleMsg(2, UserIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        If HayAgua(map, x, Y) Then
                            Call WriteConsoleMsg(2, UserIndex, "Comienzas a trabajar...", FontTypeNames.FONTTYPE_BROWNI)
                        
                            .flags.Trabajando = True
                            Exit Sub
                        Else
                            Call WriteConsoleMsg(2, UserIndex, "No hay agua donde pescar. Busca un lago, rio o mar.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        
                    Case PIQUETE_MINERO
                        DummyInt = MapData(.Pos.map, x, Y).ObjInfo.ObjIndex
                
                        If DummyInt > 0 Then
                            'Check distance
                            If Abs(.Pos.x - x) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteConsoleMsg(2, UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            DummyInt = MapData(.Pos.map, x, Y).ObjInfo.ObjIndex 'CHECK
                            '¿Hay un yacimiento donde clickeo?
                            If ObjData(DummyInt).OBJType = eOBJType.otYacimiento Then
                                .flags.Trabajando = True
                                Call WriteConsoleMsg(2, UserIndex, "Comienzas a trabajar...", FontTypeNames.FONTTYPE_BROWNI)
                            Else
                                Call WriteConsoleMsg(2, UserIndex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
                            End If
                        Else
                            Call WriteConsoleMsg(2, UserIndex, "Ahí no hay ningun yacimiento.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        
                    Case HACHA_LEÑADOR
                        DummyInt = MapData(.Pos.map, x, Y).ObjInfo.ObjIndex
                        If DummyInt > 0 Then
                            If Abs(.Pos.x - x) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteConsoleMsg(2, UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            If MapInfo(.Pos.map).Pk = False Then
                                Call WriteConsoleMsg(2, UserIndex, "No puedes talar en zona segura.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            '¿Hay un arbol donde clickeo?
                            If ObjData(DummyInt).OBJType = eOBJType.otArboles Then
                                
                                .flags.Trabajando = True
                                Call WriteConsoleMsg(2, UserIndex, "Comienzas a trabajar...", FontTypeNames.FONTTYPE_BROWNI)
                            End If
                        End If
                        
                        
                    Case TIJERAS
                        DummyInt = MapData(.Pos.map, x, Y).ObjInfo.ObjIndex
                        If DummyInt > 0 Then
                            If Abs(.Pos.x - x) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteConsoleMsg(2, UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            If MapInfo(.Pos.map).Pk = False Then
                                Call WriteConsoleMsg(2, UserIndex, "No puedes juntar raices en zona segura.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                    
                            '¿Hay un arbol donde clickeo?
                            If ObjData(DummyInt).OBJType = eOBJType.otArboles Then
                                
                                .flags.Trabajando = True
                                Call WriteConsoleMsg(2, UserIndex, "Comienzas a trabajar...", FontTypeNames.FONTTYPE_BROWNI)
                            End If
                        End If
                        
                        
                    Case iMinerales.PlataCruda, iMinerales.HierroCrudo, iMinerales.OroCrudo
                        'Check there is a proper item there
                        If .flags.TargetObj > 0 Then
                            If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then
                                'Validate other items
                                If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > MAX_INVENTORY_SLOTS Then
                                    Exit Sub
                                End If
                                
                                ''chequeamos que no se zarpe duplicando oro
                                If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
                                    If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).Amount = 0 Then
                                        Call WriteConsoleMsg(2, UserIndex, "No tienes más minerales", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub
                                    End If
                                    
                                    ''FUISTE
                                    Call WriteErrorMsg(UserIndex, "Has sido expulsado por el sistema anti cheats.")
                                    Call FlushBuffer(UserIndex)
                                    Call CloseSocket(UserIndex)
                                    Exit Sub
                                End If
                                
                                .flags.Trabajando = True
                                Call WriteConsoleMsg(2, UserIndex, "Comienzas a trabajar...", FontTypeNames.FONTTYPE_BROWNI)
                            Else
                                Call WriteConsoleMsg(2, UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                            End If
                        Else
                            Call WriteConsoleMsg(2, UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                        End If
                
                End Select

                        
                If MapData(map, x, Y).ObjInfo.ObjIndex Then
                    If ObjData(MapData(map, x, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otYunque Then
                        If .Invent.AnilloEqpObjIndex = MARTILLO_HERRERO Then
                             Call EnivarArmadurasConstruibles(UserIndex)
                             Call EnivarArmasConstruibles(UserIndex)
                             Call WriteShowBlacksmithForm(UserIndex)
                         End If
                    End If
                End If
            End If 'fin  .Invent.AnilloEqpSlot <> 0 Then
                        
            If MapData(map, x, Y).ObjInfo.ObjIndex Then
                If ObjData(MapData(map, x, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otFragua Then
                    If .flags.Lingoteando <> 0 Then
                        .flags.Trabajando = True
                
                        Call WriteConsoleMsg(2, UserIndex, "Comienzas a trabajar...", FontTypeNames.FONTTYPE_BROWNI)
                    End If
                End If
            End If
        
            If MapData(map, x, Y).NpcIndex > 0 Then     'Acciones NPCs
                tempIndex = MapData(map, x, Y).NpcIndex
                
                'Set the target NPC
                .flags.TargetNPC = tempIndex
                
                If Npclist(tempIndex).Comercia = 1 Then
                    '¿Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteMsg(UserIndex, 1)
                        Exit Sub
                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(2, UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Iniciamos la rutina pa' comerciar.
                    Call IniciarComercioNPC(UserIndex)
                
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Banquero Then
                    '¿Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteMsg(UserIndex, 1)
                        Exit Sub
                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(2, UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    Call IniciarDeposito(UserIndex, True)
                
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Revividor Or Npclist(tempIndex).NPCtype = eNPCType.ResucitadorNewbie Then
                    If Distancia(.Pos, Npclist(tempIndex).Pos) > 10 Then
                        Call WriteConsoleMsg(2, UserIndex, "El sacerdote no puede curarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Revivimos si es necesario
                    If .flags.Muerto = 1 And .flags.Resucitando = 0 Then
                        Call SendData(SendTarget.ToNPCArea, tempIndex, PrepareMessageChatOverHead("AHIL KNÄ XÄR", Npclist(tempIndex).Char.CharIndex, RGB(128, 128, 0)))
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(240, .Pos.x, .Pos.Y))
                        Call RevivirUsuario(UserIndex)
                    End If
                    
                    If (.flags.Resucitando = 0 And .flags.Muerto = 0) Then
                        If .Stats.MinHP < .Stats.MaxHP Or .flags.Envenenado <> 0 Or .flags.Incinerado = 1 Then
                        
                            .Stats.MinHP = .Stats.MaxHP
                            .flags.Envenenado = 0
                            .flags.Incinerado = 0
                            .flags.Ceguera = 0
                            Call WriteUpdateUserStats(UserIndex)
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateCharParticle(.Char.CharIndex, 119))
                            Call SendData(SendTarget.ToNPCArea, tempIndex, PrepareMessageChatOverHead("Nihil Vitae", Npclist(tempIndex).Char.CharIndex, RGB(128, 128, 0)))
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(236, .Pos.x, .Pos.Y)) 'Sum Corp Sanctis
                        End If
                    End If
                 
                 ElseIf Npclist(tempIndex).NPCtype = 11 Then 'Veterinarias
                    If UserList(UserIndex).masc.TieneFamiliar = 1 Then
                        If UserList(UserIndex).masc.MinHP <= 0 Then
                            UserList(UserIndex).masc.MinHP = UserList(UserIndex).masc.MaxHP
                            
                            UpdateFamiliar UserIndex, True
                            
                            Call WriteChatOverHead(UserIndex, "¡He resucitado a tu mascota!", Npclist(tempIndex).Char.CharIndex, vbWhite)
                            Exit Sub
                        ElseIf UserList(UserIndex).masc.MinHP <> UserList(UserIndex).masc.MaxHP Then
                            UserList(UserIndex).masc.MinHP = UserList(UserIndex).masc.MaxHP
                            
                            UpdateFamiliar UserIndex, True
                            
                            Call WriteChatOverHead(UserIndex, "Cure las heridas de tu familiar ¡¡Suerte aventurero!!", Npclist(tempIndex).Char.CharIndex, vbWhite)
                            Exit Sub
                        End If
                    Else
                        'adoptar mascota
                        If UserList(UserIndex).Stats.UserSkills(eSkill.Domar) >= 65 Then
                            Call WriteShowFamiliarForm(UserIndex)
                            Exit Sub
                        End If
                    End If
                 
                 ElseIf Npclist(tempIndex).NPCtype = 18 Then 'Enlistador
                    If Distancia(.Pos, Npclist(tempIndex).Pos) > 3 Then
                        Call WriteConsoleMsg(2, UserIndex, "Estas lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    If UserList(UserIndex).GuildIndex > 0 Then
                        If modGuilds.GuildFounder(UserList(UserIndex).GuildIndex) = UserList(UserIndex).Name Then
                            Call WriteConsoleMsg(1, UserIndex, "¡¡¡Eres el fundador de un Clan. No puedes cambiar de faccion!!!", FontTypeNames.FONTTYPE_GUILD)
                            Exit Sub
                        End If
                    End If
                    
                    If Npclist(tempIndex).flags.faccion = 1 Then 'Ciuda
                        If UserList(UserIndex).faccion.ArmadaMatados > 0 Or UserList(UserIndex).faccion.CiudadanosMatados > 0 Then
                            
                            If UserList(UserIndex).GuildIndex > 0 Then
                                Call modGuilds.m_EcharMiembroDeClan(-1, UserList(UserIndex).Name)
                                Call WriteConsoleMsg(1, UserIndex, "Has sido expulsado del clan por tu nueva facción.", FontTypeNames.FONTTYPE_GUILD)
                            End If
                            
                            UserList(UserIndex).faccion.Renegado = 0
                            UserList(UserIndex).faccion.Ciudadano = 1
                            UserList(UserIndex).faccion.Republicano = 0

                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))
                            'Call WriteChatOverHead(UserIndex, "Has asesino ciudadanos del Imperio. Para regresar, una parte de tu alma fue necesaria. Tu experiencia para subir al siguiente nivel se aumentó en 1%", Npclist(tempIndex).Char.CharIndex, vbWhite)

                        ElseIf Not (esMili(UserIndex) Or esArmada(UserIndex) Or esCaos(UserIndex)) Then
                            
                            If UserList(UserIndex).GuildIndex > 0 Then
                                Call modGuilds.m_EcharMiembroDeClan(-1, UserList(UserIndex).Name)
                                Call WriteConsoleMsg(1, UserIndex, "Has sido expulsado del clan por tu nueva facción.", FontTypeNames.FONTTYPE_GUILD)
                            End If
                            
                            UserList(tempIndex).faccion.Rango = 0
                            UserList(UserIndex).faccion.Renegado = 0
                            UserList(UserIndex).faccion.Ciudadano = 1
                            UserList(UserIndex).faccion.Republicano = 0
                        
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))
                        End If
                    
                    ElseIf Npclist(tempIndex).flags.faccion = 2 Then 'Repu
                        If UserList(UserIndex).faccion.MilicianosMatados > 0 Or UserList(UserIndex).faccion.RepublicanosMatados > 0 Then
                            
                            If UserList(UserIndex).GuildIndex > 0 Then
                                Call modGuilds.m_EcharMiembroDeClan(-1, UserList(UserIndex).Name)
                                Call WriteConsoleMsg(1, UserIndex, "Has sido expulsado del clan por tu nueva facción.", FontTypeNames.FONTTYPE_GUILD)
                            End If
                            
                            UserList(UserIndex).faccion.Renegado = 0
                            UserList(UserIndex).faccion.Ciudadano = 0
                            UserList(UserIndex).faccion.Republicano = 1
                            
                            UserList(UserIndex).faccion.MilicianosMatados = 0
                            UserList(UserIndex).faccion.RepublicanosMatados = 0

                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))
                            'Call WriteChatOverHead(UserIndex, "Has asesino ciudadanos de la Republica. Para regresar, una parte de tu alma fue necesaria. Tu experiencia para subir al siguiente nivel se aumentó en 1%", Npclist(tempIndex).Char.CharIndex, vbWhite)
                       
                       ElseIf Not (esMili(UserIndex) Or esArmada(UserIndex) Or esCaos(UserIndex)) Then
                            
                            If UserList(UserIndex).GuildIndex > 0 Then
                                Call modGuilds.m_EcharMiembroDeClan(-1, UserList(UserIndex).Name)
                                Call WriteConsoleMsg(1, UserIndex, "Has sido expulsado del clan por tu nueva facción.", FontTypeNames.FONTTYPE_GUILD)
                            End If
                            
                            UserList(UserIndex).faccion.Rango = 0
                            UserList(UserIndex).faccion.Renegado = 0
                            UserList(UserIndex).faccion.Ciudadano = 0
                            UserList(UserIndex).faccion.Republicano = 1
                            
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))
                        End If
                    End If
                ElseIf Npclist(tempIndex).NPCtype = 16 Then
                    Call WriteSubastRequest(UserIndex)
                End If
                
            '¿Es un obj?
            ElseIf MapData(map, x, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map, x, Y).ObjInfo.ObjIndex
                
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, x, Y, UserIndex)
                    Case eOBJType.otForos 'Foro
                        Call AccionParaCorreo(map, x, Y, UserIndex)
                    Case eOBJType.otCorreo 'Correo
                        Call AccionParaCorreo(map, x, Y, UserIndex)
                    Case eOBJType.otLeña    'Leña
                        If tempIndex = FOGATA_APAG And .flags.Muerto = 0 Then
                            Call AccionParaRamita(map, x, Y, UserIndex)
                        End If
                End Select
            '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
            ElseIf MapData(map, x + 1, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map, x + 1, Y).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, x + 1, Y, UserIndex)
                    
                End Select
            
            ElseIf MapData(map, x + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map, x + 1, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
        
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, x + 1, Y + 1, UserIndex)
                End Select
            
            ElseIf MapData(map, x, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map, x, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, x, Y + 1, UserIndex)
                End Select
            End If
        End With
    End If
    
    
    Exit Sub
    
hayerror:
    LogError ("Error en Accion: " & err.Number & " Desc: " & err.description & " Name: " & UserList(UserIndex).Name & " Pos: " & map & " " & x & " " & Y)
End Sub

Sub AccionParaForo(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim Pos As WorldPos
Pos.map = map
Pos.x = x
Pos.Y = Y

If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
    Call WriteConsoleMsg(1, UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

'¿Hay mensajes?
Dim F As String, tit As String, men As String, Base As String, auxcad As String
F = App.Path & "\foros\" & UCase$(ObjData(MapData(map, x, Y).ObjInfo.ObjIndex).ForoID) & ".for"
If FileExist(F, vbNormal) Then
    Dim num As Integer
    num = val(GetVar(F, "INFO", "CantMSG"))
    Base = Left$(F, Len(F) - 4)
    Dim i As Integer
    Dim N As Integer
    For i = 1 To num
        N = FreeFile
        F = Base & i & ".for"
        Open F For Input Shared As #N
        Input #N, tit
        men = vbNullString
        auxcad = vbNullString
        Do While Not EOF(N)
            Input #N, auxcad
            men = men & vbCrLf & auxcad
        Loop
        Close #N
        Call WriteAddForumMsg(UserIndex, tit, men)
        
    Next
End If
Call WriteShowForumForm(UserIndex)
End Sub


Sub AccionParaCorreo(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim Pos As WorldPos
    Pos.map = map
    Pos.x = x
    Pos.Y = Y
    
    If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
        Call WriteConsoleMsg(1, UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    '¿Hay mensajes?
    
    Dim loopC As Long
    
    If UserList(UserIndex).cant_mensajes > MENSAJES_TOPE_CORREO Then UserList(UserIndex).cant_mensajes = 20
    
    For loopC = 1 To UserList(UserIndex).cant_mensajes
       ' If UserList(UserIndex).Correos(LoopC).De <> "" Then _ ' No hay necesidad ya que se sabe con el recordcount
            Call WriteAddCorreoMsg(UserIndex, loopC)
    Next loopC
    
    Call WriteShowCorreoForm(UserIndex)
    
    UserList(UserIndex).cVer = 0
    Call WriteMensajeSigno(UserIndex)
    
    FlushBuffer UserIndex
    DoEvents
End Sub


Sub AccionParaPuerta(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

If Not (Distance(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, x, Y) > 2) Then
    If ObjData(MapData(map, x, Y).ObjInfo.ObjIndex).Llave = 0 Then
        If ObjData(MapData(map, x, Y).ObjInfo.ObjIndex).Cerrada = 1 Then
                'Abre la puerta
                If ObjData(MapData(map, x, Y).ObjInfo.ObjIndex).Llave = 0 Then
                    
                    MapData(map, x, Y).ObjInfo.ObjIndex = ObjData(MapData(map, x, Y).ObjInfo.ObjIndex).IndexAbierta
                    
                    Call modSendData.SendToAreaByPos(map, x, Y, PrepareMessageObjectCreate(x, Y, MapData(map, x, Y).ObjInfo.ObjIndex, 0))
                    
                    'Desbloquea
                    MapData(map, x, Y).Blocked = 0
                    MapData(map, x - 1, Y).Blocked = 0
                    
                    'Bloquea todos los mapas
                    Call Bloquear(True, map, x, Y, 0)
                    Call Bloquear(True, map, x - 1, Y, 0)
                    
                      
                    'Sonido
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, x, Y))
                    
                Else
                     Call WriteConsoleMsg(1, UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
                End If
        Else
            'Add Marius Las puertas abren pero no cierran, a no ser que seas gm ;)
            '¿Se te ocurre una solucion mejor al bug? Ilustrame.
            If EsCONSE(UserIndex) Then
                'Cierra puerta
                MapData(map, x, Y).ObjInfo.ObjIndex = ObjData(MapData(map, x, Y).ObjInfo.ObjIndex).IndexCerrada
                
                Call modSendData.SendToAreaByPos(map, x, Y, PrepareMessageObjectCreate(x, Y, MapData(map, x, Y).ObjInfo.ObjIndex, 0))
                                
                MapData(map, x, Y).Blocked = 1
                MapData(map, x - 1, Y).Blocked = 1
                
                
                Call Bloquear(True, map, x - 1, Y, 1)
                Call Bloquear(True, map, x, Y, 1)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, x, Y))
            End If
            
        End If
        
        UserList(UserIndex).flags.TargetObj = MapData(map, x, Y).ObjInfo.ObjIndex
    Else
        Call WriteConsoleMsg(1, UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(1, UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub



Sub AccionParaRamita(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim Suerte As Byte
Dim exito As Byte
Dim obj As obj

Dim Pos As WorldPos
    Pos.map = map
    Pos.x = x
    Pos.Y = Y
    
    If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
        Call WriteConsoleMsg(1, UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If MapData(map, x, Y).Trigger = eTrigger.ZONASEGURA Or MapInfo(map).Pk = False Then
        Call WriteConsoleMsg(1, UserIndex, "En zona segura no puedes hacer fogatas.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).Stats.UserSkills(Supervivencia) > 1 And UserList(UserIndex).Stats.UserSkills(Supervivencia) < 6 Then
        Suerte = 3
    ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 10 Then
        Suerte = 2
    ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 10 And UserList(UserIndex).Stats.UserSkills(Supervivencia) Then
        Suerte = 1
    End If
    
    exito = RandomNumber(1, Suerte)
    
    If exito = 1 Then
        If MapInfo(UserList(UserIndex).Pos.map).Zona <> Ciudad Then
            obj.ObjIndex = FOGATA
            obj.Amount = 1
            
            Call WriteConsoleMsg(1, UserIndex, "Has prendido la fogata.", FontTypeNames.FONTTYPE_INFO)
            
            Call MakeObj(obj, map, x, Y)
            
            'Las fogatas prendidas se deben eliminar
            Dim Fogatita As New cGarbage
            Fogatita.map = map
            Fogatita.x = x
            Fogatita.Y = Y
            Call TrashCollector.Add(Fogatita)
        Else
            Call WriteConsoleMsg(1, UserIndex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        Call WriteConsoleMsg(1, UserIndex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    Call SubirSkill(UserIndex, Supervivencia)

End Sub
