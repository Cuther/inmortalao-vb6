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

Sub Accion(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer)
    Dim tempIndex As Integer
    Dim DummyInt As Integer
    
On Error GoTo hayerror


    '¿Rango Visión? (ToxicWaste)
    If (Abs(UserList(userindex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(userindex).Pos.x - x) > RANGO_VISION_X) Then
        Exit Sub
    End If

    If UserList(userindex).flags.Trabajando = True Then
        UserList(userindex).flags.Trabajando = False
        
        Call WriteConsoleMsg(1, userindex, "Dejas de trabajar.", FontTypeNames.FONTTYPE_BROWNI)
    End If

    '¿Posicion valida?
    If InMapBounds(map, x, Y) Then
        With UserList(userindex)
            'Trabajo
            If .Invent.AnilloEqpSlot <> 0 Then
                Select Case .Invent.AnilloEqpObjIndex
                    Case RED_PESCA, CAÑA_PESCA
                        If MapData(.Pos.map, .Pos.x, .Pos.Y).Trigger = 1 Then
                            Call WriteConsoleMsg(2, userindex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        If HayAgua(map, x, Y) Then
                            Call WriteConsoleMsg(2, userindex, "Comienzas a trabajar...", FontTypeNames.FONTTYPE_BROWNI)
                        
                            .flags.Trabajando = True
                            Exit Sub
                        Else
                            Call WriteConsoleMsg(2, userindex, "No hay agua donde pescar. Busca un lago, rio o mar.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        
                    Case PIQUETE_MINERO
                        DummyInt = MapData(.Pos.map, x, Y).ObjInfo.ObjIndex
                
                        If DummyInt > 0 Then
                            'Check distance
                            If Abs(.Pos.x - x) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteConsoleMsg(2, userindex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            DummyInt = MapData(.Pos.map, x, Y).ObjInfo.ObjIndex 'CHECK
                            '¿Hay un yacimiento donde clickeo?
                            If ObjData(DummyInt).OBJType = eOBJType.otYacimiento Then
                                .flags.Trabajando = True
                                Call WriteConsoleMsg(2, userindex, "Comienzas a trabajar...", FontTypeNames.FONTTYPE_BROWNI)
                            Else
                                Call WriteConsoleMsg(2, userindex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
                            End If
                        Else
                            Call WriteConsoleMsg(2, userindex, "Ahí no hay ningun yacimiento.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    Case HACHA_LEÑADOR
                        DummyInt = MapData(.Pos.map, x, Y).ObjInfo.ObjIndex
                        If DummyInt > 0 Then
                            If Abs(.Pos.x - x) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteConsoleMsg(2, userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            If MapInfo(.Pos.map).Pk = False Then
                                Call WriteConsoleMsg(2, userindex, "No puedes talar en zona segura.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            '¿Hay un arbol donde clickeo?
                            If ObjData(DummyInt).OBJType = eOBJType.otArboles Then
                                
                                .flags.Trabajando = True
                                Call WriteConsoleMsg(2, userindex, "Comienzas a trabajar...", FontTypeNames.FONTTYPE_BROWNI)
                            End If
                        End If
                        
                        
                Case TIJERAS
                        DummyInt = MapData(.Pos.map, x, Y).ObjInfo.ObjIndex
                        If DummyInt > 0 Then
                            If Abs(.Pos.x - x) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteConsoleMsg(2, userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                    
                            '¿Hay un arbol donde clickeo?
                            If ObjData(DummyInt).OBJType = eOBJType.otArboles Then
                                
                                .flags.Trabajando = True
                                Call WriteConsoleMsg(2, userindex, "Comienzas a trabajar...", FontTypeNames.FONTTYPE_BROWNI)
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
                                        Call WriteConsoleMsg(2, userindex, "No tienes más minerales", FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub
                                    End If
                                    
                                    ''FUISTE
                                    Call WriteErrorMsg(userindex, "Has sido expulsado por el sistema anti cheats.")
                                    Call FlushBuffer(userindex)
                                    Call CloseSocket(userindex)
                                    Exit Sub
                                End If
                                
                                .flags.Trabajando = True
                                Call WriteConsoleMsg(2, userindex, "Comienzas a trabajar...", FontTypeNames.FONTTYPE_BROWNI)
                            Else
                                Call WriteConsoleMsg(2, userindex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                            End If
                        Else
                            Call WriteConsoleMsg(2, userindex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                        End If
                
                End Select

                        
                                    If MapData(map, x, Y).ObjInfo.ObjIndex Then
                If ObjData(MapData(map, x, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otYunque Then
               If .Invent.AnilloEqpObjIndex = MARTILLO_HERRERO Then
                    Call EnivarArmadurasConstruibles(userindex)
                    Call EnivarArmasConstruibles(userindex)
                    Call WriteShowBlacksmithForm(userindex)
                End If
                End If
                        End If
                        End If
                        
            If MapData(map, x, Y).ObjInfo.ObjIndex Then
                If ObjData(MapData(map, x, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otFragua Then
                    If .flags.Lingoteando <> 0 Then
                        .flags.Trabajando = True
                
                        Call WriteConsoleMsg(2, userindex, "Comienzas a trabajar...", FontTypeNames.FONTTYPE_BROWNI)
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
                        Call WriteMsg(userindex, 1)
                        Exit Sub
                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(2, userindex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Iniciamos la rutina pa' comerciar.
                    Call IniciarComercioNPC(userindex)
                
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Banquero Then
                    '¿Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteMsg(userindex, 1)
                        Exit Sub
                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(2, userindex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    Call IniciarDeposito(userindex, True)
                
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Revividor Or Npclist(tempIndex).NPCtype = eNPCType.ResucitadorNewbie Then
                    If Distancia(.Pos, Npclist(tempIndex).Pos) > 10 Then
                        Call WriteConsoleMsg(2, userindex, "El sacerdote no puede curarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Revivimos si es necesario
                    If .flags.Muerto = 1 And .flags.Resucitando = 0 Then
                        Call SendData(SendTarget.ToNPCArea, tempIndex, PrepareMessageChatOverHead("AHIL KNÄ XÄR", Npclist(tempIndex).Char.CharIndex, RGB(128, 128, 0)))
                        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(240, .Pos.x, .Pos.Y))
                        Call RevivirUsuario(userindex)
                    End If
                    
                    If (.flags.Resucitando = 0 And .flags.Muerto = 0) Then
                        If .Stats.MinHP < .Stats.MaxHP Or .flags.Envenenado <> 0 Or .flags.Incinerado = 1 Then
                        
                            .Stats.MinHP = .Stats.MaxHP
                            .flags.Envenenado = 0
                            .flags.Incinerado = 0
                            .flags.Ceguera = 0
                            Call WriteUpdateUserStats(userindex)
                            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateCharParticle(.Char.CharIndex, 119))
                            Call SendData(SendTarget.ToNPCArea, tempIndex, PrepareMessageChatOverHead("Nihil Vitae", Npclist(tempIndex).Char.CharIndex, RGB(128, 128, 0)))
                            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(236, .Pos.x, .Pos.Y)) 'Sum Corp Sanctis
                        End If
                    End If
                 
                 ElseIf Npclist(tempIndex).NPCtype = 11 Then 'Veterinarias
                    If UserList(userindex).masc.TieneFamiliar = 1 Then
                        If UserList(userindex).masc.MinHP <= 0 Then
                            UserList(userindex).masc.MinHP = UserList(userindex).masc.MaxHP
                            
                            UpdateFamiliar userindex, True
                            
                            Call WriteChatOverHead(userindex, "¡He resucitado a tu mascota!", Npclist(tempIndex).Char.CharIndex, vbWhite)
                            Exit Sub
                        ElseIf UserList(userindex).masc.MinHP <> UserList(userindex).masc.MaxHP Then
                            UserList(userindex).masc.MinHP = UserList(userindex).masc.MaxHP
                            
                            UpdateFamiliar userindex, True
                            
                            Call WriteChatOverHead(userindex, "Cure las heridas de tu familiar ¡¡Suerte aventurero!!", Npclist(tempIndex).Char.CharIndex, vbWhite)
                            Exit Sub
                        End If
                    Else
                        'adoptar mascota
                        If UserList(userindex).Stats.UserSkills(eSkill.Domar) >= 65 Then
                            Call WriteShowFamiliarForm(userindex)
                            Exit Sub
                        End If
                    End If
                 
                 ElseIf Npclist(tempIndex).NPCtype = 18 Then
                    If Distancia(.Pos, Npclist(tempIndex).Pos) > 3 Then
                        Call WriteConsoleMsg(2, userindex, "Estas lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    If Npclist(tempIndex).flags.Faccion = 1 Then 'Ciuda
                        If UserList(tempIndex).Faccion.ArmadaMatados > 0 Or UserList(tempIndex).Faccion.CiudadanosMatados > 0 Then
                           
                            Call WriteChatOverHead(userindex, "Has asesino ciudadanos del Imperio. Si realmente quieres regresar, una parte de tu alma sera necesaria. Tu experiencia para subir al siguiente nivel sera aumentada en 0.7%, escribiendo /PERDON", Npclist(tempIndex).Char.CharIndex, vbWhite)
                            Exit Sub
                        End If
                        
                        If Not (esMili(userindex) Or esArmada(userindex) Or esCaos(userindex)) Then
                            
                            UserList(tempIndex).Faccion.Rango = 0
                            UserList(userindex).Faccion.Renegado = 0
                            UserList(userindex).Faccion.Ciudadano = 1
                            UserList(userindex).Faccion.Republicano = 0
                        
                            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCharStatus(UserList(userindex).Char.CharIndex, UserTypeColor(userindex)))
                        End If
                    
                    ElseIf Npclist(tempIndex).flags.Faccion = 2 Then
                        If UserList(userindex).Faccion.MilicianosMatados > 0 Or UserList(userindex).Faccion.RepublicanosMatados > 0 Then
                            Call WriteChatOverHead(userindex, "Has asesino ciudadanos de la Republica. Si realmente quieres regresar, una parte de tu alma sera necesaria. Tu experiencia para subir al siguiente nivel sera aumentada en 0.7%, escribiendo /PERDON", Npclist(tempIndex).Char.CharIndex, vbWhite)
                            Exit Sub
                        End If
                        
                        If Not (esMili(userindex) Or esArmada(userindex) Or esCaos(userindex)) Then
                            UserList(userindex).Faccion.Rango = 0
                            UserList(userindex).Faccion.Renegado = 0
                            UserList(userindex).Faccion.Ciudadano = 0
                            UserList(userindex).Faccion.Republicano = 1
                            
                            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCharStatus(UserList(userindex).Char.CharIndex, UserTypeColor(userindex)))
                        End If
                    End If
                ElseIf Npclist(tempIndex).NPCtype = 16 Then
                    Call WriteSubastRequest(userindex)
                End If
                
            '¿Es un obj?
            ElseIf MapData(map, x, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map, x, Y).ObjInfo.ObjIndex
                
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, x, Y, userindex)
                    Case eOBJType.otForos 'Foro
                        Call AccionParaCorreo(map, x, Y, userindex)
                    Case eOBJType.otCorreo 'Correo
                        Call AccionParaCorreo(map, x, Y, userindex)
                    Case eOBJType.otLeña    'Leña
                        If tempIndex = FOGATA_APAG And .flags.Muerto = 0 Then
                            Call AccionParaRamita(map, x, Y, userindex)
                        End If
                End Select
            '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
            ElseIf MapData(map, x + 1, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map, x + 1, Y).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, x + 1, Y, userindex)
                    
                End Select
            
            ElseIf MapData(map, x + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map, x + 1, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
        
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, x + 1, Y + 1, userindex)
                End Select
            
            ElseIf MapData(map, x, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map, x, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, x, Y + 1, userindex)
                End Select
            End If
        End With
    End If
    
    
    Exit Sub
    
hayerror:
    LogError ("Error en Accion: " & err.Number & " Desc: " & err.description)
    
    
End Sub

Sub AccionParaForo(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal userindex As Integer)
On Error Resume Next

Dim Pos As WorldPos
Pos.map = map
Pos.x = x
Pos.Y = Y

If Distancia(Pos, UserList(userindex).Pos) > 2 Then
    Call WriteConsoleMsg(1, userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
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
        Call WriteAddForumMsg(userindex, tit, men)
        
    Next
End If
Call WriteShowForumForm(userindex)
End Sub


Sub AccionParaCorreo(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal userindex As Integer)
On Error Resume Next

Dim Pos As WorldPos
Pos.map = map
Pos.x = x
Pos.Y = Y

If Distancia(Pos, UserList(userindex).Pos) > 2 Then
    Call WriteConsoleMsg(1, userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

'¿Hay mensajes?

Dim LoopC As Long

If UserList(userindex).cant_mensajes > MENSAJES_TOPE_CORREO Then UserList(userindex).cant_mensajes = 20

For LoopC = 1 To UserList(userindex).cant_mensajes
   ' If UserList(UserIndex).Correos(LoopC).De <> "" Then _ ' No hay necesidad ya que se sabe con el recordcount
        Call WriteAddCorreoMsg(userindex, LoopC)
Next LoopC

Call WriteShowCorreoForm(userindex)

UserList(userindex).cVer = 0
Call WriteMensajeSigno(userindex)

FlushBuffer userindex
DoEvents
End Sub


Sub AccionParaPuerta(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal userindex As Integer)
On Error Resume Next

If Not (Distance(UserList(userindex).Pos.x, UserList(userindex).Pos.Y, x, Y) > 2) Then
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
                    Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_PUERTA, x, Y))
                    
                Else
                     Call WriteConsoleMsg(1, userindex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
                End If
        Else
                'Cierra puerta
                MapData(map, x, Y).ObjInfo.ObjIndex = ObjData(MapData(map, x, Y).ObjInfo.ObjIndex).IndexCerrada
                
                Call modSendData.SendToAreaByPos(map, x, Y, PrepareMessageObjectCreate(x, Y, MapData(map, x, Y).ObjInfo.ObjIndex, 0))
                                
                MapData(map, x, Y).Blocked = 1
                MapData(map, x - 1, Y).Blocked = 1
                
                
                Call Bloquear(True, map, x - 1, Y, 1)
                Call Bloquear(True, map, x, Y, 1)
                
                Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_PUERTA, x, Y))
        End If
        
        UserList(userindex).flags.TargetObj = MapData(map, x, Y).ObjInfo.ObjIndex
    Else
        Call WriteConsoleMsg(1, userindex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(1, userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub



Sub AccionParaRamita(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal userindex As Integer)
On Error Resume Next

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj

Dim Pos As WorldPos
Pos.map = map
Pos.x = x
Pos.Y = Y

If Distancia(Pos, UserList(userindex).Pos) > 2 Then
    Call WriteConsoleMsg(1, userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If MapData(map, x, Y).Trigger = eTrigger.ZONASEGURA Or MapInfo(map).Pk = False Then
    Call WriteConsoleMsg(1, userindex, "En zona segura no puedes hacer fogatas.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(userindex).Stats.UserSkills(Supervivencia) > 1 And UserList(userindex).Stats.UserSkills(Supervivencia) < 6 Then
            Suerte = 3
ElseIf UserList(userindex).Stats.UserSkills(Supervivencia) >= 6 And UserList(userindex).Stats.UserSkills(Supervivencia) <= 10 Then
            Suerte = 2
ElseIf UserList(userindex).Stats.UserSkills(Supervivencia) >= 10 And UserList(userindex).Stats.UserSkills(Supervivencia) Then
            Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    If MapInfo(UserList(userindex).Pos.map).Zona <> Ciudad Then
        Obj.ObjIndex = FOGATA
        Obj.Amount = 1
        
        Call WriteConsoleMsg(1, userindex, "Has prendido la fogata.", FontTypeNames.FONTTYPE_INFO)
        
        Call MakeObj(Obj, map, x, Y)
        
        'Las fogatas prendidas se deben eliminar
        Dim Fogatita As New cGarbage
        Fogatita.map = map
        Fogatita.x = x
        Fogatita.Y = Y
        Call TrashCollector.Add(Fogatita)
    Else
        Call WriteConsoleMsg(1, userindex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
Else
    Call WriteConsoleMsg(1, userindex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)
End If

Call SubirSkill(userindex, Supervivencia)

End Sub
