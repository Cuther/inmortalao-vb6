Attribute VB_Name = "modHechizos"

Option Explicit

Public Const HELEMENTAL_FUEGO As Integer = 26
Public Const HELEMENTAL_TIERRA As Integer = 28
Public Const ANILLO_ESPECTRAL As Integer = 1329
Public Const ANILLO_PENUMBRAS As Integer = 1330

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 13/02/2009
'13/02/2009: ZaMa - Los npcs que tiren magias, no podran hacerlo en mapas donde no se permita usarla.
'***************************************************
On Error GoTo hayerror

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then Exit Sub

'Add Marius Solucion chota, fixeamos el bug que los npc te atacan en mapas que ni siquiera estas, el tema de los clones
'Si el npc y el pj no estan en el mismo mapa, no hace nada ;)
If Npclist(NpcIndex).Pos.map <> UserList(UserIndex).Pos.map Then Exit Sub
'\Add

' Si no se peude usar magia en el mapa, no le deja hacerlo.
If MapInfo(UserList(UserIndex).Pos.map).MagiaSinEfecto > 0 Then Exit Sub

'Mannakia
If UserList(UserIndex).Invent.MagicIndex > 0 Then
    If ObjData(UserList(UserIndex).Invent.MagicIndex).EfectoMagico = eMagicType.MagicasNoAtacan Then
        Exit Sub
    End If
End If
'Mannakia

Npclist(NpcIndex).CanAttack = 0
Dim Daño As Integer

If Hechizos(Spell).SubeHP = 1 Then

    Daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
    If Hechizos(Spell).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
    If Hechizos(Spell).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateCharParticle(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).Particle))
    
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + Daño
    If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    
    Call WriteConsoleMsg(2, UserIndex, Npclist(NpcIndex).Name & " te ha quitado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    Call WriteUpdateUserStats(UserIndex)

    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Hechizos(Spell).PalabrasMagicas, Npclist(NpcIndex).Char.CharIndex, RGB(128, 128, 0)))
    
ElseIf Hechizos(Spell).SubeHP = 2 Then
    
    If UserList(UserIndex).flags.Privilegios And (PlayerType.User Or PlayerType.VIP) Then
    
        Daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
            Daño = Daño - RandomNumber(ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMax)
            ' Jose Castelli / Resistencia Magica (RM)
            Daño = Daño - ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).ResistenciaMagica
            ' Jose Castelli / Resistencia Magica (RM)
        End If
        
        ' Jose Castelli / Resistencia Magica (RM)
        
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
            Daño = Daño - ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ResistenciaMagica
        End If
        
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            Daño = Daño - ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).ResistenciaMagica
        End If
        
        If UserList(UserIndex).Invent.MonturaObjIndex > 0 Then
            Daño = Daño - ObjData(UserList(UserIndex).Invent.MonturaObjIndex).ResistenciaMagica
        End If
        
        ' Jose Castelli / Resistencia Magica (RM)
        
        If Daño < 0 Then Daño = 0
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
        If Hechizos(Spell).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
        If Hechizos(Spell).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateCharParticle(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).Particle))

        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - Daño
        
        Call WriteConsoleMsg(2, UserIndex, Npclist(NpcIndex).Name & " te ha quitado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteUpdateUserStats(UserIndex)
        
        Call SubirSkill(UserIndex, eSkill.Resistencia)
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Hechizos(Spell).PalabrasMagicas, Npclist(NpcIndex).Char.CharIndex, RGB(128, 128, 0)))
        
        'Muere
        If UserList(UserIndex).Stats.MinHP < 1 Then
            UserList(UserIndex).Stats.MinHP = 0
            Call UserDie(UserIndex)
            '[Barrin 1-12-03]
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call ContarMuerte(UserIndex, Npclist(NpcIndex).MaestroUser)
                Call ActStats(UserIndex, Npclist(NpcIndex).MaestroUser)
            End If
            '[/Barrin]
        End If
    
    End If
    
End If

If Hechizos(Spell).Paraliza = 1 Or Hechizos(Spell).Inmoviliza = 1 Then
    If UserList(UserIndex).flags.Paralizado = 0 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
        If Hechizos(Spell).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
        If Hechizos(Spell).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateCharParticle(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).Particle))

        If Hechizos(Spell).Inmoviliza = 1 Then
            UserList(UserIndex).flags.Inmovilizado = 1
        End If
          
        UserList(UserIndex).flags.Paralizado = 1
        UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
          
        Call WriteParalizeOK(UserIndex)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Hechizos(Spell).PalabrasMagicas, Npclist(NpcIndex).Char.CharIndex, RGB(128, 128, 0)))
        
    End If
End If

If Hechizos(Spell).Estupidez = 1 Then   ' turbacion
     If UserList(UserIndex).flags.Estupidez = 0 Then
          Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
          If Hechizos(Spell).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
          If Hechizos(Spell).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateCharParticle(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).Particle))
          
          UserList(UserIndex).flags.Estupidez = 1
          UserList(UserIndex).Counters.Ceguera = IntervaloInvisible
                  
        Call WriteDumb(UserIndex)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Hechizos(Spell).PalabrasMagicas, Npclist(NpcIndex).Char.CharIndex, RGB(128, 128, 0)))
     End If
End If


Exit Sub

hayerror:
    LogError ("Error en Npclanzaspellsobreuser: " & err.description)




End Sub


Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
'solo hechizos ofensivos!

    If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
    Npclist(NpcIndex).CanAttack = 0
    
    Dim Daño As Integer
    
    If Hechizos(Spell).SubeHP = 2 Then
        
            Daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
            Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).WAV, Npclist(TargetNPC).Pos.x, Npclist(TargetNPC).Pos.Y))
            If Hechizos(Spell).FXgrh <> 0 Then Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
            If Hechizos(Spell).Particle <> 0 Then Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateCharParticle(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).Particle))
            
            Npclist(TargetNPC).Stats.MinHP = Npclist(TargetNPC).Stats.MinHP - Daño
            
            'Muere
            If Npclist(TargetNPC).Stats.MinHP < 1 Then
                Npclist(TargetNPC).Stats.MinHP = 0
                If Npclist(NpcIndex).MaestroUser > 0 Then
                    Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
                Else
                    Call MuereNpc(TargetNPC, 0)
                End If
            End If
    End If
    
    If Hechizos(Spell).Inmoviliza = 1 Then
        If Npclist(TargetNPC).flags.AfectaParalisis = 0 And Npclist(TargetNPC).flags.Inmovilizado = 0 Then
            Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).WAV, Npclist(TargetNPC).Pos.x, Npclist(TargetNPC).Pos.Y))
            If Hechizos(Spell).FXgrh <> 0 Then Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
            If Hechizos(Spell).Particle <> 0 Then Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateCharParticle(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).Particle))
            Npclist(TargetNPC).flags.Inmovilizado = 1
            Npclist(TargetNPC).flags.Paralizado = 0
            Npclist(TargetNPC).Contadores.Paralisis = IntervaloParalizado
        End If
    End If
    
    If Hechizos(Spell).Paraliza = 1 Then
        If Npclist(TargetNPC).flags.AfectaParalisis = 0 And Npclist(TargetNPC).flags.Paralizado = 0 Then
            Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).WAV, Npclist(TargetNPC).Pos.x, Npclist(TargetNPC).Pos.Y))
            If Hechizos(Spell).FXgrh <> 0 Then Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
            If Hechizos(Spell).Particle <> 0 Then Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateCharParticle(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).Particle))
            Npclist(TargetNPC).flags.Paralizado = 1
            Npclist(TargetNPC).flags.Inmovilizado = 0
            Npclist(TargetNPC).Contadores.Paralisis = IntervaloParalizado
        End If
    End If
End Sub



Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean

On Error GoTo Errhandler
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
Errhandler:

End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal Slot As Integer)
Dim hIndex As Integer
Dim j As Integer
    hIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).HechizoIndex
    
    
    If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).CPO <> "" Then
        If UCase$(ListaClases(UserList(UserIndex).Clase)) <> ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).CPO Then
            Call WriteConsoleMsg(1, UserIndex, "No puedes comprender el hechizo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    If Not TieneHechizo(hIndex, UserIndex) Then
        'Buscamos un slot vacio
        For j = 1 To MAXUSERHECHIZOS
            If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
        Next j
            
        If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
            Call WriteConsoleMsg(1, UserIndex, "No tenes espacio para mas hechizos.", FontTypeNames.FONTTYPE_INFO)
        Else
            UserList(UserIndex).Stats.UserHechizos(j) = hIndex
            Call UpdateUserHechizos(False, UserIndex, CByte(j))
            'Quitamos del inv el item
            Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)
        End If
    Else
        Call WriteConsoleMsg(1, UserIndex, "Ya tenes ese hechizo.", FontTypeNames.FONTTYPE_INFO)
    End If

End Sub
            
Sub DecirPalabrasMagicas(ByVal S As String, ByVal UserIndex As Integer)
On Error GoTo hayerror
    'Mannakia
    If UserList(UserIndex).Invent.MagicIndex <> 0 Then
        If ObjData(UserList(UserIndex).Invent.MagicIndex).EfectoMagico = eMagicType.Silencio Then
            Exit Sub
        End If
    End If
    'Mannakia
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(S, UserList(UserIndex).Char.CharIndex, RGB(128, 128, 0)))
    Exit Sub


hayerror:
    LogError ("Error en DecirPalabrasMagicas: " & err.description)



End Sub

''
' Check if an user can cast a certain spell
'
' @param UserIndex Specifies reference to user
' @param HechizoIndex Specifies reference to spell
' @return   True if the user can cast the spell, otherwise returns false
Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 11/09/08
'Last Modification By: Marco Vanotti (Marco)
' - 11/09/08 Now Druid have mana bonus while casting summoning spells having a magic flute equipped (Marco)
'***************************************************
Dim DruidManaBonus As Single

    If HechizoIndex = 0 Then Exit Function
    
    If UserList(UserIndex).flags.Muerto Then
        Call WriteMsg(UserIndex, 6)
        PuedeLanzar = False
        Exit Function
    End If
    
    'Add Marius Pusimos esto para verificar que solo ciertas clases tiren ciertos hechizos
    If Not (InStr(Hechizos(HechizoIndex).ExclusivoClase, UCase$(ListaClases(UserList(UserIndex).Clase))) <> 0 Or _
    Len(Hechizos(HechizoIndex).ExclusivoClase) = 0) Then
        Call WriteConsoleMsg(2, UserIndex, "Tú clase no puede lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
        PuedeLanzar = False
        Exit Function
    End If
    '\Add
    
    If UserList(UserIndex).Stats.UserSkills(eSkill.Magia) < Hechizos(HechizoIndex).MinSkill Then
        Call WriteConsoleMsg(2, UserIndex, "No tenes suficientes puntos de magia para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
        PuedeLanzar = False
        Exit Function
    End If
    
    If UserList(UserIndex).Stats.MinSTA < Hechizos(HechizoIndex).StaRequerido Then
        If UserList(UserIndex).Genero = eGenero.Hombre Then
            Call WriteConsoleMsg(2, UserIndex, "Estás muy cansado para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(2, UserIndex, "Estás muy cansada para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
        End If
        PuedeLanzar = False
        Exit Function
    End If
    
    If UserList(UserIndex).Stats.MinMAN < Hechizos(HechizoIndex).ManaRequerido Then
        Call WriteConsoleMsg(2, UserIndex, "No tenes suficiente mana.", FontTypeNames.FONTTYPE_INFO)
        PuedeLanzar = False
        Exit Function
    End If
        
    PuedeLanzar = True
End Function

Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef b As Boolean)
    Dim PosCasteadaX As Integer
    Dim PosCasteadaY As Integer
    Dim PosCasteadaM As Integer
    Dim h As Integer
    Dim TempX As Integer
    Dim TempY As Integer
    Dim TargetUser As Integer
    Dim TargetNPC As Integer
    Dim Daño As Long

    PosCasteadaX = UserList(UserIndex).flags.TargetX
    PosCasteadaY = UserList(UserIndex).flags.TargetY
    PosCasteadaM = UserList(UserIndex).flags.TargetMap
    
    'Distribucion de daño
    'Daño = Porcentaje(Daño, 40 + (100 / Abs(TempX - PosCasteadaX)))
    
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    If Hechizos(h).HechizoDeArea Then
        If Hechizos(h).AreaEfecto <> 0 Then
            b = True
            For TempX = PosCasteadaX - Hechizos(h).AreaEfecto To PosCasteadaX + Hechizos(h).AreaEfecto
                For TempY = PosCasteadaY - Hechizos(h).AreaEfecto To PosCasteadaY + Hechizos(h).AreaEfecto
                    If InMapBounds(PosCasteadaM, TempX, TempY) Then
                        TargetUser = MapData(PosCasteadaM, TempX, TempY).UserIndex
                        If TargetUser > 0 And Not TargetUser = UserIndex Then
                            If UserList(TargetUser).flags.Muerto = 0 Then
                                If Hechizos(h).SubeHP = 1 Then
                                    If Not PuedeAyudar(UserIndex, TargetUser) Then
                                           
                                        Daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
                                        Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
                                        Daño = Porcentaje(Daño, 40 + (100 / Abs(TempX - PosCasteadaX)))
                                        If Daño < 0 Then Daño = 0
                                        
                                        UserList(TargetUser).Stats.MinHP = UserList(TargetUser).Stats.MinHP + Daño
                                        If UserList(TargetUser).Stats.MinHP > UserList(TargetUser).Stats.MaxHP Then _
                                            UserList(TargetUser).Stats.MinHP = UserList(TargetUser).Stats.MaxHP
                                        
                                        Call WriteUpdateHP(TargetUser)
                                        
                                        Call WriteConsoleMsg(1, TargetUser, UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
                                    End If
                                ElseIf Hechizos(h).SubeHP = 2 Then 'Complertar
                                    If TriggerZonaPelea(UserIndex, TargetUser) Then
                                        Daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
                                        
                                        Daño = Daño + Porcentaje(Daño, 2 * UserList(UserIndex).Stats.ELV)
                                        
                                        'Baculos DM + X
                                        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                                            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).EfectoMagico = eMagicType.DañoMagico Then
                                                Daño = Daño + (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).CuantoAumento)
                                            End If
                                        End If
                                        
                                        If (UserList(TargetUser).Invent.CascoEqpObjIndex > 0) Then _
                                            Daño = Daño - ObjData(UserList(TargetUser).Invent.CascoEqpObjIndex).ResistenciaMagica
                                            
                                        If UserList(TargetUser).Invent.EscudoEqpObjIndex > 0 Then _
                                            Daño = Daño - ObjData(UserList(TargetUser).Invent.EscudoEqpObjIndex).ResistenciaMagica
                     
                                        If UserList(TargetUser).Invent.ArmourEqpObjIndex > 0 Then _
                                            Daño = Daño - ObjData(UserList(TargetUser).Invent.ArmourEqpObjIndex).ResistenciaMagica
                    
                                        If UserList(TargetUser).Invent.MonturaObjIndex > 0 Then _
                                            Daño = Daño - ObjData(UserList(TargetUser).Invent.MonturaObjIndex).ResistenciaMagica
                                        
                                        Daño = Porcentaje(Daño, 20 + (100 / IIf(Abs(TempX - PosCasteadaX) = 0, 1, Abs(TempX - PosCasteadaX))))
                                        If Daño < 0 Then Daño = 0
                                        
                                        If Not PuedeAtacar(UserIndex, TargetUser) Then Exit Sub
                                        
                                        If UserIndex <> TargetUser Then
                                            Call UsuarioAtacadoPorUsuario(UserIndex, TargetUser)
                                        End If
                                        
        
                                        UserList(TargetUser).Stats.MinHP = UserList(TargetUser).Stats.MinHP - Daño
                                        
                                        Call WriteUpdateHP(TargetUser)

                                        'Muere
                                        If UserList(TargetUser).Stats.MinHP < 1 Then
                                            Call ContarMuerte(TargetUser, UserIndex)
                                            UserList(TargetUser).Stats.MinHP = 0
                                            Call ActStats(TargetUser, UserIndex)
                                            Call UserDie(TargetUser)
                                        End If
                                        b = True
                                    End If
                                    
                                    If Hechizos(h).Envenena > 0 Then
                                        Call UsuarioAtacadoPorUsuario(UserIndex, TargetUser)
                                        UserList(TargetUser).flags.Envenenado = Hechizos(h).Envenena
                                    End If
                                End If
                            End If
                        End If
                        TargetNPC = MapData(PosCasteadaM, TempX, TempY).NpcIndex
                        If TargetNPC <> 0 Then
                            If PuedeAtacarNPC(UserIndex, TargetNPC) Then

                                Call NPCAtacado(TargetNPC, UserIndex)
                                Daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
                                Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
                                Daño = Porcentaje(Daño, 40 + (100 / IIf(Abs(TempX - PosCasteadaX) = 0, 1, Abs(TempX - PosCasteadaX))))
                                
                                'Baculos DM + X
                                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                                    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).EfectoMagico = eMagicType.DañoMagico Then
                                        Daño = Daño + (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).CuantoAumento)
                                    End If
                                End If
    
                                If Npclist(TargetNPC).flags.Snd2 > 0 Then
                                    Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Npclist(TargetNPC).flags.Snd2, Npclist(TargetNPC).Pos.x, Npclist(TargetNPC).Pos.Y))
                                End If
                            
                                'Quizas tenga defenza magica el NPC. Pablo (ToxicWaste)
                                Daño = Daño - Npclist(TargetNPC).Stats.defM
                                If Daño < 0 Then Daño = 0
                            
                                Npclist(TargetNPC).Stats.MinHP = Npclist(TargetNPC).Stats.MinHP - Daño
                                Call CalcularDarExp(UserIndex, TargetNPC, Daño)
                            
                                If Npclist(TargetNPC).IsFamiliar Then
                                    If Npclist(TargetNPC).MaestroUser > 0 Then
                                        UpdateFamiliar Npclist(TargetNPC).MaestroUser, False
                                    End If
                                End If
                            
                                If Npclist(TargetNPC).Stats.MinHP < 1 Then
                                    Npclist(TargetNPC).Stats.MinHP = 0
                                    Call MuereNpc(TargetNPC, UserIndex)
                                End If
                            End If
                        End If
                    End If
                Next TempY
            Next TempX
        End If
    End If
    
    If Hechizos(h).RemueveInvisibilidadParcial = 1 Then
        b = True
        For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
            For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
                If InMapBounds(PosCasteadaM, TempX, TempY) Then
                    If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then
                        'hay un user
                        If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.Invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.AdminInvisible = 0 Then
                            If Hechizos(h).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))
                            If Hechizos(h).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateCharParticle(UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.CharIndex, Hechizos(h).Particle))
                        End If
                    End If
                End If
            Next TempY
        Next TempX
    
        Call InfoHechizo(UserIndex)
    End If
    
    If Hechizos(h).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticle(PosCasteadaX, PosCasteadaY, Hechizos(h).Particle))
    If Hechizos(h).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFXMap(PosCasteadaX, PosCasteadaY, Hechizos(h).FXgrh, IIf(Hechizos(h).loops < 1, 1, Hechizos(h).loops)))
    If Hechizos(h).WAV <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).WAV, PosCasteadaX, PosCasteadaY))

End Sub

''
' Le da propiedades al nuevo npc
'
' @param UserIndex  Indice del usuario que invoca.
' @param b  Indica si se termino la operación.

Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef b As Boolean)
'***************************************************
'Author: Uknown
'Last modification: 06/15/2008 (NicoNZ)
'Sale del sub si no hay una posición valida.
'***************************************************
If UserList(UserIndex).NroMascotas >= MAXMASCOTAS Then Exit Sub

'No permitimos se invoquen criaturas en zonas seguras
If MapInfo(UserList(UserIndex).Pos.map).Pk = False Or MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).Trigger = eTrigger.ZONASEGURA Then
    Call WriteConsoleMsg(1, UserIndex, "En zona segura no puedes invocar criaturas.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).Pos.map = Prision.map Or UserList(UserIndex).Pos.map = Bandera_mapa Then
    Exit Sub
End If

Dim h As Integer, j As Integer, ind As Integer, index As Integer
Dim TargetPos As WorldPos


TargetPos.map = UserList(UserIndex).flags.TargetMap
TargetPos.x = UserList(UserIndex).flags.TargetX
TargetPos.Y = UserList(UserIndex).flags.TargetY

h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
For j = 1 To Hechizos(h).Cant
    
    If UserList(UserIndex).NroMascotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(h).NumNpc, TargetPos, True, False)
        If ind > 0 Then
            UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas + 1
            
            index = FreeMascotaIndex(UserIndex)
            
            UserList(UserIndex).MascotasIndex(index) = ind
            UserList(UserIndex).MascotasType(index) = Npclist(ind).Numero
            
            Npclist(ind).MaestroUser = UserIndex
            Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
            Npclist(ind).GiveGLD = 0
            
            Call FollowAmo(ind)
        Else
            Exit Sub
        End If
            
    Else
        Exit For
    End If
    
Next j

If ind <> 0 Then
    If Hechizos(h).Particle <> 0 Then Call SendData(SendTarget.ToNPCArea, ind, PrepareMessageCreateParticle(TargetPos.x, TargetPos.Y, Hechizos(h).Particle))
    If Hechizos(h).FXgrh <> 0 Then _
        Call SendData(SendTarget.ToNPCArea, ind, PrepareMessageCreateFXMap(TargetPos.x, TargetPos.Y, Hechizos(h).FXgrh, IIf(Hechizos(h).loops < 1, 1, Hechizos(h).loops)))
    If Hechizos(h).WAV <> 0 Then Call SendData(SendTarget.ToNPCArea, ind, PrepareMessagePlayWave(Hechizos(h).WAV, TargetPos.x, TargetPos.Y))
End If

Call InfoHechizo(UserIndex)
b = True


End Sub

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal uh As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 05/01/08
'
'***************************************************
    If UserList(UserIndex).flags.ModoCombate = False Then
        Call WriteConsoleMsg(1, UserIndex, "Debes estar en modo de combate para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    Dim b As Boolean
    
    Select Case Hechizos(uh).tipo
        Case TipoHechizo.uInvocacion
            Call HechizoInvocacion(UserIndex, b)
        Case TipoHechizo.uEstado, TipoHechizo.uPropiedades
            Call HechizoTerrenoEstado(UserIndex, b)
        Case TipoHechizo.uCreateTelep
            Call HechizoCreateTelep(UserIndex, b)
        Case TipoHechizo.uMaterializa
            Call HechizoMaterializa(UserIndex, b)
        Case TipoHechizo.uFamiliar
            Call HechizoFamiliar(UserIndex, b)
        'Add Nod Kopfnickend Detectar Invis
        Case TipoHechizo.uDetectarInvis
            Call HechizoDetectaInvis(UserIndex, b)
        '\Add
        
    End Select
    
    If b Then
        Call SubirSkill(UserIndex, Magia)
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    
        If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        UserList(UserIndex).Stats.MinSTA = UserList(UserIndex).Stats.MinSTA - Hechizos(uh).StaRequerido
        
        If UserList(UserIndex).Stats.MinSTA < 0 Then UserList(UserIndex).Stats.MinSTA = 0
        Call WriteUpdateUserStats(UserIndex)
    End If


End Sub
Sub HechizoFamiliar(ByVal UserIndex As Integer, ByVal uh As Boolean)
'***************************************************
'Author: Mannakia
'Last Modification: 25/09/10
'
'***************************************************
Dim TargetPos As WorldPos
Dim ind As Integer
    'WTF como hizo para tener este hechi ?
    If UserList(UserIndex).masc.TieneFamiliar = 0 Then
        Exit Sub
    End If
    
    'Add Marius no mascotas en pision o captura la bandera
    If UserList(UserIndex).Pos.map = Prision.map Or UserList(UserIndex).Pos.map = Bandera_mapa Then
        Exit Sub
    End If
    
    TargetPos.map = UserList(UserIndex).Pos.map
    TargetPos.x = UserList(UserIndex).flags.TargetX
    TargetPos.Y = UserList(UserIndex).flags.TargetY
    
    'Desinvocamos
    If UserList(UserIndex).masc.invocado = True Then
    
        'Castelli
        Call desinvocarfami(UserIndex)
        'Castelli
    
    Else 'Invocamos
        If UserList(UserIndex).masc.MinHP > 0 Then
            ind = SpawnNpc(npcFamiToTipe(UserList(UserIndex).masc.tipo), TargetPos, True, False)
            If ind > 0 Then
                
            
                Npclist(ind).MaestroUser = UserIndex
                Npclist(ind).IsFamiliar = True
                
                Call SendData(SendTarget.ToNPCArea, ind, PrepareMessageCreateParticle(TargetPos.x, TargetPos.Y, 116))
                
                UserList(UserIndex).masc.NpcIndex = ind
                
                Call UpdateFamiliar(UserIndex, True)
                
                Call FollowAmo(ind)
                
                Npclist(ind).Movement = TipoAI.NpcFamiliar
                
                UserList(UserIndex).masc.invocado = True
                
                'Actualizamos las habilidades
                If UserList(UserIndex).Clase = eClass.Druida Or _
                    UserList(UserIndex).Clase = eClass.Cazador Or _
                    UserList(UserIndex).Clase = eClass.Mago Then
                        
                    If UserList(UserIndex).masc.ELV >= 10 Then
                        If UserList(UserIndex).masc.tipo = eMascota.Ely Then
                            UserList(UserIndex).masc.Curar = 1
                        ElseIf UserList(UserIndex).masc.tipo = eMascota.Fatuo Then
                            UserList(UserIndex).masc.Misil = 1
                        End If
                    End If
                    
                    If UserList(UserIndex).masc.ELV >= 15 Then
                        If UserList(UserIndex).masc.tipo = eMascota.Ely Or UserList(UserIndex).masc.tipo = eMascota.Fatuo Then
                            UserList(UserIndex).masc.Inmoviliza = 1
                        ElseIf UserList(UserIndex).masc.tipo = eMascota.Lobo Or UserList(UserIndex).masc.tipo = eMascota.Tigre Then
                            UserList(UserIndex).masc.gEntorpece = 1
                        ElseIf UserList(UserIndex).masc.tipo = eMascota.Ent Then
                            UserList(UserIndex).masc.gEnvenena = 1
                        End If
                    End If
                    
                    If UserList(UserIndex).masc.ELV >= 20 Then
                        If UserList(UserIndex).masc.tipo = eMascota.Tigre Or UserList(UserIndex).masc.tipo = eMascota.Ent Or UserList(UserIndex).masc.tipo = eMascota.Lobo Then
                            UserList(UserIndex).masc.gParaliza = 1
                        ElseIf UserList(UserIndex).masc.tipo = eMascota.Ely Or UserList(UserIndex).masc.tipo = eMascota.Fatuo Then
                            UserList(UserIndex).masc.Descargas = 1
                        ElseIf UserList(UserIndex).masc.tipo = eMascota.Fuego Then
                            UserList(UserIndex).masc.Tormentas = 1
                        ElseIf UserList(UserIndex).masc.tipo = eMascota.Agua Then
                            UserList(UserIndex).masc.Paraliza = 1
                        ElseIf UserList(UserIndex).masc.tipo = eMascota.Tierra Then
                            UserList(UserIndex).masc.Inmoviliza = 1
                        End If
                    End If
                    
                    If UserList(UserIndex).masc.ELV >= 30 Then
                        If UserList(UserIndex).masc.tipo = eMascota.Ely Then
                            UserList(UserIndex).masc.Desencanta = 1
                        ElseIf UserList(UserIndex).masc.tipo = eMascota.Fatuo Then
                            UserList(UserIndex).masc.DetecInvi = 1
                        ElseIf UserList(UserIndex).masc.tipo = eMascota.Oso Or UserList(UserIndex).masc.tipo = eMascota.Ent Then
                            UserList(UserIndex).masc.gDesarma = 1
                        ElseIf UserList(UserIndex).masc.tipo = eMascota.Tigre Or UserList(UserIndex).masc.tipo = eMascota.Lobo Then
                            UserList(UserIndex).masc.gEnseguece = 1
                        End If
                    End If
                End If
            End If
        Else
            Call WriteConsoleMsg(1, UserIndex, "Tu familiar esta muerto ¡¡Puedes llevarlo a un sacerdorte para resucitarlo!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
End Sub
Sub UpdateFamiliar(ByVal UserIndex As Integer, ByVal flag As Boolean)
    Dim ind As Integer
    ind = UserList(UserIndex).masc.NpcIndex
    
    If flag = True Then
        If ind > 0 Then
            ind = UserList(UserIndex).masc.NpcIndex
            
            Npclist(ind).Stats.MaxHit = UserList(UserIndex).masc.MaxHit
            Npclist(ind).Stats.MinHit = UserList(UserIndex).masc.MinHit
            
            Npclist(ind).Stats.MaxHP = UserList(UserIndex).masc.MaxHP
            Npclist(ind).Stats.MinHP = UserList(UserIndex).masc.MinHP
            
            Npclist(ind).Name = UserList(UserIndex).masc.Nombre
        End If
    Else
        If ind > 0 Then
            UserList(UserIndex).masc.MinHP = Npclist(ind).Stats.MinHP
            
            If UserList(UserIndex).masc.MinHP <= 0 Then
                UserList(UserIndex).masc.MinHP = 0
                UserList(UserIndex).masc.invocado = False
            End If
        End If
    End If
End Sub
Sub CheckFamiLevel(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If .masc.invocado Then
            If .masc.ELV = 50 Then
                .masc.Exp = 0
                .masc.ELU = 0
                Exit Sub
            End If
            
            Do While .masc.Exp > .masc.ELU
                .masc.Exp = .masc.Exp - .masc.ELU
                If .masc.Exp < 0 Then
                    .masc.Exp = 0
                End If
                
                .masc.ELU = .masc.ELU * 1.2
                
                .masc.ELV = .masc.ELV + 1
                Select Case .masc.tipo
                    Case eMascota.Ely, eMascota.Oso
                        .masc.MaxHP = .masc.MaxHP + 20
                    Case eMascota.Fuego, eMascota.Agua, eMascota.Tierra
                        .masc.MaxHP = .masc.MaxHP + 25
                    Case eMascota.Fatuo
                        .masc.MaxHP = .masc.MaxHP + 15
                        
                    Case eMascota.Tigre
                        .masc.MaxHP = .masc.MaxHP + 18
                    Case eMascota.Lobo
                        .masc.MaxHP = .masc.MaxHP + 35
                    Case eMascota.Ent
                        .masc.MaxHP = .masc.MaxHP + 23
                        
                End Select
                
                .masc.MinHP = .masc.MaxHP
                
                Select Case .masc.tipo
                    Case eMascota.Ely
                        .masc.MaxHit = .masc.MaxHit + 3
                        .masc.MinHit = .masc.MinHit + 3
                    Case eMascota.Fuego, eMascota.Agua, eMascota.Tierra
                        .masc.MaxHit = .masc.MaxHit + 4
                        .masc.MinHit = .masc.MinHit + 4
                        
                    Case eMascota.Fatuo
                        .masc.MaxHit = .masc.MaxHit + 2
                        .masc.MinHit = .masc.MinHit + 2
                        
                    Case eMascota.Tigre, eMascota.Oso
                        .masc.MaxHit = .masc.MaxHit + 6
                        .masc.MinHit = .masc.MinHit + 6
                    Case eMascota.Lobo, eMascota.Ent
                        .masc.MaxHit = .masc.MaxHit + 5
                        .masc.MinHit = .masc.MinHit + 5

                End Select
            Loop
        End If
    End With
End Sub
Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal uh As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 29/10/10
'BY: Jose Ignacio Castelli
'***************************************************


If UserList(UserIndex).flags.ModoCombate = False Then
    Call WriteConsoleMsg(1, UserIndex, "Debes estar en modo de combate para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

Dim b As Boolean

Select Case Hechizos(uh).tipo

    Case TipoHechizo.uEstado, TipoHechizo.uPropEsta, TipoHechizo.uPropiedades ' Afectan estados (por ejem : Envenenamiento)
      
       Call HechizoEsUsuario(UserIndex, b)
       
    Case TipoHechizo.uCreateMagic
        Call HechizoCreateMagic(UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)

    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido

    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSTA = UserList(UserIndex).Stats.MinSTA - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSTA < 0 Then UserList(UserIndex).Stats.MinSTA = 0
    Call WriteUpdateUserStats(UserIndex)

    If Hechizos(uh).AutoLanzar = 0 Then
        Call WriteUpdateUserStats(UserList(UserIndex).flags.TargetUser)
    End If

    UserList(UserIndex).flags.TargetUser = 0
End If

End Sub
Sub HechizoCreateMagic(ByVal UserIndex As Integer, b As Boolean)
'***************************************************
'Author: Leandro Mendoza
'Last Modification: 11/12/2010

'***************************************************
Dim h As Integer

b = False
h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
If h = 0 Then Exit Sub

With UserList(UserIndex)
    If Hechizos(h).CreaAlgo = 1 Then 'Crea arma magica
        If .flags.Muerto <> 0 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        
        .Stats.eMaxHit = Hechizos(h).MaxHit
        .Stats.eMinHit = Hechizos(h).MinHit
  
        
        .Stats.eCreateTipe = 1
        
        If Hechizos(h).Particle <> 0 Then _
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateCharParticle(.Char.CharIndex, Hechizos(h).Particle))
        
        If Hechizos(h).FXgrh <> 0 Then _
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).WAV))
        
        If Hechizos(h).WAV <> 0 Then _
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).WAV, .Pos.x, .Pos.x))
        
        b = True
    ElseIf Hechizos(h).CreaAlgo = 2 Then 'Crea aura sagrada
        If .flags.Muerto <> 0 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        
        .Stats.eMaxDef = Hechizos(h).MaxDef
        .Stats.eMinDef = Hechizos(h).MinDef
        
        
        .Stats.eCreateTipe = 2
        
        If Hechizos(h).Particle <> 0 Then _
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateCharParticle(.Char.CharIndex, Hechizos(h).Particle))
        
        If Hechizos(h).FXgrh <> 0 Then _
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).WAV))
        
        If Hechizos(h).WAV <> 0 Then _
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).WAV, .Pos.x, .Pos.x))
        
        b = True
    ElseIf Hechizos(h).CreaAlgo = 3 Then 'Menos defensa
        If .flags.Muerto <> 0 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        
        .Stats.dMaxDef = Hechizos(h).MaxDef
        .Stats.dMinDef = Hechizos(h).MinDef
        
        
        .Stats.eCreateTipe = 3
        
        If Hechizos(h).Particle <> 0 Then _
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateCharParticle(.Char.CharIndex, Hechizos(h).Particle))
        
        If Hechizos(h).FXgrh <> 0 Then _
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).WAV))
        
        If Hechizos(h).WAV <> 0 Then _
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).WAV, .Pos.x, .Pos.x))
        
        b = True
    End If
End With

End Sub

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal uh As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 13/02/2009
'***************************************************
Dim b As Boolean

Select Case Hechizos(uh).tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
        Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNPC, uh, b, UserIndex)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
        Call HechizoPropNPC(uh, UserList(UserIndex).flags.TargetNPC, UserIndex, b)
End Select


If b Then
    Call SubirSkill(UserIndex, Magia)
    UserList(UserIndex).flags.TargetNPC = 0
    
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido

    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSTA = UserList(UserIndex).Stats.MinSTA - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSTA < 0 Then UserList(UserIndex).Stats.MinSTA = 0
    Call WriteUpdateUserStats(UserIndex)
End If

End Sub


Sub LanzarHechizo(index As Integer, UserIndex As Integer)

Dim uh As Integer

uh = UserList(UserIndex).Stats.UserHechizos(index)

' Dioses pueden tirar hechi sin mana y sin skills
If PuedeLanzar(UserIndex, uh) Or EsCONSE(UserIndex) Then

    Select Case Hechizos(uh).Target
        Case TargetType.uUsuarios
            If UserList(UserIndex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(UserIndex, uh)
                Else
                    Call WriteConsoleMsg(1, UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            Else
                Call WriteConsoleMsg(1, UserIndex, "Este hechizo actúa solo sobre usuarios.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Case TargetType.uNPC
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(UserIndex, uh)
                Else
                    Call WriteConsoleMsg(1, UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            Else
                Call WriteConsoleMsg(1, UserIndex, "Este hechizo solo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Case TargetType.uUsuariosYnpc
       
            If UserList(UserIndex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(UserIndex, uh)
                Else
                    Call WriteConsoleMsg(1, UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(UserIndex, uh)
                Else
                    Call WriteConsoleMsg(1, UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            Else
                Call WriteConsoleMsg(1, UserIndex, "Target invalido.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Case TargetType.uTerreno
            Call HandleHechizoTerreno(UserIndex, uh)
    End Select
    
End If

If UserList(UserIndex).Counters.Trabajando Then _
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1

If UserList(UserIndex).Counters.Ocultando Then _
    UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
    
End Sub



Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 07/07/2008
'Handles the Spells that afect the Stats of an NPC
'04/13/2008 NicoNZ - Guardias Faccionarios pueden ser
'removidos por users de su misma faccion.
'07/07/2008: NicoNZ - Solo se puede mimetizar con npcs si es druida
'***************************************************
If Hechizos(hIndex).Invisibilidad = 1 Then
    Call InfoHechizo(UserIndex)
    Npclist(NpcIndex).flags.Invisible = 1
    b = True
End If

If Hechizos(hIndex).Envenena > 0 Then
    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    Call NPCAtacado(NpcIndex, UserIndex)
    Call InfoHechizo(UserIndex)
    Npclist(NpcIndex).flags.Envenenado = Hechizos(hIndex).Envenena
    b = True
End If

If Hechizos(hIndex).CuraVeneno = 1 Then
    Call InfoHechizo(UserIndex)
    Npclist(NpcIndex).flags.Envenenado = 0
    b = True
End If

If Hechizos(hIndex).Paraliza = 1 Then
    If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
            b = False
            Exit Sub
        End If
        Call NPCAtacado(NpcIndex, UserIndex)
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Paralizado = 1
        'Des Marius bug de inmo y paralizar
        'Npclist(NpcIndex).flags.Inmovilizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        b = True
    Else
        Call WriteConsoleMsg(1, UserIndex, "El NPC es inmune a este hechizo.", FontTypeNames.FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
End If

If Hechizos(hIndex).RemoverParalisis = 1 Then
    If Npclist(NpcIndex).flags.Paralizado = 1 Or Npclist(NpcIndex).flags.Inmovilizado = 1 Then
        If Npclist(NpcIndex).MaestroUser = UserIndex Then
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
            b = True
        Else
            If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                If esArmada(UserIndex) Then
                    Call InfoHechizo(UserIndex)
                    Npclist(NpcIndex).flags.Paralizado = 0
                    Npclist(NpcIndex).Contadores.Paralisis = 0
                    b = True
                    Exit Sub
                Else
                    Call WriteConsoleMsg(1, UserIndex, "Solo puedes Remover la Parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
                
                Call WriteConsoleMsg(1, UserIndex, "Solo puedes Remover la Parálisis de los NPCs que te consideren su amo", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            Else
                If Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos Then
                    If esCaos(UserIndex) Then
                        Call InfoHechizo(UserIndex)
                        Npclist(NpcIndex).flags.Paralizado = 0
                        Npclist(NpcIndex).Contadores.Paralisis = 0
                        b = True
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(1, UserIndex, "Solo puedes Remover la Parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    End If
                End If
            End If
        End If
   Else
      Call WriteConsoleMsg(1, UserIndex, "Este NPC no esta Paralizado", FontTypeNames.FONTTYPE_INFO)
      b = False
      Exit Sub
   End If
End If
 
If Hechizos(hIndex).Inmoviliza = 1 Then
    If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
            b = False
            Exit Sub
        End If
        Call NPCAtacado(NpcIndex, UserIndex)
        Npclist(NpcIndex).flags.Inmovilizado = 1
        'Des Marius bug de inmo y paralizar
        'Npclist(NpcIndex).flags.Paralizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        Call InfoHechizo(UserIndex)
        b = True
    Else
        Call WriteConsoleMsg(1, UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
    End If
End If

End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 14/08/2007
'Handles the Spells that afect the Life NPC
'14/08/2007 Pablo (ToxicWaste) - Orden general.
'***************************************************

Dim Daño As Long

'Salud
If Hechizos(hIndex).SubeHP = 1 Then
    Daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
    
    Call InfoHechizo(UserIndex)
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP + Daño
    If Npclist(NpcIndex).Stats.MinHP > Npclist(NpcIndex).Stats.MaxHP Then _
        Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
    Call WriteConsoleMsg(1, UserIndex, "Has curado " & Daño & " puntos de salud a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
    b = True
    
ElseIf Hechizos(hIndex).SubeHP = 2 Then
    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    Call NPCAtacado(NpcIndex, UserIndex)
    Daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)

    'Baculos DM + X
    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).EfectoMagico = eMagicType.DañoMagico Then
            Daño = Daño + (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).CuantoAumento)
        End If
    End If

    Call InfoHechizo(UserIndex)
    b = True
    
    If Npclist(NpcIndex).flags.Snd2 > 0 Then
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y))
    End If
    
    'Quizas tenga defenza magica el NPC. Pablo (ToxicWaste)
    Daño = Daño - Npclist(NpcIndex).Stats.defM
    If Daño < 0 Then Daño = 0
    
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - Daño
    Call WriteConsoleMsg(2, UserIndex, "¡Le has causado " & Daño & " puntos de daño a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
    Call CalcularDarExp(UserIndex, NpcIndex, Daño)
    
    If Npclist(NpcIndex).IsFamiliar Then
        If Npclist(NpcIndex).MaestroUser > 0 Then
            UpdateFamiliar Npclist(NpcIndex).MaestroUser, False
        End If
    End If
    
    If Npclist(NpcIndex).Stats.MinHP < 1 Then
        Npclist(NpcIndex).Stats.MinHP = 0
        Call MuereNpc(NpcIndex, UserIndex)
    End If
End If

End Sub

Sub InfoHechizo(ByVal UserIndex As Integer)


    Dim h As Integer
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    
    Call DecirPalabrasMagicas(Hechizos(h).PalabrasMagicas, UserIndex)
    
    If UserList(UserIndex).flags.TargetUser > 0 Then
        If Hechizos(h).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessageCreateFX(UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))
        Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessagePlayWave(Hechizos(h).WAV, UserList(UserList(UserIndex).flags.TargetUser).Pos.x, UserList(UserList(UserIndex).flags.TargetUser).Pos.Y))  'Esta linea faltaba. Pablo (ToxicWaste)
        If Hechizos(h).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessageCreateCharParticle(UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex, Hechizos(h).Particle))
    ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
        If Hechizos(h).FXgrh <> 0 Then Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageCreateFX(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))
        Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessagePlayWave(Hechizos(h).WAV, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.x, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y))
        If Hechizos(h).Particle <> 0 Then Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageCreateCharParticle(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, Hechizos(h).Particle))
    End If
    
 '   If UserList(userindex).flags.TargetUser > 0 Then
 '       If userindex <> UserList(userindex).flags.TargetUser Then
 '           If UserList(userindex).showName Then
 '               Call WriteConsoleMsg(2, userindex, Hechizos(h).HechizeroMsg & " " & UserList(UserList(userindex).flags.TargetUser).Name, FontTypeNames.FONTTYPE_FIGHT)
 '           Else
 '               Call WriteConsoleMsg(2, userindex, Hechizos(h).HechizeroMsg & " alguien.", FontTypeNames.FONTTYPE_FIGHT)
 '           End If
 '           Call WriteConsoleMsg(2, UserList(userindex).flags.TargetUser, UserList(userindex).Name & " " & Hechizos(h).TargetMsg, FontTypeNames.FONTTYPE_FIGHT)
 '       Else
 '           Call WriteConsoleMsg(2, userindex, Hechizos(h).PropioMsg, FontTypeNames.FONTTYPE_FIGHT)
 '       End If
 '   ElseIf UserList(userindex).flags.TargetNPC > 0 Then
 '       Call WriteConsoleMsg(2, userindex, Hechizos(h).HechizeroMsg & " " & "la criatura.", FontTypeNames.FONTTYPE_FIGHT)
 '   End If

End Sub

Sub HechizoEsUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 02/01/2008
'02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
'***************************************************

Dim h As Integer
Dim Daño As Long
Dim tU As Integer

h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
tU = UserList(UserIndex).flags.TargetUser
b = False

If UserList(UserIndex).flags.Muerto = 1 Then
    Call WriteMsg(UserIndex, 0)
    b = False
    Exit Sub
End If

If Hechizos(h).ReviveFamiliar = 1 Then
    If Not PuedeAyudar(UserIndex, tU) Then
        Call WriteMsg(UserIndex, 33)
        b = False
        Exit Sub
    End If
    
    If UserList(tU).masc.TieneFamiliar = 1 Then
        If UserList(tU).masc.MinHP <= 0 Then
            UserList(tU).masc.MinHP = UserList(tU).masc.MaxHP
            
            UpdateFamiliar tU, True
            
            b = True
        End If
    End If
End If

If Hechizos(h).Resurreccion = 1 Then
    If UserList(tU).flags.Muerto = 1 Then
        'No usar resu en mapas con ResuSinEfecto
        If MapInfo(UserList(tU).Pos.map).ResuSinEfecto > 0 Then
            Call WriteConsoleMsg(1, UserIndex, "¡Revivir no está permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        
        'Para poder tirar revivir a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
            If Not PuedeAyudar(UserIndex, tU) Then
                Call WriteMsg(UserIndex, 33)
                b = False
                Exit Sub
            End If
        End If
        
        'Add Nod Kopfnickend Solo revive si no esta en modo combate
        If UserList(tU).flags.ModoCombate Then
            Call WriteConsoleMsg(1, UserIndex, "¡No puedes revivir a alguien que esta en modo combate.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(1, tU, "¡" & UserList(UserIndex).Name & " esta intentando revivirte, desactiva el modo combate si quieres que te reviva.", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        '\Add

        DarVida tU
        
        UserList(tU).Stats.MinAGU = 100
        UserList(tU).flags.Sed = 0
        UserList(tU).Stats.MinHAM = 100
        UserList(tU).flags.Hambre = 0
        
        UserList(tU).Stats.MinHP = UserList(tU).Stats.MaxHP
        UserList(tU).Stats.MinMAN = UserList(tU).Stats.MaxMAN
        UserList(tU).Stats.MinSTA = UserList(tU).Stats.MaxSTA
        
        Call WriteUpdateUserStats(tU)
        
        Call InfoHechizo(UserIndex)
        b = True
        Exit Sub
    Else
        b = False
    End If
End If

If Hechizos(h).Revivir = 1 Then
    If UserList(tU).flags.Muerto = 1 Then
        'No usar resu en mapas con ResuSinEfecto
        If MapInfo(UserList(tU).Pos.map).ResuSinEfecto > 0 Then
            Call WriteConsoleMsg(1, UserIndex, "¡Revivir no está permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        
        'Para poder tirar revivir a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
            If Not PuedeAyudar(UserIndex, tU) Then
                Call WriteMsg(UserIndex, 33)
                b = False
                Exit Sub
            End If
        End If
        
        'Add Nod Kopfnickend Solo revive si no esta en modo combate
        If UserList(tU).flags.ModoCombate Then
            Call WriteConsoleMsg(1, UserIndex, "¡No puedes revivir a alguien que esta en modo combate.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(1, tU, "¡" & UserList(UserIndex).Name & " esta intentando revivirte, desactiva el modo combate si quieres que te reviva.", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        '\Add

        UserList(tU).Stats.MinAGU = 0
        UserList(tU).flags.Sed = 1
        UserList(tU).Stats.MinHAM = 0
        UserList(tU).flags.Hambre = 1
        UserList(tU).Stats.MinMAN = 0
        UserList(tU).Stats.MinSTA = 0
                    
        Call RevivirUsuario(tU)
        
        Call WriteUpdateHungerAndThirst(tU)
        Call InfoHechizo(UserIndex)
        b = True
        Exit Sub
    Else
        b = False
    End If

End If

If UserList(tU).flags.Muerto Then
    Call WriteMsg(UserIndex, 28)
    b = False
    Exit Sub
End If

If Hechizos(h).Sanacion = 1 Then
    If UserList(tU).flags.Incinerado = 1 Then _
        UserList(tU).flags.Incinerado = 0
    
    If UserList(tU).flags.Envenenado Then _
        UserList(tU).flags.Envenenado = 0

    If UserList(tU).flags.Estupidez = 1 Then
        UserList(tU).flags.Estupidez = 0
        Call WriteDumb(tU)
    End If
    
    If UserList(tU).flags.Paralizado = 1 Then
        UserList(tU).flags.Inmovilizado = 0
        UserList(tU).flags.Paralizado = 0
        Call WriteParalizeOK(tU)
    End If
    
    UserList(tU).flags.Envenenado = 0
    
    b = True
End If

If Hechizos(h).Certero = 1 Then
    UserList(UserIndex).flags.NoFalla = 1
    
    If Hechizos(h).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, tU, PrepareMessageCreateCharParticle(UserList(tU).Char.CharIndex, Hechizos(h).Particle))
    If Hechizos(h).WAV <> 0 Then Call SendData(SendTarget.ToPCArea, tU, PrepareMessagePlayWave(Hechizos(h).WAV, UserList(tU).Pos.x, UserList(tU).Pos.Y))
    b = True
End If

If Hechizos(h).Desencantar Then
    
    If UserList(tU).flags.Incinerado = 1 Then
        UserList(tU).flags.Incinerado = 0
    End If
    
    If UserList(tU).flags.Envenenado Then
        UserList(tU).flags.Envenenado = 0
    End If

    If UserList(tU).flags.Estupidez = 1 Then
        UserList(tU).flags.Estupidez = 0
        Call WriteDumb(tU)
    End If
        
    If UserList(tU).flags.Ceguera = 1 Then
        UserList(UserIndex).flags.Ceguera = 0
        Call WriteBlindNoMore(tU)
    End If
    
    If UserList(tU).flags.Metamorfosis = 1 Then
        UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(UserIndex).Char.body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(UserIndex)
        End If
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
     
        UserList(UserIndex).flags.Metamorfosis = 0
        UserList(UserIndex).Counters.Metamorfosis = 0
     
        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        
        'Marius movi esto dentro del If para que no se sarpen tirandose el hechi en segura y dar lag solo para que ellos se hagan los pro
        If Hechizos(h).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, tU, PrepareMessageCreateCharParticle(UserList(tU).Char.CharIndex, Hechizos(h).Particle))
        If Hechizos(h).WAV <> 0 Then Call SendData(SendTarget.ToPCArea, tU, PrepareMessagePlayWave(Hechizos(h).WAV, UserList(tU).Pos.x, UserList(tU).Pos.Y))
    End If
    
End If

If Hechizos(h).Invisibilidad = 1 Then
    If UserList(tU).flags.Navegando = 1 Then
        Call WriteMsg(UserIndex, 29)
        b = False
        Exit Sub
    End If
    
    If UserList(tU).Counters.Saliendo Then
        If UserIndex <> tU Then
            Call WriteMsg(UserIndex, 30)
            b = False
            Exit Sub
        Else
            Call WriteMsg(UserIndex, 31)
            b = False
            Exit Sub
        End If
    End If
    
    'No usar invi mapas InviSinEfecto
    If MapInfo(UserList(tU).Pos.map).InviSinEfecto > 0 Then
        Call WriteMsg(UserIndex, 32)
        b = False
        Exit Sub
    End If
    
    'Para poder tirar invi a un pk en el ring
    If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
        If Not PuedeAyudar(UserIndex, tU) Then
            Call WriteMsg(UserIndex, 33)
            b = False
            Exit Sub
        End If
    End If

    UserList(tU).flags.Invisible = 1
    Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.CharIndex, True))

    b = True
End If

If Hechizos(h).Envenena > 0 Then
    If UserIndex = tU Then
        Call WriteMsg(UserIndex, 34)
        b = False
        Exit Sub
    End If
    
    If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
    If UserIndex <> tU Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tU)
    End If
    UserList(tU).flags.Envenenado = Hechizos(h).Envenena
    b = True
End If

If Hechizos(h).Incinera = 1 Then
    If UserIndex = tU Then
        Call WriteMsg(UserIndex, 34)
        b = False
        Exit Sub
    End If
    
    If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
    If UserIndex <> tU Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tU)
    End If
    UserList(tU).flags.Incinerado = 1
    b = True
End If

If Hechizos(h).CuraVeneno = 1 Then
    'Para poder tirar curar veneno a un pk en el ring
    If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
        If Not PuedeAyudar(UserIndex, tU) Then
            Call WriteMsg(UserIndex, 33)
            b = False
            Exit Sub
        End If
    End If
        
    'Si sos user, no uses este hechizo con GMS.
    If UserList(UserIndex).flags.Privilegios And (PlayerType.User Or PlayerType.VIP) Then
        If EsCONSE(tU) Then
            Exit Sub
        End If
    End If
        
    UserList(tU).flags.Envenenado = 0
    b = True
End If

If Hechizos(h).Paraliza = 1 Or Hechizos(h).Inmoviliza = 1 Then
    If UserIndex = tU Then
        Call WriteMsg(UserIndex, 34)
        Exit Sub
    End If
    
     If UserList(tU).flags.Paralizado = 0 Then
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)
        End If
            
        b = True
            
        If Hechizos(h).Inmoviliza = 1 Then UserList(tU).flags.Inmovilizado = 1
        UserList(tU).flags.Paralizado = 1
        UserList(tU).Counters.Paralisis = IntervaloParalizado
            
        Call WriteParalizeOK(tU)
        Call FlushBuffer(tU)
      
    End If
End If


If Hechizos(h).RemoverParalisis = 1 Then
    If UserList(tU).flags.Paralizado = 1 Then
        'Para poder tirar remo a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
            If Not PuedeAyudar(UserIndex, tU) Then
                Call WriteMsg(UserIndex, 33)
                b = False
                Exit Sub
            End If
        End If
        
        UserList(tU).flags.Inmovilizado = 0
        UserList(tU).flags.Paralizado = 0
        'no need to crypt this
        Call WriteParalizeOK(tU)
        b = True
    End If
End If

If Hechizos(h).Ceguera = 1 Then
    If UserIndex = tU Then
        Call WriteMsg(UserIndex, 34)
        Exit Sub
    End If
    
    If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
    
    If UserIndex <> tU Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tU)
    End If
    UserList(tU).flags.Ceguera = 1
    UserList(tU).Counters.Ceguera = IntervaloParalizado / 3

    Call WriteBlind(tU)
    Call FlushBuffer(tU)
    b = True
End If

If Hechizos(h).Estupidez = 1 Then
    If UserIndex = tU Then
        Call WriteMsg(UserIndex, 34)
        Exit Sub
    End If
        
    If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
    If UserIndex <> tU Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tU)
    End If
    If UserList(tU).flags.Estupidez = 0 Then
        UserList(tU).flags.Estupidez = 1
        UserList(tU).Counters.Ceguera = IntervaloParalizado
    End If
    Call WriteDumb(tU)
    Call FlushBuffer(tU)

    b = True
End If


If Hechizos(h).Metamorfosis = 1 Then
    If UserList(UserIndex).flags.Montando = 1 Then
        Call WriteConsoleMsg(1, UserIndex, "Estas montando!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.Navegando = 1 Then
        Call WriteConsoleMsg(1, UserIndex, "Estas navegando!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    Debug.Print ">>>Metanorfosis"
    
        
    If Hechizos(h).MetaObj = 0 Or TieneObjetos(Hechizos(h).MetaObj, 1, UserIndex) Then
        
        Call DoMetamorfosis(UserIndex, Hechizos(h).body, Hechizos(h).Head)
        
        If Hechizos(h).WAV <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).WAV, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY))
        If Hechizos(h).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(h).FXgrh, 1))
        
        b = True
    Else
        Call WriteConsoleMsg(1, UserIndex, "Necesitas " & ObjData(Hechizos(h).MetaObj).Name & " para usar este hechizo!", FontTypeNames.FONTTYPE_INFO)
    End If


End If


' <-------- Agilidad ---------->
If Hechizos(h).SubeAgilidad = 1 Then
    
    'Para poder tirar cl a un pk en el ring
    If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
        If Not PuedeAyudar(UserIndex, tU) Then
            Call WriteMsg(UserIndex, 34)
            b = False
            Exit Sub
        End If
    End If
    
    Daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
    
    UserList(tU).flags.DuracionEfecto = 1200
    UserList(tU).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tU).Stats.UserAtributos(eAtributos.Agilidad) + Daño
    If UserList(tU).Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then _
        UserList(tU).Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
    Call WriteAgilidad(tU)
    
    UserList(tU).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(h).SubeAgilidad = 2 Then
    
    If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
    
    If UserIndex <> tU Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tU)
    End If
    
    UserList(tU).flags.TomoPocion = True
    Daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
    UserList(tU).flags.DuracionEfecto = 700
    UserList(tU).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tU).Stats.UserAtributos(eAtributos.Agilidad) - Daño
    If UserList(tU).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(tU).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
    Call WriteAgilidad(tU)
    b = True
    
End If

' <-------- Fuerza ---------->
If Hechizos(h).SubeFuerza = 1 Then
    'Para poder tirar fuerza a un pk en el ring
    If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
        If Not PuedeAyudar(UserIndex, tU) Then
            Call WriteConsoleMsg(1, UserIndex, "No puedes beneficiar a ese tipo de gente.", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
    End If
    
    Daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    
    UserList(tU).flags.DuracionEfecto = 1200

    UserList(tU).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tU).Stats.UserAtributos(eAtributos.Fuerza) + Daño
    If UserList(tU).Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then _
        UserList(tU).Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
    Call WriteFuerza(tU)
    UserList(tU).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(h).SubeFuerza = 2 Then

    If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
    
    If UserIndex <> tU Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tU)
    End If
    
    UserList(tU).flags.TomoPocion = True
    
    Daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    UserList(tU).flags.DuracionEfecto = 700
    UserList(tU).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tU).Stats.UserAtributos(eAtributos.Fuerza) - Daño
    If UserList(tU).Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then UserList(tU).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
    b = True
    Call WriteFuerza(tU)
End If

'Salud
If Hechizos(h).SubeHP = 1 Then
    'Para poder tirar curar a un pk en el ring
    If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
        If Not PuedeAyudar(UserIndex, tU) Then
            Call WriteMsg(UserIndex, 34)
            b = False
            Exit Sub
        End If
    End If
       
    Daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
    Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
    
    UserList(tU).Stats.MinHP = UserList(tU).Stats.MinHP + Daño
    If UserList(tU).Stats.MinHP > UserList(tU).Stats.MaxHP Then _
        UserList(tU).Stats.MinHP = UserList(tU).Stats.MaxHP
    
    Call WriteUpdateHP(tU)
    
    If UserIndex <> tU Then
        Call WriteConsoleMsg(2, UserIndex, "Le has restaurado " & Daño & " puntos de vida a " & UserList(tU).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(2, tU, UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(2, UserIndex, "Te has restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    b = True
ElseIf Hechizos(h).SubeHP = 2 Then
    
    If UserIndex = tU Then
        Call WriteMsg(UserIndex, 34)
        Exit Sub
    End If
    
    Daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
    
    Daño = Daño + Porcentaje(Daño, 2 * UserList(UserIndex).Stats.ELV)
    
    'Baculos DM + X
    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).EfectoMagico = eMagicType.DañoMagico Then
            Daño = Daño + (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).CuantoAumento)
        End If
    End If
    
    'cascos antimagia
    If (UserList(tU).Invent.CascoEqpObjIndex > 0) Then
        Daño = Daño - RandomNumber(ObjData(UserList(tU).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tU).Invent.CascoEqpObjIndex).DefensaMagicaMax)
        
        Daño = Daño - ObjData(UserList(tU).Invent.CascoEqpObjIndex).ResistenciaMagica
    End If

    If UserList(tU).Invent.EscudoEqpObjIndex > 0 Then
        Daño = Daño - ObjData(UserList(tU).Invent.EscudoEqpObjIndex).ResistenciaMagica
    End If
    
    If UserList(tU).Invent.ArmourEqpObjIndex > 0 Then
        Daño = Daño - ObjData(UserList(tU).Invent.ArmourEqpObjIndex).ResistenciaMagica
    End If
    
    If UserList(tU).Invent.MonturaObjIndex > 0 Then
        Daño = Daño - ObjData(UserList(tU).Invent.MonturaObjIndex).ResistenciaMagica
    End If
        
    
    If Daño < 0 Then Daño = 0
    
    If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
    
    If UserIndex <> tU Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tU)
    End If
    
    UserList(tU).Stats.MinHP = UserList(tU).Stats.MinHP - Daño
    
    Call SubirSkill(tU, eSkill.Resistencia)
    
    Call WriteUpdateHP(tU)
    
    Call WriteMsg(UserIndex, 39, CStr(UserList(tU).Char.CharIndex), CStr(Daño))
    Call WriteMsg(tU, 38, CStr(UserList(UserIndex).Char.CharIndex), CStr(h))
    
    'Muere
    If UserList(tU).Stats.MinHP < 1 Then
        Call ContarMuerte(tU, UserIndex)
        UserList(tU).Stats.MinHP = 0
        Call ActStats(tU, UserIndex)
        Call UserDie(tU)
    End If
    
    b = True
End If

If b = True Then
    InfoHechizo UserIndex
End If
    
FlushBuffer UserIndex
FlushBuffer tU

End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

'Call LogTarea("Sub UpdateUserHechizos")

Dim loopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
        Call ChangeUserHechizo(UserIndex, Slot, UserList(UserIndex).Stats.UserHechizos(Slot))
    Else
        Call ChangeUserHechizo(UserIndex, Slot, 0)
    End If

Else

'Actualiza todos los slots
For loopC = 1 To MAXUSERHECHIZOS

        'Actualiza el inventario
        If UserList(UserIndex).Stats.UserHechizos(loopC) > 0 Then
            Call ChangeUserHechizo(UserIndex, loopC, UserList(UserIndex).Stats.UserHechizos(loopC))
        Else
            Call ChangeUserHechizo(UserIndex, loopC, 0)
        End If

Next loopC

End If

End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)

'Call LogTarea("ChangeUserHechizo")

UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo


If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
    
    Call WriteChangeSpellSlot(UserIndex, Slot)

Else

    Call WriteChangeSpellSlot(UserIndex, Slot)

End If


End Sub


Public Sub DesplazarHechizo(ByVal UserIndex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)

If (Dire <> 1 And Dire <> -1) Then Exit Sub
If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

Dim TempHechizo As Integer

If Dire = 1 Then 'Mover arriba
    If CualHechizo = 1 Then
        Call WriteConsoleMsg(1, UserIndex, "No puedes mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo

        'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
        If UserList(UserIndex).flags.Hechizo > 0 Then
            UserList(UserIndex).flags.Hechizo = UserList(UserIndex).flags.Hechizo - 1
        End If
    End If
Else 'mover abajo
    If CualHechizo = MAXUSERHECHIZOS Then
        Call WriteConsoleMsg(1, UserIndex, "No puedes mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo

        'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
        If UserList(UserIndex).flags.Hechizo > 0 Then
            UserList(UserIndex).flags.Hechizo = UserList(UserIndex).flags.Hechizo + 1
        End If
    End If
End If
End Sub


 Sub HechizoCreateTelep(UserIndex As Integer, b As Boolean)
  
    Dim tU As Integer
    Dim h As Integer
    Dim i As Integer
    Dim PosTIROTELEPORT As WorldPos
  
      h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
      
      'Add Marius =) Comela puto si te digo que vas preso es para que te quedes adentro. Puto!
      'Add Marius Captura la bandera
      'Add Marius no crear tp en Intermundia
      If UserList(UserIndex).Pos.map = Prision.map Or UserList(UserIndex).Pos.map = Bandera_mapa Or UserList(UserIndex).Pos.map = 49 Then
        Call WriteConsoleMsg(2, UserIndex, "Una fuerza superior te impide crear un portal.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
      End If
      '\Add
      
      If MapInfo(UserList(UserIndex).Pos.map).Pk = False Then
            Exit Sub
      End If
      
      If Hechizos(h).Nombre = "Portal planar" Then
        If UserList(UserIndex).flags.TiroPortalL = 1 Then
          Call WriteConsoleMsg(2, UserIndex, "Ya tienes un Portal creado.", FontTypeNames.FONTTYPE_INFO)
          Exit Sub
        End If
      End If
 
    
     PosTIROTELEPORT.x = UserList(UserIndex).flags.TargetX
     PosTIROTELEPORT.Y = UserList(UserIndex).flags.TargetY
     PosTIROTELEPORT.map = UserList(UserIndex).flags.TargetMap
         
    If Not LegalPos(PosTIROTELEPORT.map, PosTIROTELEPORT.x, PosTIROTELEPORT.Y) Then
        Exit Sub
    End If
    
    If MapInfo(PosTIROTELEPORT.map).Pk = False Then
        Exit Sub
    End If
         
     UserList(UserIndex).flags.DondeTiroMap = PosTIROTELEPORT.map
 
     UserList(UserIndex).flags.DondeTiroX = PosTIROTELEPORT.x
     UserList(UserIndex).flags.DondeTiroY = PosTIROTELEPORT.Y
    
     If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).ObjInfo.ObjIndex Then
         Exit Sub
     End If
    
     If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).TileExit.map Then
         Exit Sub
     End If
    
     If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked Then
         Exit Sub
     End If
     If Not MapaValido(UserList(UserIndex).Pos.map) Or Not InMapBounds(UserList(UserIndex).flags.TargetMap, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY) Then Exit Sub
 
   Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).WAV, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY))
   Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticle(UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY, Hechizos(h).Particle))

   UserList(UserIndex).Counters.TimeTeleport = 0
   UserList(UserIndex).Counters.CreoTeleport = True
   UserList(UserIndex).flags.TiroPortalL = 1
   Call InfoHechizo(UserIndex)
   
   b = True
  
 End Sub

Sub HechizoDetectaInvis(UserIndex As Integer, b As Boolean)
'Add Nod Kopfnickend
    Dim h As Integer
    
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    
    
'Funcion tomada del sentopcarea
    Dim loopC As Long
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = UserList(UserIndex).Pos.map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(map) Then Exit Sub
    
    For loopC = 1 To ConnGroups(map).CountEntrys
        If loopC <= ConnGroups(map).CountEntrys Then
            tempIndex = ConnGroups(map).UserEntrys(loopC)
            
            If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                    If UserList(tempIndex).ConnIDValida Then
                        
                        'Si esta oculto o invi lo volvemos visible
                        If UserList(tempIndex).flags.Oculto = 1 Or UserList(tempIndex).flags.Invisible = 1 Then
                            UserList(tempIndex).flags.Invisible = 0
                            UserList(tempIndex).Counters.Invisibilidad = 0
                            
                            UserList(tempIndex).flags.Oculto = 0
                            UserList(tempIndex).Counters.Ocultando = 0
                            UserList(tempIndex).Counters.TiempoOculto = 0
                    
                            Call WriteConsoleMsg(1, tempIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                            Call SendData(SendTarget.ToPCArea, tempIndex, PrepareMessageSetInvisible(UserList(tempIndex).Char.CharIndex, False))
                        End If
                        
                    End If
                End If
            End If
        Else
            Exit For
        End If
    Next loopC
    DoEvents
    
    
    If Hechizos(h).Particle <> 0 Then _
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateCharParticle(UserList(UserIndex).Char.CharIndex, Hechizos(h).Particle))
        
    If Hechizos(h).FXgrh <> 0 Then _
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).WAV))
        
    If Hechizos(h).WAV <> 0 Then _
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).WAV, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.x))
        
     b = True
End Sub
Sub HechizoMaterializa(UserIndex As Integer, b As Boolean)
    Dim M As Integer
    Dim x As Integer
    Dim Y As Integer
    Dim obj As obj
    Dim h As Integer
    
    obj.Amount = 1
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
      
    If Hechizos(h).Nombre = "Materializar: Comida" Then
         obj.ObjIndex = 1
    ElseIf Hechizos(h).Nombre = "Materializar: Bebida" Then
         obj.ObjIndex = 43
    End If

    'Exit Sub
    If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).ObjInfo.ObjIndex Then
        Exit Sub
    End If
    
    If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).TileExit.map Then
        Exit Sub
    End If
        
    If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked Then
        Exit Sub
    End If
         
    If Not MapaValido(UserList(UserIndex).Pos.map) Or Not InMapBounds(UserList(UserIndex).flags.TargetMap, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY) Then Exit Sub
    
    M = UserList(UserIndex).flags.TargetMap
    x = UserList(UserIndex).flags.TargetX
    Y = UserList(UserIndex).flags.TargetY

    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).WAV, x, Y))
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticle(x, Y, Hechizos(h).Particle))
       
    Call MakeObj(obj, M, x, Y)
    
    b = True
End Sub

Public Sub desinvocarfami(UserIndex As Integer)
    'Por la dudas le mandamos un update para que llene
    Call UpdateFamiliar(UserIndex, False)
    
    Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).masc.NpcIndex, PrepareMessageCreateParticle(Npclist(UserList(UserIndex).masc.NpcIndex).Pos.x, Npclist(UserList(UserIndex).masc.NpcIndex).Pos.Y, 117))
    
    UserList(UserIndex).masc.invocado = False
    Call QuitarNPC(UserList(UserIndex).masc.NpcIndex)
    
    UserList(UserIndex).masc.NpcIndex = 0
End Sub
