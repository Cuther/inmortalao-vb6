Attribute VB_Name = "modHechizos"

Option Explicit

Public Const HELEMENTAL_FUEGO As Integer = 26
Public Const HELEMENTAL_TIERRA As Integer = 28
Public Const ANILLO_ESPECTRAL As Integer = 1329
Public Const ANILLO_PENUMBRAS As Integer = 1330

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal userindex As Integer, ByVal Spell As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 13/02/2009
'13/02/2009: ZaMa - Los npcs que tiren magias, no podran hacerlo en mapas donde no se permita usarla.
'***************************************************
On Error GoTo hayerror

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
If UserList(userindex).flags.Invisible = 1 Or UserList(userindex).flags.Oculto = 1 Then Exit Sub

' Si no se peude usar magia en el mapa, no le deja hacerlo.
If MapInfo(UserList(userindex).Pos.map).MagiaSinEfecto > 0 Then Exit Sub

'Mannakia
If UserList(userindex).Invent.MagicIndex > 0 Then
    If ObjData(UserList(userindex).Invent.MagicIndex).EfectoMagico = eMagicType.MagicasNoAtacan Then
        Exit Sub
    End If
End If
'Mannakia

Npclist(NpcIndex).CanAttack = 0
Dim Daño As Integer

If Hechizos(Spell).SubeHP = 1 Then

    Daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(Hechizos(Spell).WAV, UserList(userindex).Pos.x, UserList(userindex).Pos.Y))
    If Hechizos(Spell).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(UserList(userindex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
    If Hechizos(Spell).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateCharParticle(UserList(userindex).Char.CharIndex, Hechizos(Spell).Particle))
    
    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP + Daño
    If UserList(userindex).Stats.MinHP > UserList(userindex).Stats.MaxHP Then UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
    
    Call WriteConsoleMsg(2, userindex, Npclist(NpcIndex).Name & " te ha quitado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    Call WriteUpdateUserStats(userindex)

    Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageChatOverHead(Hechizos(Spell).PalabrasMagicas, Npclist(NpcIndex).Char.CharIndex, RGB(128, 128, 0)))
    
ElseIf Hechizos(Spell).SubeHP = 2 Then
    
    If UserList(userindex).flags.Privilegios And (PlayerType.User Or PlayerType.VIP) Then
    
        Daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        
        If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
            Daño = Daño - RandomNumber(ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMax)
            ' Jose Castelli / Resistencia Magica (RM)
            Daño = Daño - ObjData(UserList(userindex).Invent.CascoEqpObjIndex).ResistenciaMagica
            ' Jose Castelli / Resistencia Magica (RM)
        End If
        
        ' Jose Castelli / Resistencia Magica (RM)
        
        If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then
            Daño = Daño - ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).ResistenciaMagica
        End If
        
        If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
            Daño = Daño - ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).ResistenciaMagica
        End If
        
        If UserList(userindex).Invent.MonturaObjIndex > 0 Then
            Daño = Daño - ObjData(UserList(userindex).Invent.MonturaObjIndex).ResistenciaMagica
        End If
        
        ' Jose Castelli / Resistencia Magica (RM)
        
        If Daño < 0 Then Daño = 0
        
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(Hechizos(Spell).WAV, UserList(userindex).Pos.x, UserList(userindex).Pos.Y))
        If Hechizos(Spell).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(UserList(userindex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
        If Hechizos(Spell).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateCharParticle(UserList(userindex).Char.CharIndex, Hechizos(Spell).Particle))

        UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - Daño
        
        Call WriteConsoleMsg(2, userindex, Npclist(NpcIndex).Name & " te ha quitado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteUpdateUserStats(userindex)
        
        Call SubirSkill(userindex, eSkill.Resistencia)
        
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageChatOverHead(Hechizos(Spell).PalabrasMagicas, Npclist(NpcIndex).Char.CharIndex, RGB(128, 128, 0)))
        
        'Muere
        If UserList(userindex).Stats.MinHP < 1 Then
            UserList(userindex).Stats.MinHP = 0
            Call UserDie(userindex)
            '[Barrin 1-12-03]
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call ContarMuerte(userindex, Npclist(NpcIndex).MaestroUser)
                Call ActStats(userindex, Npclist(NpcIndex).MaestroUser)
            End If
            '[/Barrin]
        End If
    
    End If
    
End If

If Hechizos(Spell).Paraliza = 1 Or Hechizos(Spell).Inmoviliza = 1 Then
    If UserList(userindex).flags.Paralizado = 0 Then
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(Hechizos(Spell).WAV, UserList(userindex).Pos.x, UserList(userindex).Pos.Y))
        If Hechizos(Spell).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(UserList(userindex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
        If Hechizos(Spell).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateCharParticle(UserList(userindex).Char.CharIndex, Hechizos(Spell).Particle))

        If Hechizos(Spell).Inmoviliza = 1 Then
            UserList(userindex).flags.Inmovilizado = 1
        End If
          
        UserList(userindex).flags.Paralizado = 1
        UserList(userindex).Counters.Paralisis = IntervaloParalizado
          
        Call WriteParalizeOK(userindex)
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageChatOverHead(Hechizos(Spell).PalabrasMagicas, Npclist(NpcIndex).Char.CharIndex, RGB(128, 128, 0)))
        
    End If
End If

If Hechizos(Spell).Estupidez = 1 Then   ' turbacion
     If UserList(userindex).flags.Estupidez = 0 Then
          Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(Hechizos(Spell).WAV, UserList(userindex).Pos.x, UserList(userindex).Pos.Y))
          If Hechizos(Spell).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(UserList(userindex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
          If Hechizos(Spell).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateCharParticle(UserList(userindex).Char.CharIndex, Hechizos(Spell).Particle))
          
          UserList(userindex).flags.Estupidez = 1
          UserList(userindex).Counters.Ceguera = IntervaloInvisible
                  
        Call WriteDumb(userindex)
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageChatOverHead(Hechizos(Spell).PalabrasMagicas, Npclist(NpcIndex).Char.CharIndex, RGB(128, 128, 0)))
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



Function TieneHechizo(ByVal i As Integer, ByVal userindex As Integer) As Boolean

On Error GoTo Errhandler
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(userindex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
Errhandler:

End Function

Sub AgregarHechizo(ByVal userindex As Integer, ByVal Slot As Integer)
Dim hIndex As Integer
Dim j As Integer
hIndex = ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).HechizoIndex


If ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).CPO <> "" Then
If UCase$(ListaClases(UserList(userindex).Clase)) <> ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).CPO Then
Call WriteConsoleMsg(1, userindex, "No puedes comprender el hechizo.", FontTypeNames.FONTTYPE_INFO)
Exit Sub
End If
End If

If Not TieneHechizo(hIndex, userindex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(userindex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
        
    If UserList(userindex).Stats.UserHechizos(j) <> 0 Then
        Call WriteConsoleMsg(1, userindex, "No tenes espacio para mas hechizos.", FontTypeNames.FONTTYPE_INFO)
    Else
        UserList(userindex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, userindex, CByte(j))
        'Quitamos del inv el item
        Call QuitarUserInvItem(userindex, CByte(Slot), 1)
    End If
Else
    Call WriteConsoleMsg(1, userindex, "Ya tenes ese hechizo.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub
            
Sub DecirPalabrasMagicas(ByVal S As String, ByVal userindex As Integer)
On Error GoTo hayerror
    'Mannakia
    If UserList(userindex).Invent.MagicIndex <> 0 Then
        If ObjData(UserList(userindex).Invent.MagicIndex).EfectoMagico = eMagicType.Silencio Then
            Exit Sub
        End If
    End If
    'Mannakia
    Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageChatOverHead(S, UserList(userindex).Char.CharIndex, RGB(128, 128, 0)))
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
Function PuedeLanzar(ByVal userindex As Integer, ByVal HechizoIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 11/09/08
'Last Modification By: Marco Vanotti (Marco)
' - 11/09/08 Now Druid have mana bonus while casting summoning spells having a magic flute equipped (Marco)
'***************************************************
Dim DruidManaBonus As Single

    If HechizoIndex = 0 Then Exit Function
    
    If UserList(userindex).flags.Muerto Then
        Call WriteMsg(userindex, 6)
        PuedeLanzar = False
        Exit Function
    End If
        
    If UserList(userindex).Stats.UserSkills(eSkill.Magia) < Hechizos(HechizoIndex).MinSkill Then
        Call WriteConsoleMsg(1, userindex, "No tenes suficientes puntos de magia para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
        PuedeLanzar = False
        Exit Function
    End If
    
    If UserList(userindex).Stats.MinSTA < Hechizos(HechizoIndex).StaRequerido Then
        If UserList(userindex).Genero = eGenero.Hombre Then
            Call WriteConsoleMsg(1, userindex, "Estás muy cansado para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(1, userindex, "Estás muy cansada para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
        End If
        PuedeLanzar = False
        Exit Function
    End If
    
    If UserList(userindex).Stats.MinMAN < Hechizos(HechizoIndex).ManaRequerido Then
        Call WriteConsoleMsg(1, userindex, "No tenes suficiente mana.", FontTypeNames.FONTTYPE_INFO)
        PuedeLanzar = False
        Exit Function
    End If
        
    PuedeLanzar = True
End Function

Sub HechizoTerrenoEstado(ByVal userindex As Integer, ByRef b As Boolean)
    Dim PosCasteadaX As Integer
    Dim PosCasteadaY As Integer
    Dim PosCasteadaM As Integer
    Dim h As Integer
    Dim TempX As Integer
    Dim TempY As Integer
    Dim TargetUser As Integer
    Dim TargetNPC As Integer
    Dim Daño As Long

    PosCasteadaX = UserList(userindex).flags.TargetX
    PosCasteadaY = UserList(userindex).flags.TargetY
    PosCasteadaM = UserList(userindex).flags.TargetMap
    
    'Distribucion de daño
    'Daño = Porcentaje(Daño, 40 + (100 / Abs(TempX - PosCasteadaX)))
    
    h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
    
    If Hechizos(h).HechizoDeArea Then
        If Hechizos(h).AreaEfecto <> 0 Then
            b = True
            For TempX = PosCasteadaX - Hechizos(h).AreaEfecto To PosCasteadaX + Hechizos(h).AreaEfecto
                For TempY = PosCasteadaY - Hechizos(h).AreaEfecto To PosCasteadaY + Hechizos(h).AreaEfecto
                    If InMapBounds(PosCasteadaM, TempX, TempY) Then
                        TargetUser = MapData(PosCasteadaM, TempX, TempY).userindex
                        If TargetUser > 0 And Not TargetUser = userindex Then
                            If UserList(TargetUser).flags.Muerto = 0 Then
                                If Hechizos(h).SubeHP = 1 Then
                                    If Not PuedeAyudar(userindex, TargetUser) Then
                                           
                                        Daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
                                        Daño = Daño + Porcentaje(Daño, 3 * UserList(userindex).Stats.ELV)
                                        Daño = Porcentaje(Daño, 40 + (100 / Abs(TempX - PosCasteadaX)))
                                        If Daño < 0 Then Daño = 0
                                        
                                        UserList(TargetUser).Stats.MinHP = UserList(TargetUser).Stats.MinHP + Daño
                                        If UserList(TargetUser).Stats.MinHP > UserList(TargetUser).Stats.MaxHP Then _
                                            UserList(TargetUser).Stats.MinHP = UserList(TargetUser).Stats.MaxHP
                                        
                                        Call WriteUpdateHP(TargetUser)
                                        
                                        Call WriteConsoleMsg(1, TargetUser, UserList(userindex).Name & " te ha restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
                                    End If
                                ElseIf Hechizos(h).SubeHP = 2 Then 'Complertar
                                    If TriggerZonaPelea(userindex, TargetUser) Then
                                        Daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
                                        
                                        Daño = Daño + Porcentaje(Daño, 2 * UserList(userindex).Stats.ELV)
                                        
                                        'Baculos DM + X
                                        If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                                            If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).EfectoMagico = eMagicType.DañoMagico Then
                                                Daño = Daño + (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).CuantoAumento)
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
                                        
                                        If Not PuedeAtacar(userindex, TargetUser) Then Exit Sub
                                        
                                        If userindex <> TargetUser Then
                                            Call UsuarioAtacadoPorUsuario(userindex, TargetUser)
                                        End If
                                        
        
                                        UserList(TargetUser).Stats.MinHP = UserList(TargetUser).Stats.MinHP - Daño
                                        
                                        Call WriteUpdateHP(TargetUser)

                                        'Muere
                                        If UserList(TargetUser).Stats.MinHP < 1 Then
                                            Call ContarMuerte(TargetUser, userindex)
                                            UserList(TargetUser).Stats.MinHP = 0
                                            Call ActStats(TargetUser, userindex)
                                            Call UserDie(TargetUser)
                                        End If
                                        b = True
                                    End If
                                    
                                    If Hechizos(h).Envenena > 0 Then
                                        Call UsuarioAtacadoPorUsuario(userindex, TargetUser)
                                        UserList(TargetUser).flags.Envenenado = Hechizos(h).Envenena
                                    End If
                                End If
                            End If
                        End If
                        TargetNPC = MapData(PosCasteadaM, TempX, TempY).NpcIndex
                        If TargetNPC <> 0 Then
                            If PuedeAtacarNPC(userindex, TargetNPC) Then

                                Call NPCAtacado(TargetNPC, userindex)
                                Daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
                                Daño = Daño + Porcentaje(Daño, 3 * UserList(userindex).Stats.ELV)
                                Daño = Porcentaje(Daño, 40 + (100 / IIf(Abs(TempX - PosCasteadaX) = 0, 1, Abs(TempX - PosCasteadaX))))
                                
                                'Baculos DM + X
                                If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                                    If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).EfectoMagico = eMagicType.DañoMagico Then
                                        Daño = Daño + (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).CuantoAumento)
                                    End If
                                End If
    
                                If Npclist(TargetNPC).flags.Snd2 > 0 Then
                                    Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Npclist(TargetNPC).flags.Snd2, Npclist(TargetNPC).Pos.x, Npclist(TargetNPC).Pos.Y))
                                End If
                            
                                'Quizas tenga defenza magica el NPC. Pablo (ToxicWaste)
                                Daño = Daño - Npclist(TargetNPC).Stats.defM
                                If Daño < 0 Then Daño = 0
                            
                                Npclist(TargetNPC).Stats.MinHP = Npclist(TargetNPC).Stats.MinHP - Daño
                                Call CalcularDarExp(userindex, TargetNPC, Daño)
                            
                                If Npclist(TargetNPC).IsFamiliar Then
                                    If Npclist(TargetNPC).MaestroUser > 0 Then
                                        UpdateFamiliar Npclist(TargetNPC).MaestroUser, False
                                    End If
                                End If
                            
                                If Npclist(TargetNPC).Stats.MinHP < 1 Then
                                    Npclist(TargetNPC).Stats.MinHP = 0
                                    Call MuereNpc(TargetNPC, userindex)
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
                    If MapData(PosCasteadaM, TempX, TempY).userindex > 0 Then
                        'hay un user
                        If UserList(MapData(PosCasteadaM, TempX, TempY).userindex).flags.Invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).userindex).flags.AdminInvisible = 0 Then
                            If Hechizos(h).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(UserList(MapData(PosCasteadaM, TempX, TempY).userindex).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))
                            If Hechizos(h).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateCharParticle(UserList(MapData(PosCasteadaM, TempX, TempY).userindex).Char.CharIndex, Hechizos(h).Particle))
                        End If
                    End If
                End If
            Next TempY
        Next TempX
    
        Call InfoHechizo(userindex)
    End If
    
    If Hechizos(h).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateParticle(PosCasteadaX, PosCasteadaY, Hechizos(h).Particle))
    If Hechizos(h).FXgrh <> 0 Then _
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFXMap(PosCasteadaX, PosCasteadaY, Hechizos(h).FXgrh, IIf(Hechizos(h).loops < 1, 1, Hechizos(h).loops)))
    If Hechizos(h).WAV <> 0 Then Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(Hechizos(h).WAV, PosCasteadaX, PosCasteadaY))

End Sub

''
' Le da propiedades al nuevo npc
'
' @param UserIndex  Indice del usuario que invoca.
' @param b  Indica si se termino la operación.

Sub HechizoInvocacion(ByVal userindex As Integer, ByRef b As Boolean)
'***************************************************
'Author: Uknown
'Last modification: 06/15/2008 (NicoNZ)
'Sale del sub si no hay una posición valida.
'***************************************************
If UserList(userindex).NroMascotas >= MAXMASCOTAS Then Exit Sub

'No permitimos se invoquen criaturas en zonas seguras
If MapInfo(UserList(userindex).Pos.map).Pk = False Or MapData(UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.Y).Trigger = eTrigger.ZONASEGURA Then
    Call WriteConsoleMsg(1, userindex, "En zona segura no puedes invocar criaturas.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

Dim h As Integer, j As Integer, ind As Integer, index As Integer
Dim TargetPos As WorldPos


TargetPos.map = UserList(userindex).flags.TargetMap
TargetPos.x = UserList(userindex).flags.TargetX
TargetPos.Y = UserList(userindex).flags.TargetY

h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
    
For j = 1 To Hechizos(h).Cant
    
    If UserList(userindex).NroMascotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(h).NumNpc, TargetPos, True, False)
        If ind > 0 Then
            UserList(userindex).NroMascotas = UserList(userindex).NroMascotas + 1
            
            index = FreeMascotaIndex(userindex)
            
            UserList(userindex).MascotasIndex(index) = ind
            UserList(userindex).MascotasType(index) = Npclist(ind).Numero
            
            Npclist(ind).MaestroUser = userindex
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

Call InfoHechizo(userindex)
b = True


End Sub

Sub HandleHechizoTerreno(ByVal userindex As Integer, ByVal uh As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 05/01/08
'
'***************************************************
If UserList(userindex).flags.ModoCombate = False Then
    Call WriteConsoleMsg(1, userindex, "Debes estar en modo de combate para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uInvocacion
        Call HechizoInvocacion(userindex, b)
    Case TipoHechizo.uEstado, TipoHechizo.uPropiedades
        Call HechizoTerrenoEstado(userindex, b)
    Case TipoHechizo.uCreateTelep
        Call HechizoCreateTelep(userindex, b)
   Case TipoHechizo.uMaterializa
      Call HechizoMaterializa(userindex, b)
    Case TipoHechizo.uFamiliar
        Call HechizoFamiliar(userindex, b)
End Select

If b Then
    Call SubirSkill(userindex, Magia)
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido

    If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
    UserList(userindex).Stats.MinSTA = UserList(userindex).Stats.MinSTA - Hechizos(uh).StaRequerido
    If UserList(userindex).Stats.MinSTA < 0 Then UserList(userindex).Stats.MinSTA = 0
    Call WriteUpdateUserStats(userindex)
End If


End Sub
Sub HechizoFamiliar(ByVal userindex As Integer, ByVal uh As Boolean)
'***************************************************
'Author: Mannakia
'Last Modification: 25/09/10
'
'***************************************************
Dim TargetPos As WorldPos
Dim ind As Integer
'WTF como hizo para tener este hechi ?
If UserList(userindex).masc.TieneFamiliar = 0 Then
    Exit Sub
End If

TargetPos.map = UserList(userindex).Pos.map
TargetPos.x = UserList(userindex).flags.TargetX
TargetPos.Y = UserList(userindex).flags.TargetY

'Desinvocamos
If UserList(userindex).masc.invocado = True Then

'Castelli
Call desinvocarfami(userindex)
'Castelli

Else 'Invocamos
    If UserList(userindex).masc.MinHP > 0 Then
        ind = SpawnNpc(npcFamiToTipe(UserList(userindex).masc.Tipo), TargetPos, True, False)
        If ind > 0 Then
            
        
            Npclist(ind).MaestroUser = userindex
            Npclist(ind).IsFamiliar = True
            
            Call SendData(SendTarget.ToNPCArea, ind, PrepareMessageCreateParticle(TargetPos.x, TargetPos.Y, 116))
            
            UserList(userindex).masc.NpcIndex = ind
            
            Call UpdateFamiliar(userindex, True)
            
            Call FollowAmo(ind)
            
            Npclist(ind).Movement = TipoAI.NpcFamiliar
            
            UserList(userindex).masc.invocado = True
            
            'Actualizamos las habilidades
            If UserList(userindex).Clase = eClass.Druida Or _
                UserList(userindex).Clase = eClass.Cazador Or _
                UserList(userindex).Clase = eClass.Mago Then
                    
                If UserList(userindex).masc.ELV >= 10 Then
                    If UserList(userindex).masc.Tipo = eMascota.Ely Then
                        UserList(userindex).masc.Curar = 1
                    ElseIf UserList(userindex).masc.Tipo = eMascota.Fatuo Then
                        UserList(userindex).masc.Misil = 1
                    End If
                End If
                
                If UserList(userindex).masc.ELV >= 15 Then
                    If UserList(userindex).masc.Tipo = eMascota.Ely Or UserList(userindex).masc.Tipo = eMascota.Fatuo Then
                        UserList(userindex).masc.Inmoviliza = 1
                    ElseIf UserList(userindex).masc.Tipo = eMascota.Lobo Or UserList(userindex).masc.Tipo = eMascota.Tigre Then
                        UserList(userindex).masc.gEntorpece = 1
                    ElseIf UserList(userindex).masc.Tipo = eMascota.Ent Then
                        UserList(userindex).masc.gEnvenena = 1
                    End If
                End If
                
                If UserList(userindex).masc.ELV >= 20 Then
                    If UserList(userindex).masc.Tipo = eMascota.Tigre Or UserList(userindex).masc.Tipo = eMascota.Ent Or UserList(userindex).masc.Tipo = eMascota.Lobo Then
                        UserList(userindex).masc.gParaliza = 1
                    ElseIf UserList(userindex).masc.Tipo = eMascota.Ely Or UserList(userindex).masc.Tipo = eMascota.Fatuo Then
                        UserList(userindex).masc.Descargas = 1
                    ElseIf UserList(userindex).masc.Tipo = eMascota.Fuego Then
                        UserList(userindex).masc.Tormentas = 1
                    ElseIf UserList(userindex).masc.Tipo = eMascota.Agua Then
                        UserList(userindex).masc.Paraliza = 1
                    ElseIf UserList(userindex).masc.Tipo = eMascota.Tierra Then
                        UserList(userindex).masc.Inmoviliza = 1
                    End If
                End If
                
                If UserList(userindex).masc.ELV >= 30 Then
                    If UserList(userindex).masc.Tipo = eMascota.Ely Then
                        UserList(userindex).masc.Desencanta = 1
                    ElseIf UserList(userindex).masc.Tipo = eMascota.Fatuo Then
                        UserList(userindex).masc.DetecInvi = 1
                    ElseIf UserList(userindex).masc.Tipo = eMascota.Oso Or UserList(userindex).masc.Tipo = eMascota.Ent Then
                        UserList(userindex).masc.gDesarma = 1
                    ElseIf UserList(userindex).masc.Tipo = eMascota.Tigre Or UserList(userindex).masc.Tipo = eMascota.Lobo Then
                        UserList(userindex).masc.gEnseguece = 1
                    End If
                End If
            End If
        End If
    Else
        Call WriteConsoleMsg(1, userindex, "Tu familiar esta muerto ¡¡Puedes llevarlo a un sacerdorte para resucitarlo!!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
End If
End Sub
Sub UpdateFamiliar(ByVal userindex As Integer, ByVal flag As Boolean)
    Dim ind As Integer
    ind = UserList(userindex).masc.NpcIndex
    
    If flag = True Then
        If ind > 0 Then
            ind = UserList(userindex).masc.NpcIndex
            
            Npclist(ind).Stats.MaxHit = UserList(userindex).masc.MaxHit
            Npclist(ind).Stats.MinHit = UserList(userindex).masc.MinHit
            
            Npclist(ind).Stats.MaxHP = UserList(userindex).masc.MaxHP
            Npclist(ind).Stats.MinHP = UserList(userindex).masc.MinHP
            
            Npclist(ind).Name = UserList(userindex).masc.Nombre
        End If
    Else
        If ind > 0 Then
            UserList(userindex).masc.MinHP = Npclist(ind).Stats.MinHP
            
            If UserList(userindex).masc.MinHP <= 0 Then
                UserList(userindex).masc.MinHP = 0
                UserList(userindex).masc.invocado = False
            End If
        End If
    End If
End Sub
Sub CheckFamiLevel(ByVal userindex As Integer)
    With UserList(userindex)
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
                Select Case .masc.Tipo
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
                
                Select Case .masc.Tipo
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
Sub HandleHechizoUsuario(ByVal userindex As Integer, ByVal uh As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 29/10/10
'BY: Jose Ignacio Castelli
'***************************************************


If UserList(userindex).flags.ModoCombate = False Then
    Call WriteConsoleMsg(1, userindex, "Debes estar en modo de combate para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

Dim b As Boolean

Select Case Hechizos(uh).Tipo

    Case TipoHechizo.uEstado, TipoHechizo.uPropEsta, TipoHechizo.uPropiedades ' Afectan estados (por ejem : Envenenamiento)
      
       Call HechizoEsUsuario(userindex, b)
       
    Case TipoHechizo.uCreateMagic
        Call HechizoCreateMagic(userindex, b)
End Select

If b Then
    Call SubirSkill(userindex, Magia)

    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido

    If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
    UserList(userindex).Stats.MinSTA = UserList(userindex).Stats.MinSTA - Hechizos(uh).StaRequerido
    If UserList(userindex).Stats.MinSTA < 0 Then UserList(userindex).Stats.MinSTA = 0
    Call WriteUpdateUserStats(userindex)

    If Hechizos(uh).AutoLanzar = 0 Then
        Call WriteUpdateUserStats(UserList(userindex).flags.TargetUser)
    End If

    UserList(userindex).flags.TargetUser = 0
End If

End Sub
Sub HechizoCreateMagic(ByVal userindex As Integer, b As Boolean)
'***************************************************
'Author: Leandro Mendoza
'Last Modification: 11/12/2010

'***************************************************
Dim h As Integer

b = False
h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
If h = 0 Then Exit Sub

With UserList(userindex)
    If Hechizos(h).CreaAlgo = 1 Then 'Crea arma magica
        If .flags.Muerto <> 0 Then
            Call WriteMsg(userindex, 1)
            Exit Sub
        End If
        
        .Stats.eMaxHit = Hechizos(h).MaxHit
        .Stats.eMinHit = Hechizos(h).MinHit
  
        
        .Stats.eCreateTipe = 1
        
        If Hechizos(h).Particle <> 0 Then _
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateCharParticle(.Char.CharIndex, Hechizos(h).Particle))
        
        If Hechizos(h).FXgrh <> 0 Then _
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).WAV))
        
        If Hechizos(h).WAV <> 0 Then _
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(Hechizos(h).WAV, .Pos.x, .Pos.x))
        
        b = True
    ElseIf Hechizos(h).CreaAlgo = 2 Then 'Crea aura sagrada
        If .flags.Muerto <> 0 Then
            Call WriteMsg(userindex, 1)
            Exit Sub
        End If
        
        .Stats.eMaxDef = Hechizos(h).MaxDef
        .Stats.eMinDef = Hechizos(h).MinDef
        
        
        .Stats.eCreateTipe = 2
        
        If Hechizos(h).Particle <> 0 Then _
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateCharParticle(.Char.CharIndex, Hechizos(h).Particle))
        
        If Hechizos(h).FXgrh <> 0 Then _
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).WAV))
        
        If Hechizos(h).WAV <> 0 Then _
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(Hechizos(h).WAV, .Pos.x, .Pos.x))
        
        b = True
    ElseIf Hechizos(h).CreaAlgo = 3 Then 'Menos defensa
        If .flags.Muerto <> 0 Then
            Call WriteMsg(userindex, 1)
            Exit Sub
        End If
        
        .Stats.dMaxDef = Hechizos(h).MaxDef
        .Stats.dMinDef = Hechizos(h).MinDef
        
        
        .Stats.eCreateTipe = 3
        
        If Hechizos(h).Particle <> 0 Then _
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateCharParticle(.Char.CharIndex, Hechizos(h).Particle))
        
        If Hechizos(h).FXgrh <> 0 Then _
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).WAV))
        
        If Hechizos(h).WAV <> 0 Then _
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(Hechizos(h).WAV, .Pos.x, .Pos.x))
        
        b = True
    End If
End With

End Sub

Sub HandleHechizoNPC(ByVal userindex As Integer, ByVal uh As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 13/02/2009
'***************************************************
Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
        Call HechizoEstadoNPC(UserList(userindex).flags.TargetNPC, uh, b, userindex)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
        Call HechizoPropNPC(uh, UserList(userindex).flags.TargetNPC, userindex, b)
End Select


If b Then
    Call SubirSkill(userindex, Magia)
    UserList(userindex).flags.TargetNPC = 0
    
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido

    If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
    UserList(userindex).Stats.MinSTA = UserList(userindex).Stats.MinSTA - Hechizos(uh).StaRequerido
    If UserList(userindex).Stats.MinSTA < 0 Then UserList(userindex).Stats.MinSTA = 0
    Call WriteUpdateUserStats(userindex)
End If

End Sub


Sub LanzarHechizo(index As Integer, userindex As Integer)

Dim uh As Integer

uh = UserList(userindex).Stats.UserHechizos(index)

' Dioses pueden tirar hechi sin mana
If PuedeLanzar(userindex, uh) _
Or (UserList(userindex).flags.Privilegios And (PlayerType.Dios Or PlayerType.VIP)) Then

    Select Case Hechizos(uh).Target
        Case TargetType.uUsuarios
        
            If UserList(userindex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(userindex).flags.TargetUser).Pos.Y - UserList(userindex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(userindex, uh)
                Else
                    Call WriteConsoleMsg(1, userindex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            Else
                Call WriteConsoleMsg(1, userindex, "Este hechizo actúa solo sobre usuarios.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Case TargetType.uNPC
            If UserList(userindex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(userindex).flags.TargetNPC).Pos.Y - UserList(userindex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(userindex, uh)
                Else
                    Call WriteConsoleMsg(1, userindex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            Else
                Call WriteConsoleMsg(1, userindex, "Este hechizo solo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Case TargetType.uUsuariosYnpc
       
            If UserList(userindex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(userindex).flags.TargetUser).Pos.Y - UserList(userindex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(userindex, uh)
                Else
                    Call WriteConsoleMsg(1, userindex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            ElseIf UserList(userindex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(userindex).flags.TargetNPC).Pos.Y - UserList(userindex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(userindex, uh)
                Else
                    Call WriteConsoleMsg(1, userindex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            Else
                Call WriteConsoleMsg(1, userindex, "Target invalido.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Case TargetType.uTerreno
            Call HandleHechizoTerreno(userindex, uh)
    End Select
    
End If

If UserList(userindex).Counters.Trabajando Then _
    UserList(userindex).Counters.Trabajando = UserList(userindex).Counters.Trabajando - 1

If UserList(userindex).Counters.Ocultando Then _
    UserList(userindex).Counters.Ocultando = UserList(userindex).Counters.Ocultando - 1
    
End Sub



Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal userindex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 07/07/2008
'Handles the Spells that afect the Stats of an NPC
'04/13/2008 NicoNZ - Guardias Faccionarios pueden ser
'removidos por users de su misma faccion.
'07/07/2008: NicoNZ - Solo se puede mimetizar con npcs si es druida
'***************************************************
If Hechizos(hIndex).Invisibilidad = 1 Then
    Call InfoHechizo(userindex)
    Npclist(NpcIndex).flags.Invisible = 1
    b = True
End If

If Hechizos(hIndex).Envenena > 0 Then
    If Not PuedeAtacarNPC(userindex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    Call NPCAtacado(NpcIndex, userindex)
    Call InfoHechizo(userindex)
    Npclist(NpcIndex).flags.Envenenado = Hechizos(hIndex).Envenena
    b = True
End If

If Hechizos(hIndex).CuraVeneno = 1 Then
    Call InfoHechizo(userindex)
    Npclist(NpcIndex).flags.Envenenado = 0
    b = True
End If

If Hechizos(hIndex).Paraliza = 1 Then
    If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        If Not PuedeAtacarNPC(userindex, NpcIndex) Then
            b = False
            Exit Sub
        End If
        Call NPCAtacado(NpcIndex, userindex)
        Call InfoHechizo(userindex)
        Npclist(NpcIndex).flags.Paralizado = 1
        Npclist(NpcIndex).flags.Inmovilizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        b = True
    Else
        Call WriteConsoleMsg(1, userindex, "El NPC es inmune a este hechizo.", FontTypeNames.FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
End If

If Hechizos(hIndex).RemoverParalisis = 1 Then
    If Npclist(NpcIndex).flags.Paralizado = 1 Or Npclist(NpcIndex).flags.Inmovilizado = 1 Then
        If Npclist(NpcIndex).MaestroUser = userindex Then
            Call InfoHechizo(userindex)
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
            b = True
        Else
            If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                If esArmada(userindex) Then
                    Call InfoHechizo(userindex)
                    Npclist(NpcIndex).flags.Paralizado = 0
                    Npclist(NpcIndex).Contadores.Paralisis = 0
                    b = True
                    Exit Sub
                Else
                    Call WriteConsoleMsg(1, userindex, "Solo puedes Remover la Parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
                
                Call WriteConsoleMsg(1, userindex, "Solo puedes Remover la Parálisis de los NPCs que te consideren su amo", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            Else
                If Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos Then
                    If esCaos(userindex) Then
                        Call InfoHechizo(userindex)
                        Npclist(NpcIndex).flags.Paralizado = 0
                        Npclist(NpcIndex).Contadores.Paralisis = 0
                        b = True
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(1, userindex, "Solo puedes Remover la Parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    End If
                End If
            End If
        End If
   Else
      Call WriteConsoleMsg(1, userindex, "Este NPC no esta Paralizado", FontTypeNames.FONTTYPE_INFO)
      b = False
      Exit Sub
   End If
End If
 
If Hechizos(hIndex).Inmoviliza = 1 Then
    If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        If Not PuedeAtacarNPC(userindex, NpcIndex) Then
            b = False
            Exit Sub
        End If
        Call NPCAtacado(NpcIndex, userindex)
        Npclist(NpcIndex).flags.Inmovilizado = 1
        Npclist(NpcIndex).flags.Paralizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        Call InfoHechizo(userindex)
        b = True
    Else
        Call WriteConsoleMsg(1, userindex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
    End If
End If

End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal userindex As Integer, ByRef b As Boolean)
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
    Daño = Daño + Porcentaje(Daño, 3 * UserList(userindex).Stats.ELV)
    
    Call InfoHechizo(userindex)
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP + Daño
    If Npclist(NpcIndex).Stats.MinHP > Npclist(NpcIndex).Stats.MaxHP Then _
        Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
    Call WriteConsoleMsg(1, userindex, "Has curado " & Daño & " puntos de salud a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
    b = True
    
ElseIf Hechizos(hIndex).SubeHP = 2 Then
    If Not PuedeAtacarNPC(userindex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    Call NPCAtacado(NpcIndex, userindex)
    Daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    Daño = Daño + Porcentaje(Daño, 3 * UserList(userindex).Stats.ELV)

    'Baculos DM + X
    If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).EfectoMagico = eMagicType.DañoMagico Then
            Daño = Daño + (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).CuantoAumento)
        End If
    End If

    Call InfoHechizo(userindex)
    b = True
    
    If Npclist(NpcIndex).flags.Snd2 > 0 Then
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y))
    End If
    
    'Quizas tenga defenza magica el NPC. Pablo (ToxicWaste)
    Daño = Daño - Npclist(NpcIndex).Stats.defM
    If Daño < 0 Then Daño = 0
    
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - Daño
    Call WriteConsoleMsg(2, userindex, "¡Le has causado " & Daño & " puntos de daño a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
    Call CalcularDarExp(userindex, NpcIndex, Daño)
    
    If Npclist(NpcIndex).IsFamiliar Then
        If Npclist(NpcIndex).MaestroUser > 0 Then
            UpdateFamiliar Npclist(NpcIndex).MaestroUser, False
        End If
    End If
    
    If Npclist(NpcIndex).Stats.MinHP < 1 Then
        Npclist(NpcIndex).Stats.MinHP = 0
        Call MuereNpc(NpcIndex, userindex)
    End If
End If

End Sub

Sub InfoHechizo(ByVal userindex As Integer)


    Dim h As Integer
    h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
    
    
    Call DecirPalabrasMagicas(Hechizos(h).PalabrasMagicas, userindex)
    
    If UserList(userindex).flags.TargetUser > 0 Then
        If Hechizos(h).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserList(userindex).flags.TargetUser, PrepareMessageCreateFX(UserList(UserList(userindex).flags.TargetUser).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))
        Call SendData(SendTarget.ToPCArea, UserList(userindex).flags.TargetUser, PrepareMessagePlayWave(Hechizos(h).WAV, UserList(UserList(userindex).flags.TargetUser).Pos.x, UserList(UserList(userindex).flags.TargetUser).Pos.Y))  'Esta linea faltaba. Pablo (ToxicWaste)
        If Hechizos(h).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserList(userindex).flags.TargetUser, PrepareMessageCreateCharParticle(UserList(UserList(userindex).flags.TargetUser).Char.CharIndex, Hechizos(h).Particle))
    ElseIf UserList(userindex).flags.TargetNPC > 0 Then
        If Hechizos(h).FXgrh <> 0 Then Call SendData(SendTarget.ToNPCArea, UserList(userindex).flags.TargetNPC, PrepareMessageCreateFX(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))
        Call SendData(SendTarget.ToNPCArea, UserList(userindex).flags.TargetNPC, PrepareMessagePlayWave(Hechizos(h).WAV, Npclist(UserList(userindex).flags.TargetNPC).Pos.x, Npclist(UserList(userindex).flags.TargetNPC).Pos.Y))
        If Hechizos(h).Particle <> 0 Then Call SendData(SendTarget.ToNPCArea, UserList(userindex).flags.TargetNPC, PrepareMessageCreateCharParticle(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex, Hechizos(h).Particle))
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

Sub HechizoEsUsuario(ByVal userindex As Integer, ByRef b As Boolean)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 02/01/2008
'02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
'***************************************************

Dim h As Integer
Dim Daño As Long
Dim tU As Integer

h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
tU = UserList(userindex).flags.TargetUser
b = False

If UserList(userindex).flags.Muerto = 1 Then
    Call WriteMsg(userindex, 0)
    b = False
    Exit Sub
End If

If Hechizos(h).ReviveFamiliar = 1 Then
    If Not PuedeAyudar(userindex, tU) Then
        Call WriteMsg(userindex, 33)
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
            Call WriteConsoleMsg(1, userindex, "¡Revivir no está permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        
        'Para poder tirar revivir a un pk en el ring
        If (TriggerZonaPelea(userindex, tU) <> TRIGGER6_PERMITE) Then
            If Not PuedeAyudar(userindex, tU) Then
                Call WriteMsg(userindex, 33)
                b = False
                Exit Sub
            End If
        End If

        DarVida tU
        
        UserList(tU).Stats.MinAGU = 100
        UserList(tU).flags.Sed = 0
        UserList(tU).Stats.MinHAM = 100
        UserList(tU).flags.Hambre = 0
        
        UserList(tU).Stats.MinHP = UserList(tU).Stats.MaxHP
        UserList(tU).Stats.MinMAN = UserList(tU).Stats.MaxMAN
        UserList(tU).Stats.MinSTA = UserList(tU).Stats.MaxSTA
        
        Call WriteUpdateUserStats(tU)
        
        Call InfoHechizo(userindex)
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
            Call WriteConsoleMsg(1, userindex, "¡Revivir no está permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        
        'Para poder tirar revivir a un pk en el ring
        If (TriggerZonaPelea(userindex, tU) <> TRIGGER6_PERMITE) Then
            If Not PuedeAyudar(userindex, tU) Then
                Call WriteMsg(userindex, 33)
                b = False
                Exit Sub
            End If
        End If

        UserList(tU).Stats.MinAGU = 0
        UserList(tU).flags.Sed = 1
        UserList(tU).Stats.MinHAM = 0
        UserList(tU).flags.Hambre = 1
        UserList(tU).Stats.MinMAN = 0
        UserList(tU).Stats.MinSTA = 0
                    
        Call RevivirUsuario(tU)
        
        Call WriteUpdateHungerAndThirst(tU)
        Call InfoHechizo(userindex)
        b = True
        Exit Sub
    Else
        b = False
    End If

End If

If UserList(tU).flags.Muerto Then
    Call WriteMsg(userindex, 28)
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
    
    b = True
End If

If Hechizos(h).Certero = 1 Then
    UserList(userindex).flags.NoFalla = 1
    
    If Hechizos(h).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, tU, PrepareMessageCreateCharParticle(UserList(tU).Char.CharIndex, Hechizos(h).Particle))
    If Hechizos(h).WAV <> 0 Then Call SendData(SendTarget.ToPCArea, tU, PrepareMessagePlayWave(Hechizos(h).WAV, UserList(tU).Pos.x, UserList(tU).Pos.Y))
    b = True
End If

If Hechizos(h).Desencantar Then
    
    If UserList(tU).flags.Incinerado = 1 Then _
        UserList(tU).flags.Incinerado = 0
    
    If UserList(tU).flags.Envenenado Then _
        UserList(tU).flags.Envenenado = 0

    If UserList(tU).flags.Estupidez = 1 Then
        UserList(tU).flags.Estupidez = 0
        Call WriteDumb(tU)
    End If

    If UserList(tU).flags.Metamorfosis = 1 Then
        UserList(userindex).Char.Head = UserList(userindex).OrigChar.Head
        If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(userindex).Char.body = ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(userindex)
        End If
        If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then _
            UserList(userindex).Char.ShieldAnim = ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then _
            UserList(userindex).Char.WeaponAnim = ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then _
            UserList(userindex).Char.CascoAnim = ObjData(UserList(userindex).Invent.CascoEqpObjIndex).CascoAnim
     
        UserList(userindex).flags.Metamorfosis = 0
        UserList(userindex).Counters.Metamorfosis = 0
     
        Call ChangeUserChar(userindex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
    End If
    
    If UserList(tU).flags.Ceguera = 1 Then
        UserList(userindex).flags.Ceguera = 0
        Call WriteBlindNoMore(tU)
    End If
    
    If Hechizos(h).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, tU, PrepareMessageCreateCharParticle(UserList(tU).Char.CharIndex, Hechizos(h).Particle))
    If Hechizos(h).WAV <> 0 Then Call SendData(SendTarget.ToPCArea, tU, PrepareMessagePlayWave(Hechizos(h).WAV, UserList(tU).Pos.x, UserList(tU).Pos.Y))
End If

If Hechizos(h).Invisibilidad = 1 Then
    If UserList(tU).flags.Navegando = 1 Then
        Call WriteMsg(userindex, 29)
        b = False
        Exit Sub
    End If
    
    If UserList(tU).Counters.Saliendo Then
        If userindex <> tU Then
            Call WriteMsg(userindex, 30)
            b = False
            Exit Sub
        Else
            Call WriteMsg(userindex, 31)
            b = False
            Exit Sub
        End If
    End If
    
    'No usar invi mapas InviSinEfecto
    If MapInfo(UserList(tU).Pos.map).InviSinEfecto > 0 Then
        Call WriteMsg(userindex, 32)
        b = False
        Exit Sub
    End If
    
    'Para poder tirar invi a un pk en el ring
    If (TriggerZonaPelea(userindex, tU) <> TRIGGER6_PERMITE) Then
        If Not PuedeAyudar(userindex, tU) Then
            Call WriteMsg(userindex, 33)
            b = False
            Exit Sub
        End If
    End If

    UserList(tU).flags.Invisible = 1
    Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.CharIndex, True))

    b = True
End If

If Hechizos(h).Envenena > 0 Then
    If userindex = tU Then
        Call WriteMsg(userindex, 34)
        b = False
        Exit Sub
    End If
    
    If Not PuedeAtacar(userindex, tU) Then Exit Sub
    If userindex <> tU Then
        Call UsuarioAtacadoPorUsuario(userindex, tU)
    End If
    UserList(tU).flags.Envenenado = Hechizos(h).Envenena
    b = True
End If

If Hechizos(h).Incinera = 1 Then
    If userindex = tU Then
        Call WriteMsg(userindex, 34)
        b = False
        Exit Sub
    End If
    
    If Not PuedeAtacar(userindex, tU) Then Exit Sub
    If userindex <> tU Then
        Call UsuarioAtacadoPorUsuario(userindex, tU)
    End If
    UserList(tU).flags.Incinerado = 1
    b = True
End If

If Hechizos(h).CuraVeneno = 1 Then
    'Para poder tirar curar veneno a un pk en el ring
    If (TriggerZonaPelea(userindex, tU) <> TRIGGER6_PERMITE) Then
        If Not PuedeAyudar(userindex, tU) Then
            Call WriteMsg(userindex, 33)
            b = False
            Exit Sub
        End If
    End If
        
    'Si sos user, no uses este hechizo con GMS.
    If UserList(userindex).flags.Privilegios And (PlayerType.User Or PlayerType.VIP) Then
        If UserList(tU).flags.Privilegios And (PlayerType.Dios) Then
            Exit Sub
        End If
    End If
        
    UserList(tU).flags.Envenenado = 0
    b = True
End If

If Hechizos(h).Paraliza = 1 Or Hechizos(h).Inmoviliza = 1 Then
    If userindex = tU Then
        Call WriteMsg(userindex, 34)
        Exit Sub
    End If
    
     If UserList(tU).flags.Paralizado = 0 Then
        If Not PuedeAtacar(userindex, tU) Then Exit Sub
            
        If userindex <> tU Then
            Call UsuarioAtacadoPorUsuario(userindex, tU)
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
        If (TriggerZonaPelea(userindex, tU) <> TRIGGER6_PERMITE) Then
            If Not PuedeAyudar(userindex, tU) Then
                Call WriteMsg(userindex, 33)
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
    If userindex = tU Then
        Call WriteMsg(userindex, 34)
        Exit Sub
    End If
    
    If Not PuedeAtacar(userindex, tU) Then Exit Sub
    
    If userindex <> tU Then
        Call UsuarioAtacadoPorUsuario(userindex, tU)
    End If
    UserList(tU).flags.Ceguera = 1
    UserList(tU).Counters.Ceguera = IntervaloParalizado / 3

    Call WriteBlind(tU)
    Call FlushBuffer(tU)
    b = True
End If

If Hechizos(h).Estupidez = 1 Then
    If userindex = tU Then
        Call WriteMsg(userindex, 34)
        Exit Sub
    End If
        
    If Not PuedeAtacar(userindex, tU) Then Exit Sub
    If userindex <> tU Then
        Call UsuarioAtacadoPorUsuario(userindex, tU)
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
    If UserList(userindex).flags.Montando = 1 Then
        Call WriteConsoleMsg(1, userindex, "Estas montando!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Hechizos(h).WAV <> 0 Then Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(Hechizos(h).WAV, UserList(userindex).flags.DondeTiroX, UserList(userindex).flags.DondeTiroY))
    If Hechizos(h).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(UserList(userindex).Char.CharIndex, Hechizos(h).FXgrh, 1))
 
    Call DoMetamorfosis(userindex, Hechizos(h).body, Hechizos(h).Head)
    b = True
End If


' <-------- Agilidad ---------->
If Hechizos(h).SubeAgilidad = 1 Then
    
    'Para poder tirar cl a un pk en el ring
    If (TriggerZonaPelea(userindex, tU) <> TRIGGER6_PERMITE) Then
        If Not PuedeAyudar(userindex, tU) Then
            Call WriteMsg(userindex, 34)
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
    
    If Not PuedeAtacar(userindex, tU) Then Exit Sub
    
    If userindex <> tU Then
        Call UsuarioAtacadoPorUsuario(userindex, tU)
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
    If (TriggerZonaPelea(userindex, tU) <> TRIGGER6_PERMITE) Then
        If Not PuedeAyudar(userindex, tU) Then
            Call WriteConsoleMsg(1, userindex, "No puedes beneficiar a ese tipo de gente.", FontTypeNames.FONTTYPE_INFO)
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

    If Not PuedeAtacar(userindex, tU) Then Exit Sub
    
    If userindex <> tU Then
        Call UsuarioAtacadoPorUsuario(userindex, tU)
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
    If (TriggerZonaPelea(userindex, tU) <> TRIGGER6_PERMITE) Then
        If Not PuedeAyudar(userindex, tU) Then
            Call WriteMsg(userindex, 34)
            b = False
            Exit Sub
        End If
    End If
       
    Daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
    Daño = Daño + Porcentaje(Daño, 3 * UserList(userindex).Stats.ELV)
    
    UserList(tU).Stats.MinHP = UserList(tU).Stats.MinHP + Daño
    If UserList(tU).Stats.MinHP > UserList(tU).Stats.MaxHP Then _
        UserList(tU).Stats.MinHP = UserList(tU).Stats.MaxHP
    
    Call WriteUpdateHP(tU)
    
    If userindex <> tU Then
        Call WriteConsoleMsg(2, userindex, "Le has restaurado " & Daño & " puntos de vida a " & UserList(tU).Name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(2, tU, UserList(userindex).Name & " te ha restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(2, userindex, "Te has restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    b = True
ElseIf Hechizos(h).SubeHP = 2 Then
    
    If userindex = tU Then
        Call WriteMsg(userindex, 34)
        Exit Sub
    End If
    
    Daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
    
    Daño = Daño + Porcentaje(Daño, 2 * UserList(userindex).Stats.ELV)
    
    'Baculos DM + X
    If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).EfectoMagico = eMagicType.DañoMagico Then
            Daño = Daño + (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).CuantoAumento)
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
    
    If Not PuedeAtacar(userindex, tU) Then Exit Sub
    
    If userindex <> tU Then
        Call UsuarioAtacadoPorUsuario(userindex, tU)
    End If
    
    UserList(tU).Stats.MinHP = UserList(tU).Stats.MinHP - Daño
    
    Call SubirSkill(tU, eSkill.Resistencia)
    
    Call WriteUpdateHP(tU)
    
    Call WriteMsg(userindex, 39, CStr(UserList(tU).Char.CharIndex), CStr(Daño))
    Call WriteMsg(tU, 38, CStr(UserList(userindex).Char.CharIndex), CStr(h))
    
    'Muere
    If UserList(tU).Stats.MinHP < 1 Then
        Call ContarMuerte(tU, userindex)
        UserList(tU).Stats.MinHP = 0
        Call ActStats(tU, userindex)
        Call UserDie(tU)
    End If
    
    b = True
End If

If b = True Then
    InfoHechizo userindex
End If
    
FlushBuffer userindex
FlushBuffer tU

End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal userindex As Integer, ByVal Slot As Byte)

'Call LogTarea("Sub UpdateUserHechizos")

Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(userindex).Stats.UserHechizos(Slot) > 0 Then
        Call ChangeUserHechizo(userindex, Slot, UserList(userindex).Stats.UserHechizos(Slot))
    Else
        Call ChangeUserHechizo(userindex, Slot, 0)
    End If

Else

'Actualiza todos los slots
For LoopC = 1 To MAXUSERHECHIZOS

        'Actualiza el inventario
        If UserList(userindex).Stats.UserHechizos(LoopC) > 0 Then
            Call ChangeUserHechizo(userindex, LoopC, UserList(userindex).Stats.UserHechizos(LoopC))
        Else
            Call ChangeUserHechizo(userindex, LoopC, 0)
        End If

Next LoopC

End If

End Sub

Sub ChangeUserHechizo(ByVal userindex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)

'Call LogTarea("ChangeUserHechizo")

UserList(userindex).Stats.UserHechizos(Slot) = Hechizo


If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
    
    Call WriteChangeSpellSlot(userindex, Slot)

Else

    Call WriteChangeSpellSlot(userindex, Slot)

End If


End Sub


Public Sub DesplazarHechizo(ByVal userindex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)

If (Dire <> 1 And Dire <> -1) Then Exit Sub
If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

Dim TempHechizo As Integer

If Dire = 1 Then 'Mover arriba
    If CualHechizo = 1 Then
        Call WriteConsoleMsg(1, userindex, "No puedes mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(userindex).Stats.UserHechizos(CualHechizo)
        UserList(userindex).Stats.UserHechizos(CualHechizo) = UserList(userindex).Stats.UserHechizos(CualHechizo - 1)
        UserList(userindex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo

        'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
        If UserList(userindex).flags.Hechizo > 0 Then
            UserList(userindex).flags.Hechizo = UserList(userindex).flags.Hechizo - 1
        End If
    End If
Else 'mover abajo
    If CualHechizo = MAXUSERHECHIZOS Then
        Call WriteConsoleMsg(1, userindex, "No puedes mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(userindex).Stats.UserHechizos(CualHechizo)
        UserList(userindex).Stats.UserHechizos(CualHechizo) = UserList(userindex).Stats.UserHechizos(CualHechizo + 1)
        UserList(userindex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo

        'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
        If UserList(userindex).flags.Hechizo > 0 Then
            UserList(userindex).flags.Hechizo = UserList(userindex).flags.Hechizo + 1
        End If
    End If
End If
End Sub


 Sub HechizoCreateTelep(userindex As Integer, b As Boolean)
  
    Dim tU As Integer
   Dim h As Integer
   Dim i As Integer
   Dim PosTIROTELEPORT As WorldPos
  
      h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
  
      If Hechizos(h).Nombre = "Portal planar" Then
        If UserList(userindex).flags.TiroPortalL = 1 Then
          Call WriteConsoleMsg(2, userindex, "Ya tienes un Portal creado.", FontTypeNames.FONTTYPE_INFO)
          Exit Sub
        End If
      End If
 
    
     PosTIROTELEPORT.x = UserList(userindex).flags.TargetX
     PosTIROTELEPORT.Y = UserList(userindex).flags.TargetY
     PosTIROTELEPORT.map = UserList(userindex).flags.TargetMap
         
    If Not LegalPos(PosTIROTELEPORT.map, PosTIROTELEPORT.x, PosTIROTELEPORT.Y) Then
        Exit Sub
    End If
    
    If MapInfo(PosTIROTELEPORT.map).Pk = False Then
        Exit Sub
    End If
         
     UserList(userindex).flags.DondeTiroMap = PosTIROTELEPORT.map
 
     UserList(userindex).flags.DondeTiroX = PosTIROTELEPORT.x
     UserList(userindex).flags.DondeTiroY = PosTIROTELEPORT.Y
    
     If MapData(UserList(userindex).Pos.map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).ObjInfo.ObjIndex Then
         Exit Sub
     End If
    
     If MapData(UserList(userindex).Pos.map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).TileExit.map Then
         Exit Sub
     End If
    
     If MapData(UserList(userindex).Pos.map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).Blocked Then
         Exit Sub
     End If
     If Not MapaValido(UserList(userindex).Pos.map) Or Not InMapBounds(UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY) Then Exit Sub
 
   Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(Hechizos(h).WAV, UserList(userindex).flags.DondeTiroX, UserList(userindex).flags.DondeTiroY))
   Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateParticle(UserList(userindex).flags.DondeTiroX, UserList(userindex).flags.DondeTiroY, Hechizos(h).Particle))

   UserList(userindex).Counters.TimeTeleport = 0
   UserList(userindex).Counters.CreoTeleport = True
   UserList(userindex).flags.TiroPortalL = 1
   Call InfoHechizo(userindex)
   
   b = True
  
 End Sub
 
Sub HechizoMaterializa(userindex As Integer, b As Boolean)
    Dim M As Integer
    Dim x As Integer
    Dim Y As Integer
    Dim Obj As Obj
    Dim h As Integer
    
    Obj.Amount = 1
    h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
      
    If Hechizos(h).Nombre = "Materializar: Comida" Then
         Obj.ObjIndex = 1
    ElseIf Hechizos(h).Nombre = "Materializar: Bebida" Then
         Obj.ObjIndex = 43
    End If

    'Exit Sub
    If MapData(UserList(userindex).Pos.map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).ObjInfo.ObjIndex Then
        Exit Sub
    End If
    
    If MapData(UserList(userindex).Pos.map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).TileExit.map Then
        Exit Sub
    End If
        
    If MapData(UserList(userindex).Pos.map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).Blocked Then
        Exit Sub
    End If
         
    If Not MapaValido(UserList(userindex).Pos.map) Or Not InMapBounds(UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY) Then Exit Sub
    
    M = UserList(userindex).flags.TargetMap
    x = UserList(userindex).flags.TargetX
    Y = UserList(userindex).flags.TargetY

    Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(Hechizos(h).WAV, x, Y))
    Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateParticle(x, Y, Hechizos(h).Particle))
       
    Call MakeObj(Obj, M, x, Y)
    
    b = True
End Sub

Public Sub desinvocarfami(userindex As Integer)
    'Por la dudas le mandamos un update para que llene
    Call UpdateFamiliar(userindex, False)
    
    Call SendData(SendTarget.ToNPCArea, UserList(userindex).masc.NpcIndex, PrepareMessageCreateParticle(Npclist(UserList(userindex).masc.NpcIndex).Pos.x, Npclist(UserList(userindex).masc.NpcIndex).Pos.Y, 117))
    
    UserList(userindex).masc.invocado = False
    Call QuitarNPC(UserList(userindex).masc.NpcIndex)
    
    UserList(userindex).masc.NpcIndex = 0
End Sub
