Attribute VB_Name = "SistemaCombate"


Option Explicit

Public Const MAXDISTANCIAARCO As Byte = 18
Public Const MAXDISTANCIAMAGIA As Byte = 18

Public Const NPC_DEMONIO As Integer = 1
Public Const ARCO_DEMONIO As Integer = 666

Public Function MinimoInt(ByVal A As Integer, ByVal b As Integer) As Integer
    If A > b Then
        MinimoInt = b
    Else
        MinimoInt = A
    End If
End Function

Public Function MaximoInt(ByVal A As Integer, ByVal b As Integer) As Integer
    If A > b Then
        MaximoInt = A
    Else
        MaximoInt = b
    End If
End Function

Private Function PoderEvasionEscudo(ByVal userindex As Integer) As Long
On Error GoTo err
    PoderEvasionEscudo = (UserList(userindex).Stats.UserSkills(eSkill.DefensaEscudos) * ModClase(UserList(userindex).Clase).Evasion) * 0.5
Exit Function
err:

End Function

Private Function PoderEvasion(ByVal userindex As Integer) As Long
On Error GoTo err

    Dim lTemp As Long
    With UserList(userindex)
        lTemp = (.Stats.UserSkills(eSkill.Tacticas) + _
          .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).Evasion
       
        PoderEvasion = (lTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
    
    Exit Function
err:
    
End Function

Private Function PoderAtaqueArma(ByVal userindex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With UserList(userindex)
        If .Stats.UserSkills(eSkill.armas) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.armas) * ModClase(.Clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.armas) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.armas) + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.armas) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.armas) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        Else
           PoderAtaqueTemp = (.Stats.UserSkills(eSkill.armas) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        End If
        
        PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Private Function PoderAtaqueProyectil(ByVal userindex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With UserList(userindex)
        If .Stats.UserSkills(eSkill.Proyectiles) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Proyectiles) * ModClase(.Clase).AtaqueProyectiles
        ElseIf .Stats.UserSkills(eSkill.Proyectiles) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueProyectiles
        ElseIf .Stats.UserSkills(eSkill.Proyectiles) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueProyectiles
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueProyectiles
        End If
        
        PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Private Function PoderAtaqueWrestling(ByVal userindex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With UserList(userindex)
        If .Stats.UserSkills(eSkill.artes) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.artes) * ModClase(.Clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.artes) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.artes) + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.artes) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.artes) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.artes) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).AtaqueArmas
        End If
        
        PoderAtaqueWrestling = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
End Function

Public Function UserImpactoNpc(ByVal userindex As Integer, ByVal NpcIndex As Integer) As Boolean
    Dim PoderAtaque As Long
    Dim Arma As Integer
    Dim Skill As eSkill
    Dim ProbExito As Long
    
    If UserList(userindex).flags.NoFalla = 1 Then
        UserList(userindex).flags.NoFalla = 0
        UserImpactoNpc = True
        Exit Function
    End If
    
    Arma = UserList(userindex).Invent.WeaponEqpObjIndex
    
    If UserList(userindex).Invent.NudiEqpIndex > 0 Then
        PoderAtaque = PoderAtaqueWrestling(userindex)
        Skill = eSkill.artes
    ElseIf Arma > 0 Then 'Usando un arma
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = PoderAtaqueProyectil(userindex)
            Skill = eSkill.Proyectiles
        ElseIf ObjData(Arma).SubTipo = 5 Or ObjData(Arma).SubTipo = 6 Then
            PoderAtaque = PoderAtaqueArma(userindex)
            Skill = eSkill.arrojadizas
        Else
            PoderAtaque = PoderAtaqueArma(userindex)
            Skill = eSkill.armas
        End If
    Else 'Peleando con puños
        PoderAtaque = PoderAtaqueWrestling(userindex)
        Skill = eSkill.artes
    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))
    If UserList(userindex).flags.Privilegios And PlayerType.Dios Then
        ProbExito = 95
    End If
    
    UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
    
    If UserImpactoNpc Then
        Call SubirSkill(userindex, Skill)
    End If
End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal userindex As Integer) As Boolean
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Revisa si un NPC logra impactar a un user o no
'03/15/2006 Maraxus - Evité una división por cero que eliminaba NPCs
'*************************************************
    Dim Rechazo As Boolean
    Dim ProbRechazo As Long
    Dim ProbExito As Long
    Dim UserEvasion As Long
    Dim NpcPoderAtaque As Long
    Dim PoderEvasioEscudo As Long
    Dim SkillTacticas As Long
    Dim SkillDefensa As Long
    
    UserEvasion = PoderEvasion(userindex)
    NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
    PoderEvasioEscudo = PoderEvasionEscudo(userindex)
    
    SkillTacticas = UserList(userindex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(userindex).Stats.UserSkills(eSkill.DefensaEscudos)
    
    'Esta usando un escudo ???
    If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))
    
    NpcImpacto = (RandomNumber(1, 100) <= ProbExito)
    
    ' el usuario esta usando un escudo ???
    If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then
        If Not NpcImpacto Then
            If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
                ' Chances are rounded
                ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
                Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
                
                If Rechazo Then
                    'Se rechazo el ataque con el escudo
                    Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_ESCUDO, UserList(userindex).Pos.x, UserList(userindex).Pos.Y))
                    Call WriteBlockedWithShieldUser(userindex)
                    Call SubirSkill(userindex, DefensaEscudos)
                End If
            End If
        End If
    End If
End Function

Public Function CalcularDaño(ByVal userindex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
    Dim DañoArma As Long
    Dim DañoUsuario As Long
    Dim Arma As ObjData
    Dim ModifClase As Single
    Dim proyectil As ObjData
    Dim DañoMaxArma As Long
    
    ''sacar esto si no queremos q la matadracos mate el Dragon si o si
    Dim matoDragon As Boolean
    matoDragon = False
    
    With UserList(userindex)
        If .Invent.WeaponEqpObjIndex > 0 And .Invent.NudiEqpSlot = 0 Then
            Arma = ObjData(.Invent.WeaponEqpObjIndex)
            
            ' Ataca a un npc?
            If NpcIndex > 0 Then
                If Arma.proyectil = 1 Then
                    ModifClase = ModClase(.Clase).DañoProyectiles
                    
                    If .Invent.WeaponEqpObjIndex = ARCO_DEMONIO Then ' Usa la mata Dragones?
                        If Npclist(NpcIndex).Numero = NPC_DEMONIO Then
                            DañoArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                            DañoMaxArma = Arma.MaxHit
                        Else
                            DañoArma = 1
                            DañoMaxArma = 1
                        End If
                    Else
                        DañoArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                        DañoMaxArma = Arma.MaxHit
                    End If
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        DañoArma = DañoArma + RandomNumber(proyectil.MinHit, proyectil.MaxHit)
                    End If
                Else
                    ModifClase = ModClase(.Clase).DañoArmas
                    
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then ' Usa la mata Dragones?
                        If Npclist(NpcIndex).NPCtype = DRAGON Then 'Ataca Dragon?
                            DañoArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                            DañoMaxArma = Arma.MaxHit
                            matoDragon = True ''sacar esto si no queremos q la matadracos mate el Dragon si o si
                        Else ' Sino es Dragon daño es 1
                            DañoArma = 1
                            DañoMaxArma = 1
                        End If
                    Else
                        DañoArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                        DañoMaxArma = Arma.MaxHit
                    End If
                End If
            Else ' Ataca usuario
                If Arma.proyectil = 1 Then
                    ModifClase = ModClase(.Clase).DañoProyectiles
                    If .Invent.WeaponEqpObjIndex = ARCO_DEMONIO Then
                        DañoArma = 1
                        DañoMaxArma = 1
                    Else
                        DañoArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                        DañoMaxArma = Arma.MaxHit
                    End If
                     
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        DañoArma = DañoArma + RandomNumber(proyectil.MinHit, proyectil.MaxHit)
                    End If
                Else
                    ModifClase = ModClase(.Clase).DañoArmas
                    
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                        DañoArma = 1 ' Si usa la espada mataDragones daño es 1
                        DañoMaxArma = 1
                    Else
                        DañoArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                        DañoMaxArma = Arma.MaxHit
                    End If
                End If
            End If
        ElseIf .Stats.eCreateTipe = 1 Then
            'Arma Magica
            ModifClase = ModClase(.Clase).DañoArmas
            DañoArma = RandomNumber(.Stats.eMinHit, .Stats.eMaxHit)
            DañoMaxArma = .Stats.eMaxHit
        Else
            ModifClase = ModClase(.Clase).DañoWrestling
            If .Invent.NudiEqpIndex > 0 Then
                Arma = ObjData(.Invent.NudiEqpIndex)
                DañoArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                DañoMaxArma = Arma.MaxHit
            End If
            DañoArma = DañoArma + RandomNumber(1, 3) 'Hacemos que sea "tipo" una daga el ataque de Wrestling
            DañoMaxArma = DañoMaxArma + 3
        End If
        
        DañoUsuario = RandomNumber(.Stats.MinHit, .Stats.MaxHit)
        
        ''sacar esto si no queremos q la matadracos mate el Dragon si o si
        CalcularDaño = (3 * DañoArma + ((DañoMaxArma * 0.2) * MaximoInt(0, .Stats.UserAtributos(eAtributos.Fuerza) - 15)) + DañoUsuario) * ModifClase
        
        'Mannakia
        If UserList(userindex).Invent.MagicIndex > 0 And NpcIndex <> 0 Then
            If ObjData(UserList(userindex).Invent.MagicIndex).EfectoMagico = eMagicType.AumentaGolpe Then
                CalcularDaño = CalcularDaño + ObjData(UserList(userindex).Invent.MagicIndex).CuantoAumento
            End If
        End If
        'Mannakia
    End With
End Function

Public Sub UserDañoNpc(ByVal userindex As Integer, ByVal NpcIndex As Integer)
    Dim Daño As Long
    
    Daño = CalcularDaño(userindex, NpcIndex)
    
    'esta navegando? si es asi le sumamos el daño del barco
    If UserList(userindex).flags.Navegando = 1 And UserList(userindex).Invent.BarcoObjIndex > 0 Then
        Daño = Daño + RandomNumber(ObjData(UserList(userindex).Invent.BarcoObjIndex).MinHit, ObjData(UserList(userindex).Invent.BarcoObjIndex).MaxHit)
    End If
    
    If UserList(userindex).flags.Montando = 1 Then
        Daño = Daño + RandomNumber(ObjData(UserList(userindex).Invent.MonturaObjIndex).MinHit, ObjData(UserList(userindex).Invent.MonturaObjIndex).MaxHit)
    End If
    
    With Npclist(NpcIndex)
        Daño = Daño - .Stats.def
        
        If Daño < 0 Then Daño = 0
        
        Call WriteUserHitNPC(userindex, Daño)
        Call WriteMsg(userindex, 24, UserList(userindex).Char.CharIndex, CStr(Daño))
        
        Call CalcularDarExp(userindex, NpcIndex, Daño)
        .Stats.MinHP = .Stats.MinHP - Daño
        
        If .Stats.MinHP > 0 Then
            'Trata de apuñalar por la espalda al enemigo
            If PuedeApuñalar(userindex) Then
               Call DoApuñalar(userindex, NpcIndex, 0, Daño)
               Call SubirSkill(userindex, Apuñalar)
            End If
        End If
        
        
        If .Stats.MinHP <= 0 Then
            ' Si era un Dragon perdemos la espada mataDragones
            If .NPCtype = DRAGON Then
                'Si tiene equipada la matadracos se la sacamos
                If UserList(userindex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                    Call QuitarObjetos(EspadaMataDragonesIndex, 1, userindex)
                End If
             End If
            
            ' Para que las mascotas no sigan intentando luchar y
            ' comiencen a seguir al amo
            Dim j As Integer
            For j = 1 To MAXMASCOTAS
                If UserList(userindex).MascotasIndex(j) > 0 Then
                    If Npclist(UserList(userindex).MascotasIndex(j)).TargetNPC = NpcIndex Then
                        Npclist(UserList(userindex).MascotasIndex(j)).TargetNPC = 0
                        Npclist(UserList(userindex).MascotasIndex(j)).Movement = TipoAI.SigueAmo
                    End If
                End If
            Next j
            
            Call MuereNpc(NpcIndex, userindex)
        End If
    End With
End Sub

Public Sub NpcDaño(ByVal NpcIndex As Integer, ByVal userindex As Integer)
    Dim Daño As Integer
    Dim Lugar As Integer
    Dim absorbido As Integer
    Dim defbarco As Integer
    Dim defmontura As Integer
    Dim Obj As ObjData
    
    Daño = RandomNumber(Npclist(NpcIndex).Stats.MinHit, Npclist(NpcIndex).Stats.MaxHit)
    
    With UserList(userindex)
        If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
            Obj = ObjData(.Invent.BarcoObjIndex)
            defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        If .flags.Montando = 1 Then
            Obj = ObjData(.Invent.MonturaObjIndex)
            defmontura = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case Lugar
            Case PartesCuerpo.bCabeza
                'Si tiene casco absorbe el golpe
                If .Invent.CascoEqpObjIndex > 0 Then
                   Obj = ObjData(.Invent.CascoEqpObjIndex)
                   absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                End If
          Case Else
                'Si tiene armadura absorbe el golpe
                If .Invent.ArmourEqpObjIndex > 0 Then
                    Dim Obj2 As ObjData
                    Obj = ObjData(.Invent.ArmourEqpObjIndex)
                    If .Invent.EscudoEqpObjIndex Then
                        Obj2 = ObjData(.Invent.EscudoEqpObjIndex)
                        absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
                    Else
                        absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                   End If
                End If
        End Select
        
        absorbido = absorbido + defbarco
        absorbido = absorbido + defmontura
        Daño = Daño - absorbido
        If Daño < 1 Then Daño = 1
        
        Call WriteNPCHitUser(userindex, Lugar, Daño)
        
        Call WriteMsg(userindex, 24, CStr(Npclist(NpcIndex).Char.CharIndex), CStr(Daño))
        
        'Add Nod Kopfnickend
        'If Not .flags.Privilegios And PlayerType.Dios Then .Stats.MinHP = .Stats.MinHP - Daño
        .Stats.MinHP = .Stats.MinHP - Daño
        
        
        
        If .flags.Meditando Then
            If Daño > Fix(.Stats.MinHP * 0.01 * .Stats.UserAtributos(eAtributos.Inteligencia) * .Stats.UserSkills(eSkill.Meditar) * 0.01 * 12 / (RandomNumber(0, 5) + 7)) Then
                .flags.Meditando = False
                Call WriteMeditateToggle(userindex)
                Call WriteConsoleMsg(1, userindex, "Dejas de meditar.", FontTypeNames.FONTTYPE_BROWNI)
                Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageDestCharParticle(UserList(userindex).Char.CharIndex, ParticleToLevel(userindex)))
            End If
        End If
        
        'Muere el usuario
        If .Stats.MinHP <= 0 Then
            Call WriteNPCKillUser(userindex) ' Le informamos que ha muerto ;)
            
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
            Else
                'Al matarlo no lo sigue mas
                If Npclist(NpcIndex).Stats.Alineacion = 0 Then
                    Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
                    Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
                    Npclist(NpcIndex).flags.AttackedBy = 0
                End If
            End If
            
            Call UserDie(userindex)
        End If
    End With
End Sub

Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal userindex As Integer, Optional ByVal CheckElementales As Boolean = True)
    Dim j As Integer
    
    For j = 1 To MAXMASCOTAS
        If UserList(userindex).MascotasIndex(j) > 0 Then
            If UserList(userindex).MascotasIndex(j) <> NpcIndex Then
                If Npclist(UserList(userindex).MascotasIndex(j)).TargetNPC = 0 Then
                    Npclist(UserList(userindex).MascotasIndex(j)).TargetNPC = NpcIndex
                    Npclist(UserList(userindex).MascotasIndex(j)).Movement = TipoAI.NpcAtacaNpc
                End If
            End If
        End If
    Next j
    
    If UserList(userindex).masc.TieneFamiliar = 1 Then
        If UserList(userindex).masc.NpcIndex > 0 And UserList(userindex).masc.invocado Then
            If Npclist(UserList(userindex).masc.NpcIndex).TargetNPC = 0 Then
                Npclist(UserList(userindex).masc.NpcIndex).TargetNPC = NpcIndex
                Npclist(UserList(userindex).masc.NpcIndex).Movement = TipoAI.NpcAtacaNpc
                Npclist(UserList(userindex).masc.NpcIndex).Hostile = 1
            End If
        End If
    End If
End Sub

Public Sub AllFollowAmo(ByVal userindex As Integer)
    Dim j As Integer
    
    For j = 1 To MAXMASCOTAS
        If UserList(userindex).MascotasIndex(j) > 0 Then
            Call FollowAmo(UserList(userindex).MascotasIndex(j))
        End If
    Next j
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal userindex As Integer) As Boolean

    If UserList(userindex).flags.AdminInvisible = 1 Then Exit Function
    If (UserList(userindex).flags.Privilegios And PlayerType.Dios) And Not UserList(userindex).flags.AdminPerseguible Then Exit Function
    
    With Npclist(NpcIndex)
        ' El npc puede atacar ???
        If .CanAttack = 1 Then
            NpcAtacaUser = True
            Call CheckPets(NpcIndex, userindex, False)
            
            If .Target = 0 Then .Target = userindex
            
            If UserList(userindex).flags.AtacadoPorNpc = 0 And UserList(userindex).flags.AtacadoPorUser = 0 Then
                UserList(userindex).flags.AtacadoPorNpc = NpcIndex
            End If
        Else
            NpcAtacaUser = False
            Exit Function
        End If
        
        .CanAttack = 0
        
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(.flags.Snd1, .Pos.x, .Pos.Y))
        End If
    End With
    
    If NpcImpacto(NpcIndex, userindex) Then
        With UserList(userindex)
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.x, .Pos.Y))
            
            If .flags.Meditando = False Then
                If .flags.Navegando = 0 And .flags.Montando = 0 Then
                    Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0))
                End If
            End If
            
            Call NpcDaño(NpcIndex, userindex)
            Call WriteUpdateHP(userindex)
            
            '¿Puede envenenar?
            If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(userindex)
            
            If Npclist(NpcIndex).IsFamiliar And Npclist(NpcIndex).MaestroUser <> 0 Then
                If UserList(Npclist(NpcIndex).MaestroUser).masc.gDesarma Then
                    'Aca desarma con probabilidad
                End If
                
                If UserList(Npclist(NpcIndex).MaestroUser).masc.gEnseguece Then
                    'Aca enseguece con probabilidad
                End If
                
                If UserList(Npclist(NpcIndex).MaestroUser).masc.gEntorpece Then
                    'Aca entorpece con probabilidad
                End If
                
                If UserList(Npclist(NpcIndex).MaestroUser).masc.gEnvenena Then
                    'Aca envenena con probabilidad
                End If
                
                If UserList(Npclist(NpcIndex).MaestroUser).masc.gParaliza Then
                    'Aca paraliza con probabilidad
                End If
            End If
        End With
    Else
        Call WriteNPCSwing(userindex)
        Call WriteMsg(userindex, 25, Npclist(NpcIndex).Char.CharIndex)
    End If
    
    '-----Tal vez suba los skills------
    Call SubirSkill(userindex, Tacticas)
    
    'Controla el nivel del usuario
    Call CheckUserLevel(userindex)
End Function

Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
    Dim PoderAtt As Long
    Dim PoderEva As Long
    Dim ProbExito As Long
    
    PoderAtt = Npclist(Atacante).PoderAtaque
    PoderEva = Npclist(Victima).PoderEvasion
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtt - PoderEva) * 0.4))
    NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
End Function

Public Sub NpcDañoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
    Dim Daño As Integer
    Dim ExpGanada As Integer
    
    With Npclist(Atacante)
        Daño = RandomNumber(.Stats.MinHit, .Stats.MaxHit)
        Npclist(Victima).Stats.MinHP = Npclist(Victima).Stats.MinHP - Daño
        ExpGanada = (((Daño * Npclist(Victima).GiveEXP) / Npclist(Victima).Stats.MaxHP) * 5) / 2
        
        If .IsFamiliar = True Then
            If .MaestroUser > 0 Then
                UserList(.MaestroUser).masc.Exp = UserList(.MaestroUser).masc.Exp + ExpGanada
                ' (Daño * ((Npclist(Victima).GiveEXP / 600) / Npclist(Victima).Stats.MaxHP))
                UserList(.MaestroUser).Stats.Exp = UserList(.MaestroUser).Stats.Exp + ExpGanada
           
           Call WriteMsg(.MaestroUser, 21, CStr(ExpGanada))
           
              CheckFamiLevel .MaestroUser
            End If
        End If
        
        If Npclist(Victima).Stats.MinHP < 1 Then
            .Movement = .flags.OldMovement
            
            If LenB(.flags.AttackedBy) <> 0 Then
                .Hostile = .flags.OldHostil
            End If
            
            If .MaestroUser > 0 Then
                Call FollowAmo(Atacante)
            End If
            
            Call MuereNpc(Victima, 0)
        Else
            If Npclist(Victima).IsFamiliar Then
                If Npclist(Victima).MaestroUser > 0 Then
                    UpdateFamiliar Npclist(Victima).MaestroUser, False
                End If
            End If
        End If
    End With
End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)
'*************************************************
'Author: Unknown
'Last modified: 01/03/2009
'01/03/2009: ZaMa - Las mascotas no pueden atacar al rey si quedan pretorianos vivos.
'*************************************************
    
    With Npclist(Atacante)
        
        ' El npc puede atacar ???
        If .CanAttack = 1 Then
            .CanAttack = 0
            If cambiarMOvimiento Then
                Npclist(Victima).TargetNPC = Atacante
                Npclist(Victima).Movement = TipoAI.NpcAtacaNpc
            End If
        Else
            Exit Sub
        End If
        
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(.flags.Snd1, .Pos.x, .Pos.Y))
        End If
        
        If NpcImpactoNpc(Atacante, Victima) Then
            If Npclist(Victima).flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(Npclist(Victima).flags.Snd2, Npclist(Victima).Pos.x, Npclist(Victima).Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(Victima).Pos.x, Npclist(Victima).Pos.Y))
            End If
        
            If .MaestroUser > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_IMPACTO, .Pos.x, .Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, Npclist(Victima).Pos.x, Npclist(Victima).Pos.Y))
            End If
            
            Call NpcDañoNpc(Atacante, Victima)
        Else
            If .MaestroUser > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_SWING, .Pos.x, .Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_SWING, Npclist(Victima).Pos.x, Npclist(Victima).Pos.Y))
            End If
        End If
    End With
End Sub
Public Sub UsuarioAtacaPuesto(ByVal userindex As Integer)
    Dim Arma As Integer
    Dim ProbExito As Integer
    Dim PoderAtaque As Integer
    Dim UserImpacto As Boolean
    Dim Skill As Integer
 
    Arma = UserList(userindex).Invent.WeaponEqpObjIndex
    
    If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = PoderAtaqueProyectil(userindex)
            Skill = eSkill.Proyectiles
        ElseIf ObjData(Arma).SubTipo = 5 Or ObjData(Arma).SubTipo = 6 Then
            PoderAtaque = PoderAtaqueArma(userindex)
            Skill = eSkill.arrojadizas
        Else
            Skill = eSkill.armas
            PoderAtaque = PoderAtaqueArma(userindex)
        End If
    Else
        PoderAtaque = PoderAtaqueWrestling(userindex)
        Skill = eSkill.artes
    End If
    
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + PoderAtaque * 0.4))
    UserImpacto = (RandomNumber(1, 100) <= ProbExito)

    If UserImpacto Then
        Dim Daño As Long
    
        Daño = CalcularDaño(userindex)
        
        'esta navegando? si es asi le sumamos el daño del barco
        If UserList(userindex).flags.Navegando = 1 And UserList(userindex).Invent.BarcoObjIndex > 0 Then
            Daño = Daño + RandomNumber(ObjData(UserList(userindex).Invent.BarcoObjIndex).MinHit, ObjData(UserList(userindex).Invent.BarcoObjIndex).MaxHit)
        End If
        
        If UserList(userindex).flags.Montando = 1 Then
            Daño = Daño + RandomNumber(ObjData(UserList(userindex).Invent.MonturaObjIndex).MinHit, ObjData(UserList(userindex).Invent.MonturaObjIndex).MaxHit)
        End If
        
        If Daño < 0 Then Daño = 0
            
        Call WriteUserHitNPC(userindex, Daño)
        Call WriteMsg(userindex, 24, UserList(userindex).Char.CharIndex, CStr(Daño))

        If UserList(userindex).flags.Entrenando = 0 Then
            Call WriteConsoleMsg(1, userindex, "Comienzas a trabajar...", FontTypeNames.FONTTYPE_BROWNI)
            UserList(userindex).flags.Entrenando = 1
        End If
        
        Dim ExpaDar As Long
        ExpaDar = Daño * 1.58396226415094
        
        If ExpaDar > 0 Then
            If UserList(userindex).GrupoIndex > 0 Then
                Call mdGrupo.ObtenerExito(userindex, ExpaDar, UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.Y)
            Else
                UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpaDar
            '    If UserList(Userindex).Stats.Exp > MAXEXP Then UserList(Userindex).Stats.Exp = MAXEXP
                Call WriteMsg(userindex, 21, CStr(ExpaDar))
            End If
            
            Call CheckUserLevel(userindex)
        End If
        
        Call SubirSkill(userindex, Skill)
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_IMPACTO2, UserList(userindex).Pos.x, UserList(userindex).Pos.Y))
    
    Else
    
    'Castelli // Mensaje FAllas al fallar en puesto NW, se quita tmb_
    ' el sound Swing diciendo q si estas entrenando no manda el sonido de falla...
Call WriteMsg(userindex, 26)
      'Castelli // Mensaje FAllas al fallar en puesto NW, se quita tmb_
    ' el sound Swing diciendo q si estas entrenando no manda el sonido de falla...
 
    
    End If
    
End Sub
Public Sub UsuarioAtacaNpc(ByVal userindex As Integer, ByVal NpcIndex As Integer)
    If Not PuedeAtacarNPC(userindex, NpcIndex) Then
        Exit Sub
    End If
    
    Call NPCAtacado(NpcIndex, userindex)
    
    If UserImpactoNpc(userindex, NpcIndex) Then
        If Npclist(NpcIndex).flags.Snd2 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y))
        Else
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y))
        End If
        
        Call UserDañoNpc(userindex, NpcIndex)
        
        GolpeInmovilizaNpc userindex, NpcIndex
    Else
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_SWING, UserList(userindex).Pos.x, UserList(userindex).Pos.Y))
        Call WriteUserSwing(userindex)
        Call WriteMsg(userindex, 26)
    End If

    If UserList(userindex).flags.Oculto = 1 Or UserList(userindex).flags.Invisible = 1 Then
        UserList(userindex).flags.Invisible = 0
        UserList(userindex).Counters.Invisibilidad = 0
        
        UserList(userindex).flags.Oculto = 0
        UserList(userindex).Counters.Ocultando = 0
        UserList(userindex).Counters.TiempoOculto = 0

        Call WriteConsoleMsg(1, userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageSetInvisible(UserList(userindex).Char.CharIndex, False))
    End If
End Sub

Public Sub UsuarioAtaca(ByVal userindex As Integer)
    Dim index As Integer
    Dim AttackPos As WorldPos



    With UserList(userindex)
        'Quitamos stamina
        If .Stats.MinSTA >= 10 Then
            Call QuitarSta(userindex, RandomNumber(1, 10))
        Else
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(1, userindex, "Estas muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, userindex, "Estas muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)
            End If
            Exit Sub
        End If
        
        AttackPos = .Pos
        Call HeadtoPos(.Char.heading, AttackPos)
        
        'Exit if not legal
        If AttackPos.x < XMinMapSize Or AttackPos.x > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_SWING, .Pos.x, .Pos.Y))
            Exit Sub
        End If
        
        index = MapData(AttackPos.map, AttackPos.x, AttackPos.Y).userindex
        
        'Look for user
        If index > 0 Then
            Call UsuarioAtacaUsuario(userindex, index)
            Call WriteUpdateUserStats(userindex)
            Call WriteUpdateUserStats(index)
            Exit Sub
        End If
        
        index = MapData(AttackPos.map, AttackPos.x, AttackPos.Y).NpcIndex
        
        'Look for NPC
        If index > 0 Then
            If Npclist(index).Attackable Then
                If Npclist(index).MaestroUser > 0 And MapInfo(Npclist(index).Pos.map).Pk = False Then
                    Call WriteConsoleMsg(1, userindex, "No podés atacar mascotas en zonas seguras", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
                
                Call UsuarioAtacaNpc(userindex, index)
            Else
                Call WriteConsoleMsg(1, userindex, "No podés atacar a este NPC", FontTypeNames.FONTTYPE_FIGHT)
            End If
            
            Call WriteUpdateUserStats(userindex)
            
            Exit Sub
        End If
        
        index = MapData(AttackPos.map, AttackPos.x, AttackPos.Y).ObjInfo.ObjIndex
        
        If index > 0 Then
            If ObjData(index).OBJType = otPuestos Then
                Call UsuarioAtacaPuesto(userindex)
            End If
        End If
        
        
        'CASTELLI // SAle sonido fallas si no esta entrenando en puesto_
        'de la dungeon newbie
        If UserList(userindex).flags.Entrenando = 0 Then
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_SWING, .Pos.x, .Pos.Y))
        End If
        'CASTELLI // SAle sonido fallas si no esta entrenando en puesto_
        'de la dungeon newbie
        
        Call WriteUpdateUserStats(userindex)
        
        If .Counters.Trabajando Then .Counters.Trabajando = .Counters.Trabajando - 1
            
        If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1
    End With
End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
    Dim ProbRechazo As Long
    Dim Rechazo As Boolean
    Dim ProbExito As Long
    Dim PoderAtaque As Long
    Dim UserPoderEvasion As Long
    Dim UserPoderEvasionEscudo As Long
    Dim Arma As Integer
    Dim SkillTacticas As Long
    Dim SkillDefensa As Long
    
    If UserList(AtacanteIndex).flags.NoFalla = 1 Then
        UserList(AtacanteIndex).flags.NoFalla = 0
        UsuarioImpacto = True
        Exit Function
    End If
    
    SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(eSkill.DefensaEscudos)
    
    Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    
    'Calculamos el poder de evasion...
    UserPoderEvasion = PoderEvasion(VictimaIndex)
    
    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
       UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
       UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
    Else
        UserPoderEvasionEscudo = 0
    End If
    
    'Esta usando un arma ???
    If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
        Else
            PoderAtaque = PoderAtaqueArma(AtacanteIndex)
        End If
    Else
        PoderAtaque = PoderAtaqueWrestling(AtacanteIndex)
    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtaque - UserPoderEvasion) * 0.4))
    
    UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)
    
    ' el usuario esta usando un escudo ???
    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
        'Fallo ???
        If Not UsuarioImpacto Then
            ' Chances are rounded
            If SkillDefensa = 0 And SkillTacticas = 0 Then
                Rechazo = False
                Exit Function
            Else
                ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
            End If
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
            If Rechazo = True Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(VictimaIndex).Pos.x, UserList(VictimaIndex).Pos.Y))
                  
                Call WriteBlockedWithShieldOther(AtacanteIndex)
                Call WriteBlockedWithShieldUser(VictimaIndex)
                
                Call SubirSkill(VictimaIndex, DefensaEscudos)
            End If
        End If
    End If
    
    Call FlushBuffer(VictimaIndex)
End Function

Public Sub UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)

    If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Sub
    
    With UserList(AtacanteIndex)
        If Distancia(.Pos, UserList(VictimaIndex).Pos) > MAXDISTANCIAARCO Then
           Call WriteConsoleMsg(1, AtacanteIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
           Exit Sub
        End If
        
        Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)
        
        If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
            Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.x, .Pos.Y))
            
            If UserList(VictimaIndex).flags.Navegando = 0 And UserList(VictimaIndex).flags.Montando = 0 Then
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, FXSANGRE, 0))
            End If
            
            Call GolpeInmoviliza(AtacanteIndex, VictimaIndex)
            Call GolpeDesarma(AtacanteIndex, VictimaIndex)
            
            Call UserDañoUser(AtacanteIndex, VictimaIndex)
        Else
            Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_SWING, .Pos.x, .Pos.Y))
            Call WriteUserSwing(AtacanteIndex)
            Call WriteUserAttackedSwing(VictimaIndex, AtacanteIndex)
            
            Call WriteMsg(VictimaIndex, 25, UserList(AtacanteIndex).Char.CharIndex)
            Call WriteMsg(AtacanteIndex, 26)
        End If
        
        If UserList(AtacanteIndex).flags.Oculto = 1 Or UserList(AtacanteIndex).flags.Invisible = 1 Then
            UserList(AtacanteIndex).flags.Invisible = 0
            UserList(AtacanteIndex).Counters.Invisibilidad = 0
            
            UserList(AtacanteIndex).flags.Oculto = 0
            UserList(AtacanteIndex).Counters.Ocultando = 0
            UserList(AtacanteIndex).Counters.TiempoOculto = 0

            Call WriteConsoleMsg(1, AtacanteIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessageSetInvisible(UserList(AtacanteIndex).Char.CharIndex, False))
        End If
        
    End With
End Sub

Public Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    Dim Daño As Long
    Dim Lugar As Integer
    Dim absorbido As Long
    Dim defbarco As Integer
    Dim defmontura As Integer
    Dim Obj As ObjData
    Dim Resist As Byte
    
    Daño = CalcularDaño(AtacanteIndex)
    
    Call GolpeEnvenena(AtacanteIndex, VictimaIndex)
    
    With UserList(AtacanteIndex)
        If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
             Obj = ObjData(.Invent.BarcoObjIndex)
             Daño = Daño + RandomNumber(Obj.MinHit, Obj.MaxHit)
        End If
        
        If UserList(VictimaIndex).flags.Navegando = 1 And UserList(VictimaIndex).Invent.BarcoObjIndex > 0 Then
             Obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
             defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        If .flags.Montando = 1 Then
             Obj = ObjData(.Invent.MonturaObjIndex)
             Daño = Daño + RandomNumber(Obj.MinHit, Obj.MaxHit)
        End If
         
        If UserList(VictimaIndex).flags.Montando = 1 Then
             Obj = ObjData(UserList(VictimaIndex).Invent.MonturaObjIndex)
             defmontura = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        If .Invent.WeaponEqpObjIndex > 0 Then
            Resist = ObjData(.Invent.WeaponEqpObjIndex).Refuerzo
        End If
        
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case Lugar
            Case PartesCuerpo.bCabeza
                'Si tiene casco absorbe el golpe
                If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
                    absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                    absorbido = absorbido + defbarco - Resist
                    absorbido = absorbido + defmontura - Resist
                    Daño = Daño - absorbido
                    If Daño < 0 Then Daño = 1
                End If
            
            Case Else
                'Si tiene armadura absorbe el golpe
                If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
                    Dim Obj2 As ObjData
                    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex Then
                        Obj2 = ObjData(UserList(VictimaIndex).Invent.EscudoEqpObjIndex)
                        absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
                    Else
                        absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                    End If
                    absorbido = absorbido + defbarco - Resist
                    absorbido = absorbido + defmontura - Resist
                    Daño = Daño - absorbido
                    If Daño < 0 Then Daño = 1
                End If
        End Select
        
        If UserList(VictimaIndex).Stats.eCreateTipe = 2 Then
            'Aura protectora
            absorbido = RandomNumber(UserList(VictimaIndex).Stats.eMinDef, UserList(VictimaIndex).Stats.eMaxDef)
            Daño = Daño - absorbido
            If Daño < 0 Then Daño = 1
        End If
        
        Call WriteUserHittedUser(AtacanteIndex, Lugar, UserList(VictimaIndex).Char.CharIndex, Daño)
        Call WriteUserHittedByUser(VictimaIndex, Lugar, .Char.CharIndex, Daño)
        
        Call WriteMsg(AtacanteIndex, 24, UserList(AtacanteIndex).Char.CharIndex, CStr(Daño))
        Call WriteMsg(VictimaIndex, 24, UserList(AtacanteIndex).Char.CharIndex, CStr(Daño))

        UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - Daño
        
        Call SubirSkill(VictimaIndex, Tacticas)
        
        If .flags.Hambre = 0 And .flags.Sed = 0 Then
            'Si usa un arma quizas suba "Combate con armas"
            If .Invent.NudiEqpIndex > 0 Then
                Call SubirSkill(AtacanteIndex, artes)
            ElseIf .Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(.Invent.WeaponEqpObjIndex).proyectil Then
                    'es un Arco. Sube Armas a Distancia
                    Call SubirSkill(AtacanteIndex, Proyectiles)
                ElseIf ObjData(.Invent.WeaponEqpObjIndex).SubTipo = 5 Or ObjData(.Invent.WeaponEqpObjIndex).SubTipo = 6 Then
                    Call SubirSkill(AtacanteIndex, arrojadizas)
                Else
                    'Sube combate con armas.
                    Call SubirSkill(AtacanteIndex, armas)
                End If
            Else
                'sino tal vez lucha libre
                Call SubirSkill(AtacanteIndex, artes)
            End If
                    
            'Trata de apuñalar por la espalda al enemigo
            If PuedeApuñalar(AtacanteIndex) Then
                Call DoApuñalar(AtacanteIndex, 0, VictimaIndex, Daño)
                Call SubirSkill(AtacanteIndex, Apuñalar)
            End If
        End If
        
        If UserList(VictimaIndex).Stats.MinHP <= 0 Then
            Call ContarMuerte(VictimaIndex, AtacanteIndex)
            
            ' Para que las mascotas no sigan intentando luchar y
            ' comiencen a seguir al amo
            Dim j As Integer
            For j = 1 To MAXMASCOTAS
                If .MascotasIndex(j) > 0 Then
                    If Npclist(.MascotasIndex(j)).Target = VictimaIndex Then
                        Npclist(.MascotasIndex(j)).Target = 0
                        Call FollowAmo(.MascotasIndex(j))
                    End If
                End If
            Next j
            
            Call ActStats(VictimaIndex, AtacanteIndex)
            Call UserDie(VictimaIndex)
        Else
            'Está vivo - Actualizamos el HP
            Call WriteUpdateHP(VictimaIndex)
        End If
    End With
    
    'Controla el nivel del usuario
    Call CheckUserLevel(AtacanteIndex)
    
    Call FlushBuffer(VictimaIndex)
End Sub

Sub UsuarioAtacadoPorUsuario(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 10/01/08
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
'***************************************************

    If TriggerZonaPelea(attackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub
    
    Dim EraCriminal As Boolean
    
    If UserList(VictimIndex).flags.Meditando Then
        UserList(VictimIndex).flags.Meditando = False
        Call WriteMeditateToggle(VictimIndex)
        Call WriteConsoleMsg(1, VictimIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_BROWNI)
        Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageDestCharParticle(UserList(VictimIndex).Char.CharIndex, ParticleToLevel(VictimIndex)))
    End If
    
    Call AllMascotasAtacanUser(attackerIndex, VictimIndex)
    Call AllMascotasAtacanUser(VictimIndex, attackerIndex)
    
    'Si la victima esta saliendo se cancela la salida
    Call CancelExit(VictimIndex)
    Call FlushBuffer(VictimIndex)
End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
    'Reaccion de las mascotas
    Dim iCount As Integer
    
    For iCount = 1 To MAXMASCOTAS
        If UserList(Maestro).MascotasIndex(iCount) > 0 Then
            Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = victim
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = TipoAI.NPCDEFENSA
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
        End If
    Next iCount
    
    If UserList(Maestro).masc.TieneFamiliar = 1 Then
        If UserList(Maestro).masc.NpcIndex > 0 And UserList(Maestro).masc.invocado Then
            Npclist(UserList(Maestro).masc.NpcIndex).flags.AttackedBy = victim
            Npclist(UserList(Maestro).masc.NpcIndex).Movement = TipoAI.NPCDEFENSA
            Npclist(UserList(Maestro).masc.NpcIndex).Hostile = 1
        End If
    End If
End Sub

Public Function PuedeAtacar(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
'***************************************************
'Autor: Unknown
'Last Modification: 24/02/2009
'Returns true if the AttackerIndex is allowed to attack the VictimIndex.
'24/01/2007 Pablo (ToxicWaste) - Ordeno todo y agrego situacion de Defensa en ciudad Armada y Caos.
'24/02/2009: ZaMa - Los usuarios pueden atacarse entre si.
'***************************************************
    'MUY importante el orden de estos "IF"...
    
    'Esta muerto no podes atacar
    If UserList(attackerIndex).flags.Muerto = 1 Then
        Call WriteMsg(attackerIndex, 4)
        PuedeAtacar = False
        Exit Function
    End If
    
    'No podes atacar a alguien muerto
    If UserList(VictimIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(1, attackerIndex, "No podés atacar a un espiritu", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    
    If (UserList(VictimIndex).GrupoIndex = UserList(attackerIndex).GrupoIndex) And UserList(VictimIndex).GrupoIndex <> 0 Then
        PuedeAtacar = False
        Exit Function
    End If
    
    If UserList(VictimIndex).flags.inDuelo = 1 Then
        If UserList(VictimIndex).flags.vicDuelo = attackerIndex Then
            PuedeAtacar = True
            Exit Function
        Else
            Call WriteConsoleMsg(1, attackerIndex, "El objetivo esta en duelo!!", FontTypeNames.FONTTYPE_BROWNI)
        End If
    ElseIf UserList(attackerIndex).flags.inDuelo Then
        PuedeAtacar = False
        Call WriteConsoleMsg(1, attackerIndex, "No puedes atacar a otras personas que no sean tu oponente!!", FontTypeNames.FONTTYPE_BROWNI)
        Exit Function
    End If
    
    'Estamos en una Arena? o un trigger zona segura?
    Select Case TriggerZonaPelea(attackerIndex, VictimIndex)
        Case eTrigger6.TRIGGER6_PERMITE
            PuedeAtacar = True
            Exit Function
        
        Case eTrigger6.TRIGGER6_PROHIBE
            PuedeAtacar = False
            Exit Function
        
        Case eTrigger6.TRIGGER6_AUSENTE
            'Si no estamos en el Trigger 6 entonces es imposible atacar un gm
            If UserList(VictimIndex).flags.Privilegios And PlayerType.Dios Then
                If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(1, attackerIndex, "El ser es demasiado poderoso", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = False
                Exit Function
            End If
    End Select
    
    If (esCiuda(attackerIndex) Or esArmada(attackerIndex)) And (esCiuda(VictimIndex) Or esArmada(VictimIndex)) Then
        Call WriteConsoleMsg(1, attackerIndex, "Para poder atacar ciudadanos de un mismo ejercito escribe /RETIRAR. Este como consecuencia te quedaras en estado renegado.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
    
    If (esRepu(VictimIndex) Or esMili(VictimIndex)) And (esMili(attackerIndex) Or esRepu(attackerIndex)) Then
        Call WriteConsoleMsg(1, attackerIndex, "Para poder atacar ciudadanos de un mismo ejercito escribe /RETIRAR. Este como consecuencia te quedaras en estado renegado.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
    
    'Estas en un Mapa Seguro?
    If MapInfo(UserList(VictimIndex).Pos.map).Pk = False Then
        Call WriteConsoleMsg(1, attackerIndex, "Esta es una zona segura, aqui no podes atacar otros usuarios.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
    
    'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
    If MapData(UserList(VictimIndex).Pos.map, UserList(VictimIndex).Pos.x, UserList(VictimIndex).Pos.Y).Trigger = eTrigger.ZONASEGURA Or _
        MapData(UserList(attackerIndex).Pos.map, UserList(attackerIndex).Pos.x, UserList(attackerIndex).Pos.Y).Trigger = eTrigger.ZONASEGURA Then
        Call WriteConsoleMsg(1, attackerIndex, "No podes pelear aqui.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
    
    PuedeAtacar = True
End Function
Public Function PuedeRobar(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
    'Esta muerto no podes atacar
    If UserList(attackerIndex).flags.Muerto = 1 Then
        Call WriteMsg(attackerIndex, 5)
        PuedeRobar = False
        Exit Function
    End If
    
    'No podes atacar a alguien muerto
    If UserList(VictimIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(1, attackerIndex, "No podés robarle a un espiritu", FontTypeNames.FONTTYPE_INFO)
        PuedeRobar = False
        Exit Function
    End If
    
    If (UserList(VictimIndex).GrupoIndex = UserList(attackerIndex).GrupoIndex) And UserList(VictimIndex).GrupoIndex <> 0 Then
        PuedeRobar = False
        Exit Function
    End If
    
    If (esCiuda(attackerIndex) Or esArmada(attackerIndex)) And (esCiuda(VictimIndex) Or esArmada(VictimIndex)) Then
        Call WriteConsoleMsg(1, attackerIndex, "Para poder robar a ciudadanos de un mismo ejercito escribe /RETIRAR. Este como consecuencia te quedaras en estado renegado.", FontTypeNames.FONTTYPE_WARNING)
        PuedeRobar = False
        Exit Function
    End If
    
    If (esRepu(VictimIndex) Or esMili(VictimIndex)) And (esMili(attackerIndex) Or esRepu(attackerIndex)) Then
        Call WriteConsoleMsg(1, attackerIndex, "Para poder robar a ciudadanos de un mismo ejercito escribe /RETIRAR. Este como consecuencia te quedaras en estado renegado.", FontTypeNames.FONTTYPE_WARNING)
        PuedeRobar = False
        Exit Function
    End If
    
    'Estas en un Mapa Seguro?
    If MapInfo(UserList(VictimIndex).Pos.map).Pk = False Then
        Call WriteConsoleMsg(1, attackerIndex, "Esta es una zona segura, aqui no podes robarles a otros usuarios.", FontTypeNames.FONTTYPE_WARNING)
        PuedeRobar = False
        Exit Function
    End If
    
    'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
    If MapData(UserList(VictimIndex).Pos.map, UserList(VictimIndex).Pos.x, UserList(VictimIndex).Pos.Y).Trigger = eTrigger.ZONASEGURA Or _
        MapData(UserList(attackerIndex).Pos.map, UserList(attackerIndex).Pos.x, UserList(attackerIndex).Pos.Y).Trigger = eTrigger.ZONASEGURA Then
        Call WriteConsoleMsg(1, attackerIndex, "No podes robar aqui.", FontTypeNames.FONTTYPE_WARNING)
        PuedeRobar = False
        Exit Function
    End If
    
    PuedeRobar = True
End Function
Public Function PuedeAyudar(ByVal userindex As Integer, ByVal tU As Integer) As Boolean
    
    If UserList(userindex).Faccion.Renegado = 1 Then
        PuedeAyudar = True
        Exit Function
    End If
    
   ' If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    '    If Not (esArmada(tU) Or esMili(tU)) Then
    '        PuedeAyudar = False
    '        Exit Function
    '    End If
   ' End If
    
    If UserList(userindex).Faccion.ArmadaReal = 1 Or _
       UserList(userindex).Faccion.Ciudadano = 1 Then
            
        If Not (esArmada(tU) Or esCiuda(tU)) Then
            PuedeAyudar = False
            Exit Function
        End If
    End If
    
    If UserList(userindex).Faccion.Republicano = 1 Or _
       UserList(userindex).Faccion.Milicia = 1 Then
        
        If Not (esMili(tU) Or esRepu(tU)) Then
            PuedeAyudar = False
            Exit Function
        End If
    End If
    
    PuedeAyudar = True
End Function

Public Function PuedeAtacarNPC(ByVal attackerIndex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Autor: Unknown Author (Original version)
'Returns True if AttackerIndex can attack the NpcIndex
'Last Modification: 24/01/2007
'24/01/2007 Pablo (ToxicWaste) - Orden y corrección de ataque sobre una mascota y guardias
'14/08/2007 Pablo (ToxicWaste) - Reescribo y agrego TODOS los casos posibles cosa de usar
'esta función para todo lo referente a ataque a un NPC. Ya sea Magia, Físico o a Distancia.
'***************************************************
    'Esta muerto?
    If UserList(attackerIndex).flags.Muerto = 1 Then
        Call WriteMsg(attackerIndex, 4)
        PuedeAtacarNPC = False
        Exit Function
    End If
    
    'Estas en modo Combate?
    If Not UserList(attackerIndex).flags.ModoCombate Then
        Call WriteConsoleMsg(1, attackerIndex, "Debes estar en modo de combate poder atacar al NPC.", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacarNPC = False
        Exit Function
    End If
    
    'Es una criatura atacable?
    If Npclist(NpcIndex).Attackable = 0 Then
        Call WriteConsoleMsg(1, attackerIndex, "Objetivo ínvalido.", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacarNPC = False
        Exit Function
    End If
    
    'Es valida la distancia a la cual estamos atacando?
    If Distancia(UserList(attackerIndex).Pos, Npclist(NpcIndex).Pos) >= MAXDISTANCIAARCO Then
       Call WriteConsoleMsg(1, attackerIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
       PuedeAtacarNPC = False
       Exit Function
    End If
    
    'Es una criatura No-Hostil?
    If Npclist(NpcIndex).Hostile = 0 Then
        'Es Guardia del Caos?
        If Npclist(NpcIndex).Faccion = 3 Then
            If esCaos(attackerIndex) Then
                Call WriteConsoleMsg(1, attackerIndex, "Objetivo ínvalido.", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function
            End If
        ElseIf Npclist(NpcIndex).Faccion = 2 Then
            If esMili(attackerIndex) Or esRepu(attackerIndex) Then
                Call WriteConsoleMsg(1, attackerIndex, "Objetivo ínvalido.", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function
            End If
        ElseIf Npclist(NpcIndex).Faccion = 1 Then
            If esArmada(attackerIndex) Or esCiuda(attackerIndex) Then
                Call WriteConsoleMsg(1, attackerIndex, "Objetivo ínvalido.", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function
            End If
        End If
    End If
    
    'Es el NPC mascota de alguien?
    If Not (esCaos(attackerIndex) Or esRene(attackerIndex)) And Npclist(NpcIndex).MaestroUser > 0 Then
        If esCiuda(Npclist(NpcIndex).MaestroUser) Or esArmada(Npclist(NpcIndex).MaestroUser) Then
            If esCiuda(attackerIndex) Or esArmada(attackerIndex) Then
                Call WriteConsoleMsg(1, attackerIndex, "Los imperiales no pueden atacar mascotas de Ciudadanos o Armadas. Para retirarte las tropas del imperio tipee '/RETIRAR'", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function
            End If
        Else
            If esMili(attackerIndex) Or esRepu(attackerIndex) Then
                Call WriteConsoleMsg(1, attackerIndex, "Los republicanos no pueden atacar mascotas de ciudadanos o milicianos.", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function
            End If
        End If
    End If
    
    PuedeAtacarNPC = True
End Function

Sub CalcularDarExp(ByVal userindex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/09/06 Nacho
'Reescribi gran parte del Sub
'Ahora, da toda la experiencia del npc mientras este vivo.
'***************************************************
    Dim ExpaDar As Long
    

    
    '[Nacho] Chekeamos que las variables sean validas para las operaciones
    If ElDaño <= 0 Then ElDaño = 0
    If Npclist(NpcIndex).Stats.MaxHP <= 0 Then Exit Sub
    If ElDaño > Npclist(NpcIndex).Stats.MinHP Then ElDaño = Npclist(NpcIndex).Stats.MinHP
    
    '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
    ExpaDar = CLng(ElDaño * (Npclist(NpcIndex).GiveEXP / Npclist(NpcIndex).Stats.MaxHP))
    If ExpaDar <= 0 Then Exit Sub
    
    '[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
            'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
            'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
    If ExpaDar > Npclist(NpcIndex).flags.ExpCount Then
        ExpaDar = Npclist(NpcIndex).flags.ExpCount
        Npclist(NpcIndex).flags.ExpCount = 0
    Else
        Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).flags.ExpCount - ExpaDar
    End If
    
    '[Nacho] Le damos la exp al user
    If ExpaDar > 0 Then
        If UserList(userindex).GrupoIndex > 0 Then
            Call mdGrupo.ObtenerExito(userindex, ExpaDar, Npclist(NpcIndex).Pos.map, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y)
        Else
            UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpaDar
          '  If UserList(Userindex).Stats.Exp > MAXEXP Then _
                UserList(Userindex).Stats.Exp = MAXEXP
        Call WriteMsg(userindex, 21, CStr(ExpaDar))
        End If
        
        Call CheckUserLevel(userindex)
    End If
End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6
'TODO: Pero que rebuscado!!
'Nigo:  Te lo rediseñe, pero no te borro el TODO para que lo revises.
On Error GoTo Errhandler
    Dim tOrg As eTrigger
    Dim tDst As eTrigger
    
    tOrg = MapData(UserList(Origen).Pos.map, UserList(Origen).Pos.x, UserList(Origen).Pos.Y).Trigger
    tDst = MapData(UserList(Destino).Pos.map, UserList(Destino).Pos.x, UserList(Destino).Pos.Y).Trigger
    
    If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
        If tOrg = tDst Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE
        End If
    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE
    End If

Exit Function
Errhandler:
    TriggerZonaPelea = TRIGGER6_AUSENTE
    LogError ("Error en TriggerZonaPelea - " & err.description)
End Function

Sub GolpeEnvenena(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    Dim ObjInd As Integer
    
    ObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    
    If ObjInd > 0 Then
        If ObjData(ObjInd).proyectil = 1 Then
            ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
            If ObjData(ObjInd).SubTipo = 3 Then
                GoTo Envenena
            End If
        End If
    End If
    
    ObjInd = UserList(AtacanteIndex).Invent.MagicIndex
    If ObjInd > 0 Then
        If ObjData(ObjInd).EfectoMagico = eMagicType.Envenena Then
            GoTo Envenena
        End If
    End If
    
    Exit Sub
    
Envenena:
    If RandomNumber(1, 35) < 3 Then
        UserList(VictimaIndex).flags.Envenenado = 3
        Call WriteConsoleMsg(2, VictimaIndex, UserList(AtacanteIndex).Name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(2, AtacanteIndex, "Has envenenado a " & UserList(VictimaIndex).Name & "!!", FontTypeNames.FONTTYPE_FIGHT)
    End If
    Call FlushBuffer(VictimaIndex)
    Exit Sub
    
End Sub
Sub GolpeIncinera(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    Dim ObjInd As Integer
    
    ObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    
    If ObjInd > 0 Then
        If ObjData(ObjInd).proyectil = 1 Then
            ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
            If ObjData(ObjInd).SubTipo = 2 Then
                GoTo Incinera
            End If
        End If
    End If
    
    ObjInd = UserList(AtacanteIndex).Invent.MagicIndex
    If ObjInd > 0 Then
        If ObjData(ObjInd).EfectoMagico = eMagicType.Incinera Then
            GoTo Incinera
        End If
    End If
    
    Exit Sub
    
Incinera:
    If RandomNumber(1, 35) < 2 Then
        UserList(VictimaIndex).flags.Incinerado = 1
        Call WriteConsoleMsg(2, VictimaIndex, UserList(AtacanteIndex).Name & " te ha incinerado!!", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(2, AtacanteIndex, "Has incinerado a " & UserList(VictimaIndex).Name & "!!", FontTypeNames.FONTTYPE_FIGHT)
    End If
    Call FlushBuffer(VictimaIndex)
    Exit Sub
    
End Sub
Public Sub GolpeInmoviliza(ByVal userindex As Integer, ByVal VictimaIndex As Integer)
'*********************************************************************
'Author: Leandro Mendoza (Mannakia)
'Desc: The coup can look up to the victim
'Last Modify: 21/10/10
'*********************************************************************
    Dim res As Byte, probm As Integer, orbe As Boolean
    If UserList(VictimaIndex).flags.Paralizado = 1 Then Exit Sub
    
    If UserList(userindex).Invent.WeaponEqpObjIndex <> 0 Then
        If UserList(userindex).Invent.MagicIndex = 0 Then
            Exit Sub
        Else
            If ObjData(UserList(userindex).Invent.MagicIndex).EfectoMagico = eMagicType.Paraliza Then _
                 orbe = True
        End If
    End If
    
    If Not orbe And Not UserList(userindex).Invent.WeaponEqpObjIndex <> 0 Then
        res = RandomNumber(0, ObtenerSuerte(UserList(userindex).Stats.UserSkills(eSkill.artes)))
        
        If UserList(userindex).Invent.NudiEqpIndex <> 0 Then probm = 10
        If UserList(userindex).Clase = eClass.Gladiador Then probm = probm + 10
        res = res - Porcentaje(res, probm)
        
        If res < 5 Then
            GoTo Paraliza
        End If
    Else
        res = RandomNumber(1, 35)
        If res < 3 Then
            GoTo Paraliza
        End If
    End If
    Exit Sub
    
Paraliza:
    UserList(VictimaIndex).flags.Paralizado = 1
    UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado * 0.5
    Call WriteParalizeOK(VictimaIndex)
    Call WriteConsoleMsg(2, userindex, "Tu golpe ha dejado inmóvil a tu oponente", FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(2, VictimaIndex, "¡El golpe te ha dejado inmóvil!", FontTypeNames.FONTTYPE_INFO)
    
End Sub
Public Sub GolpeInmovilizaNpc(ByVal userindex As Integer, ByVal NpcIndex As Integer)
'*********************************************************************
'Author: Leandro Mendoza (Mannakia)
'Desc: The coup can look up to the victim
'Last Modify: 21/10/10
'*********************************************************************
    Dim res As Byte, probm As Integer, orbe As Boolean
    If Npclist(NpcIndex).flags.Paralizado = 1 Then Exit Sub

    If UserList(userindex).Invent.WeaponEqpObjIndex <> 0 Then
        If UserList(userindex).Invent.MagicIndex = 0 Then
            Exit Sub
        Else
            If ObjData(UserList(userindex).Invent.MagicIndex).EfectoMagico = eMagicType.Paraliza Then _
                 orbe = True
        End If
    End If
    
    If Not orbe And UserList(userindex).Invent.WeaponEqpObjIndex = 0 Then
        res = RandomNumber(0, ObtenerSuerte(UserList(userindex).Stats.UserSkills(eSkill.artes)))
        
        If UserList(userindex).Invent.NudiEqpIndex <> 0 Then probm = 10
        If UserList(userindex).Clase = eClass.Gladiador Then probm = probm + 10
        res = res - Porcentaje(res, probm)
        
        If res < 5 Then
            GoTo Paraliza
        End If
    ElseIf orbe Then
        res = RandomNumber(1, 35)
        If res < 5 Then
            GoTo Paraliza
        End If
    End If
    Exit Sub
    
Paraliza:
    Npclist(NpcIndex).flags.Paralizado = 1
    Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
    Call WriteConsoleMsg(2, userindex, "Tu golpe ha dejado inmóvil a la criatura", FontTypeNames.FONTTYPE_INFO)

End Sub
Public Sub GolpeDesarma(ByVal userindex As Integer, ByVal VictimIndex As Integer)
'*********************************************************************
'Author: Leandro Mendoza (Mannakia)
'Desc: The coup can disarm to the victim
'Last Modify: 21/10/10
'*********************************************************************
    Dim res As Byte, probm As Integer
    
    If UserList(userindex).Invent.WeaponEqpSlot <> 0 Then Exit Sub
    
    If UserList(userindex).Invent.NudiEqpIndex <> 0 Then probm = 10
    If UserList(userindex).Clase = eClass.Gladiador Then probm = probm + 10
    
    res = RandomNumber(1, ObtenerSuerte(UserList(userindex).Stats.UserSkills(eSkill.artes)))
    res = res - Porcentaje(res, probm)
    
    If res < 3 Then
        Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
        Call WriteConsoleMsg(2, userindex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
        
        Call FlushBuffer(VictimIndex)
    End If
End Sub

Public Sub GolpeEstupidiza(ByVal userindex As Integer, ByVal VictimIndex As Integer)
'*********************************************************************
'Author: Leandro Mendoza (Mannakia)
'Desc: The coup can dumbed to the victim
'Last Modify: 21/10/10
'*********************************************************************
    Dim res As Byte, probm As Integer
    If UserList(userindex).Invent.WeaponEqpSlot <> 0 Then Exit Sub
    
    If UserList(userindex).Invent.NudiEqpIndex <> 0 Then probm = 10
    If UserList(userindex).Clase = eClass.Gladiador Then probm = probm + 10
    
    res = RandomNumber(1, ObtenerSuerte(UserList(userindex).Stats.UserSkills(eSkill.artes)))
    res = res - Porcentaje(res, probm)
    
    If res < 5 Then
        If UserList(userindex).flags.Estupidez = 0 Then
            UserList(userindex).flags.Estupidez = 1
            UserList(userindex).Counters.Ceguera = IntervaloParalizado
        End If
        Call WriteDumb(userindex)
        Call WriteConsoleMsg(2, userindex, "Has dejado estúpido a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
        
        Call FlushBuffer(VictimIndex)
    End If
End Sub

