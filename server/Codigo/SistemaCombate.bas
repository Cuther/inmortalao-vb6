Attribute VB_Name = "SistemaCombate"
Option Explicit

Public Const MAXDISTANCIAARCO As Byte = 18
Public Const MAXDISTANCIAMAGIA As Byte = 18

Public Const NPC_DEMONIO As Integer = 1
Public Const ARCO_DEMONIO As Integer = 666

Public Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer
    If a > b Then
        MinimoInt = b
    Else
        MinimoInt = a
    End If
End Function

Public Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer
    If a > b Then
        MaximoInt = a
    Else
        MaximoInt = b
    End If
End Function

Private Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long
On Error GoTo err
    PoderEvasionEscudo = (UserList(UserIndex).Stats.UserSkills(eSkill.DefensaEscudos) * ModClase(UserList(UserIndex).Clase).Evasion) * 0.5
Exit Function
err:

End Function

Private Function PoderEvasion(ByVal UserIndex As Integer) As Long
On Error GoTo err

    Dim lTemp As Long
    With UserList(UserIndex)
        lTemp = (.Stats.UserSkills(eSkill.Tacticas) + _
          .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.Clase).Evasion
       
        PoderEvasion = (lTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
    End With
    
    Exit Function
err:
    
End Function

Private Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
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

Private Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
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

Private Function PoderAtaqueWrestling(ByVal UserIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
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

Public Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
    Dim PoderAtaque As Long
    Dim Arma As Integer
    Dim Skill As eSkill
    Dim ProbExito As Long
    
    If UserList(UserIndex).flags.NoFalla = 1 Then
        UserList(UserIndex).flags.NoFalla = 0
        UserImpactoNpc = True
        Exit Function
    End If
    
    Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
    
    If UserList(UserIndex).Invent.NudiEqpIndex > 0 Then
        PoderAtaque = PoderAtaqueWrestling(UserIndex)
        Skill = eSkill.artes
    ElseIf Arma > 0 Then 'Usando un arma
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = PoderAtaqueProyectil(UserIndex)
            Skill = eSkill.Proyectiles
        ElseIf ObjData(Arma).SubTipo = 5 Or ObjData(Arma).SubTipo = 6 Then
            PoderAtaque = PoderAtaqueArma(UserIndex)
            Skill = eSkill.arrojadizas
        Else
            PoderAtaque = PoderAtaqueArma(UserIndex)
            Skill = eSkill.armas
        End If
    Else 'Peleando con puños
        PoderAtaque = PoderAtaqueWrestling(UserIndex)
        Skill = eSkill.artes
    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))
    If EsDIOS(UserIndex) Then
        ProbExito = 95
    End If
    
    UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
    
    If UserImpactoNpc Then
        Call SubirSkill(UserIndex, Skill)
    End If
End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
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
    
    UserEvasion = PoderEvasion(UserIndex)
    NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
    PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)
    
    SkillTacticas = UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(UserIndex).Stats.UserSkills(eSkill.DefensaEscudos)
    
    'Esta usando un escudo ???
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))
    
    NpcImpacto = (RandomNumber(1, 100) <= ProbExito)
    
    ' el usuario esta usando un escudo ???
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
        If Not NpcImpacto Then
            If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
                ' Chances are rounded
                ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
                Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
                
                If Rechazo Then
                    'Se rechazo el ataque con el escudo
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                    Call WriteBlockedWithShieldUser(UserIndex)
                    Call SubirSkill(UserIndex, DefensaEscudos)
                End If
            End If
        End If
    End If
End Function

Public Function CalcularDaño(ByVal UserIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
    Dim DañoArma As Long
    Dim DañoUsuario As Long
    Dim Arma As ObjData
    Dim ModifClase As Single
    Dim proyectil As ObjData
    Dim DañoMaxArma As Long
    
    ''sacar esto si no queremos q la matadracos mate el Dragon si o si
    Dim matoDragon As Boolean
    matoDragon = False
    
    With UserList(UserIndex)
        If .Invent.WeaponEqpObjIndex > 0 And .Invent.NudiEqpSlot = 0 Then
            Arma = ObjData(.Invent.WeaponEqpObjIndex)
            
            ' Ataca a un npc?
            If NpcIndex > 0 Then
                If Arma.proyectil = 1 Then
                    ModifClase = ModClase(.Clase).DañoProyectiles
                    
                    If .Invent.WeaponEqpObjIndex = ARCO_DEMONIO Then ' Usa la arco mata Demonios?
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
        If UserList(UserIndex).Invent.MagicIndex > 0 And NpcIndex <> 0 Then
            If ObjData(UserList(UserIndex).Invent.MagicIndex).EfectoMagico = eMagicType.AumentaGolpe Then
                CalcularDaño = CalcularDaño + ObjData(UserList(UserIndex).Invent.MagicIndex).CuantoAumento
            End If
        End If
        'Mannakia
    End With
End Function

Public Sub UserDañoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    Dim Daño As Long
    
    Daño = CalcularDaño(UserIndex, NpcIndex)
    
    'esta navegando? si es asi le sumamos el daño del barco
    If UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).Invent.BarcoObjIndex > 0 Then
        Daño = Daño + RandomNumber(ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MinHit, ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MaxHit)
    End If
    
    If UserList(UserIndex).flags.Montando = 1 Then
        Daño = Daño + RandomNumber(ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MinHit, ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MaxHit)
    End If
    
    With Npclist(NpcIndex)
        Daño = Daño - .Stats.def
        
        If Daño < 0 Then Daño = 0
        
        Call WriteUserHitNPC(UserIndex, Daño)
        Call WriteMsg(UserIndex, 24, UserList(UserIndex).Char.CharIndex, CStr(Daño))

        Call CalcularDarExp(UserIndex, NpcIndex, Daño)
        .Stats.MinHP = .Stats.MinHP - Daño
        
        If .Stats.MinHP > 0 Then
            'Trata de apuñalar por la espalda al enemigo
            If PuedeApuñalar(UserIndex) Then
               Call DoApuñalar(UserIndex, NpcIndex, 0, Daño)
               Call SubirSkill(UserIndex, Apuñalar)
            End If
        End If
        
        
        If .Stats.MinHP <= 0 Then
            ' Si era un Dragon perdemos la espada mataDragones
            If .NPCtype = DRAGON Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(32, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                'Si tiene equipada la matadracos se la sacamos
            '    If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
            '        Call QuitarObjetos(EspadaMataDragonesIndex, 1, UserIndex)
            '    End If
             End If
            
            ' Para que las mascotas no sigan intentando luchar y
            ' comiencen a seguir al amo
            Dim j As Integer
            For j = 1 To MAXMASCOTAS
                If UserList(UserIndex).MascotasIndex(j) > 0 Then
                    If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex Then
                        Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0
                        Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.SigueAmo
                    End If
                End If
            Next j

            Call MuereNpc(NpcIndex, UserIndex)
        End If
    End With
End Sub

Public Sub NpcDaño(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    Dim Daño As Integer
    Dim Lugar As Integer
    Dim absorbido As Integer
    Dim defbarco As Integer
    Dim defmontura As Integer
    Dim obj As ObjData
    
    Daño = RandomNumber(Npclist(NpcIndex).Stats.MinHit, Npclist(NpcIndex).Stats.MaxHit)
    
    With UserList(UserIndex)
        If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
            obj = ObjData(.Invent.BarcoObjIndex)
            defbarco = RandomNumber(obj.MinDef, obj.MaxDef)
        End If
        
        If .flags.Montando = 1 Then
            obj = ObjData(.Invent.MonturaObjIndex)
            defmontura = RandomNumber(obj.MinDef, obj.MaxDef)
        End If
        
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case Lugar
            Case PartesCuerpo.bCabeza
                'Si tiene casco absorbe el golpe
                If .Invent.CascoEqpObjIndex > 0 Then
                   obj = ObjData(.Invent.CascoEqpObjIndex)
                   absorbido = RandomNumber(obj.MinDef, obj.MaxDef)
                End If
          Case Else
                'Si tiene armadura absorbe el golpe
                If .Invent.ArmourEqpObjIndex > 0 Then
                    Dim Obj2 As ObjData
                    obj = ObjData(.Invent.ArmourEqpObjIndex)
                    If .Invent.EscudoEqpObjIndex Then
                        Obj2 = ObjData(.Invent.EscudoEqpObjIndex)
                        absorbido = RandomNumber(obj.MinDef + Obj2.MinDef, obj.MaxDef + Obj2.MaxDef)
                    Else
                        absorbido = RandomNumber(obj.MinDef, obj.MaxDef)
                   End If
                End If
        End Select
        
        absorbido = absorbido + defbarco
        absorbido = absorbido + defmontura
        Daño = Daño - absorbido
        If Daño < 1 Then Daño = 1
        
        Call WriteNPCHitUser(UserIndex, Lugar, Daño)
        
        Call WriteMsg(UserIndex, 24, CStr(Npclist(NpcIndex).Char.CharIndex), CStr(Daño))
        
        'Add Nod Kopfnickend
        'If Not .flags.Privilegios And PlayerType.Dios Then .Stats.MinHP = .Stats.MinHP - Daño
        .Stats.MinHP = .Stats.MinHP - Daño
        
        
        
        If .flags.Meditando Then
            If Daño > Fix(.Stats.MinHP * 0.01 * .Stats.UserAtributos(eAtributos.Inteligencia) * .Stats.UserSkills(eSkill.Meditar) * 0.01 * 12 / (RandomNumber(0, 5) + 7)) Then
                .flags.Meditando = False
                Call WriteMeditateToggle(UserIndex)
                Call WriteConsoleMsg(1, UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_BROWNI)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageDestCharParticle(UserList(UserIndex).Char.CharIndex, ParticleToLevel(UserIndex)))
            End If
        End If
        
        'Muere el usuario
        If .Stats.MinHP <= 0 Then
            Call WriteNPCKillUser(UserIndex) ' Le informamos que ha muerto ;)
            
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
            
            Call UserDie(UserIndex)
        End If
    End With
End Sub

Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, Optional ByVal CheckElementales As Boolean = True)
    Dim j As Integer
    
    For j = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(j) > 0 Then
            If UserList(UserIndex).MascotasIndex(j) <> NpcIndex Then
                If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0 Then
                    Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex
                    Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.NpcAtacaNpc
                End If
            End If
        End If
    Next j
    
    If UserList(UserIndex).masc.TieneFamiliar = 1 Then
        If UserList(UserIndex).masc.NpcIndex > 0 And UserList(UserIndex).masc.invocado Then
            If Npclist(UserList(UserIndex).masc.NpcIndex).TargetNPC = 0 Then
                Npclist(UserList(UserIndex).masc.NpcIndex).TargetNPC = NpcIndex
                Npclist(UserList(UserIndex).masc.NpcIndex).Movement = TipoAI.NpcAtacaNpc
                Npclist(UserList(UserIndex).masc.NpcIndex).Hostile = 1
            End If
        End If
    End If
End Sub

Public Sub AllFollowAmo(ByVal UserIndex As Integer)
    Dim j As Integer
    
    For j = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(j) > 0 Then
            Call FollowAmo(UserList(UserIndex).MascotasIndex(j))
        End If
    Next j
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

    If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Function
    If EsDIOS(UserIndex) And Not UserList(UserIndex).flags.AdminPerseguible Then Exit Function
    
    'Add Marius
    If UserList(UserIndex).flags.Oculto = 1 Or UserList(UserIndex).flags.Invisible = 1 Then Exit Function
    '\Add
    
    With Npclist(NpcIndex)
        ' El npc puede atacar ???
        If .CanAttack = 1 Then
            NpcAtacaUser = True
            Call CheckPets(NpcIndex, UserIndex, False)
            
            If .Target = 0 Then .Target = UserIndex
            
            If UserList(UserIndex).flags.AtacadoPorNpc = 0 And UserList(UserIndex).flags.AtacadoPorUser = 0 Then
                UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex
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
    
    If NpcImpacto(NpcIndex, UserIndex) Then
        With UserList(UserIndex)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.x, .Pos.Y))
            
            If .flags.Meditando = False Then
                If .flags.Navegando = 0 And .flags.Montando = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0))
                End If
            End If
            
            Call NpcDaño(NpcIndex, UserIndex)
            Call WriteUpdateHP(UserIndex)
            
            '¿Puede envenenar?
            If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(UserIndex)
            
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
        Call WriteNPCSwing(UserIndex)
        Call WriteMsg(UserIndex, 25, Npclist(NpcIndex).Char.CharIndex)
    End If
    
    '-----Tal vez suba los skills------
    Call SubirSkill(UserIndex, Tacticas)
    
    'Controla el nivel del usuario
    Call CheckUserLevel(UserIndex)
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

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMovimiento As Boolean = True)
'*************************************************
'Author: Unknown
'Last modified: 01/03/2009
'01/03/2009: ZaMa - Las mascotas no pueden atacar al rey si quedan pretorianos vivos.
'*************************************************
    
    With Npclist(Atacante)
        
        ' El npc puede atacar ???
        If .CanAttack = 1 Then
            .CanAttack = 0
            If cambiarMovimiento Then
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
        
        'Add Marius para que npc mate cuando ataque a otro npc pasaba con las mascotas
        'Muere
        If Npclist(Victima).Stats.MinHP < 1 Then
            Npclist(Victima).Stats.MinHP = 0
            If Npclist(Atacante).MaestroUser > 0 Then
                Call MuereNpc(Victima, Npclist(Atacante).MaestroUser)
            Else
                Call MuereNpc(Victima, 0)
            End If
        End If
        '\Add
    End With
End Sub
Public Sub UsuarioAtacaPuesto(ByVal UserIndex As Integer)
    Dim Arma As Integer
    Dim ProbExito As Integer
    Dim PoderAtaque As Integer
    Dim UserImpacto As Boolean
    Dim Skill As Integer
 
    Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
    
    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = PoderAtaqueProyectil(UserIndex)
            Skill = eSkill.Proyectiles
        ElseIf ObjData(Arma).SubTipo = 5 Or ObjData(Arma).SubTipo = 6 Then
            PoderAtaque = PoderAtaqueArma(UserIndex)
            Skill = eSkill.arrojadizas
        Else
            Skill = eSkill.armas
            PoderAtaque = PoderAtaqueArma(UserIndex)
        End If
    Else
        PoderAtaque = PoderAtaqueWrestling(UserIndex)
        Skill = eSkill.artes
    End If
    
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + PoderAtaque * 0.4))
    UserImpacto = (RandomNumber(1, 100) <= ProbExito)

    If UserImpacto Then
        Dim Daño As Long
    
        Daño = CalcularDaño(UserIndex)
        
        'esta navegando? si es asi le sumamos el daño del barco
        If UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).Invent.BarcoObjIndex > 0 Then
            Daño = Daño + RandomNumber(ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MinHit, ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MaxHit)
        End If
        
        If UserList(UserIndex).flags.Montando = 1 Then
            Daño = Daño + RandomNumber(ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MinHit, ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MaxHit)
        End If
        
        If Daño < 0 Then Daño = 0
            
        Call WriteUserHitNPC(UserIndex, Daño)
        Call WriteMsg(UserIndex, 24, UserList(UserIndex).Char.CharIndex, CStr(Daño))

        If UserList(UserIndex).flags.Entrenando = 0 Then
            Call WriteConsoleMsg(1, UserIndex, "Comienzas a trabajar...", FontTypeNames.FONTTYPE_BROWNI)
            UserList(UserIndex).flags.Entrenando = 1
        End If
        
        Dim ExpaDar As Long
        ExpaDar = Daño * 1.58396226415094
        
        If ExpaDar > 0 Then
            If UserList(UserIndex).GrupoIndex > 0 Then
                Call mdGrupo.ObtenerExito(UserIndex, ExpaDar, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y)
            Else
                UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar
            '    If UserList(Userindex).Stats.Exp > MAXEXP Then UserList(Userindex).Stats.Exp = MAXEXP
                Call WriteMsg(UserIndex, 21, CStr(ExpaDar))
            End If
            
            Call CheckUserLevel(UserIndex)
        End If
        
        Call SubirSkill(UserIndex, Skill)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO2, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
    
    Else
    
    'Castelli // Mensaje FAllas al fallar en puesto NW, se quita tmb_
    ' el sound Swing diciendo q si estas entrenando no manda el sonido de falla...
Call WriteMsg(UserIndex, 26)
      'Castelli // Mensaje FAllas al fallar en puesto NW, se quita tmb_
    ' el sound Swing diciendo q si estas entrenando no manda el sonido de falla...
 
    
    End If
    
End Sub
Public Sub UsuarioAtacaNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
        Exit Sub
    End If
    
    Call NPCAtacado(NpcIndex, UserIndex)
    
    If UserImpactoNpc(UserIndex, NpcIndex) Then
        
        
        If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then 'Cuandno atacamos con la mata drako a un users
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(149, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y))
        Else
            
            If Npclist(NpcIndex).flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y))
            End If
            
        End If
        Call UserDañoNpc(UserIndex, NpcIndex)
        
        GolpeInmovilizaNpc UserIndex, NpcIndex
    Else
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
        Call WriteUserSwing(UserIndex)
        Call WriteMsg(UserIndex, 26)
    End If

    If UserList(UserIndex).flags.Oculto = 1 Or UserList(UserIndex).flags.Invisible = 1 Then
        UserList(UserIndex).flags.Invisible = 0
        UserList(UserIndex).Counters.Invisibilidad = 0
        
        UserList(UserIndex).flags.Oculto = 0
        UserList(UserIndex).Counters.Ocultando = 0
        UserList(UserIndex).Counters.TiempoOculto = 0

        Call WriteConsoleMsg(1, UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
    End If
End Sub

Public Sub UsuarioAtaca(ByVal UserIndex As Integer)
    Dim index As Integer
    Dim AttackPos As WorldPos
    
    'Check bow's interval
    'If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
    
    'Check Spell-Magic interval
    'If Not IntervaloPermiteMagiaGolpe(UserIndex) Then
        'Check Attack interval
        If Not IntervaloPermiteAtacar(UserIndex) Then
            Exit Sub
        End If
    'End If
    
    With UserList(UserIndex)
        'Quitamos stamina
        If .Stats.MinSTA >= 10 Then
            Call QuitarSta(UserIndex, RandomNumber(1, 10))
        Else
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(1, UserIndex, "Estas muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, UserIndex, "Estas muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            Exit Sub
        End If
        
        AttackPos = .Pos
        Call HeadtoPos(.Char.heading, AttackPos)
        
        'Exit if not legal
        If AttackPos.x < XMinMapSize Or AttackPos.x > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, .Pos.x, .Pos.Y))
            Exit Sub
        End If
        
        index = MapData(AttackPos.map, AttackPos.x, AttackPos.Y).UserIndex
        
        'Look for user
        If index > 0 Then
            Call UsuarioAtacaUsuario(UserIndex, index)
            Call WriteUpdateUserStats(UserIndex)
            Call WriteUpdateUserStats(index)
            Exit Sub
        End If
        
        index = MapData(AttackPos.map, AttackPos.x, AttackPos.Y).NpcIndex
        
        'Look for NPC
        If index > 0 Then
            If Npclist(index).Attackable Then
                If Npclist(index).MaestroUser > 0 And MapInfo(Npclist(index).Pos.map).Pk = False Then
                    Call WriteConsoleMsg(1, UserIndex, "No podés atacar mascotas en zonas seguras", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
                
                Call UsuarioAtacaNpc(UserIndex, index)
            Else
                Call WriteConsoleMsg(1, UserIndex, "No podés atacar a este NPC", FontTypeNames.FONTTYPE_FIGHT)
            End If
            
            Call WriteUpdateUserStats(UserIndex)
            
            Exit Sub
        End If
        
        index = MapData(AttackPos.map, AttackPos.x, AttackPos.Y).ObjInfo.ObjIndex
        
        If index > 0 Then
            If ObjData(index).OBJType = otPuestos Then
                Call UsuarioAtacaPuesto(UserIndex)
            End If
        End If
        
        
        'CASTELLI // SAle sonido fallas si no esta entrenando en puesto_
        'de la dungeon newbie
        If UserList(UserIndex).flags.Entrenando = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, .Pos.x, .Pos.Y))
        End If
        'CASTELLI // SAle sonido fallas si no esta entrenando en puesto_
        'de la dungeon newbie
        
        Call WriteUpdateUserStats(UserIndex)
        
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
            
            If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then 'Cuandno atacamos con la mata drako a un users
                Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(149, .Pos.x, .Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.x, .Pos.Y))
            End If
            
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
    Dim obj As ObjData
    Dim Resist As Byte
    
    Daño = CalcularDaño(AtacanteIndex)
    
    Call GolpeEnvenena(AtacanteIndex, VictimaIndex)
    
    With UserList(AtacanteIndex)
        If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
             obj = ObjData(.Invent.BarcoObjIndex)
             Daño = Daño + RandomNumber(obj.MinHit, obj.MaxHit)
        End If
        
        If UserList(VictimaIndex).flags.Navegando = 1 And UserList(VictimaIndex).Invent.BarcoObjIndex > 0 Then
             obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
             defbarco = RandomNumber(obj.MinDef, obj.MaxDef)
        End If
        
        If .flags.Montando = 1 Then
             obj = ObjData(.Invent.MonturaObjIndex)
             Daño = Daño + RandomNumber(obj.MinHit, obj.MaxHit)
        End If
         
        If UserList(VictimaIndex).flags.Montando = 1 Then
             obj = ObjData(UserList(VictimaIndex).Invent.MonturaObjIndex)
             defmontura = RandomNumber(obj.MinDef, obj.MaxDef)
        End If
        
        If .Invent.WeaponEqpObjIndex > 0 Then
            Resist = ObjData(.Invent.WeaponEqpObjIndex).Refuerzo
        End If
        
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case Lugar
            Case PartesCuerpo.bCabeza
                'Si tiene casco absorbe el golpe
                If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
                    obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
                    absorbido = RandomNumber(obj.MinDef, obj.MaxDef)
                    absorbido = absorbido + defbarco - Resist
                    absorbido = absorbido + defmontura - Resist
                    Daño = Daño - absorbido
                    If Daño < 0 Then Daño = 1
                End If
            
            Case Else
                'Si tiene armadura absorbe el golpe
                If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
                    obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
                    Dim Obj2 As ObjData
                    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex Then
                        Obj2 = ObjData(UserList(VictimaIndex).Invent.EscudoEqpObjIndex)
                        absorbido = RandomNumber(obj.MinDef + Obj2.MinDef, obj.MaxDef + Obj2.MaxDef)
                    Else
                        absorbido = RandomNumber(obj.MinDef, obj.MaxDef)
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
        
    'Add Marius Carreras
    If UserList(attackerIndex).Pos.map = MapaCarrera Then
        'Si empezo la carrera, se pueden dar murra entre ellos.
        If Carrera_puestos = 255 Then
            'La cortamos aca por que sino verificaria toda la mierda de facciones que no quiero.
            PuedeAtacar = True
            Exit Function
        Else
            Call WriteConsoleMsg(1, attackerIndex, "Debes esperar a que comience la carrera para poder atacar a los demas.", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacar = False
            Exit Function
        End If
    End If
    '\Add
        
    If (UserList(VictimIndex).GrupoIndex = UserList(attackerIndex).GrupoIndex) And UserList(VictimIndex).GrupoIndex <> 0 Then
        PuedeAtacar = False
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
            'Des Nod Kopfnickend
            'Se puede atacar a un conse siempre y cuando este en modo combate, igual la proteccion divina no lo deja piñatear
            If EsFacc(VictimIndex) And Not UserList(VictimIndex).flags.ModoCombate Then
                If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(1, attackerIndex, UserList(VictimIndex).Name & " no esta de humor para pelear.", FontTypeNames.FONTTYPE_WARNING)
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
Public Function PuedeAyudar(ByVal UserIndex As Integer, ByVal tU As Integer) As Boolean
    
    If UserList(UserIndex).faccion.Renegado = 1 Then
        PuedeAyudar = True
        Exit Function
    End If
    
   ' If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    '    If Not (esArmada(tU) Or esMili(tU)) Then
    '        PuedeAyudar = False
    '        Exit Function
    '    End If
   ' End If
    
    If UserList(UserIndex).faccion.ArmadaReal = 1 Or _
       UserList(UserIndex).faccion.Ciudadano = 1 Then
            
        If Not (esArmada(tU) Or esCiuda(tU)) Then
            PuedeAyudar = False
            Exit Function
        End If
    End If
    
    If UserList(UserIndex).faccion.Republicano = 1 Or _
       UserList(UserIndex).faccion.Milicia = 1 Then
        
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
    
    'Add Marius Carrera
    If UserList(attackerIndex).Pos.map = Bandera_mapa Then
        PuedeAtacarNPC = False
        Exit Function
    End If
    '\Add
    
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
        If Npclist(NpcIndex).faccion = 3 Then
            If esCaos(attackerIndex) Then
                Call WriteConsoleMsg(1, attackerIndex, "Objetivo ínvalido.", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function
            End If
        ElseIf Npclist(NpcIndex).faccion = 2 Then
            If esMili(attackerIndex) Or esRepu(attackerIndex) Then
                Call WriteConsoleMsg(1, attackerIndex, "Objetivo ínvalido.", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function
            End If
        ElseIf Npclist(NpcIndex).faccion = 1 Then
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

Sub CalcularDarExp(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)
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
        If UserList(UserIndex).GrupoIndex > 0 Then
            Call mdGrupo.ObtenerExito(UserIndex, ExpaDar, Npclist(NpcIndex).Pos.map, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y)
        Else
            UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar
            'If UserList(UserIndex).Stats.Exp > MAXEXP Then _
                UserList(UserIndex).Stats.Exp = MAXEXP
        Call WriteMsg(UserIndex, 21, CStr(ExpaDar))
        End If
        
        Call CheckUserLevel(UserIndex)
    End If
End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6
'TODO: Pero que rebuscado!!
'Nigo:  Te lo rediseñe, pero no te borro el TODO para que lo revises.
On Error GoTo Errhandler
    Dim tOrg As eTrigger
    Dim tDst As eTrigger
    
    If UserList(Origen).Pos.map = UserList(Destino).Pos.map Then
        
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
    If RandomNumber(1, 40) < 3 Then
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
Public Sub GolpeInmoviliza(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)
'*********************************************************************
'Author: Leandro Mendoza (Mannakia)
'Desc: The coup can look up to the victim
'Last Modify: 21/10/10
'*********************************************************************
    Dim res As Byte, probm As Integer, orbe As Boolean
    If UserList(VictimaIndex).flags.Paralizado = 1 Then Exit Sub
    
    If UserList(UserIndex).Invent.WeaponEqpObjIndex <> 0 Then
        If UserList(UserIndex).Invent.MagicIndex = 0 Then
            Exit Sub
        Else
            If ObjData(UserList(UserIndex).Invent.MagicIndex).EfectoMagico = eMagicType.Paraliza Then _
                 orbe = True
        End If
    End If
    
    If Not orbe And Not UserList(UserIndex).Invent.WeaponEqpObjIndex <> 0 Then
        res = RandomNumber(0, ObtenerSuerte(UserList(UserIndex).Stats.UserSkills(eSkill.artes)))
        
        If UserList(UserIndex).Invent.NudiEqpIndex <> 0 Then probm = 10
        If UserList(UserIndex).Clase = eClass.Gladiador Then probm = probm + 10
        res = res - Porcentaje(res, probm)
        
        If res < 5 Then
            UserList(VictimaIndex).flags.Paralizado = 1
            UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado * 0.5
            Call WriteParalizeOK(VictimaIndex)
            
            'Add Marius animacion de paralizar
            If Hechizos(9).WAV <> 0 Then Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(Hechizos(9).WAV, UserList(VictimaIndex).Pos.x, UserList(VictimaIndex).Pos.Y))
            If Hechizos(9).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, Hechizos(9).FXgrh, Hechizos(9).loops))
            If Hechizos(9).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateCharParticle(UserList(VictimaIndex).Char.CharIndex, Hechizos(9).Particle))
            '\add
            
            Call WriteConsoleMsg(2, UserIndex, "Tu golpe ha dejado inmóvil a tu oponente", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(2, VictimaIndex, "¡El golpe te ha dejado inmóvil!", FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        res = RandomNumber(1, 40)
        If res < 3 Then
            UserList(VictimaIndex).flags.Paralizado = 1
            UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado * 0.5
            Call WriteParalizeOK(VictimaIndex)
            
            'Add Marius animacion de paralizar
            If Hechizos(9).WAV <> 0 Then Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(Hechizos(9).WAV, UserList(VictimaIndex).Pos.x, UserList(VictimaIndex).Pos.Y))
            If Hechizos(9).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, Hechizos(9).FXgrh, Hechizos(9).loops))
            If Hechizos(9).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateCharParticle(UserList(VictimaIndex).Char.CharIndex, Hechizos(9).Particle))
            '\add
            
            Call WriteConsoleMsg(2, UserIndex, "Tu golpe ha dejado inmóvil a tu oponente", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(2, VictimaIndex, "¡El golpe te ha dejado inmóvil!", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
    
End Sub
Public Sub GolpeInmovilizaNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'*********************************************************************
'Author: Leandro Mendoza (Mannakia)
'Desc: The coup can look up to the victim
'Last Modify: 21/10/10
'*********************************************************************
    Dim res As Byte, probm As Integer, orbe As Boolean
    If Npclist(NpcIndex).flags.Paralizado = 1 Then Exit Sub

    If UserList(UserIndex).Invent.WeaponEqpObjIndex <> 0 Then
        If UserList(UserIndex).Invent.MagicIndex = 0 Then
            Exit Sub
        Else
            If ObjData(UserList(UserIndex).Invent.MagicIndex).EfectoMagico = eMagicType.Paraliza Then _
                 orbe = True
        End If
    End If
    
    If Not orbe And UserList(UserIndex).Invent.WeaponEqpObjIndex = 0 Then
        res = RandomNumber(0, ObtenerSuerte(UserList(UserIndex).Stats.UserSkills(eSkill.artes)))
        
        If UserList(UserIndex).Invent.NudiEqpIndex <> 0 Then probm = 10
        If UserList(UserIndex).Clase = eClass.Gladiador Then probm = probm + 10
        res = res - Porcentaje(res, probm)
        
        If res < 5 Then
            Npclist(NpcIndex).flags.Paralizado = 1
            Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
            Call WriteConsoleMsg(2, UserIndex, "Tu golpe ha dejado inmóvil a la criatura", FontTypeNames.FONTTYPE_INFO)
            
            'Add Nod kopfniceknd animacion de paralizar
            If Hechizos(9).WAV <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(9).WAV, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y))
            If Hechizos(9).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(Npclist(NpcIndex).Char.CharIndex, Hechizos(9).FXgrh, Hechizos(9).loops))
            If Hechizos(9).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateCharParticle(Npclist(NpcIndex).Char.CharIndex, Hechizos(9).Particle))
            '\add
            
        End If
    ElseIf orbe Then
        res = RandomNumber(1, 35)
        If res < 5 Then
            Npclist(NpcIndex).flags.Paralizado = 1
            Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
            Call WriteConsoleMsg(2, UserIndex, "Tu golpe ha dejado inmóvil a la criatura", FontTypeNames.FONTTYPE_INFO)
            
            'Add Nod kopfniceknd animacion de paralizar
            If Hechizos(9).WAV <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(9).WAV, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y))
            If Hechizos(9).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(Npclist(NpcIndex).Char.CharIndex, Hechizos(9).FXgrh, Hechizos(9).loops))
            If Hechizos(9).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateCharParticle(Npclist(NpcIndex).Char.CharIndex, Hechizos(9).Particle))
            '\add
            
        End If
    End If

End Sub
Public Sub GolpeDesarma(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
'*********************************************************************
'Author: Leandro Mendoza (Mannakia)
'Desc: The coup can disarm to the victim
'Last Modify: 21/10/10
'*********************************************************************
    Dim res As Byte, probm As Integer
    
    If UserList(UserIndex).Invent.WeaponEqpSlot <> 0 Then Exit Sub
    
    If UserList(UserIndex).Invent.NudiEqpIndex <> 0 Then probm = 10
    If UserList(UserIndex).Clase = eClass.Gladiador Then probm = probm + 10
    
    res = RandomNumber(1, ObtenerSuerte(UserList(UserIndex).Stats.UserSkills(eSkill.artes)))
    res = res - Porcentaje(res, probm)
    
    If res < 3 Then
        Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
        Call WriteConsoleMsg(2, UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
        
        Call FlushBuffer(VictimIndex)
    End If
End Sub

Public Sub GolpeEstupidiza(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
'*********************************************************************
'Author: Leandro Mendoza (Mannakia)
'Desc: The coup can dumbed to the victim
'Last Modify: 21/10/10
'*********************************************************************
    Dim res As Byte, probm As Integer
    If UserList(UserIndex).Invent.WeaponEqpSlot <> 0 Then Exit Sub
    
    If UserList(UserIndex).Invent.NudiEqpIndex <> 0 Then probm = 10
    If UserList(UserIndex).Clase = eClass.Gladiador Then probm = probm + 10
    
    res = RandomNumber(1, ObtenerSuerte(UserList(UserIndex).Stats.UserSkills(eSkill.artes)))
    res = res - Porcentaje(res, probm)
    
    If res < 5 Then
        If UserList(UserIndex).flags.Estupidez = 0 Then
            UserList(UserIndex).flags.Estupidez = 1
            UserList(UserIndex).Counters.Ceguera = IntervaloParalizado
        End If
        Call WriteDumb(UserIndex)
        Call WriteConsoleMsg(2, UserIndex, "Has dejado estúpido a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
        
        Call FlushBuffer(VictimIndex)
    End If
End Sub

