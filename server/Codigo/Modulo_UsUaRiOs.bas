Attribute VB_Name = "UsUaRiOs"

Option Explicit

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Sub ActStats(ByVal VictimIndex As Integer, ByVal attackerIndex As Integer)
    Dim DaExp As Integer
    Dim EraCriminal As Boolean
    
    DaExp = CInt(UserList(VictimIndex).Stats.ELV) * 2
    
    With UserList(attackerIndex)
        .Stats.Exp = .Stats.Exp + DaExp
       ' If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
        
        'Lo mata
        Call WriteConsoleMsg(2, attackerIndex, "Has matado a " & UserList(VictimIndex).Name & "!", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteMsg(attackerIndex, 21, CStr(DaExp))
              
        Call WriteConsoleMsg(2, VictimIndex, "¡" & .Name & " te ha matado!", FontTypeNames.FONTTYPE_FIGHT)

        Call FlushBuffer(VictimIndex)
    End With
End Sub
Sub DoResucitar(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If .flags.Resucitando = 0 Then Exit Sub
        
        Dim TActual As Long
        TActual = GetTickCount() And &H7FFFFFFF
        If TActual - UserList(UserIndex).Counters.IntervaloRevive < 2500 Then
            Exit Sub
        Else
            Call DarVida(UserIndex)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_REVIVE, .Pos.x, .Pos.Y))
        End If
        
    End With
End Sub

Sub RevivirUsuario(ByVal UserIndex As Integer)
    If UserList(UserIndex).flags.Resucitando <> 1 Then
        UserList(UserIndex).flags.Resucitando = 1
        UserList(UserIndex).Counters.IntervaloRevive = GetTickCount
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateCharParticle(UserList(UserIndex).Char.CharIndex, 22))
    End If
End Sub

Sub DarVida(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        .flags.Muerto = 0
        
        .flags.Resucitando = 0
        
        'If .Stats.MinHP > .Stats.MaxHP Then
            .Stats.MinHP = .Stats.MaxHP
        'End If
        
        If .flags.Navegando = 1 Then
            Dim Barco As ObjData
            Barco = ObjData(.Invent.BarcoObjIndex)
            .Char.Head = 0
            .Char.body = 84 'Barco.Ropaje
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
        Else
            Call DarCuerpoDesnudo(UserIndex)
            
            .Char.Head = .OrigChar.Head
        End If
        
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        Call WriteUpdateUserStats(UserIndex)
    End With
End Sub

Sub ChangeUserChar(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal casco As Integer)

    With UserList(UserIndex).Char
        .body = body
        .Head = Head
        .heading = heading
        .WeaponAnim = Arma
        .ShieldAnim = Escudo
        .CascoAnim = casco
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(body, Head, heading, .CharIndex, Arma, Escudo, .fx, .loops, casco))
    End With
End Sub


Sub EraseUserChar(ByVal UserIndex As Integer)
On Error GoTo ErrorHandler
    'Add Marius Para los errores, solo debug, despeus cuando no pase mas este error puto habria que sacarlo.
    Dim UserName As String
    Dim CharIndex As Integer
    
    If UserIndex > 0 Then
        UserName = UserList(UserIndex).Name
        CharIndex = UserList(UserIndex).Char.CharIndex
    End If
    '\Add
    

    
    With UserList(UserIndex)
        Debug.Print "MAPAAA:" & .Pos.map
        Debug.Print "XXXX:" & .Pos.x
        Debug.Print "YYYYY:" & .Pos.Y
        Debug.Print "USERINDEX:" & UserIndex
        Debug.Print "LASTCHAR:" & LastChar
        
        
        
        If .Char.CharIndex = LastChar Then
            Do Until CharList(LastChar) > 0
                LastChar = LastChar - 1
                If LastChar <= 1 Then Exit Do
            Loop
        End If
        
        CharList(.Char.CharIndex) = 0

        'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
        Call QuitarUser(UserIndex, .Pos.map)
        
        MapData(.Pos.map, .Pos.x, .Pos.Y).UserIndex = 0
        .Char.CharIndex = 0
    End With
    
    NumChars = NumChars - 1
Exit Sub
    
ErrorHandler:
    Call LogError("Error en EraseUserchar " & err.Number & ": " & err.description & ". User: " & UserName & "(UI: " & UserIndex & " - CI: " & CharIndex & " - LastChar: " & LastChar & " - NumChars:" & NumChars & ")")
End Sub

Sub RefreshCharStatus(ByVal UserIndex As Integer)
'*************************************************
'Author: Tararira
'Last modified: 04/21/2008 (NicoNZ)
'Refreshes the status and tag of UserIndex.
'*************************************************
    Dim klan As String
    Dim Barco As ObjData

    With UserList(UserIndex)
        If .GuildIndex > 0 Then
            klan = modGuilds.GuildName(.GuildIndex)
            klan = " <" & klan & ">"
        End If
        
        If EsFacc(UserIndex) And Not EsVIP(UserIndex) Then
            klan = " <Staff>"
        End If
        
        If .showName Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, UserTypeColor(UserIndex), .Name & klan))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, UserTypeColor(UserIndex), vbNullString))
        End If
        
        'Si esta navengando, se cambia la barca.
        If .flags.Navegando Then
            Barco = ObjData(.Invent.Object(.Invent.BarcoSlot).ObjIndex)
            .Char.body = Barco.Ropaje
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
    End With
End Sub
Sub MakeUserChar(ByVal ToMap As Boolean, ByVal sndIndex As Integer, ByVal UserIndex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer)

On Error GoTo hayerror
    Dim CharIndex As Integer

    If InMapBounds(map, x, Y) Then
        'If needed make a new character in list
        If UserList(UserIndex).Char.CharIndex = 0 Then
            CharIndex = NextOpenCharIndex
            UserList(UserIndex).Char.CharIndex = CharIndex
            CharList(CharIndex) = UserIndex
        End If
        
        'Place character on map if needed
        If ToMap Then MapData(map, x, Y).UserIndex = UserIndex
         If UserList(UserIndex).Char.heading = 0 Then
            UserList(UserIndex).Char.heading = 3
        End If
        
        'Send make character command to clients
        Dim klan As String
        If UserList(UserIndex).GuildIndex > 0 Then
            klan = modGuilds.GuildName(UserList(UserIndex).GuildIndex)
        End If
        
        If EsFacc(UserIndex) And Not EsVIP(UserIndex) Then
            klan = "Staff"
        End If

        If LenB(klan) <> 0 Then
            If Not ToMap Then
                If UserList(UserIndex).showName Then
                    Call WriteCharacterCreate(sndIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.CharIndex, x, Y, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.fx, 999, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Name & " <" & klan & ">", UserTypeColor(UserIndex), UserList(UserIndex).donador, UserList(UserIndex).bandera)
                Else
                    'Hide the name and clan - set privs as normal user
                    Call WriteCharacterCreate(sndIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.CharIndex, x, Y, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.fx, 999, UserList(UserIndex).Char.CascoAnim, vbNullString, UserTypeColor(UserIndex), UserList(UserIndex).donador, 0)
                End If
            Else
                Call AgregarUser(UserIndex, UserList(UserIndex).Pos.map)
            End If
        Else 'if tiene clan
            If Not ToMap Then
                If UserList(UserIndex).showName Then
                    Call WriteCharacterCreate(sndIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.CharIndex, x, Y, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.fx, 999, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Name, UserTypeColor(UserIndex), UserList(UserIndex).donador, UserList(UserIndex).bandera)
                Else
                    Call WriteCharacterCreate(sndIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.CharIndex, x, Y, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.fx, 999, UserList(UserIndex).Char.CascoAnim, vbNullString, UserTypeColor(UserIndex), UserList(UserIndex).donador, 0)
                End If
            Else
                Call AgregarUser(UserIndex, UserList(UserIndex).Pos.map)
            End If
        End If 'if clan
    End If
Exit Sub

hayerror:
    LogError ("MakeUserChar: num: " & err.Number & " desc: " & err.description)
    'Resume Next
    'Call CloseSocket(UserIndex)
End Sub

''
' Checks if the user gets the next level.
'
' @param UserIndex Specifies reference to user

Sub CheckUserLevel(ByVal UserIndex As Integer)
    Dim Pts As Integer
    Dim AumentoHIT As Integer
    Dim AumentoMANA As Integer
    Dim AumentoSTA As Integer
    Dim AumentoHP As Integer
    Dim WasNewbie As Boolean
    Dim Promedio As Double
    Dim aux As Integer
    Dim DistVida(1 To 5) As Integer
    Dim GI As Integer 'Guild Index
    Dim mando As Boolean
    
On Error GoTo Errhandler
    
    mando = False
    
    WasNewbie = EsNewbie(UserIndex)
    'Checkea si alcanzó el máximo nivel
    If UserList(UserIndex).Stats.ELV >= STAT_MAXELV Then
        UserList(UserIndex).Stats.Exp = 0
        UserList(UserIndex).Stats.ELU = 0
        Exit Sub
    End If
            
    With UserList(UserIndex)
        Do While .Stats.Exp >= .Stats.ELU
            
            'Checkea si alcanzó el máximo nivel
            If .Stats.ELV >= STAT_MAXELV Then
                .Stats.Exp = 0
                .Stats.ELU = 0
                Exit Sub
            End If
            
            If .Stats.ELV = 1 Then
                Pts = 10
            Else
                'For multiple levels being rised at once
                Pts = Pts + 5
            End If
            
            .Stats.ELV = .Stats.ELV + 1
            
            .Stats.Exp = .Stats.Exp - .Stats.ELU
            
            'Nueva subida de exp x lvl. Pablo (ToxicWaste)
            If .Stats.ELV < 15 Then
                .Stats.ELU = .Stats.ELU * 1.5
                
            ElseIf .Stats.ELV < 21 Then
                .Stats.ELU = .Stats.ELU * 1.35
            ElseIf .Stats.ELV < 33 Then
                .Stats.ELU = .Stats.ELU * 1.3
            ElseIf .Stats.ELV < 41 Then
                .Stats.ELU = .Stats.ELU * 1.225
            
            'Add Nod kopfnickend
            'Hacemos mas dificiles los ultimos levels
            ElseIf .Stats.ELV < 46 Then
                .Stats.ELU = .Stats.ELU * 1.35
            '/add
            
            ElseIf .Stats.ELV = 50 Then
                .Stats.ELU = 0
                .Stats.Exp = 0
            Else
                .Stats.ELU = .Stats.ELU * 1.25
            End If
            
            'Calculo subida de vida
            Promedio = ModVida(.Clase) - (21 - .Stats.UserAtributos(eAtributos.constitucion)) * 0.5
            aux = RandomNumber(0, 100)
        
            'Es promedio semientero
            DistVida(1) = DistribucionSemienteraVida(1)
            DistVida(2) = DistVida(1) + DistribucionSemienteraVida(2)
            DistVida(3) = DistVida(2) + DistribucionSemienteraVida(3)
            DistVida(4) = DistVida(3) + DistribucionSemienteraVida(4)
            
            If aux <= DistVida(1) Then
                AumentoHP = Promedio + 1.5
            ElseIf aux <= DistVida(2) Then
                AumentoHP = Promedio + 0.5
            ElseIf aux <= DistVida(3) Then
                AumentoHP = Promedio - 0.5
            Else
                AumentoHP = Promedio - 1.5
            End If
            
            AumentoSTA = AumentoSTDef
            AumentoHIT = 1
            AumentoMANA = 0
            
            If mando = False Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, .Pos.x, .Pos.Y))
                Call WriteConsoleMsg(2, UserIndex, "¡Has subido de nivel!", FontTypeNames.FONTTYPE_INFO)
                mando = True
            End If
            
            Select Case .Clase
                Case eClass.Guerrero, eClass.Cazador
                    AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
                
                Case eClass.Mercenario, eClass.Gladiador
                    AumentoHIT = 3

                Case eClass.Paladin
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = 0.94 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Ladron
                    AumentoSTA = AumentoSTLadron
                
                Case eClass.Mago
                    AumentoMANA = 3 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTMago
                
                Case eClass.Leñador
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTLeñador
                
                Case eClass.Minero
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTMinero
                
                Case eClass.Pescador
                    AumentoSTA = AumentoSTPescador
                
                Case eClass.Clerigo
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Druida
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Asesino
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = 0.93 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Bardo
                    AumentoHIT = 2
                    AumentoMANA = 1.685 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Herrero, eClass.Carpintero
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Nigromante
                    AumentoHIT = 2
                    AumentoMANA = 2.4 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                    
                Case Else
                    AumentoHIT = 2

            End Select
            
            'Actualizamos HitPoints
            .Stats.MaxHP = .Stats.MaxHP + AumentoHP
            If .Stats.MaxHP > STAT_MAXHP Then .Stats.MaxHP = STAT_MAXHP
            
            'Actualizamos Stamina
            .Stats.MaxSTA = .Stats.MaxSTA + AumentoSTA
            If .Stats.MaxSTA > STAT_MAXSTA Then .Stats.MaxSTA = STAT_MAXSTA
            
            'Actualizamos Mana
            .Stats.MaxMAN = .Stats.MaxMAN + AumentoMANA
            If .Stats.MaxMAN > STAT_MAXMAN Then .Stats.MaxMAN = STAT_MAXMAN
            
            'Actualizamos Golpe Máximo
            .Stats.MaxHit = .Stats.MaxHit + AumentoHIT
            If .Stats.ELV < 36 Then
                If .Stats.MaxHit > STAT_MAXHIT_UNDER36 Then _
                    .Stats.MaxHit = STAT_MAXHIT_UNDER36
            Else
                If .Stats.MaxHit > STAT_MAXHIT_OVER36 Then _
                    .Stats.MaxHit = STAT_MAXHIT_OVER36
            End If
            
            'Actualizamos Golpe Mínimo
            .Stats.MinHit = .Stats.MinHit + AumentoHIT
            If .Stats.ELV < 36 Then
                If .Stats.MinHit > STAT_MAXHIT_UNDER36 Then _
                    .Stats.MinHit = STAT_MAXHIT_UNDER36
            Else
                If .Stats.MinHit > STAT_MAXHIT_OVER36 Then _
                    .Stats.MinHit = STAT_MAXHIT_OVER36
            End If
            
            'Notificamos al user
            If AumentoHP > 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Has ganado " & AumentoHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
            End If
            If AumentoSTA > 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Has ganado " & AumentoSTA & " puntos de vitalidad.", FontTypeNames.FONTTYPE_INFO)
            End If
            If AumentoMANA > 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Has ganado " & AumentoMANA & " puntos de magia.", FontTypeNames.FONTTYPE_INFO)
            End If
            If AumentoHIT > 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Tu golpe máximo aumentó en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(1, UserIndex, "Tu golpe minimo aumentó en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
            End If

            .Stats.MinHP = .Stats.MaxHP

            If .GrupoIndex > 0 Then _
                Parties(.GrupoIndex).UpdateLevels
            
            If .Stats.ELV < 5 Then
                .Stats.GLD = .Stats.GLD + 80 * ModOroX
            ElseIf .Stats.ELV < 10 Then
                .Stats.GLD = .Stats.GLD + 160 * ModOroX
            ElseIf .Stats.ELV < 14 Then
                .Stats.GLD = .Stats.GLD + 240 * ModOroX
            End If
            
            Call FlushBuffer(UserIndex)
            DoEvents
        Loop
        
        'If it ceased to be a newbie, remove newbie items and get char away from newbie dungeon
        If Not EsNewbie(UserIndex) And WasNewbie Then
            

            Call QuitarNewbieObj(UserIndex)
            
            If .Pos.map = 37 Or .Pos.map = 208 Then
                If UserList(UserIndex).faccion.Ciudadano = 1 Then
                    Call WarpUserChar(UserIndex, 34, 40, 87, True)
                ElseIf UserList(UserIndex).faccion.Republicano = 1 Then
                    Call WarpUserChar(UserIndex, 185, 50, 78, True)
                Else
                    Call WarpUserChar(UserIndex, 20, 50, 50, True)
                End If
                
                Call WriteConsoleMsg(1, UserIndex, "Debes abandonar el Dungeon Newbie.", FontTypeNames.FONTTYPE_INFO)
            End If
                
        End If
        
        'Send all gained skill points at once (if any)
        If Pts > 0 Then
            Call WriteUpdateUserStats(UserIndex)
            Call WriteUpdateExp(UserIndex)
            
            .Stats.SkillPts = .Stats.SkillPts + Pts
            
            Call WriteConsoleMsg(1, UserIndex, "Has ganado un total de " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)
        
            Call FlushBuffer(UserIndex)
        End If
        
    End With

Exit Sub

Errhandler:
    Call LogError("Error en la subrutina CheckUserLevel - Error : " & err.Number & " - Description : " & err.description)
End Sub
'Add Marius Rebalanceo de pjs ya creados
Sub ReBalanceo(ByVal UserIndex As Integer)
    Dim Pts As Integer
    Dim AumentoHIT As Integer
    Dim AumentoMANA As Integer
    Dim AumentoSTA As Integer
    Dim AumentoHP As Integer
    Dim WasNewbie As Boolean
    Dim Promedio As Double
    Dim aux As Integer
    Dim DistVida(1 To 5) As Integer
    Dim GI As Integer 'Guild Index
    Dim mando As Boolean
    
    Dim level As Integer
    
On Error GoTo Errhandler

    With UserList(UserIndex)
        
        level = .Stats.ELV
        'Reseteamos todo como al primer level
        Dim MiInt As Long
        MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.constitucion) \ 3)
        
        .Stats.MaxHP = 15 + MiInt
        .Stats.MinHP = 15 + MiInt
        
        MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Agilidad) \ 6)
        If MiInt = 1 Then MiInt = 2
        
        .Stats.MaxSTA = 20 * MiInt
        .Stats.MinSTA = 20 * MiInt
        
        .Stats.MaxAGU = 100
        .Stats.MinAGU = 100
        
        .Stats.MaxHAM = 100
        .Stats.MinHAM = 100
        
        .Stats.MaxHit = 2
        .Stats.MinHit = 1
    
        .Stats.ELV = 1
        
        If .Clase = eClass.Mago Then  'Cambio en mana inicial (ToxicWaste)
            MiInt = .Stats.UserAtributos(eAtributos.Inteligencia) * 3
            .Stats.MaxMAN = MiInt
            .Stats.MinMAN = MiInt
        ElseIf .Clase = eClass.Clerigo Or .Clase = eClass.Druida _
            Or .Clase = eClass.Bardo Or .Clase = eClass.Asesino _
            Or .Clase = eClass.Nigromante Then
                .Stats.MaxMAN = 50
                .Stats.MinMAN = 50
        Else
            .Stats.MaxMAN = 0
            .Stats.MinMAN = 0
        End If
        
        Do While .Stats.ELV >= level
                        
            .Stats.ELV = .Stats.ELV + 1
                        
            'Calculo subida de vida
            Promedio = ModVida(.Clase) - (21 - .Stats.UserAtributos(eAtributos.constitucion)) * 0.5
            aux = RandomNumber(0, 100)
        
            'Es promedio semientero
            DistVida(1) = DistribucionSemienteraVida(1)
            DistVida(2) = DistVida(1) + DistribucionSemienteraVida(2)
            DistVida(3) = DistVida(2) + DistribucionSemienteraVida(3)
            DistVida(4) = DistVida(3) + DistribucionSemienteraVida(4)
            
            If aux <= DistVida(1) Then
                AumentoHP = Promedio + 1.5
            ElseIf aux <= DistVida(2) Then
                AumentoHP = Promedio + 0.5
            ElseIf aux <= DistVida(3) Then
                AumentoHP = Promedio - 0.5
            Else
                AumentoHP = Promedio - 1.5
            End If
            
            AumentoSTA = AumentoSTDef
            AumentoHIT = 1
            AumentoMANA = 0
            
            Select Case .Clase
                Case eClass.Guerrero, eClass.Cazador
                    AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
                
                Case eClass.Mercenario, eClass.Gladiador
                    AumentoHIT = 3

                Case eClass.Paladin
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = 0.94 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Ladron
                    AumentoSTA = AumentoSTLadron
                
                Case eClass.Mago
                    AumentoMANA = 3 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTMago
                
                Case eClass.Leñador
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTLeñador
                
                Case eClass.Minero
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTMinero
                
                Case eClass.Pescador
                    AumentoSTA = AumentoSTPescador
                
                Case eClass.Clerigo
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Druida
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Asesino
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = 0.93 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Bardo
                    AumentoHIT = 2
                    AumentoMANA = 1.685 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Herrero, eClass.Carpintero
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Nigromante
                    AumentoHIT = 2
                    AumentoMANA = 2.4 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                    
                Case Else
                    AumentoHIT = 2

            End Select
            
            'Actualizamos HitPoints
            .Stats.MaxHP = .Stats.MaxHP + AumentoHP
            If .Stats.MaxHP > STAT_MAXHP Then .Stats.MaxHP = STAT_MAXHP
            
            'Actualizamos Stamina
            .Stats.MaxSTA = .Stats.MaxSTA + AumentoSTA
            If .Stats.MaxSTA > STAT_MAXSTA Then .Stats.MaxSTA = STAT_MAXSTA
            
            'Actualizamos Mana
            .Stats.MaxMAN = .Stats.MaxMAN + AumentoMANA
            If .Stats.MaxMAN > STAT_MAXMAN Then .Stats.MaxMAN = STAT_MAXMAN
            
            'Actualizamos Golpe Máximo
            .Stats.MaxHit = .Stats.MaxHit + AumentoHIT
            If .Stats.ELV < 36 Then
                If .Stats.MaxHit > STAT_MAXHIT_UNDER36 Then _
                    .Stats.MaxHit = STAT_MAXHIT_UNDER36
            Else
                If .Stats.MaxHit > STAT_MAXHIT_OVER36 Then _
                    .Stats.MaxHit = STAT_MAXHIT_OVER36
            End If
            
            'Actualizamos Golpe Mínimo
            .Stats.MinHit = .Stats.MinHit + AumentoHIT
            If .Stats.ELV < 36 Then
                If .Stats.MinHit > STAT_MAXHIT_UNDER36 Then _
                    .Stats.MinHit = STAT_MAXHIT_UNDER36
            Else
                If .Stats.MinHit > STAT_MAXHIT_OVER36 Then _
                    .Stats.MinHit = STAT_MAXHIT_OVER36
            End If
            

            .Stats.MinHP = .Stats.MaxHP
            
            Call FlushBuffer(UserIndex)
            DoEvents
        Loop
        
    End With

Exit Sub

Errhandler:
    Call LogError("Error en la subrutina ReBalanceo - Error : " & err.Number & " - Description : " & err.description)
End Sub
'\Add Marius
Function ClaseVida(ByVal Clase As eClass, ByVal constucion As Byte) As Byte
    Dim min As Single
    Dim max As Single
    
    'If constucion >= 22 Then min = 0: max = 0
    'If constucion = 21 Then min = 0: max = 0
    'If constucion = 20 Then min = 0: max = 0
    'If constucion = 19 Then min = 0: max = 0
    'If constucion = 18 Then min = 0: max = 0
    'If constucion <= 17 Then min = 0: max = 0
    
    Select Case Clase
        Case eClass.Asesino
            If constucion >= 22 Then min = 7: max = 11
            If constucion = 21 Then min = 7: max = 10.5
            If constucion = 20 Then min = 6: max = 10
            If constucion = 19 Then min = 6: max = 9
            If constucion = 18 Then min = 6: max = 8.5
            If constucion <= 17 Then min = 5: max = 8
        Case eClass.Bardo
            If constucion >= 22 Then min = 7: max = 10
            If constucion = 21 Then min = 6: max = 10
            If constucion = 20 Then min = 6: max = 9
            If constucion = 19 Then min = 5: max = 9
            If constucion = 18 Then min = 5: max = 8.5
            If constucion <= 17 Then min = 5: max = 8
        Case eClass.Cazador
            If constucion >= 22 Then min = 8: max = 12
            If constucion = 21 Then min = 8: max = 11.5
            If constucion = 20 Then min = 8: max = 11.33
            If constucion = 19 Then min = 8: max = 11
            If constucion = 18 Then min = 8: max = 10.5
            If constucion <= 17 Then min = 8: max = 10
        Case eClass.Gladiador
            If constucion >= 22 Then min = 8: max = 11
            If constucion = 21 Then min = 8: max = 10.5
            If constucion = 20 Then min = 8: max = 10
            If constucion = 19 Then min = 7: max = 10
            If constucion = 18 Then min = 7: max = 9
            If constucion <= 17 Then min = 6: max = 9
        Case eClass.Clerigo
            If constucion >= 22 Then min = 7: max = 10
            If constucion = 21 Then min = 6: max = 10
            If constucion = 20 Then min = 6: max = 9
            If constucion = 19 Then min = 5: max = 9
            If constucion = 18 Then min = 5: max = 8.5
            If constucion <= 17 Then min = 5: max = 8
        Case eClass.Druida
            If constucion >= 22 Then min = 7: max = 10
            If constucion = 21 Then min = 6: max = 10
            If constucion = 20 Then min = 6: max = 9
            If constucion = 19 Then min = 5: max = 9
            If constucion = 18 Then min = 5: max = 8.5
            If constucion <= 17 Then min = 5: max = 8
        Case eClass.Guerrero
            If constucion >= 22 Then min = 8: max = 12
            If constucion = 21 Then min = 8: max = 11.5
            If constucion = 20 Then min = 8: max = 11.33
            If constucion = 19 Then min = 8: max = 11
            If constucion = 18 Then min = 8: max = 10.5
            If constucion <= 17 Then min = 8: max = 10
        Case eClass.Ladron
            If constucion >= 22 Then min = 7: max = 10
            If constucion = 21 Then min = 7: max = 9.5
            If constucion = 20 Then min = 7: max = 9
            If constucion = 19 Then min = 6: max = 9
            If constucion = 18 Then min = 6: max = 8
            If constucion <= 17 Then min = 5: max = 8
        Case eClass.Mago
            If constucion >= 22 Then min = 5: max = 9
            If constucion = 21 Then min = 5: max = 8.5
            If constucion = 20 Then min = 5: max = 8
            If constucion = 19 Then min = 4: max = 8
            If constucion = 18 Then min = 4: max = 7.5
            If constucion <= 17 Then min = 4: max = 7
        Case eClass.Nigromante
            If constucion >= 22 Then min = 6: max = 10
            If constucion = 21 Then min = 6: max = 9
            If constucion = 20 Then min = 5: max = 9
            If constucion = 19 Then min = 5: max = 8.5
            If constucion = 18 Then min = 5: max = 8
            If constucion <= 17 Then min = 5: max = 7.5
        Case eClass.Paladin
            If constucion >= 22 Then min = 8: max = 11.5
            If constucion = 21 Then min = 8: max = 11.33
            If constucion = 20 Then min = 8: max = 11
            If constucion = 19 Then min = 7: max = 11
            If constucion = 18 Then min = 7: max = 10.5
            If constucion <= 17 Then min = 7: max = 10
        Case eClass.Mercenario
            If constucion >= 22 Then min = 8: max = 11
            If constucion = 21 Then min = 8: max = 10.5
            If constucion = 20 Then min = 8: max = 10
            If constucion = 19 Then min = 7: max = 10
            If constucion = 18 Then min = 7: max = 9
            If constucion <= 17 Then min = 6: max = 9
        Case eClass.Pescador
            If constucion >= 22 Then min = 5: max = 9
            If constucion = 21 Then min = 5: max = 8.5
             If constucion = 20 Then min = 5: max = 8
             If constucion = 19 Then min = 5: max = 7.5
             If constucion = 18 Then min = 5: max = 7
            If constucion <= 17 Then min = 5: max = 6.5
        Case eClass.Leñador
            If constucion >= 22 Then min = 6: max = 9
             If constucion = 21 Then min = 6: max = 8.5
             If constucion = 20 Then min = 6: max = 8
             If constucion = 19 Then min = 6: max = 7.5
             If constucion = 18 Then min = 6: max = 7
            If constucion <= 17 Then min = 6: max = 6.5
        Case eClass.Minero
            If constucion >= 22 Then min = 5: max = 8
             If constucion = 21 Then min = 5: max = 7.5
             If constucion = 20 Then min = 5: max = 7
             If constucion = 19 Then min = 5: max = 6.5
             If constucion = 18 Then min = 5: max = 6
            If constucion <= 17 Then min = 5: max = 5.5
        Case eClass.Sastre, eClass.Herrero
            If constucion >= 22 Then min = 6: max = 8
             If constucion = 21 Then min = 6: max = 7.5
             If constucion = 20 Then min = 6: max = 7
             If constucion = 19 Then min = 5: max = 6.5
             If constucion = 18 Then min = 5: max = 6
            If constucion <= 17 Then min = 5: max = 5.5
    End Select
    ClaseVida = RandomNumber(min, max)
End Function
Public Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean
    PuedeAtravesarAgua = UserList(UserIndex).flags.Navegando = 1
End Function

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As eHeading)
'*************************************************
'Author: Unknown
'Last modified: 30/03/2009
'Moves the char, sending the message to everyone in range.
'30/03/2009: ZaMa - Now it's legal to move where a casper is, changing its pos to where the moving char was.
'*************************************************
    Dim nPos As WorldPos
    Dim sailing As Boolean
    Dim CasperIndex As Integer
    Dim CasperHeading As eHeading
    Dim CasPerPos As WorldPos
    Dim isAdminInvi As Boolean
    
    sailing = PuedeAtravesarAgua(UserIndex)
    nPos = UserList(UserIndex).Pos
    Call HeadtoPos(nHeading, nPos)
    
    isAdminInvi = (UserList(UserIndex).flags.AdminInvisible = 1)
    
    'Add Marius No monturas en dungeons
    'Debug.Print "Zona:" & MapInfo(UserList(UserIndex).Pos.map).Zona
    'Debug.Print "Montado:" & UserList(UserIndex).flags.Montando
    If UserList(UserIndex).flags.Montando = 1 And (MapInfo(UserList(UserIndex).Pos.map).Zona = "DUNGEON" Or UserList(UserIndex).Pos.map = Bandera_mapa) Then
        
        UserList(UserIndex).flags.Montando = 0
        If UserList(UserIndex).flags.Muerto = 0 Then
            UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
            If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                UserList(UserIndex).Char.body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
            Else
                Call DarCuerpoDesnudo(UserIndex)
            End If
        Else
            UserList(UserIndex).Char.body = iCuerpoMuerto
            UserList(UserIndex).Char.Head = iCabezaMuerto
            UserList(UserIndex).Char.ShieldAnim = NingunEscudo
            UserList(UserIndex).Char.WeaponAnim = NingunArma
            UserList(UserIndex).Char.CascoAnim = NingunCasco
        End If
        UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MonturaSlot).Equipped = 0
        Call UpdateUserInv(False, UserIndex, UserList(UserIndex).Invent.MonturaSlot)
            
        UserList(UserIndex).Invent.MonturaObjIndex = 0
        UserList(UserIndex).Invent.MonturaSlot = 0
        
        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        Call WriteEquitateToggle(UserIndex)
        
    End If
    '\Add
    
    
    
        
   If MoveToLegalPos(UserList(UserIndex).Pos.map, nPos.x, nPos.Y, sailing, Not sailing) Then
        'si no estoy solo en el mapa...
        If MapInfo(UserList(UserIndex).Pos.map).NumUsers > 1 Then
            
            CasperIndex = MapData(UserList(UserIndex).Pos.map, nPos.x, nPos.Y).UserIndex
            'Si hay un usuario, y paso la validacion, entonces es un casper
            If CasperIndex > 0 Then
                ' Los admins invisibles no pueden patear caspers
                
                If Not isAdminInvi Then
                
                    If UserList(CasperIndex).flags.Muerto Then
                        CasperHeading = InvertHeading(nHeading)
                        CasPerPos = UserList(CasperIndex).Pos
                        Call HeadtoPos(CasperHeading, CasPerPos)
        
                        With UserList(CasperIndex)
                            
                            ' Si es un admin invisible, no se avisa a los demas clientes
                            If Not .flags.AdminInvisible = 1 Then _
                                Call SendData(SendTarget.ToPCAreaButIndex, CasperIndex, PrepareMessageCharacterMove(.Char.CharIndex, CasPerPos.x, CasPerPos.Y))
                            
                            Call WriteForceCharMove(CasperIndex, CasperHeading)
                                
                            'Update map and user pos
                            .Pos = CasPerPos
                            .Char.heading = CasperHeading
                            MapData(.Pos.map, CasPerPos.x, CasPerPos.Y).UserIndex = CasperIndex
                        End With
                    
                        'Actualizamos las áreas de ser necesario
                        Call ModAreas.CheckUpdateNeededUser(CasperIndex, CasperHeading)
                    End If
                    
                End If
            End If

            ' Si es un admin invisible, no se avisa a los demas clientes
            If Not isAdminInvi Then _
                Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, nPos.x, nPos.Y))
            
        End If
        
        ' Los admins invisibles no pueden patear caspers
        If Not (isAdminInvi And (CasperIndex <> 0)) Then
            Dim oldUserIndex As Integer
            With UserList(UserIndex)
                oldUserIndex = MapData(.Pos.map, .Pos.x, .Pos.Y).UserIndex
                
                ' Si no hay intercambio de pos con nadie
                If oldUserIndex = UserIndex Then
                    MapData(.Pos.map, .Pos.x, .Pos.Y).UserIndex = 0
                End If
                
                .Pos = nPos
                .Char.heading = nHeading
                MapData(.Pos.map, .Pos.x, .Pos.Y).UserIndex = UserIndex
                
                Call DoTileEvents(UserIndex, .Pos.map, .Pos.x, .Pos.Y)
            End With
            
            'Actualizamos las áreas de ser necesario
            Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading)
        Else
            Call WritePosUpdate(UserIndex)
        End If
    Else
        'Des Marius dava un error en el errors.log
        'Fix Marius Clones Que paja que es estooooo!!
        'Debug.Print "UI: " & MapData(nPos.map, nPos.x, nPos.Y).UserIndex
        'If UserList(MapData(nPos.map, nPos.x, nPos.Y).UserIndex).Pos.map <> nPos.map Or UserList(MapData(nPos.map, nPos.x, nPos.Y).UserIndex).Pos.x <> nPos.x Or UserList(MapData(nPos.map, nPos.x, nPos.Y).UserIndex).Pos.Y <> nPos.Y Then
        '    Call QuitarUser(MapData(nPos.map, nPos.x, nPos.Y).UserIndex, nPos.map)
        '    MapData(nPos.map, nPos.x, nPos.Y).UserIndex = 0
        '    Call WritePosUpdate(UserIndex)
        'End If
        '\Fix
    End If
    
    If UserList(UserIndex).Counters.Trabajando Then _
        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1

    If UserList(UserIndex).Counters.Ocultando Then _
        UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
End Sub

Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading
'*************************************************
'Author: ZaMa
'Last modified: 30/03/2009
'Returns the heading opposite to the one passed by val.
'*************************************************
    Select Case nHeading
        Case eHeading.EAST
            InvertHeading = WEST
        Case eHeading.WEST
            InvertHeading = EAST
        Case eHeading.SOUTH
            InvertHeading = NORTH
        Case eHeading.NORTH
            InvertHeading = SOUTH
    End Select
End Function


Sub ChangeUserInv(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef Object As UserObj)
    UserList(UserIndex).Invent.Object(Slot) = Object
    Call WriteChangeInventorySlot(UserIndex, Slot)
End Sub

Function NextOpenCharIndex() As Integer
    Dim loopC As Long
    
    For loopC = 1 To MAXCHARS
        If CharList(loopC) = 0 Then
            NextOpenCharIndex = loopC
            NumChars = NumChars + 1
            
            If loopC > LastChar Then _
                LastChar = loopC
            
            Exit Function
        End If
    Next loopC
    Debug.Print "Lstchar:" & LastChar
End Function


Function NextOpenUser() As Integer
    '////Modificacion por Castelli
    'Se agrego UserList(LoopC).flags.UserLogged = False para asegurarce_
    'de que el usuario ya este offline y no agregar un index_
    'mal en un caso realmente extraño :S =P
    '////Modificacion por Castelli
    
    Dim loopC As Long
    
    For loopC = 1 To MaxUsers + 1
     
        If loopC > MaxUsers Then Exit For
        If (UserList(loopC).ConnID = -1 And UserList(loopC).flags.UserLogged = False) Then Exit For
    
    Next loopC
    
    NextOpenUser = loopC
End Function

'\Add Marius
Public Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim GuildI As Integer
    
    With UserList(UserIndex)
        Call WriteConsoleMsg(1, sendIndex, "Estadísticas de: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(1, sendIndex, "Nivel: " & .Stats.ELV & "  EXP: " & .Stats.Exp & "/" & .Stats.ELU, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(1, sendIndex, "Salud: " & .Stats.MinHP & "/" & .Stats.MaxHP & "  Maná: " & .Stats.MinMAN & "/" & .Stats.MaxMAN & "  Energía: " & .Stats.MinSTA & "/" & .Stats.MaxSTA, FontTypeNames.FONTTYPE_INFO)
        
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call WriteConsoleMsg(1, sendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.MinHit & "/" & .Stats.MaxHit & " (" & ObjData(.Invent.WeaponEqpObjIndex).MinHit & "/" & ObjData(.Invent.WeaponEqpObjIndex).MaxHit & ")", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(1, sendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.MinHit & "/" & .Stats.MaxHit, FontTypeNames.FONTTYPE_INFO)
        End If
        
        If .Invent.ArmourEqpObjIndex > 0 Then
            If .Invent.EscudoEqpObjIndex > 0 Then
                Call WriteConsoleMsg(1, sendIndex, "(CUERPO) Mín Def/Máx Def: " & ObjData(.Invent.ArmourEqpObjIndex).MinDef + ObjData(.Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef + ObjData(.Invent.EscudoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, sendIndex, "(CUERPO) Mín Def/Máx Def: " & ObjData(.Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(1, sendIndex, "(CUERPO) Mín Def/Máx Def: 0", FontTypeNames.FONTTYPE_INFO)
        End If
        
        If .Invent.CascoEqpObjIndex > 0 Then
            Call WriteConsoleMsg(1, sendIndex, "(CABEZA) Mín Def/Máx Def: " & ObjData(.Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(.Invent.CascoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(1, sendIndex, "(CABEZA) Mín Def/Máx Def: 0", FontTypeNames.FONTTYPE_INFO)
        End If
        
        GuildI = .GuildIndex
        If GuildI > 0 Then
            Call WriteConsoleMsg(1, sendIndex, "Clan: " & modGuilds.GuildName(GuildI), FontTypeNames.FONTTYPE_INFO)
            If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(.Name) Then
                Call WriteConsoleMsg(1, sendIndex, "Status: Líder", FontTypeNames.FONTTYPE_INFO)
            End If
            'guildpts no tienen objeto
        End If
        
        Call WriteConsoleMsg(1, sendIndex, "Oro: " & .Stats.GLD & "  Posición: " & .Pos.map & "," & .Pos.x & "," & .Pos.Y, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(1, sendIndex, "Dados: " & .Stats.UserAtributos(eAtributos.Fuerza) & ", " & .Stats.UserAtributos(eAtributos.Agilidad) & ", " & .Stats.UserAtributos(eAtributos.Inteligencia) & ", " & .Stats.UserAtributos(eAtributos.Carisma) & ", " & .Stats.UserAtributos(eAtributos.constitucion), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub
'\Add

Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
'**********************************************
'Author: Unknown
'Last Modification: 06/28/2008
'24/01/2007 -> Pablo (ToxicWaste): Agrego para que se actualize el tag si corresponde.
'24/07/2007 -> Pablo (ToxicWaste): Guardar primero que ataca NPC y el que atacas ahora.
'06/28/2008 -> NicoNZ: Los elementales al atacarlos por su amo no se paran más al lado de él sin hacer nada.
'**********************************************
    Dim EraCriminal As Boolean
    
    'Guardamos el usuario que ataco el npc.
    Npclist(NpcIndex).flags.AttackedBy = UserIndex
    
    'Npc que estabas atacando.
    Dim LastNpcHit As Integer
    LastNpcHit = UserList(UserIndex).flags.NPCAtacado
    'Guarda el NPC que estas atacando ahora.
    UserList(UserIndex).flags.NPCAtacado = NpcIndex
    
    'Revisamos robo de npc.
    'Guarda el primer nick que lo ataca.
    If Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then
        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
            End If
        End If
        Npclist(NpcIndex).flags.AttackedFirstBy = UserList(UserIndex).Name
    ElseIf Npclist(NpcIndex).flags.AttackedFirstBy <> UserList(UserIndex).Name Then
        'Estas robando NPC
        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
            End If
        End If
    End If
    
    If Npclist(NpcIndex).MaestroUser > 0 Then
        If Npclist(NpcIndex).MaestroUser <> UserIndex Then
            Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)
        End If
    End If
    

End Sub

Public Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean

    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1 Then
            PuedeApuñalar = UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR _
                        Or UserList(UserIndex).Clase = eClass.Asesino
        End If
    End If
End Function

Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)


On Error GoTo err


    With UserList(UserIndex)
        If .flags.Hambre = 0 And .flags.Sed = 0 Then
            
            If .Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
            
            Dim Lvl As Integer
            Lvl = .Stats.ELV
            
            If Lvl > UBound(LevelSkill) Then Lvl = UBound(LevelSkill)
            
            If .Stats.UserSkills(Skill) >= LevelSkill(Lvl).LevelValue Then Exit Sub
            
            Dim Prob As Integer
            
            If Lvl <= 3 Then
                Prob = 25
            ElseIf Lvl > 3 And Lvl < 6 Then
                Prob = 35
            ElseIf Lvl >= 6 And Lvl < 10 Then
                Prob = 40
            ElseIf Lvl >= 10 And Lvl < 20 Then
                Prob = 45
            Else
                Prob = 50
            End If
            
            'Mannakia
            If .Invent.MagicIndex <> 0 Then
                If ObjData(.Invent.MagicIndex).EfectoMagico = eMagicType.Experto Then
                    Prob = Prob - Porcentaje(Prob, ObjData(.Invent.MagicIndex).CuantoAumento)
                End If
            End If
            'Mannakia
            Prob = 15
            If RandomNumber(1, Prob) < 10 Then
                .Stats.UserSkills(Skill) = .Stats.UserSkills(Skill) + 1

                .Stats.Exp = .Stats.Exp + (18 * .Stats.ELV)
               ' If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
                
                Call WriteMsg(UserIndex, 40, CStr(Skill), CStr(.Stats.UserSkills(Skill)))
                Call WriteMsg(UserIndex, 21, CStr(18 * .Stats.ELV))
                
                Call WriteUpdateExp(UserIndex)
                Call CheckUserLevel(UserIndex)
                
                Call FlushBuffer(UserIndex)
            End If
        End If
    End With
    
    Exit Sub
    
    
err:
    
End Sub

Sub PerderDuelo(ByVal UserIndex As Integer)
    If UserList(UserIndex).flags.inDuelo = False Then Exit Sub
   
    Dim IDA As Integer
    IDA = UserList(UserIndex).flags.vicDuelo
    If IDA <> 0 Then
        UserList(IDA).flags.inDuelo = 0
        UserList(IDA).flags.vicDuelo = 0
        Call WriteConsoleMsg(1, UserIndex, "¡Has perdido el duelo contra " & UserList(IDA).Name & "!", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(1, IDA, "¡Has ganado el duelo contra " & UserList(UserIndex).Name & "!", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    UserList(UserIndex).flags.vicDuelo = 0
    UserList(UserIndex).flags.inDuelo = 0

End Sub
''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'

Sub UserDie(ByVal UserIndex As Integer)
'************************************************
'Author: Uknown
'Last Modified: 13/02/2009
'04/15/2008: NicoNZ - Ahora se resetea el counter del invi
'13/02/2009: Ahora se borran las mascotas cuando moris en agua.
'************************************************
On Error GoTo ErrorHandler
    Dim i As Long
    Dim aN As Integer
    Dim DropObjs As Integer
    Dim NewPos As WorldPos
    Dim Drops As obj

    With UserList(UserIndex)
        
        'Add Nod Kopfnickend
        'Solo le saca vida a los que no son Dioses
        If EsCONSE(UserIndex) And Not EsVIP(UserIndex) Then
            .Stats.MinHP = .Stats.MaxHP
            .Stats.MinMAN = .Stats.MaxMAN
            .flags.Envenenado = 0
            .Counters.Veneno = 0
            .flags.Incinerado = 0
            .flags.Ceguera = 0
            Call WriteUpdateUserStats(UserIndex)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateCharParticle(.Char.CharIndex, 119))
            Call DecirPalabrasMagicas("Divinum Protection", UserIndex)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(236, .Pos.x, .Pos.Y))
            
            Exit Sub
        End If
        
        'Sonido
        Call ReproducirSonido(SendTarget.ToPCArea, UserIndex, 11)

        'Quitar el dialogo del user muerto
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
        .Stats.MinHP = 0
        '.Stats.MinSTA = 0
        .flags.AtacadoPorUser = 0
        .flags.Envenenado = 0
        .Counters.Veneno = 0
        
        .flags.Metamorfosis = 0
        .flags.Incinerado = 0
        
        .flags.Muerto = 1

        aN = .flags.AtacadoPorNpc
        If aN > 0 Then
            Npclist(aN).Movement = Npclist(aN).flags.OldMovement
            Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
            Npclist(aN).flags.AttackedBy = 0
        End If
        
        aN = .flags.NPCAtacado
        If aN > 0 Then
            If Npclist(aN).flags.AttackedFirstBy = .Name Then
                Npclist(aN).flags.AttackedFirstBy = vbNullString
            End If
        End If
        .flags.AtacadoPorNpc = 0
        .flags.NPCAtacado = 0
        
        '<<<< Paralisis >>>>
        If .flags.Paralizado = 1 Then
            .flags.Paralizado = 0
            Call WriteParalizeOK(UserIndex)
        End If
        
        '<<< Estupidez >>>
        If .flags.Estupidez = 1 Then
            .flags.Estupidez = 0
            Call WriteDumbNoMore(UserIndex)
        End If
        
        '<<<< Descansando >>>>
        If .flags.Descansar Then
            .flags.Descansar = False
            Call WriteRestOK(UserIndex)
        End If
        
        '<<<< Meditando >>>>
        If .flags.Meditando Then
            .flags.Meditando = False
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageDestCharParticle(UserList(UserIndex).Char.CharIndex, ParticleToLevel(UserIndex)))
            Call WriteMeditateToggle(UserIndex)
        End If
        
        '<<<< Invisible >>>>
        If .flags.Invisible = 1 Or .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .flags.Invisible = 0
            .Counters.TiempoOculto = 0
            .Counters.Invisibilidad = 0
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
        End If
        
        If (TriggerZonaPelea(UserIndex, UserIndex) <> eTrigger6.TRIGGER6_PERMITE And MapInfo(UserList(UserIndex).Pos.map).Pk = True) Then
            
            'Add Nod kopfnickend
            'No se caen las cosas en mapa Captura la bandera
            'Nuevo sistema de sacris con estados
            'Piedra de resurreccion
            If Not EsCONSE(UserIndex) And UserList(UserIndex).Pos.map <> Bandera_mapa Then
                
                DropObjs = Have_Obj_Slot(1614, UserIndex)
                If DropObjs > 0 Then 'Piedra de resurreción
                    
                    Call QuitarUserInvItem(UserIndex, DropObjs, 1)
                    Call WriteConsoleMsg(1, UserIndex, "La Piedra de resurreción esta brillando!", FontTypeNames.FONTTYPE_BROWNI)
                    Call FlushBuffer(UserIndex)
                    
                    Call UpdateUserInv(False, UserIndex, DropObjs)
                    
                    Call RevivirUsuario(UserIndex)
                    .Counters.IntervaloRevive = .Counters.IntervaloRevive + 5000
                Else
                
                    DropObjs = TieneSacri(UserIndex)
                    If DropObjs = 0 Then
                        ' << Si es newbie no pierde el inventario >>
                        If Not EsNewbie(UserIndex) Then
                            Call TirarTodo(UserIndex)
                        Else
                            Call TirarTodosLosItemsNoNewbies(UserIndex)
                        End If
                    Else
                        'ReMod Marius Pendiente de sacrificio con 3 estados
                        'Tiramos el sacri 3/3 2/3 1/3
                        If .Invent.Object(DropObjs).ObjIndex = 1081 Then 'Pendiente del Sacrificio 3/3
                            'Debug.Print "2/3"
                            Drops.ObjIndex = 1602 '2/3
                            Drops.Amount = 1
                            
                            TileLibre UserList(UserIndex).Pos, NewPos, Drops, True, True
                            Call TirarItemAlPiso(NewPos, Drops)
                        ElseIf .Invent.Object(DropObjs).ObjIndex = 1602 Then 'Pendiente del Sacrificio 2/3
                            'Debug.Print "1/3"
                            Drops.ObjIndex = 1603 '1/3
                            Drops.Amount = 1
                            
                            TileLibre UserList(UserIndex).Pos, NewPos, Drops, True, True
                            Call TirarItemAlPiso(NewPos, Drops)
                        ElseIf .Invent.Object(DropObjs).ObjIndex = 1603 Then 'Pendiente del Sacrificio 1/3
                            'Debug.Print "0/3"
                            'No cae nada, se destruyó el pendiente
                        Else
                            'Debug.Print "Error"
                        End If
                        
                        'Le saca del inventario el pendiente
                        Call QuitarUserInvItem(UserIndex, DropObjs, 1)
                        Call UpdateUserInv(False, UserIndex, DropObjs)
                        
                        'Call DropObj(UserIndex, DropObjs, 1, NewPos.map, NewPos.x, NewPos.Y)
                        '\ReMod
                    End If
                
                End If
                
            End If
            
        End If
        
        ' DESEQUIPA TODOS LOS OBJETOS
        'desequipar armadura
        If .Invent.ArmourEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
        End If
        
        'desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
        End If
        
        'desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
        End If
        
        'desequipar herramienta
        If .Invent.AnilloEqpSlot > 0 Then
            Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)
        End If
        
        'desequipar municiones
        If .Invent.MunicionEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
        End If
        
        'desequipamos items macigos
        If .Invent.MagicIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.MagicSlot)
        End If
        
        'desequipar escudo
        If .Invent.EscudoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
        End If
        
        ' << Reseteamos los posibles FX sobre el personaje >>
        If .Char.loops = INFINITE_LOOPS Then
            .Char.fx = 0
            .Char.loops = 0
        End If
        
        ' << Restauramos los atributos >>
        If .flags.TomoPocion = True Then
            For i = 1 To 5
                .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
            Next i
        End If
        
        Call WriteAgilidad(UserIndex)
        Call WriteFuerza(UserIndex)
        
        '<< Cambiamos la apariencia del char >>
        If .flags.Navegando = 0 Then
            .Char.body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
        Else
            .Char.body = iFragataFantasmal
        End If
        
        If .flags.Montando = 1 Then
            .flags.Montando = 0
            Call WriteEquitateToggle(UserIndex)
        End If
        
        For i = 1 To MAXMASCOTAS
            If .MascotasIndex(i) > 0 Then
                Call MuereNpc(.MascotasIndex(i), 0)
            ' Si estan en agua o zona segura
            Else
                .MascotasType(i) = 0
            End If
        Next i
        
        .Stats.VecesMuertos = .Stats.VecesMuertos + 1
        
        .NroMascotas = 0
        
        
        '/// CASTELLI , desinvocamos a la mascota si la tiene invocada
        If .masc.invocado = True Then
            Call desinvocarfami(UserIndex)
        End If
         '/// CASTELLI , desinvocamos a la mascota si la tiene invocada
        
        If UserList(UserIndex).flags.inDuelo = 1 Then
            Call PerderDuelo(UserIndex)
        End If
        
        If UserList(UserIndex).evento <> 0 Then
            Call salir_arena(UserIndex)
        End If
        
         'Add Marius Captura la Bandera
        If UserList(UserIndex).Pos.map = Bandera_mapa Then
            Call Bandera_muere(UserIndex)
        End If
        '\Add
         
        '<< Actualizamos clientes >>
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
        Call WriteUpdateUserStats(UserIndex)
        
        '<<Castigos por Grupo>>
        'If .GrupoIndex > 0 Then
          '  Call mdGrupo.ObtenerExito(UserIndex, .Stats.ELV * -10 * mdGrupo.CantMiembros(UserIndex), .Pos.map, .Pos.X, .Pos.Y)
        'End If
        
        
        If TriggerZonaPelea(UserIndex, UserIndex) = TRIGGER6_PERMITE Then
            RevivirUsuario UserIndex
        End If
        
       Call ControlarPortalLum(UserIndex)
        
    End With
Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & err.Number & " Descripción: " & err.description)
End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)

    If EsNewbie(Muerto) Then Exit Sub
    
    'Des Marius Para mas diversion de los usrs
        'If mapasEspeciales(Muerto) Then Exit Sub
        'If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
    '\Des
    
    'Add Marius antitrucheo
    If UserList(Atacante).Pos.map = Prision.map Then Exit Sub
    If (EsCONSE(Muerto)) Then Exit Sub
    
    'Si esta denudo el muerto no cuenta la muerte
    If UserList(Muerto).flags.Desnudo = 1 Then Exit Sub
    
    Call AgregarListaMuertos(Muerto, Atacante)
    If YaMatoUsuario(Muerto, Atacante) > 1 Then Exit Sub
    '\Add
    
    With UserList(Atacante)
        
        If UserList(Muerto).faccion.ArmadaReal = 1 Then
            .faccion.ArmadaMatados = .faccion.ArmadaMatados + 1
            Exit Sub
        End If
        
        If UserList(Muerto).faccion.Milicia = 1 Then
            .faccion.MilicianosMatados = .faccion.MilicianosMatados + 1
            Exit Sub
        End If
        
        If UserList(Muerto).faccion.FuerzasCaos = 1 Then
            .faccion.CaosMatados = .faccion.CaosMatados + 1
            Exit Sub
        End If
        
        If UserList(Muerto).faccion.Renegado = 1 Then
            .faccion.RenegadosMatados = .faccion.RenegadosMatados + 1
            Exit Sub
        End If
        
        If UserList(Muerto).faccion.Republicano = 1 Then
            .faccion.RepublicanosMatados = .faccion.RepublicanosMatados + 1
            Exit Sub
        End If
        
        If UserList(Muerto).faccion.Ciudadano Then
            .faccion.CiudadanosMatados = .faccion.CiudadanosMatados + 1
            Exit Sub
        End If
        
    End With
End Sub

Sub TileLibre(ByRef Pos As WorldPos, ByRef nPos As WorldPos, ByRef obj As obj, ByRef Agua As Boolean, ByRef Tierra As Boolean)
'**************************************************************
'Author: Unknown
'Last Modify Date: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
'**************************************************************
    Dim loopC As Integer
    Dim tX As Long
    Dim tY As Long
    Dim hayobj As Boolean
    
    hayobj = False
    nPos.map = Pos.map
    nPos.x = 0
    nPos.Y = 0
    
    Do While Not LegalPos(Pos.map, nPos.x, nPos.Y, Agua, Tierra) Or hayobj
        
        If loopC > 15 Then
            Exit Do
        End If
        
        For tY = Pos.Y - loopC To Pos.Y + loopC
            For tX = Pos.x - loopC To Pos.x + loopC
                
                If LegalPos(nPos.map, tX, tY, Agua, Tierra) Then
                    'We continue if: a - the item is different from 0 and the dropped item or b - the amount dropped + amount in map exceeds MAX_INVENTORY_OBJS
                    hayobj = (MapData(nPos.map, tX, tY).ObjInfo.ObjIndex > 0 And MapData(nPos.map, tX, tY).ObjInfo.ObjIndex <> obj.ObjIndex)
                    If Not hayobj Then _
                        hayobj = (MapData(nPos.map, tX, tY).ObjInfo.Amount + obj.Amount > MAX_INVENTORY_OBJS)
                    If Not hayobj And MapData(nPos.map, tX, tY).TileExit.map = 0 Then
                        nPos.x = tX
                        nPos.Y = tY
                        
                        'break both fors
                        tX = Pos.x + loopC
                        tY = Pos.Y + loopC
                    End If
                End If
            
            Next tX
        Next tY
        
        loopC = loopC + 1
    Loop
End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal fx As Boolean)
    Dim OldMap As Integer
    Dim OldX As Integer
    Dim OldY As Integer

    With UserList(UserIndex)
    
    'If .Pos.map And .Pos.x And .Pos.Y Then Exit Sub
    
        'Quitar el dialogo
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        Call WriteRemoveAllDialogs(UserIndex)
        
        OldMap = .Pos.map
        OldX = .Pos.x
        OldY = .Pos.Y
        

        Call EraseUserChar(UserIndex)
        
        If OldMap <> map Then
            Call WriteChangeMap(UserIndex, map, MapInfo(.Pos.map).MapVersion)
            
            
            'Add Marius Cuando pasas de mapa y no esta permitido invi, te lo saca. Sacado de la 0.13.3 xD
            If .flags.Privilegios And PlayerType.User Then 'El chequeo de invi/ocultar solo afecta a Usuarios (C4b3z0n)
                'Chequeo de flags de mapa por invisibilidad (C4b3z0n)
                If MapInfo(map).InviSinEfecto > 0 And .flags.Invisible = 1 Then
                    .flags.Invisible = 0
                    .Counters.Invisibilidad = 0
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                    Call WriteConsoleMsg(2, UserIndex, "Has vuelto a ser visible ya que no esta permitida la invisibilidad en este mapa.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
            '\Add
            
            
            Call WritePlayMidi(UserIndex, val(ReadField(1, MapInfo(map).Music, 45)))
            
            'Update new Map Users
            MapInfo(map).NumUsers = MapInfo(map).NumUsers + 1
            
            'Update old Map Users
            MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
            If MapInfo(OldMap).NumUsers < 0 Then
                MapInfo(OldMap).NumUsers = 0
            End If
        End If
        
        .Pos.x = x
        .Pos.Y = Y
        .Pos.map = map
        
        Call MakeUserChar(True, map, UserIndex, map, x, Y)
        Call WriteUserCharIndexInServer(UserIndex)
        
        Call DoTileEvents(UserIndex, map, x, Y)
        
        'Force a flush, so user index is in there before it's destroyed for teleporting
        Call FlushBuffer(UserIndex)
        
        'Seguis invisible al pasar de mapa
        If (.flags.Invisible = 1 Or .flags.Oculto = 1) And (Not .flags.AdminInvisible = 1) Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
        End If
        
        If fx And .flags.AdminInvisible = 0 Then 'FX
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_WARP, x, Y))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
        End If
        
        If .NroMascotas Then Call WarpMascotas(UserIndex)
    End With
End Sub

Private Sub WarpMascotas(ByVal UserIndex As Integer)
'************************************************
'Author: Uknown
'Last Modified: 13/02/2009
'13/02/2009: ZaMa - Arreglado respawn de mascotas al cambiar de mapa.
'13/02/2009: ZaMa - Las mascotas no regeneran su vida al cambiar de mapa (Solo entre mapas inseguros).
'************************************************
    Dim i As Integer
    Dim petType As Integer
    Dim PetRespawn As Boolean
    Dim PetTiempoDeVida As Integer
    Dim NroPets As Integer
    Dim InvocadosMatados As Integer
    Dim canWarp As Boolean
    Dim index As Integer
    Dim iMinHP As Integer
    
    NroPets = UserList(UserIndex).NroMascotas
    canWarp = (MapInfo(UserList(UserIndex).Pos.map).Pk = True)
    
    For i = 1 To MAXMASCOTAS
        index = UserList(UserIndex).MascotasIndex(i)
        
        If index > 0 Then
            ' si la mascota tiene tiempo de vida > 0 significa q fue invocada => we kill it
            If Npclist(index).Contadores.TiempoExistencia > 0 Then
                Call QuitarNPC(index)
                UserList(UserIndex).MascotasIndex(i) = 0
                InvocadosMatados = InvocadosMatados + 1
                NroPets = NroPets - 1
                
                petType = 0
            Else
                'Store data and remove NPC to recreate it after warp
                'PetRespawn = Npclist(index).flags.Respawn = 0
                petType = UserList(UserIndex).MascotasType(i)
                'PetTiempoDeVida = Npclist(index).Contadores.TiempoExistencia
                
                ' Guardamos el hp, para restaurarlo uando se cree el npc
                iMinHP = Npclist(index).Stats.MinHP
                
                Call QuitarNPC(index)
                
                ' Restauramos el valor de la variable
                UserList(UserIndex).MascotasType(i) = petType

            End If
        ElseIf UserList(UserIndex).MascotasType(i) > 0 Then
            'Store data and remove NPC to recreate it after warp
            PetRespawn = True
            petType = UserList(UserIndex).MascotasType(i)
            PetTiempoDeVida = 0
        Else
            petType = 0
        End If
        
        If petType > 0 And canWarp Then
            index = SpawnNpc(petType, UserList(UserIndex).Pos, False, PetRespawn)
            UserList(UserIndex).MascotasIndex(i) = index

            ' Nos aseguramos de que conserve el hp, si estaba dañado
            Npclist(index).Stats.MinHP = IIf(iMinHP = 0, Npclist(index).Stats.MinHP, iMinHP)
            
            'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
            If index = 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Tus mascotas no pueden transitar este mapa.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            Npclist(index).MaestroUser = UserIndex
            Npclist(index).Movement = TipoAI.SigueAmo
            Npclist(index).Target = 0
            Npclist(index).TargetNPC = 0
            Npclist(index).Contadores.TiempoExistencia = PetTiempoDeVida
            Call FollowAmo(index)
        End If
    Next i
    
    If InvocadosMatados > 0 Then
        Call WriteConsoleMsg(1, UserIndex, "Pierdes el control de tus mascotas invocadas.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    If Not canWarp Then
        Call WriteConsoleMsg(1, UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    UserList(UserIndex).NroMascotas = NroPets
End Sub

''
' Se inicia la salida de un usuario.
'
' @param    UserIndex   El index del usuario que va a salir

Sub Cerrar_Usuario(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 09/04/08 (NicoNZ)
'
'***************************************************

    Dim isNotVisible As Boolean
    Dim diezSeg As Boolean
    
    If UserList(UserIndex).Counters.Saliendo = True Then
        CancelExit UserIndex
    ElseIf UserList(UserIndex).flags.UserLogged And Not UserList(UserIndex).Counters.Saliendo Then
        UserList(UserIndex).Counters.Saliendo = True

        If EsFacc(UserIndex) Or _
            UserList(UserIndex).flags.Muerto = 1 Or _
            MapInfo(UserList(UserIndex).Pos.map).Pk = False Then
            
            UserList(UserIndex).Counters.salir = 0
        
        'Add Nod Kopfnickend Donadores sale en 5 seg
        ElseIf UserList(UserIndex).donador = True Then
            UserList(UserIndex).Counters.salir = 5
        '\Add
        
        Else
            UserList(UserIndex).Counters.salir = IntervaloCerrarConexion
        End If

        
        
        isNotVisible = (UserList(UserIndex).flags.Oculto Or UserList(UserIndex).flags.Invisible)
        'Agregamos los conse se pueden deslogear invisibles sin revelar su pos
        If isNotVisible And Not EsCONSE(UserIndex) Then
            UserList(UserIndex).flags.Oculto = 0
            UserList(UserIndex).flags.Invisible = 0
            Call WriteConsoleMsg(1, UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
        End If
        
        If UserList(UserIndex).flags.Trabajando = True Then
            Call WriteConsoleMsg(1, UserIndex, "Dejas de trabajar.", FontTypeNames.FONTTYPE_BROWNI)
            UserList(UserIndex).flags.Trabajando = False
            UserList(UserIndex).flags.Lingoteando = 0
        End If
    End If
End Sub

''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.

Public Sub CancelExit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/02/08
'
'***************************************************
    If UserList(UserIndex).Counters.Saliendo Then
        ' Is the user still connected?
        If UserList(UserIndex).ConnIDValida Then
            UserList(UserIndex).Counters.Saliendo = False
            UserList(UserIndex).Counters.salir = 0
            Call WriteMsg(UserIndex, 42)
        Else
            'Simply reset
            UserList(UserIndex).Counters.salir = IIf((UserList(UserIndex).flags.Privilegios And (PlayerType.User Or PlayerType.VIP)) And MapInfo(UserList(UserIndex).Pos.map).Pk, IntervaloCerrarConexion, 0)
        End If
    End If
    
    If UserList(UserIndex).Counters.IdleCount > 0 Then
    UserList(UserIndex).Counters.IdleCount = 0
    End If
    
End Sub

Sub VolverRenegado(ByVal UserIndex As Integer)
    With UserList(UserIndex).faccion
        .ArmadaReal = 0
        .FuerzasCaos = 0
        .Milicia = 0
        .Rango = 0
        
        .Ciudadano = 0
        .Republicano = 0
        .Renegado = 1
    End With
    Call RefreshCharStatus(UserIndex)
End Sub


Public Sub TalkNormal(ByVal UserIndex As Integer, ByVal chat As String)
With UserList(UserIndex)
             
  If .Counters.IdleCount > 0 Then
    .Counters.IdleCount = 0
  End If
    'I see you....
    If .flags.Oculto > 0 Then
        .flags.Oculto = 0
        .Counters.TiempoOculto = 0
        If .flags.Invisible = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
            Call WriteConsoleMsg(1, UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
                
    If LenB(chat) <> 0 Then
        If .flags.Muerto = 1 Then
            Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, &HC0C0C0))
        Else
            If EsCONSE(UserIndex) And Not EsVIP(UserIndex) Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, &H18C10, 1))
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, vbWhite, 1))
            End If
        End If
    End If
    
End With



End Sub
Public Sub TalkGritar(ByVal UserIndex As Integer, ByVal chat As String)
With UserList(UserIndex)
    If .flags.Muerto = 1 Then
        Call WriteMsg(UserIndex, 3)
    Else
        'I see you....
        If .flags.Oculto > 0 Then
                .flags.Oculto = 0
            .Counters.TiempoOculto = 0
        If .flags.Invisible = 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                Call WriteConsoleMsg(1, UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
                    
        If LenB(chat) <> 0 Then
            If .flags.Privilegios And (PlayerType.User Or PlayerType.VIP) Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, vbRed))
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, &HF82FF))
            End If
            Call SendData(SendTarget.ToMap, .Pos.map, PrepareMessageConsoleMsg(1, "[" & .Name & "] " & chat, FontTypeNames.FONTTYPE_RED))
        End If
    End If
End With
End Sub
Public Sub TalkGlobal(ByVal UserIndex As Integer, ByVal chat As String)
With UserList(UserIndex)
    If .flags.Muerto = 1 Then
        Call WriteMsg(UserIndex, 3)
    Else
    
        If LenB(chat) <> 0 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(3, .Name & ">" & chat, FontTypeNames.FONTTYPE_GLOBAL))
        End If
        
    End If
End With
End Sub
Function UserTypeColor(ByVal UserIndex As Integer) As Byte
    If UserList(UserIndex).flags.Privilegios = PlayerType.Admin Then
        UserTypeColor = 8
    ElseIf UserList(UserIndex).flags.Privilegios = PlayerType.Dios Then
        UserTypeColor = 7
    ElseIf UserList(UserIndex).flags.Privilegios = PlayerType.Semi Then
        UserTypeColor = 6
    ElseIf UserList(UserIndex).flags.Privilegios = PlayerType.Conse Then
        UserTypeColor = 5
    
    'Add Marius Lideres faccionario
    ElseIf UserList(UserIndex).flags.Privilegios = PlayerType.FaccCaos Then
        UserTypeColor = 11
    ElseIf UserList(UserIndex).flags.Privilegios = PlayerType.FaccRepu Then
        UserTypeColor = 10
    ElseIf UserList(UserIndex).flags.Privilegios = PlayerType.FaccImpe Then
        UserTypeColor = 9
    '\Add
    
    ElseIf UserList(UserIndex).faccion.Renegado = 1 Then
        UserTypeColor = 1
    ElseIf UserList(UserIndex).faccion.ArmadaReal = 1 Or UserList(UserIndex).faccion.Ciudadano = 1 Then
        UserTypeColor = 2
    ElseIf UserList(UserIndex).faccion.FuerzasCaos = 1 Then
        UserTypeColor = 3
    ElseIf UserList(UserIndex).faccion.Milicia = 1 Or UserList(UserIndex).faccion.Republicano = 1 Then
        UserTypeColor = 4
    Else
        UserTypeColor = 1
    End If
End Function


Public Function EntregarMsgOn(ByVal UserIndex As Integer, ByVal para As Integer, ByRef Mensaje As String, ByVal Slot As Byte, ByVal Cantidad As Integer) As Boolean
'***********************************************************************
'Author: Jose Ignacio Castelli (FEDUDOK)
'***********************************************************************
On Error GoTo Errhandler

Dim ObjIndex As Integer
Dim cantmensajes As Integer
Dim loopC As Long

    If Slot > 0 And Slot < MAX_INVENTORY_SLOTS + 1 Then
        ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        
        If UserList(UserIndex).Invent.Object(Slot).Amount < Cantidad Then
            WriteMsg UserIndex, 13
            Exit Function
        End If
    Else
        ObjIndex = 0
    End If
    
    cantmensajes = Cantidadmensajes(GetIndexPJ(UserList(para).Name))
    
    If cantmensajes < (MENSAJES_TOPE_CORREO + 1) Then
    
        For loopC = 1 To MENSAJES_TOPE_CORREO
            
            If UserList(para).Correos(loopC).De = "" Then
                UserList(para).Correos(loopC).De = UserList(UserIndex).Name
                UserList(para).Correos(loopC).Item = ObjIndex
                UserList(para).Correos(loopC).Mensaje = Mensaje
                UserList(para).Correos(loopC).Cantidad = Cantidad
                
                'Add Marius Sin esto se puede duplicar items
                'Nunca se cargaba el id del mensaje nuevo, entonces cuando lo queria borrar era = 0 y no lo borraba, relogeabas y tenias otra vez el mismo objeto, pero ahi si lo borraba por que ya tenia cargado el id del mensaje.
                UserList(para).Correos(loopC).idmsj = EnviarCorreoSql(GetIndexPJ(UserList(para).Name), loopC, para)
                '\Add
        
                UserList(UserIndex).cant_mensajes = cantmensajes + 1
                
                UserList(para).cVer = 1
                EntregarMsgOn = True
                
                WriteMensajeSigno para
                Exit Function
            End If
        Next loopC
    
    End If
    
    EntregarMsgOn = False

Errhandler:
    Call LogError("Error en EntregarMsgOn. N:" & UserList(UserIndex).Name & " - " & err.Number & "-" & err.description)
    EntregarMsgOn = False

End Function
Public Function EntregarMsgOff(ByVal UserIndex As Integer, ByRef para As String, ByRef Mensaje As String, ByVal Slot As Byte, ByVal Cantidad As Integer) As Boolean
'***********************************************************************
'Author: Jose Ignacio Castelli (FEDUDOK)
'***********************************************************************

Dim ObjIndex As Integer
Dim loopC As Long
Dim ipj As Integer

Dim i As String
Dim cantmensajes As Byte

If Slot > 0 And Slot < MAX_INVENTORY_SLOTS + 1 Then
    ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    
    If UserList(UserIndex).Invent.Object(Slot).Amount < Cantidad Then
        WriteMsg UserIndex, 13
        Exit Function
    End If
Else
    ObjIndex = 0
End If

ipj = GetIndexPJ(para)

cantmensajes = Cantidadmensajes(ipj)

If cantmensajes < (MENSAJES_TOPE_CORREO + 1) Then


Dim RS As ADODB.Recordset

    Set RS = DB_Conn.Execute("SELECT * FROM `charcorreo` WHERE IndexPJ=" & ipj & " LIMIT 1")


            DB_Conn.Execute "INSERT INTO `charcorreo` SET IndexPJ=" & ipj & "," & _
                "Mensaje='" & Mensaje & "'," & _
                "De='" & UserList(UserIndex).Name & "'," & _
                "Item=" & ObjIndex & "," & _
                "Cantidad=" & Cantidad
                
            EntregarMsgOff = True
            
          Set RS = Nothing
            Exit Function



End If


EntregarMsgOff = False


Set RS = Nothing

End Function
Public Sub SwapObjects(ByVal UserIndex As Integer, ByVal ObjSlot1 As Byte, ByVal ObjSlot2 As Byte)
    Dim tmpUserObj As UserObj
 
    With UserList(UserIndex)
        If .Invent.AnilloEqpSlot = ObjSlot1 Then
            .Invent.AnilloEqpSlot = ObjSlot2
        ElseIf .Invent.AnilloEqpSlot = ObjSlot2 Then
            .Invent.AnilloEqpSlot = ObjSlot1
        End If
       
        If .Invent.ArmourEqpSlot = ObjSlot1 Then
            .Invent.ArmourEqpSlot = ObjSlot2
        ElseIf .Invent.ArmourEqpSlot = ObjSlot2 Then
            .Invent.ArmourEqpSlot = ObjSlot1
        End If
       
        If .Invent.BarcoSlot = ObjSlot1 Then
            .Invent.BarcoSlot = ObjSlot2
        ElseIf .Invent.BarcoSlot = ObjSlot2 Then
            .Invent.BarcoSlot = ObjSlot1
        End If
       
        If .Invent.CascoEqpSlot = ObjSlot1 Then
            .Invent.CascoEqpSlot = ObjSlot2
        ElseIf .Invent.CascoEqpSlot = ObjSlot2 Then
            .Invent.CascoEqpSlot = ObjSlot1
        End If
       
        If .Invent.EscudoEqpSlot = ObjSlot1 Then
            .Invent.EscudoEqpSlot = ObjSlot2
        ElseIf .Invent.EscudoEqpSlot = ObjSlot2 Then
            .Invent.EscudoEqpSlot = ObjSlot1
        End If
       
        If .Invent.MunicionEqpSlot = ObjSlot1 Then
            .Invent.MunicionEqpSlot = ObjSlot2
        ElseIf .Invent.MunicionEqpSlot = ObjSlot2 Then
            .Invent.MunicionEqpSlot = ObjSlot1
        End If
       
        If .Invent.WeaponEqpSlot = ObjSlot1 Then
            .Invent.WeaponEqpSlot = ObjSlot2
        ElseIf .Invent.WeaponEqpSlot = ObjSlot2 Then
            .Invent.WeaponEqpSlot = ObjSlot1
        End If
        
        If .Invent.NudiEqpSlot = ObjSlot1 Then
            .Invent.NudiEqpSlot = ObjSlot2
        ElseIf .Invent.NudiEqpSlot = ObjSlot2 Then
            .Invent.NudiEqpSlot = ObjSlot1
        End If
       
        If .Invent.MagicSlot = ObjSlot1 Then
            .Invent.MagicSlot = ObjSlot2
        ElseIf .Invent.MagicSlot = ObjSlot2 Then
            .Invent.MagicSlot = ObjSlot1
        End If
        
        'Hacemos el intercambio propiamente dicho
        tmpUserObj = .Invent.Object(ObjSlot1)
        .Invent.Object(ObjSlot1) = .Invent.Object(ObjSlot2)
        .Invent.Object(ObjSlot2) = tmpUserObj
 
        'Actualizamos los 2 slots que cambiamos solamente
        Call UpdateUserInv(False, UserIndex, ObjSlot1)
        Call UpdateUserInv(False, UserIndex, ObjSlot2)
    End With
End Sub
