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
Sub DoResucitar(ByVal userindex As Integer)
    With UserList(userindex)
        If .flags.Resucitando = 0 Then Exit Sub
        
        Dim TActual As Long
        TActual = GetTickCount() And &H7FFFFFFF
        If TActual - UserList(userindex).Counters.IntervaloRevive < 2500 Then
            Exit Sub
        Else
            DarVida userindex
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_REVIVE, .pos.x, .pos.Y))
        End If
        
    End With
End Sub

Sub RevivirUsuario(ByVal userindex As Integer)
    If UserList(userindex).flags.Resucitando <> 1 Then
        UserList(userindex).flags.Resucitando = 1
        UserList(userindex).Counters.IntervaloRevive = GetTickCount
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateCharParticle(UserList(userindex).Char.CharIndex, 22))
    End If
End Sub

Sub DarVida(ByVal userindex As Integer)
    With UserList(userindex)
        .flags.Muerto = 0
        'Mod nod Kopfnickend
        'Cuando revive, lo hace con toda la vida
        '.Stats.MinHP = .Stats.UserAtributos(eAtributos.constitucion)
        
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
            Call DarCuerpoDesnudo(userindex)
            
            .Char.Head = .OrigChar.Head
        End If
        
        Call ChangeUserChar(userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        Call WriteUpdateUserStats(userindex)
    End With
End Sub

Sub ChangeUserChar(ByVal userindex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal casco As Integer)

    With UserList(userindex).Char
        .body = body
        .Head = Head
        .heading = heading
        .WeaponAnim = Arma
        .ShieldAnim = Escudo
        .CascoAnim = casco
        
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCharacterChange(body, Head, heading, .CharIndex, Arma, Escudo, .FX, .loops, casco))
    End With
End Sub


Sub EraseUserChar(ByVal userindex As Integer)

On Error GoTo ErrorHandler
    
    With UserList(userindex)

        CharList(.Char.CharIndex) = 0

        
        If .Char.CharIndex = LastChar Then
            Do Until CharList(LastChar) > 0
                LastChar = LastChar - 1
                If LastChar <= 1 Then Exit Do
            Loop
        End If
        
        'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCharacterRemove(.Char.CharIndex))
        Call QuitarUser(userindex, .pos.map)
        
        MapData(.pos.map, .pos.x, .pos.Y).userindex = 0
        .Char.CharIndex = 0
    End With
    
    NumChars = NumChars - 1
Exit Sub
    
ErrorHandler:
    Call LogError("Error en EraseUserchar " & err.Number & ": " & err.description)
End Sub

Sub RefreshCharStatus(ByVal userindex As Integer)
'*************************************************
'Author: Tararira
'Last modified: 04/21/2008 (NicoNZ)
'Refreshes the status and tag of UserIndex.
'*************************************************
    Dim klan As String
    Dim Barco As ObjData

    With UserList(userindex)
        If .GuildIndex > 0 Then
            klan = modGuilds.GuildName(.GuildIndex)
            klan = " <" & klan & ">"
        End If
        
        If .showName Then
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageUpdateTagAndStatus(userindex, UserTypeColor(userindex), .Name & klan))
        Else
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageUpdateTagAndStatus(userindex, UserTypeColor(userindex), vbNullString))
        End If
        
        'Si esta navengando, se cambia la barca.
        If .flags.Navegando Then
            Barco = ObjData(.Invent.Object(.Invent.BarcoSlot).ObjIndex)
            .Char.body = Barco.Ropaje
            Call ChangeUserChar(userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
    End With
End Sub

Sub MakeUserChar(ByVal ToMap As Boolean, ByVal sndIndex As Integer, ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer)

On Error GoTo hayerror
    Dim CharIndex As Integer

    If InMapBounds(map, x, Y) Then
        'If needed make a new character in list
        If UserList(userindex).Char.CharIndex = 0 Then
            CharIndex = NextOpenCharIndex
            UserList(userindex).Char.CharIndex = CharIndex
            CharList(CharIndex) = userindex
        End If
        
        'Place character on map if needed
        If ToMap Then MapData(map, x, Y).userindex = userindex
         If UserList(userindex).Char.heading = 0 Then
            UserList(userindex).Char.heading = 3
        End If
        
        'Send make character command to clients
        Dim klan As String
        If UserList(userindex).GuildIndex > 0 Then
            klan = modGuilds.GuildName(UserList(userindex).GuildIndex)
        End If

        If LenB(klan) <> 0 Then
            If Not ToMap Then
                If UserList(userindex).showName Then
                    Call WriteCharacterCreate(sndIndex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.CharIndex, x, Y, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.FX, 999, UserList(userindex).Char.CascoAnim, UserList(userindex).Name & " <" & klan & ">", UserTypeColor(userindex), UserList(userindex).donador)
                Else
                    'Hide the name and clan - set privs as normal user
                    Call WriteCharacterCreate(sndIndex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.CharIndex, x, Y, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.FX, 999, UserList(userindex).Char.CascoAnim, vbNullString, UserTypeColor(userindex), UserList(userindex).donador)
                End If
            Else
                Call AgregarUser(userindex, UserList(userindex).pos.map)
            End If
        Else 'if tiene clan
            If Not ToMap Then
                If UserList(userindex).showName Then
                    Call WriteCharacterCreate(sndIndex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.CharIndex, x, Y, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.FX, 999, UserList(userindex).Char.CascoAnim, UserList(userindex).Name, UserTypeColor(userindex), UserList(userindex).donador)
                Else
                    Call WriteCharacterCreate(sndIndex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.CharIndex, x, Y, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.FX, 999, UserList(userindex).Char.CascoAnim, vbNullString, UserTypeColor(userindex), UserList(userindex).donador)
                End If
            Else
                Call AgregarUser(userindex, UserList(userindex).pos.map)
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

Sub CheckUserLevel(ByVal userindex As Integer)
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
    
On Error GoTo Errhandler
    
    WasNewbie = EsNewbie(userindex)
    'Checkea si alcanzó el máximo nivel
    If UserList(userindex).Stats.ELV >= STAT_MAXELV Then
        UserList(userindex).Stats.Exp = 0
        UserList(userindex).Stats.ELU = 0
        Exit Sub
    End If
            
    With UserList(userindex)
        Do While .Stats.Exp >= .Stats.ELU
            
            'Checkea si alcanzó el máximo nivel
            If .Stats.ELV >= STAT_MAXELV Then
                .Stats.Exp = 0
                .Stats.ELU = 0
                Exit Sub
            End If
            

            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_NIVEL, .pos.x, .pos.Y))
            Call WriteConsoleMsg(2, userindex, "¡Has subido de nivel!", FontTypeNames.FONTTYPE_INFO)
            
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
                Call WriteConsoleMsg(1, userindex, "Has ganado " & AumentoHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
            End If
            If AumentoSTA > 0 Then
                Call WriteConsoleMsg(1, userindex, "Has ganado " & AumentoSTA & " puntos de vitalidad.", FontTypeNames.FONTTYPE_INFO)
            End If
            If AumentoMANA > 0 Then
                Call WriteConsoleMsg(1, userindex, "Has ganado " & AumentoMANA & " puntos de magia.", FontTypeNames.FONTTYPE_INFO)
            End If
            If AumentoHIT > 0 Then
                Call WriteConsoleMsg(1, userindex, "Tu golpe máximo aumentó en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(1, userindex, "Tu golpe minimo aumentó en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
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
            
            Call FlushBuffer(userindex)
            DoEvents
        Loop
        
        'If it ceased to be a newbie, remove newbie items and get char away from newbie dungeon
        If Not EsNewbie(userindex) And WasNewbie Then
            

            Call QuitarNewbieObj(userindex)
            
            If .pos.map = 37 Or .pos.map = 208 Then
                If UserList(userindex).Faccion.Ciudadano = 1 Then
                    Call WarpUserChar(userindex, 34, 40, 87, True)
                ElseIf UserList(userindex).Faccion.Republicano = 1 Then
                    Call WarpUserChar(userindex, 185, 50, 78, True)
                Else
                    Call WarpUserChar(userindex, 20, 50, 50, True)
                End If
                
                Call WriteConsoleMsg(1, userindex, "Debes abandonar el Dungeon Newbie.", FontTypeNames.FONTTYPE_INFO)
            End If
                
        End If
        
        'Send all gained skill points at once (if any)
        If Pts > 0 Then
            Call WriteUpdateUserStats(userindex)
            Call WriteUpdateExp(userindex)
            
            .Stats.SkillPts = .Stats.SkillPts + Pts
            
            Call WriteConsoleMsg(1, userindex, "Has ganado un total de " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)
        
            Call FlushBuffer(userindex)
        End If
        
    End With

Exit Sub

Errhandler:
    Call LogError("Error en la subrutina CheckUserLevel - Error : " & err.Number & " - Description : " & err.description)
End Sub
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
Public Function PuedeAtravesarAgua(ByVal userindex As Integer) As Boolean
    PuedeAtravesarAgua = UserList(userindex).flags.Navegando = 1
End Function

Sub MoveUserChar(ByVal userindex As Integer, ByVal nHeading As eHeading)
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
    
    sailing = PuedeAtravesarAgua(userindex)
    nPos = UserList(userindex).pos
    Call HeadtoPos(nHeading, nPos)
        
   If MoveToLegalPos(UserList(userindex).pos.map, nPos.x, nPos.Y, sailing, Not sailing) Then
        'si no estoy solo en el mapa...
        If MapInfo(UserList(userindex).pos.map).NumUsers > 1 Then
               
            CasperIndex = MapData(UserList(userindex).pos.map, nPos.x, nPos.Y).userindex
            'Si hay un usuario, y paso la validacion, entonces es un casper
            If CasperIndex > 0 Then
                If UserList(CasperIndex).flags.Muerto Then
                CasperHeading = InvertHeading(nHeading)
                CasPerPos = UserList(CasperIndex).pos
                Call HeadtoPos(CasperHeading, CasPerPos)

                With UserList(CasperIndex)
                    
                    Call SendData(SendTarget.ToPCAreaButIndex, CasperIndex, PrepareMessageCharacterMove(.Char.CharIndex, CasPerPos.x, CasPerPos.Y))
                    
                    Call WriteForceCharMove(CasperIndex, CasperHeading)
                        
                    'Update map and user pos
                    .pos = CasPerPos
                    .Char.heading = CasperHeading
                    MapData(.pos.map, CasPerPos.x, CasPerPos.Y).userindex = CasperIndex
                End With
            
                'Actualizamos las áreas de ser necesario
                Call ModAreas.CheckUpdateNeededUser(CasperIndex, CasperHeading)
                End If
            End If

            
            Call SendData(SendTarget.ToPCAreaButIndex, userindex, _
                    PrepareMessageCharacterMove(UserList(userindex).Char.CharIndex, nPos.x, nPos.Y))
            
        End If
        
        Dim oldUserIndex As Integer
        
        oldUserIndex = MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex
        
        ' Si no hay intercambio de pos con nadie
        If oldUserIndex = userindex Then
            MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex = 0
        End If
        
        UserList(userindex).pos = nPos
        UserList(userindex).Char.heading = nHeading
        MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex = userindex
        
        'Actualizamos las áreas de ser necesario
        Call ModAreas.CheckUpdateNeededUser(userindex, nHeading)

    Else
        Call WritePosUpdate(userindex)
    End If
    
    If UserList(userindex).Counters.Trabajando Then _
        UserList(userindex).Counters.Trabajando = UserList(userindex).Counters.Trabajando - 1

    If UserList(userindex).Counters.Ocultando Then _
        UserList(userindex).Counters.Ocultando = UserList(userindex).Counters.Ocultando - 1
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


Sub ChangeUserInv(ByVal userindex As Integer, ByVal Slot As Byte, ByRef Object As UserObj)
    UserList(userindex).Invent.Object(Slot) = Object
    Call WriteChangeInventorySlot(userindex, Slot)
End Sub

Function NextOpenCharIndex() As Integer
    Dim LoopC As Long
    
    For LoopC = 1 To MAXCHARS
        If CharList(LoopC) = 0 Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            
            If LoopC > LastChar Then _
                LastChar = LoopC
            
            Exit Function
        End If
    Next LoopC
End Function

Function NextOpenUser() As Integer
    '////Modificacion por Castelli
    'Se agrego UserList(LoopC).flags.UserLogged = False para asegurarce_
    'de que el usuario ya este offline y no agregar un index_
    'mal en un caso realmente extraño :S =P
    '////Modificacion por Castelli
    
    Dim LoopC As Long
    
    For LoopC = 1 To MaxUsers + 1
     
        If LoopC > MaxUsers Then Exit For
        If (UserList(LoopC).ConnID = -1 And UserList(LoopC).flags.UserLogged = False) Then Exit For
        Next LoopC
    
    NextOpenUser = LoopC
End Function

Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal userindex As Integer)
'**********************************************
'Author: Unknown
'Last Modification: 06/28/2008
'24/01/2007 -> Pablo (ToxicWaste): Agrego para que se actualize el tag si corresponde.
'24/07/2007 -> Pablo (ToxicWaste): Guardar primero que ataca NPC y el que atacas ahora.
'06/28/2008 -> NicoNZ: Los elementales al atacarlos por su amo no se paran más al lado de él sin hacer nada.
'**********************************************
    Dim EraCriminal As Boolean
    
    'Guardamos el usuario que ataco el npc.
    Npclist(NpcIndex).flags.AttackedBy = userindex
    
    'Npc que estabas atacando.
    Dim LastNpcHit As Integer
    LastNpcHit = UserList(userindex).flags.NPCAtacado
    'Guarda el NPC que estas atacando ahora.
    UserList(userindex).flags.NPCAtacado = NpcIndex
    
    'Revisamos robo de npc.
    'Guarda el primer nick que lo ataca.
    If Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then
        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(userindex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
            End If
        End If
        Npclist(NpcIndex).flags.AttackedFirstBy = UserList(userindex).Name
    ElseIf Npclist(NpcIndex).flags.AttackedFirstBy <> UserList(userindex).Name Then
        'Estas robando NPC
        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(userindex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
            End If
        End If
    End If
    
    If Npclist(NpcIndex).MaestroUser > 0 Then
        If Npclist(NpcIndex).MaestroUser <> userindex Then
            Call AllMascotasAtacanUser(userindex, Npclist(NpcIndex).MaestroUser)
        End If
    End If
    

End Sub

Public Function PuedeApuñalar(ByVal userindex As Integer) As Boolean

    If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Apuñala = 1 Then
            PuedeApuñalar = UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR _
                        Or UserList(userindex).Clase = eClass.Asesino
        End If
    End If
End Function

Sub SubirSkill(ByVal userindex As Integer, ByVal Skill As Integer)


On Error GoTo err


    With UserList(userindex)
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
                
                Call WriteMsg(userindex, 40, CStr(Skill), CStr(.Stats.UserSkills(Skill)))
                Call WriteMsg(userindex, 21, CStr(18 * .Stats.ELV))
                
                Call WriteUpdateExp(userindex)
                Call CheckUserLevel(userindex)
                
                Call FlushBuffer(userindex)
            End If
        End If
    End With
    
    Exit Sub
    
    
err:
    
End Sub

Sub PerderDuelo(ByVal userindex As Integer)
    If UserList(userindex).flags.inDuelo = False Then Exit Sub
   
    Dim IDA As Integer
    IDA = UserList(userindex).flags.vicDuelo
    If IDA <> 0 Then
        UserList(IDA).flags.inDuelo = 0
        UserList(IDA).flags.vicDuelo = 0
        Call WriteConsoleMsg(1, userindex, "¡Has perdido el duelo contra " & UserList(IDA).Name & "!", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(1, IDA, "¡Has ganado el duelo contra " & UserList(userindex).Name & "!", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    UserList(userindex).flags.vicDuelo = 0
    UserList(userindex).flags.inDuelo = 0

End Sub
''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'

Sub UserDie(ByVal userindex As Integer)
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
    Dim Drops As Obj

    With UserList(userindex)
        
        'Add Nod Kopfnickend
        'Solo le saca vida a los que no son Vips o Dioses
        If (UserList(userindex).flags.Privilegios And (PlayerType.Dios Or PlayerType.VIP)) Then
            .Stats.MinHP = .Stats.MaxHP
            .Stats.MinMAN = .Stats.MaxMAN
            .flags.Envenenado = 0
            .flags.Incinerado = 0
            .flags.Ceguera = 0
            Call WriteUpdateUserStats(userindex)
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateCharParticle(.Char.CharIndex, 119))
            Call DecirPalabrasMagicas("Divinum Protection", userindex)
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(236, .pos.x, .pos.Y))
            
            Exit Sub
        End If
        
        'Sonido
        Call ReproducirSonido(SendTarget.ToPCArea, userindex, 11)

        'Quitar el dialogo del user muerto
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
        .Stats.MinHP = 0
        '.Stats.MinSTA = 0
        .flags.AtacadoPorUser = 0
        .flags.Envenenado = 0
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
            Call WriteParalizeOK(userindex)
        End If
        
        '<<< Estupidez >>>
        If .flags.Estupidez = 1 Then
            .flags.Estupidez = 0
            Call WriteDumbNoMore(userindex)
        End If
        
        '<<<< Descansando >>>>
        If .flags.Descansar Then
            .flags.Descansar = False
            Call WriteRestOK(userindex)
        End If
        
        '<<<< Meditando >>>>
        If .flags.Meditando Then
            .flags.Meditando = False
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageDestCharParticle(UserList(userindex).Char.CharIndex, ParticleToLevel(userindex)))
            Call WriteMeditateToggle(userindex)
        End If
        
        '<<<< Invisible >>>>
        If .flags.Invisible = 1 Or .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .flags.Invisible = 0
            .Counters.TiempoOculto = 0
            .Counters.Invisibilidad = 0
            
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageSetInvisible(.Char.CharIndex, False))
        End If
        
        If TriggerZonaPelea(userindex, userindex) <> eTrigger6.TRIGGER6_PERMITE _
            And Not UserList(userindex).flags.inDuelo = 1 And MapInfo(UserList(userindex).pos.map).Pk = True Then
            
            'Add Nod kopfnickend
            'Solo los que no son Vip o Dioses pierden los items al morir
            If (UserList(userindex).flags.Privilegios And (PlayerType.Dios Or PlayerType.VIP)) = 0 Then
                DropObjs = TieneSacri(userindex)
                If DropObjs = 0 Then
                    'Mod Nod kopfnickend
                    'A todos se les cae todo menos los items nws
                    ' << Si es newbie no pierde el inventario >>
                    If Not EsNewbie(userindex) Then
                        Call TirarTodo(userindex)
                    Else
                        Call TirarTodosLosItemsNoNewbies(userindex)
                    End If
                Else
                    Drops.ObjIndex = .Invent.Object(DropObjs).ObjIndex
                    Drops.Amount = 1
                
                    TileLibre UserList(userindex).pos, NewPos, Drops, True, True
                    
                    Call DropObj(userindex, DropObjs, 1, NewPos.map, NewPos.x, NewPos.Y)
                End If
            End If
            
        End If
        
        ' DESEQUIPA TODOS LOS OBJETOS
        'desequipar armadura
        If .Invent.ArmourEqpObjIndex > 0 Then
            Call Desequipar(userindex, .Invent.ArmourEqpSlot)
        End If
        
        'desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call Desequipar(userindex, .Invent.WeaponEqpSlot)
        End If
        
        'desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            Call Desequipar(userindex, .Invent.CascoEqpSlot)
        End If
        
        'desequipar herramienta
        If .Invent.AnilloEqpSlot > 0 Then
            Call Desequipar(userindex, .Invent.AnilloEqpSlot)
        End If
        
        'desequipar municiones
        If .Invent.MunicionEqpObjIndex > 0 Then
            Call Desequipar(userindex, .Invent.MunicionEqpSlot)
        End If
        
        'desequipamos items macigos
        If .Invent.MagicIndex > 0 Then
            Call Desequipar(userindex, .Invent.MagicSlot)
        End If
        
        'desequipar escudo
        If .Invent.EscudoEqpObjIndex > 0 Then
            Call Desequipar(userindex, .Invent.EscudoEqpSlot)
        End If
        
        ' << Reseteamos los posibles FX sobre el personaje >>
        If .Char.loops = INFINITE_LOOPS Then
            .Char.FX = 0
            .Char.loops = 0
        End If
        
        ' << Restauramos los atributos >>
        If .flags.TomoPocion = True Then
            For i = 1 To 5
                .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
            Next i
        End If
        
        Call WriteAgilidad(userindex)
        Call WriteFuerza(userindex)
        
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
            Call WriteEquitateToggle(userindex)
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
            Call desinvocarfami(userindex)
        End If
         '/// CASTELLI , desinvocamos a la mascota si la tiene invocada
         
          If UserList(userindex).flags.automatico = True Then
            Call Rondas_UsuarioMuere(userindex)
            End If
         
        '<< Actualizamos clientes >>
        Call ChangeUserChar(userindex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
        Call WriteUpdateUserStats(userindex)
        
        '<<Castigos por Grupo>>
        'If .GrupoIndex > 0 Then
          '  Call mdGrupo.ObtenerExito(UserIndex, .Stats.ELV * -10 * mdGrupo.CantMiembros(UserIndex), .Pos.map, .Pos.X, .Pos.Y)
        'End If
        
        If .flags.inDuelo = 1 Then
            PerderDuelo userindex
            RevivirUsuario userindex
        End If
        
        If TriggerZonaPelea(userindex, userindex) = TRIGGER6_PERMITE Then
            RevivirUsuario userindex
        End If
        
       Call ControlarPortalLum(userindex)
        
    End With
Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & err.Number & " Descripción: " & err.description)
End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)

    If EsNewbie(Muerto) Then Exit Sub
    If UserList(Muerto).flags.inDuelo = 1 Then Exit Sub
    
    With UserList(Atacante)
        If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
        
        If UserList(Muerto).Faccion.ArmadaReal = 1 Then
            .Faccion.ArmadaMatados = .Faccion.ArmadaMatados + 1
            Exit Sub
        End If
        
        If UserList(Muerto).Faccion.Milicia = 1 Then
            .Faccion.MilicianosMatados = .Faccion.MilicianosMatados + 1
            Exit Sub
        End If
        
        If UserList(Muerto).Faccion.FuerzasCaos = 1 Then
            .Faccion.CaosMatados = .Faccion.CaosMatados + 1
            Exit Sub
        End If
        
        If UserList(Muerto).Faccion.Renegado = 1 Then
            .Faccion.RenegadosMatados = .Faccion.RenegadosMatados + 1
            Exit Sub
        End If
        
        If UserList(Muerto).Faccion.Republicano = 1 Then
            .Faccion.RepublicanosMatados = .Faccion.RepublicanosMatados + 1
            Exit Sub
        End If
        
        If UserList(Muerto).Faccion.Ciudadano Then
            .Faccion.CiudadanosMatados = .Faccion.CiudadanosMatados + 1
            Exit Sub
        End If
        
    End With
End Sub

Sub TileLibre(ByRef pos As WorldPos, ByRef nPos As WorldPos, ByRef Obj As Obj, ByRef Agua As Boolean, ByRef Tierra As Boolean)
'**************************************************************
'Author: Unknown
'Last Modify Date: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
'**************************************************************
    Dim LoopC As Integer
    Dim tX As Long
    Dim tY As Long
    Dim hayobj As Boolean
    
    hayobj = False
    nPos.map = pos.map
    nPos.x = 0
    nPos.Y = 0
    
    Do While Not LegalPos(pos.map, nPos.x, nPos.Y, Agua, Tierra) Or hayobj
        
        If LoopC > 15 Then
            Exit Do
        End If
        
        For tY = pos.Y - LoopC To pos.Y + LoopC
            For tX = pos.x - LoopC To pos.x + LoopC
                
                If LegalPos(nPos.map, tX, tY, Agua, Tierra) Then
                    'We continue if: a - the item is different from 0 and the dropped item or b - the amount dropped + amount in map exceeds MAX_INVENTORY_OBJS
                    hayobj = (MapData(nPos.map, tX, tY).ObjInfo.ObjIndex > 0 And MapData(nPos.map, tX, tY).ObjInfo.ObjIndex <> Obj.ObjIndex)
                    If Not hayobj Then _
                        hayobj = (MapData(nPos.map, tX, tY).ObjInfo.Amount + Obj.Amount > MAX_INVENTORY_OBJS)
                    If Not hayobj And MapData(nPos.map, tX, tY).TileExit.map = 0 Then
                        nPos.x = tX
                        nPos.Y = tY
                        
                        'break both fors
                        tX = pos.x + LoopC
                        tY = pos.Y + LoopC
                    End If
                End If
            
            Next tX
        Next tY
        
        LoopC = LoopC + 1
    Loop
End Sub

Sub WarpUserChar(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal FX As Boolean)
    Dim OldMap As Integer
    Dim OldX As Integer
    Dim OldY As Integer

    With UserList(userindex)
    
   ' If .Pos.map And .Pos.X And .Pos.Y Then Exit Sub
    
        'Quitar el dialogo
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        Call WriteRemoveAllDialogs(userindex)
        
        OldMap = .pos.map
        OldX = .pos.x
        OldY = .pos.Y
        
        Call EraseUserChar(userindex)
        
        If OldMap <> map Then
            Call WriteChangeMap(userindex, map, MapInfo(.pos.map).MapVersion)
            Call WritePlayMidi(userindex, val(ReadField(1, MapInfo(map).Music, 45)))
            
            'Update new Map Users
            MapInfo(map).NumUsers = MapInfo(map).NumUsers + 1
            
            'Update old Map Users
            MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
            If MapInfo(OldMap).NumUsers < 0 Then
                MapInfo(OldMap).NumUsers = 0
            End If
        End If
        
        .pos.x = x
        .pos.Y = Y
        .pos.map = map
        
        Call MakeUserChar(True, map, userindex, map, x, Y)
        Call WriteUserCharIndexInServer(userindex)
        
        'Force a flush, so user index is in there before it's destroyed for teleporting
        Call FlushBuffer(userindex)
        
        'Seguis invisible al pasar de mapa
        If (.flags.Invisible = 1 Or .flags.Oculto = 1) And (Not .flags.AdminInvisible = 1) Then
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageSetInvisible(.Char.CharIndex, True))
        End If
        
        If FX And .flags.AdminInvisible = 0 Then 'FX
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_WARP, x, Y))
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
        End If
        
        If .NroMascotas Then Call WarpMascotas(userindex)
    End With
End Sub

Private Sub WarpMascotas(ByVal userindex As Integer)
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
    
    NroPets = UserList(userindex).NroMascotas
    canWarp = (MapInfo(UserList(userindex).pos.map).Pk = True)
    
    For i = 1 To MAXMASCOTAS
        index = UserList(userindex).MascotasIndex(i)
        
        If index > 0 Then
            ' si la mascota tiene tiempo de vida > 0 significa q fue invocada => we kill it
            If Npclist(index).Contadores.TiempoExistencia > 0 Then
                Call QuitarNPC(index)
                UserList(userindex).MascotasIndex(i) = 0
                InvocadosMatados = InvocadosMatados + 1
                NroPets = NroPets - 1
                
                petType = 0
            Else
                'Store data and remove NPC to recreate it after warp
                'PetRespawn = Npclist(index).flags.Respawn = 0
                petType = UserList(userindex).MascotasType(i)
                'PetTiempoDeVida = Npclist(index).Contadores.TiempoExistencia
                
                ' Guardamos el hp, para restaurarlo uando se cree el npc
                iMinHP = Npclist(index).Stats.MinHP
                
                Call QuitarNPC(index)
                
                ' Restauramos el valor de la variable
                UserList(userindex).MascotasType(i) = petType

            End If
        ElseIf UserList(userindex).MascotasType(i) > 0 Then
            'Store data and remove NPC to recreate it after warp
            PetRespawn = True
            petType = UserList(userindex).MascotasType(i)
            PetTiempoDeVida = 0
        Else
            petType = 0
        End If
        
        If petType > 0 And canWarp Then
            index = SpawnNpc(petType, UserList(userindex).pos, False, PetRespawn)
            UserList(userindex).MascotasIndex(i) = index

            ' Nos aseguramos de que conserve el hp, si estaba dañado
            Npclist(index).Stats.MinHP = IIf(iMinHP = 0, Npclist(index).Stats.MinHP, iMinHP)
            
            'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
            If index = 0 Then
                Call WriteConsoleMsg(1, userindex, "Tus mascotas no pueden transitar este mapa.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            Npclist(index).MaestroUser = userindex
            Npclist(index).Movement = TipoAI.SigueAmo
            Npclist(index).Target = 0
            Npclist(index).TargetNPC = 0
            Npclist(index).Contadores.TiempoExistencia = PetTiempoDeVida
            Call FollowAmo(index)
        End If
    Next i
    
    If InvocadosMatados > 0 Then
        Call WriteConsoleMsg(1, userindex, "Pierdes el control de tus mascotas invocadas.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    If Not canWarp Then
        Call WriteConsoleMsg(1, userindex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    UserList(userindex).NroMascotas = NroPets
End Sub

''
' Se inicia la salida de un usuario.
'
' @param    UserIndex   El index del usuario que va a salir

Sub Cerrar_Usuario(ByVal userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 09/04/08 (NicoNZ)
'
'***************************************************

    Dim isNotVisible As Boolean
    Dim diezSeg As Boolean
    
    If UserList(userindex).Counters.Saliendo = True Then
        CancelExit userindex
    ElseIf UserList(userindex).flags.UserLogged And Not UserList(userindex).Counters.Saliendo Then
        UserList(userindex).Counters.Saliendo = True

        If (UserList(userindex).flags.Privilegios And PlayerType.Dios) Or _
            UserList(userindex).flags.Muerto = 1 Or _
            MapInfo(UserList(userindex).pos.map).Pk = False Then
            
            diezSeg = False
        Else
            diezSeg = True
        End If

        UserList(userindex).Counters.Salir = IIf(diezSeg, IntervaloCerrarConexion, 0)
        
        isNotVisible = (UserList(userindex).flags.Oculto Or UserList(userindex).flags.Invisible)
        If isNotVisible Then
            UserList(userindex).flags.Oculto = 0
            UserList(userindex).flags.Invisible = 0
            Call WriteConsoleMsg(1, userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageSetInvisible(UserList(userindex).Char.CharIndex, False))
        End If
        
        If UserList(userindex).flags.Trabajando = True Then
            Call WriteConsoleMsg(1, userindex, "Dejas de trabajar.", FontTypeNames.FONTTYPE_BROWNI)
            UserList(userindex).flags.Trabajando = False
            UserList(userindex).flags.Lingoteando = 0
        End If
        

    
    End If
End Sub

''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.

Public Sub CancelExit(ByVal userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/02/08
'
'***************************************************
    If UserList(userindex).Counters.Saliendo Then
        ' Is the user still connected?
        If UserList(userindex).ConnIDValida Then
            UserList(userindex).Counters.Saliendo = False
            UserList(userindex).Counters.Salir = 0
            Call WriteMsg(userindex, 42)
        Else
            'Simply reset
            UserList(userindex).Counters.Salir = IIf((UserList(userindex).flags.Privilegios And (PlayerType.User Or PlayerType.VIP)) And MapInfo(UserList(userindex).pos.map).Pk, IntervaloCerrarConexion, 0)
        End If
    End If
    
    If UserList(userindex).Counters.IdleCount > 0 Then
    UserList(userindex).Counters.IdleCount = 0
    End If
    
End Sub

Sub VolverRenegado(ByVal userindex As Integer)
    With UserList(userindex).Faccion
        .ArmadaReal = 0
        .FuerzasCaos = 0
        .Milicia = 0
        .Rango = 0
        
        .Ciudadano = 0
        .Republicano = 0
        .Renegado = 1
    End With
    Call RefreshCharStatus(userindex)
End Sub


Public Sub TalkNormal(ByVal userindex As Integer, ByVal chat As String)
With UserList(userindex)
             
  If .Counters.IdleCount > 0 Then
  .Counters.IdleCount = 0
  End If
             
    'I see you....
    If .flags.Oculto > 0 Then
        .flags.Oculto = 0
        .Counters.TiempoOculto = 0
        If .flags.Invisible = 0 Then
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageSetInvisible(.Char.CharIndex, False))
            Call WriteConsoleMsg(1, userindex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
                
    If LenB(chat) <> 0 Then
        If .flags.Muerto = 1 Then
            Call SendData(SendTarget.ToDeadArea, userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, &HC0C0C0))
        Else
            If .flags.Privilegios And PlayerType.Dios Then
                Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, &H18C10, 1))
            Else
                Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, vbWhite, 1))
            End If
        End If
    End If
    
End With



End Sub
Public Sub TalkGritar(ByVal userindex As Integer, ByVal chat As String)
With UserList(userindex)
    If .flags.Muerto = 1 Then
        Call WriteMsg(userindex, 3)
    Else
        'I see you....
        If .flags.Oculto > 0 Then
                .flags.Oculto = 0
            .Counters.TiempoOculto = 0
        If .flags.Invisible = 0 Then
                Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                Call WriteConsoleMsg(1, userindex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
                    
        If LenB(chat) <> 0 Then
            If .flags.Privilegios And (PlayerType.User Or PlayerType.VIP) Then
                Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, vbRed))
            Else
                Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, &HF82FF))
            End If
            Call SendData(SendTarget.ToMap, .pos.map, PrepareMessageConsoleMsg(1, "[" & .Name & "] " & chat, FontTypeNames.FONTTYPE_RED))
        End If
    End If
End With
End Sub
Public Sub TalkGlobal(ByVal userindex As Integer, ByVal chat As String)
With UserList(userindex)
    If .flags.Muerto = 1 Then
        Call WriteMsg(userindex, 3)
    Else
        If LenB(chat) <> 0 Then
            Call SendData(SendTarget.toAll, 0, PrepareMessageConsoleMsg(3, .Name & ">" & chat, FontTypeNames.FONTTYPE_GLOBAL))
        End If
        
    End If
End With
End Sub
Function UserTypeColor(ByVal userindex As Integer) As Byte
    If UserList(userindex).flags.Privilegios = PlayerType.Dios Then UserTypeColor = 5: Exit Function
    
    If UserList(userindex).Faccion.Renegado = 1 Then
        UserTypeColor = 1
    ElseIf UserList(userindex).Faccion.ArmadaReal = 1 Or UserList(userindex).Faccion.Ciudadano = 1 Then
        UserTypeColor = 2
    ElseIf UserList(userindex).Faccion.FuerzasCaos = 1 Then
        UserTypeColor = 3
    ElseIf UserList(userindex).Faccion.Milicia = 1 Or UserList(userindex).Faccion.Republicano = 1 Then
        UserTypeColor = 4
    Else
        UserTypeColor = 1
    End If
End Function


Public Function EntregarMsgOn(ByVal userindex As Integer, ByVal para As Integer, ByRef Mensaje As String, ByVal Slot As Byte, ByVal Cantidad As Integer) As Boolean
'***********************************************************************
'Author: Jose Ignacio Castelli (FEDUDOK)
'***********************************************************************
Dim ObjIndex As Integer
Dim cantmensajes As Integer
Dim LoopC As Long

If Slot > 0 And Slot < MAX_INVENTORY_SLOTS + 1 Then
    ObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
    
    If UserList(userindex).Invent.Object(Slot).Amount < Cantidad Then
        WriteMsg userindex, 13
        Exit Function
    End If
Else
    ObjIndex = 0
End If

cantmensajes = Cantidadmensajes(GetIndexPJ(UserList(para).Name))

If cantmensajes < (MENSAJES_TOPE_CORREO + 1) Then


For LoopC = 1 To MENSAJES_TOPE_CORREO
    
    If UserList(para).Correos(LoopC).De = "" Then
        UserList(para).Correos(LoopC).De = UserList(userindex).Name
        UserList(para).Correos(LoopC).Item = ObjIndex
        UserList(para).Correos(LoopC).Mensaje = Mensaje
        UserList(para).Correos(LoopC).Cantidad = Cantidad
        Call EnviarCorreoSql(GetIndexPJ(UserList(para).Name), LoopC, para)

        UserList(userindex).cant_mensajes = cantmensajes + 1
        
        UserList(para).cVer = 1
        EntregarMsgOn = True
        
        WriteMensajeSigno para
        Exit Function
    End If
Next LoopC



End If

EntregarMsgOn = False


End Function
Public Function EntregarMsgOff(ByVal userindex As Integer, ByRef para As String, ByRef Mensaje As String, ByVal Slot As Byte, ByVal Cantidad As Integer) As Boolean
'***********************************************************************
'Author: Jose Ignacio Castelli (FEDUDOK)
'***********************************************************************

Dim ObjIndex As Integer
Dim LoopC As Long
Dim ipj As Integer

Dim i As String
Dim cantmensajes As Byte

If Slot > 0 And Slot < MAX_INVENTORY_SLOTS + 1 Then
    ObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
    
    If UserList(userindex).Invent.Object(Slot).Amount < Cantidad Then
        WriteMsg userindex, 13
        Exit Function
    End If
Else
    ObjIndex = 0
End If

ipj = GetIndexPJ(para)

cantmensajes = Cantidadmensajes(ipj)

If cantmensajes < (MENSAJES_TOPE_CORREO + 1) Then


Dim RS As ADODB.Recordset

    Set RS = Con.Execute("SELECT * FROM `charcorreo` WHERE IndexPJ=" & ipj & " LIMIT 1")


            Con.Execute "INSERT INTO `charcorreo` SET IndexPJ=" & ipj & "," & _
                "Mensaje='" & Mensaje & "'," & _
                "De='" & UserList(userindex).Name & "'," & _
                "Item=" & ObjIndex & "," & _
                "Cantidad=" & Cantidad
                
            EntregarMsgOff = True
            
          Set RS = Nothing
            Exit Function



End If


EntregarMsgOff = False


Set RS = Nothing

End Function
Public Sub SwapObjects(ByVal userindex As Integer, ByVal ObjSlot1 As Byte, ByVal ObjSlot2 As Byte)
    Dim tmpUserObj As UserObj
 
    With UserList(userindex)
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
        Call UpdateUserInv(False, userindex, ObjSlot1)
        Call UpdateUserInv(False, userindex, ObjSlot2)
    End With
End Sub
