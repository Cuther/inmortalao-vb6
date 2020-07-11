Attribute VB_Name = "AI"

Option Explicit

Public Enum TipoAI
    ESTATICO = 1
    NpcMaloAtacaUsersBuenos = 3
    NPCDEFENSA = 4
    SigueAmo = 8
    NpcAtacaNpc = 9
    NpcFamiliar = 11
End Enum

Public Const ELEMENTALFUEGO As Integer = 93
Public Const ELEMENTALTIERRA As Integer = 94
Public Const ELEMENTALAGUA As Integer = 92

'Damos a los NPCs el mismo rango de visión que un PJ
Public Const RANGO_VISION_X As Byte = 8
Public Const RANGO_VISION_Y As Byte = 6


'El familiar ve mas lejos que un npc normal (ve un 25% mas que un npc normal) para no lagear mucho el server
Public Const RANGO_VISION_FAMILIAR_X As Byte = 11
Public Const RANGO_VISION_FAMILIAR_Y As Byte = 9


'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo AI_NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'AI de los NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Private Sub GuardiasAI(ByVal NpcIndex As Integer, ByVal faccion As Byte)
    Dim nPos As WorldPos
    Dim tHeading As Byte
    
    Dim UI As Integer
    With Npclist(NpcIndex)
        If .Target > 0 Then
            If InRangoVision(.Target, .Pos.x, .Pos.Y) And UserList(.Target).flags.Muerto = 0 And UserList(.Target).flags.AdminInvisible = 0 And UserList(.Target).flags.Invisible = 0 And Not (EsDIOS(.Target)) Then
                If Distancia(.Pos, UserList(.Target).Pos) = 1 Then
                    If NpcAtacaUser(NpcIndex, .Target) Then
                        tHeading = FindDonde(NpcIndex, UserList(.Target).Pos)
                        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, tHeading)
                    Else
                        If .flags.LanzaSpells > 0 Then
                            Call NpcLanzaUnSpell(NpcIndex, .Target)
                        End If
                    End If
                    Exit Sub
                Else
                    If .flags.LanzaSpells > 0 Then
                        Call NpcLanzaUnSpell(NpcIndex, .Target)
                    End If
                    
                    If .flags.Inmovilizado = 1 Then Exit Sub
                    tHeading = FindDirection(NpcIndex, UserList(.Target).Pos)
                    Call MoveNPCChar(NpcIndex, tHeading)
                    Exit Sub
                End If
            Else
                .Target = 0
                GoTo alfinal
            End If
        Else
            .Target = BuscarUserFaccion(NpcIndex)
            GoTo alfinal
        End If
    End With

alfinal:
    If Npclist(NpcIndex).Pos.x <> Npclist(NpcIndex).StartPos.x Or Npclist(NpcIndex).Pos.Y <> Npclist(NpcIndex).StartPos.Y Then
        If Distancia(Npclist(NpcIndex).Pos, Npclist(NpcIndex).StartPos) = 1 Then
            tHeading = FindDonde(NpcIndex, Npclist(NpcIndex).StartPos)
        Else
            tHeading = FindDirection(NpcIndex, Npclist(NpcIndex).StartPos)
        End If
                
        Call MoveNPCChar(NpcIndex, tHeading)
    Else
        If Npclist(NpcIndex).Char.heading <> 3 Then Call ChangeNPCChar(NpcIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.Head, eHeading.SOUTH)
    End If
    
    Exit Sub
End Sub
Function BuscarUserFaccion(ByVal NpcIndex As Integer) As Integer
On Error Resume Next
    Dim i As Integer
    Dim faccion As Integer
    Dim UI As Integer

    For i = 1 To ModAreas.ConnGroups(Npclist(NpcIndex).Pos.map).CountEntrys
        UI = ModAreas.ConnGroups(Npclist(NpcIndex).Pos.map).UserEntrys(i)
        
        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Metamorfosis = 0 Then
            If Abs(UserList(UI).Pos.x - Npclist(NpcIndex).Pos.x) <= RANGO_VISION_X Then
                If Abs(UserList(UI).Pos.Y - Npclist(NpcIndex).Pos.Y) <= RANGO_VISION_Y Then
                    faccion = Npclist(NpcIndex).faccion
                    If (faccion = 1 And Not (esArmada(UI) Or esCiuda(UI))) Or _
                        (faccion = 2 And Not (esMili(UI) Or esRepu(UI))) Or _
                        (faccion = 3 And Not (esCaos(UI) Or esRene(UI))) Then
                        
                        BuscarUserFaccion = UI
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
    BuscarUserFaccion = 0
    Exit Function
End Function
''
' Handles the evil npcs' artificial intelligency.
'
' @param NpcIndex Specifies reference to the npc
Private Sub HostilMalvadoAI(ByVal NpcIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 28/04/2009
'28/04/2009: ZaMa - Now those NPCs who doble attack, have 50% of posibility of casting a spell on user.
'**************************************************************
    Dim nPos As WorldPos
    Dim headingloop As Byte
    Dim UI As Integer
    Dim NPCI As Integer
    Dim atacoPJ As Boolean
    
    atacoPJ = False
    
    With Npclist(NpcIndex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.map, nPos.x, nPos.Y) Then
                    UI = MapData(nPos.map, nPos.x, nPos.Y).UserIndex
                    NPCI = MapData(nPos.map, nPos.x, nPos.Y).NpcIndex
                    If UI > 0 And Not atacoPJ Then
                        If UserList(UI).flags.Muerto = 0 Then
                            atacoPJ = True
                            If .flags.LanzaSpells Then
                                If .flags.AtacaDoble Then
                                    If (RandomNumber(0, 1)) Then
                                        If NpcAtacaUser(NpcIndex, UI) Then
                                            Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                                        End If
                                        Exit Sub
                                    End If
                                End If
                                
                                Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                                Call NpcLanzaUnSpell(NpcIndex, UI)
                            End If
                            If NpcAtacaUser(NpcIndex, UI) Then
                                Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                            End If
                            Exit Sub
                        End If
                    ElseIf NPCI > 0 Then
                        If Npclist(NPCI).MaestroUser > 0 And Npclist(NPCI).flags.Paralizado = 0 Then
                            Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                            Call SistemaCombate.NpcAtacaNpc(NpcIndex, NPCI, False)
                            Exit Sub
                        End If
                    End If
                End If
            End If  'inmo
        Next headingloop
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub HostilBuenoAI(ByVal NpcIndex As Integer)
    Dim nPos As WorldPos
    Dim headingloop As eHeading
    Dim UI As Integer
    
    With Npclist(NpcIndex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.map, nPos.x, nPos.Y) Then
                    UI = MapData(nPos.map, nPos.x, nPos.Y).UserIndex
                    If UI > 0 Then
                        If UI = .flags.AttackedBy Then
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, UI)
                                End If
                                
                                If NpcAtacaUser(NpcIndex, UI) Then
                                    Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        Next headingloop
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub IrUsuarioCercano(ByVal NpcIndex As Integer)
    Dim tHeading As Byte
    Dim UI As Integer
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    Dim i As Long
    
    With Npclist(NpcIndex)
        If .flags.Inmovilizado = 1 Then
        
            If .flags.LanzaSpells = 0 Then Exit Sub
            
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            If .Target = 0 Then
                For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
                    UI = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                    
                    'Add Marius Agregamos el If para evitar que npc ataquen a los clones lo saque de la 0.13.3
                    ' TODO: Es temporal hatsa reparar un bug que hace que ataquen a usuarios de otros mapas
                    If UserList(UI).Pos.map = .Pos.map Then
                
                        'Is it in it's range of vision??
                        If Abs(UserList(UI).Pos.x - .Pos.x) <= RANGO_VISION_X And Sgn(UserList(UI).Pos.x - .Pos.x) = SignoEO Then
                            If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then
                                
                                If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Metamorfosis = 0 Then
                                    If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                                    .Target = UI
                                    Exit Sub
                                End If
                                
                            End If
                        End If
                        
                    ' Esto significa que esta bugueado.. Lo logueo, y "reparo" el error a mano (Todo temporal)
                    Else
                        Call LogError("El npc: " & .Name & "(Index: " & NpcIndex & " pos: " & .Pos.map & " " & .Pos.x & " " & .Pos.Y & "), intenta atacar a " & _
                                      UserList(UI).Name & "(Index: " & UI & ", pos: " & UserList(UI).Pos.map & " " & UserList(UI).Pos.x & " " & UserList(UI).Pos.Y & ")")
                        '.Owner = 0
                        'ModAreas.ConnGroups(.Pos.map).UserEntrys(i) = 0
                        Call QuitarUser(UI, .Pos.map)
                    End If
                    
                Next i
            Else
                UI = .Target
                If Abs(UserList(UI).Pos.x - .Pos.x) <= RANGO_VISION_X And Sgn(UserList(UI).Pos.x - .Pos.x) = SignoEO Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Metamorfosis = 0 Then
                            If .flags.LanzaSpells <> 0 Then
                                Call NpcLanzaUnSpell(NpcIndex, UI)
                            End If
                        Else
                            .Target = 0
                        End If
                    Else
                        .Target = 0
                    End If
                Else
                    .Target = 0
                End If
            End If
        Else
            If .Target = 0 Then
                For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
                    UI = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                    
                    'Is it in it's range of vision??
                    If Abs(UserList(UI).Pos.x - .Pos.x) <= RANGO_VISION_X Then
                        If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                            
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible And Not (EsDIOS(UI)) And UserList(UI).flags.Metamorfosis = 0 Then
                                If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                                
                                .Target = UI
                                
                                tHeading = FindDirection(NpcIndex, UserList(UI).Pos)
                                Call MoveNPCChar(NpcIndex, tHeading)
                                Exit Sub
                            End If
                        End If
                    End If
                Next i
            Else
                UI = .Target
                If Abs(UserList(UI).Pos.x - .Pos.x) <= RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible And Not (EsDIOS(UI)) And UserList(UI).flags.Metamorfosis = 0 Then
                            If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                            
                            tHeading = FindDirection(NpcIndex, UserList(UI).Pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                            Exit Sub
                        Else
                            .Target = 0
                        End If
                    Else
                        .Target = 0
                    End If
                Else
                    .Target = 0
                End If
            End If
        End If
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub SeguirAgresor(ByVal NpcIndex As Integer)
    Dim tHeading As Byte
    Dim UI As Integer
    
    Dim i As Long
    
    Dim SignoNS As Integer
    Dim SignoEO As Integer

    With Npclist(NpcIndex)
        If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            UI = .flags.AttackedBy
            If Abs(UserList(UI).Pos.x - .Pos.x) <= RANGO_VISION_X And Sgn(UserList(UI).Pos.x - .Pos.x) = SignoEO Then
                If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then
                
                    If .MaestroUser > 0 Then
                        If esMismoBando(.MaestroUser, .flags.AttackedBy) Then
                            Call WriteConsoleMsg(1, .MaestroUser, "Las mascotas no atacaran a usuario de tu mismo bando.", FontTypeNames.FONTTYPE_INFO)
                            Call FlushBuffer(.MaestroUser)
                            .flags.AttackedBy = 0
                            Exit Sub
                        End If
                    End If

                    If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.Metamorfosis = 0 Then
                        If .MaestroUser = 0 Then
                            If .flags.LanzaSpells > 0 Then
                                Call NpcLanzaUnSpell(NpcIndex, UI)
                            Else
                                If Distancia(UserList(UI).Pos, Npclist(NpcIndex).Pos) <= 1 Then
                                    Call NpcAtacaUser(NpcIndex, UI)
                                End If
                            End If
                            Exit Sub
                        Else
                            If .flags.LanzaSpells > 0 Then
                                If RandomNumber(1, 10) < 5 Then
                                    Call NpcLanzaUnSpell(NpcIndex, UI)
                                Else
                                    If Distancia(UserList(UI).Pos, Npclist(NpcIndex).Pos) <= 1 Then
                                        Call NpcAtacaUser(NpcIndex, UI)
                                    End If
                                End If
                            Else
                                If Distancia(UserList(UI).Pos, Npclist(NpcIndex).Pos) <= 1 Then
                                    Call NpcAtacaUser(NpcIndex, UI)
                                End If
                            End If
                            Exit Sub
                        End If
                    End If
                    
                End If
            End If

        Else
            UI = .flags.AttackedBy
                
            'Is it in it's range of vision??
            If Abs(UserList(UI).Pos.x - .Pos.x) <= RANGO_VISION_X Then
                If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                    If .MaestroUser > 0 Then
                        If esMismoBando(.flags.AttackedBy, .MaestroUser) Then
                            Call WriteConsoleMsg(1, .MaestroUser, "La mascota no atacará a ciudadanos si eres miembro de la Armada Real o tienes el seguro activado", FontTypeNames.FONTTYPE_INFO)
                            Call FlushBuffer(.MaestroUser)
                            .flags.AttackedBy = 0
                            Call FollowAmo(NpcIndex)
                            Exit Sub
                        End If
                    End If
                    
                    If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.Metamorfosis = 0 Then
                        Dim bolo As Boolean
                        If .flags.LanzaSpells > 0 Then
                            bolo = (RandomNumber(1, 10) > 5)
                        Else
                            bolo = False
                        End If
                         
                        If bolo Then
                            Call NpcLanzaUnSpell(NpcIndex, UI)
                        Else
                            If Distancia(UserList(UI).Pos, Npclist(NpcIndex).Pos) <= 1 Then
                                Call NpcAtacaUser(NpcIndex, UI)
                            End If
                        End If
                         
                        tHeading = FindDirection(NpcIndex, UserList(UI).Pos)
                        Call MoveNPCChar(NpcIndex, tHeading)
                         
                        Exit Sub
                    End If

                End If
            End If
            
        End If
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)
    With Npclist(NpcIndex)
        If .MaestroUser = 0 Then
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
            .flags.AttackedBy = 0
        End If
    End With
End Sub

Private Sub SeguirAmo(ByVal NpcIndex As Integer)
    Dim tHeading As Byte
    Dim UI As Integer
    
    With Npclist(NpcIndex)
        UI = .MaestroUser
        
        If UI = 0 Then Exit Sub
        
            'Is it in it's range of vision??
        If Abs(UserList(UI).Pos.x - .Pos.x) <= RANGO_VISION_FAMILIAR_X Then
            If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_FAMILIAR_Y Then
                If UserList(UI).flags.Muerto = 0 Then
                    If UserList(UI).flags.Invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                        If Distancia(.Pos, UserList(UI).Pos) > 3 Then
                            tHeading = FindDirection(NpcIndex, UserList(UI).Pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                            Exit Sub
                        End If
                    End If
                Else
                    .MaestroUser = 0
                    If .IsFamiliar = True Then
                        MuereNpc NpcIndex, 0
                    End If
                End If
            End If
        End If
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub AiNpcAtacaNpc(ByVal NpcIndex As Integer)
    Dim tHeading As Byte
    Dim x As Long
    Dim Y As Long
    Dim NI As Integer
    Dim bNoEsta As Boolean
    
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    
    With Npclist(NpcIndex)
        
        If .TargetNPC = 0 Then Exit Sub
        
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
         
            NI = .TargetNPC
            If Abs(Npclist(NI).Pos.x - .Pos.x) <= RANGO_VISION_X And Sgn(Npclist(NI).Pos.x - .Pos.x) = SignoEO Then
                If Abs(Npclist(NI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(Npclist(NI).Pos.Y - .Pos.Y) = SignoNS Then
                    bNoEsta = True
                    If .Numero = ELEMENTALFUEGO Then
                        Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
                        If Npclist(NI).NPCtype = DRAGON Then
                            Npclist(NI).CanAttack = 1
                            Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)
                        End If
                    Else
                        'aca verificamosss la distancia de ataque
                        If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
                            Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)
                        End If
                        
                    End If
                Else
                    .TargetNPC = 0
                End If
            Else
                .TargetNPC = 0
                Exit Sub
            End If
        Else
            NI = .TargetNPC
            If Abs(Npclist(NI).Pos.x - .Pos.x) <= RANGO_VISION_X Then
                If Abs(Npclist(NI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                    bNoEsta = True
                    If .Numero = ELEMENTALFUEGO Then
                        Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
                        If Npclist(NI).NPCtype = DRAGON Then
                            Npclist(NI).CanAttack = 1
                            Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)
                        End If
                    
                    Else
                        'aca verificamosss la distancia de ataque
                        If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
                            If RandomNumber(1, 10) < 3 Then
                                Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
                            Else
                                Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)
                            End If
                        Else
                            If RandomNumber(1, 10) < 5 Then
                                Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
                            End If
                        End If
                    End If
                        
                    If .flags.Inmovilizado = 1 Then Exit Sub

                    tHeading = FindDirection(NpcIndex, Npclist(NI).Pos)
                    Call MoveNPCChar(NpcIndex, tHeading)
                    Exit Sub
                End If
            End If
        End If
        
        If Not bNoEsta Then
            If .MaestroUser > 0 Then
                Call FollowAmo(NpcIndex)
            Else
                .Movement = .flags.OldMovement
                .Hostile = .flags.OldHostil
            End If
        End If
    End With
End Sub

Sub NPCAI(ByVal NpcIndex As Integer)
On Error GoTo ErrorHandler
    With Npclist(NpcIndex)
        '<<<<<<<<<<< Ataques >>>>>>>>>>>>>>>>
        If .MaestroUser = 0 Then
            If Npclist(NpcIndex).faccion = 3 Then
                Call GuardiasAI(NpcIndex, 3)
                Exit Sub
                
            ElseIf Npclist(NpcIndex).faccion = 2 Then
                Call GuardiasAI(NpcIndex, 2)
                Exit Sub
                
            ElseIf Npclist(NpcIndex).faccion = 1 Then
                Call GuardiasAI(NpcIndex, 1)
                Exit Sub
                
            ElseIf .Hostile And .Stats.Alineacion <> 0 Then
                Call HostilMalvadoAI(NpcIndex)
            ElseIf .Hostile And .Stats.Alineacion = 0 Then
                Call HostilBuenoAI(NpcIndex)
            End If
        End If
        
        
        '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
        Select Case .Movement
            Case TipoAI.NpcMaloAtacaUsersBuenos
                Call IrUsuarioCercano(NpcIndex)

            Case TipoAI.NPCDEFENSA
                Call SeguirAgresor(NpcIndex)
            
            Case TipoAI.SigueAmo
                If .flags.Inmovilizado = 1 Then Exit Sub
                Call SeguirAmo(NpcIndex)
                
            Case TipoAI.NpcAtacaNpc
                Call AiNpcAtacaNpc(NpcIndex)
            
            Case TipoAI.NpcFamiliar
                If .IsFamiliar Then
                    If .MaestroUser > 0 Then
                        If .Target = 0 Then
                            If .TargetNPC = 0 Then
                                Call SeguirAmo(NpcIndex)
                            Else
                                Call AiNpcAtacaNpc(NpcIndex)
                            End If
                        Else
                            Call SeguirAgresor(NpcIndex)
                        End If
                    Else
                       ' Castelli
                        QuitarNPC NpcIndex
                    End If
                End If
                
        End Select
    End With
Exit Sub

ErrorHandler:
    Call LogError("NPCAI " & Npclist(NpcIndex).Name & " " & Npclist(NpcIndex).MaestroUser & " " & Npclist(NpcIndex).MaestroNpc & " mapa:" & Npclist(NpcIndex).Pos.map & " x:" & Npclist(NpcIndex).Pos.x & " y:" & Npclist(NpcIndex).Pos.Y & " Mov:" & Npclist(NpcIndex).Movement & " TargU:" & Npclist(NpcIndex).Target & " TargN:" & Npclist(NpcIndex).TargetNPC)
    Dim MiNPC As npc
    MiNPC = Npclist(NpcIndex)
    Call QuitarNPC(NpcIndex)
    Call ReSpawnNpc(MiNPC)
End Sub


Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    Dim k As Integer
    
    
    If Npclist(NpcIndex).IsFamiliar Then
        Dim UI As Integer, L As Byte
        Dim Spells(1 To 5) As Integer
        L = 1
        
        UI = Npclist(NpcIndex).MaestroUser
        If UI > 0 Then
            If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then
                If UserList(UI).masc.DetecInvi Then
                    Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, 62)
                End If
            Else
                If UserList(UI).masc.Descargas Then Spells(L) = 93: L = L + 1
                If UserList(UI).masc.Misil Then Spells(L) = 92: L = L + 1
                If UserList(UI).masc.Tormentas Then Spells(L) = 15: L = L + 1
                If UserList(UI).masc.Paraliza Then Spells(L) = 9: L = L + 1
                If UserList(UI).masc.Inmoviliza Then Spells(L) = 24: L = L + 1
    
                L = L - 1
                k = RandomNumber(1, L)
                
                Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, Npclist(NpcIndex).Spells(k))
            End If
        End If
    Else
        If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then Exit Sub
        
        k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
        Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, Npclist(NpcIndex).Spells(k))
    End If
End Sub

Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer)
    Dim k As Integer
    If Npclist(NpcIndex).IsFamiliar Then
        Dim UI As Integer, L As Byte
        Dim Spells(1 To 5) As Integer
        L = 1
        
        UI = Npclist(NpcIndex).MaestroUser
        If UI > 0 Then
            If UserList(UI).masc.Descargas Then Spells(L) = 93: L = L + 1
            If UserList(UI).masc.Misil Then Spells(L) = 92: L = L + 1
            If UserList(UI).masc.Tormentas Then Spells(L) = 15: L = L + 1
            If UserList(UI).masc.Paraliza Then Spells(L) = 9: L = L + 1
            If UserList(UI).masc.Inmoviliza Then Spells(L) = 24: L = L + 1

            L = L - 1
            k = RandomNumber(1, L)
            If Not L = 0 Then Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, Spells(k))
        End If
    Else
        k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
        If Npclist(NpcIndex).flags.LanzaSpells > 0 Then Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, Npclist(NpcIndex).Spells(k))
    End If
End Sub

