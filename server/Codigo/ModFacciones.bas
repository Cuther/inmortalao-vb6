Attribute VB_Name = "ModFacciones"
'Tempest AO version 1.0
'modFacciones reescrito de 0 por mannakia basandome
'en los procesimientos de alkon 12.2
Const DIVFACCION As Byte = 4
Option Explicit
Public Sub EnlistarCaos(ByVal UserIndex As Integer)
Dim Matados As Integer
    Matados = (UserList(UserIndex).faccion.RenegadosMatados + UserList(UserIndex).faccion.ArmadaMatados + UserList(UserIndex).faccion.CiudadanosMatados + UserList(UserIndex).faccion.MilicianosMatados + UserList(UserIndex).faccion.RepublicanosMatados)
    
    If UserList(UserIndex).faccion.FuerzasCaos = 1 Then
        Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a la horda del caos!!! Ve a combatir enemigos", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If UserList(UserIndex).faccion.FuerzasCaos = 100 Then
        Call WriteChatOverHead(UserIndex, "¡¡¡Has sido expulsado y no quiero volverte a ver por aqui!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If UserList(UserIndex).faccion.Renegado = 0 Then
        Call WriteChatOverHead(UserIndex, "¡¡Sal de aquí insecto!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If Matados < 400 / DIVFACCION Then
        Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos " & 400 / DIVFACCION & " enemigos, solo has matado " & Matados, str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If UserList(UserIndex).Stats.ELV < 40 Then
        Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel 40!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    UserList(UserIndex).faccion.FuerzasCaos = 1
    UserList(UserIndex).faccion.Renegado = 0
    UserList(UserIndex).faccion.Rango = 1
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))
     
    '------- Ropa -------
    Dim MiObj As obj
    Dim bajos As Byte
    MiObj.Amount = 1
        
    If UserList(UserIndex).Raza = Enano Or UserList(UserIndex).Raza = Gnomo Then
        bajos = 1
    End If
        
    Select Case UserList(UserIndex).Clase
        Case eClass.Clerigo
            MiObj.ObjIndex = 1500 + bajos
        Case eClass.Mago
            MiObj.ObjIndex = 1502 + bajos
        Case eClass.Guerrero
            MiObj.ObjIndex = 1504 + bajos
        Case eClass.Asesino
            MiObj.ObjIndex = 1506 + bajos
        Case eClass.Bardo
            MiObj.ObjIndex = 1508 + bajos
        Case eClass.Druida
            MiObj.ObjIndex = 1510 + bajos
        Case eClass.Gladiador
            MiObj.ObjIndex = 1512 + bajos
        Case eClass.Paladin
            MiObj.ObjIndex = 1514 + bajos
        Case eClass.Cazador
            MiObj.ObjIndex = 1516 + bajos
        Case eClass.Mercenario
            MiObj.ObjIndex = 1518 + bajos
        Case eClass.Nigromante
            MiObj.ObjIndex = 1520 + bajos
    End Select
    
    'Add Marius Le da la armadura faccionaria solo si ya no la tiene.
    If Not TieneObjetos(MiObj.ObjIndex, 1, UserIndex) Then
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
    End If
    '------- Ropa -------
    
    Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido a las Fuerza del Caos!!!, aqui tienes tus vestimentas. Cumple bien tu labor exterminando enemigos y me encargaré de recompensarte.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    
    ' Lo sacamos del clan, por que para hacerse caos tiene que se rene, y si esta en un clan, es un clan rene. Con las demas facciones no pasan, por qeu lo clanes son imperiales o repu sin importar la gerarquia etc..
    If UserList(UserIndex).GuildIndex > 0 Then
        Call modGuilds.m_EcharMiembroDeClan(-1, UserList(UserIndex).Name)
        Call WriteConsoleMsg(1, UserIndex, "Has sido expulsado del clan por tu nueva facción.", FontTypeNames.FONTTYPE_GUILD)
    End If
    
End Sub
Public Sub EnlistarMilicia(ByVal UserIndex As Integer)

Dim Matados As Integer
    Matados = (UserList(UserIndex).faccion.RenegadosMatados + UserList(UserIndex).faccion.ArmadaMatados + UserList(UserIndex).faccion.CiudadanosMatados + UserList(UserIndex).faccion.CaosMatados)
    
    If UserList(UserIndex).faccion.Milicia = 1 Then
        Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a las tropas milicianas!!! Ve a combatir enemigos", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If UserList(UserIndex).faccion.Milicia = 100 Then
        Call WriteChatOverHead(UserIndex, "¡¡¡Has sido expulsado, sal de aquí!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If UserList(UserIndex).faccion.Republicano = 0 Then
        Call WriteChatOverHead(UserIndex, "No aceptamos otros seguidores de otros bandon en la armada imperial.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If UserList(UserIndex).faccion.FuerzasCaos = 1 Or UserList(UserIndex).faccion.ArmadaReal = 1 Then
        Call WriteChatOverHead(UserIndex, "¡¡Sal de aqui!! Asqueroso enemigo.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If Matados < 200 / DIVFACCION Then
        Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos " & 200 / DIVFACCION & "  enemigos, solo has matado " & Matados, str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If UserList(UserIndex).Stats.ELV < 25 Then
        Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    UserList(UserIndex).faccion.Milicia = 1
    UserList(UserIndex).faccion.Republicano = 0
    UserList(UserIndex).faccion.Rango = 1
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))
    '------- Ropa -------
    Dim MiObj As obj
    Dim bajos As Byte
    MiObj.Amount = 1
        
    If UserList(UserIndex).Raza = Enano Or UserList(UserIndex).Raza = Gnomo Then
        MiObj.ObjIndex = 1589
    Else
        MiObj.ObjIndex = 1588
    End If
    
    'Add Marius Le da la armadura faccionaria solo si ya no la tiene.
    If Not TieneObjetos(MiObj.ObjIndex, 1, UserIndex) Then
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
    End If
    '------- Ropa -------
    
    Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al Ejercito de las Milicias!!!, aqui tienes tus vestimentas. Cumple bien tu labor exterminando Criminales y me encargaré de recompensarte.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        
End Sub
Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)

Dim Matados As Integer
    Matados = (UserList(UserIndex).faccion.RenegadosMatados + UserList(UserIndex).faccion.RepublicanosMatados + UserList(UserIndex).faccion.MilicianosMatados + UserList(UserIndex).faccion.CaosMatados)
    
    If UserList(UserIndex).faccion.ArmadaReal = 1 Then
        Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a las tropas reales!!! Ve a combatir criminales", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If UserList(UserIndex).faccion.ArmadaReal = 100 Then
        Call WriteChatOverHead(UserIndex, "¡¡¡Has sido expulsado y no queremos volver a verte!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If UserList(UserIndex).faccion.Ciudadano <> 1 Then
        Call WriteChatOverHead(UserIndex, "No aceptamos otros seguidores de otros bandon en la armada imperial.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If UserList(UserIndex).faccion.FuerzasCaos = 1 Or UserList(UserIndex).faccion.Milicia = 1 Then
        Call WriteChatOverHead(UserIndex, "¡¡Sal de aqui!! Asqueroso enemigo. Guardias, matenlo!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If Matados < 250 / DIVFACCION Then
        Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos " & 250 / DIVFACCION & "  enemigos, solo has matado " & (UserList(UserIndex).faccion.RenegadosMatados + UserList(UserIndex).faccion.RepublicanosMatados + UserList(UserIndex).faccion.MilicianosMatados + UserList(UserIndex).faccion.CaosMatados), str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    If UserList(UserIndex).Stats.ELV < 25 Then
        Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
     
    UserList(UserIndex).faccion.ArmadaReal = 1
    UserList(UserIndex).faccion.Ciudadano = 0
    UserList(UserIndex).faccion.Rango = 1
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))
    Dim MiObj As obj
    Dim bajos As Byte
    MiObj.Amount = 1
        
    If UserList(UserIndex).Raza = Enano Or UserList(UserIndex).Raza = Gnomo Then
        bajos = 1
    End If
        
    Select Case UserList(UserIndex).Clase
        Case eClass.Clerigo
            MiObj.ObjIndex = 1544 + bajos
        Case eClass.Mago
            MiObj.ObjIndex = 1546 + bajos
        Case eClass.Guerrero
            MiObj.ObjIndex = 1548 + bajos
        Case eClass.Asesino
            MiObj.ObjIndex = 1550 + bajos
        Case eClass.Bardo
            MiObj.ObjIndex = 1552 + bajos
        Case eClass.Druida
            MiObj.ObjIndex = 1554 + bajos
        Case eClass.Gladiador
            MiObj.ObjIndex = 1556 + bajos
        Case eClass.Paladin
            MiObj.ObjIndex = 1558 + bajos
        Case eClass.Cazador
            MiObj.ObjIndex = 1560 + bajos
        Case eClass.Mercenario
            MiObj.ObjIndex = 1562 + bajos
        Case eClass.Nigromante
            MiObj.ObjIndex = 1564 + bajos
    End Select
    
    'Add Marius Le da la armadura faccionaria solo si ya no la tiene.
    If Not TieneObjetos(MiObj.ObjIndex, 1, UserIndex) Then
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
    End If
    
    Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al Ejército Imperial!!!, aqui tienes tus vestimentas. Cumple bien tu labor exterminando Criminales y me encargaré de recompensarte.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    
End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)
Dim Matados As Long

    If UserList(UserIndex).faccion.Rango = 10 Then
        Exit Sub
    End If
    
    Matados = UserList(UserIndex).faccion.RenegadosMatados + UserList(UserIndex).faccion.CaosMatados + UserList(UserIndex).faccion.MilicianosMatados + UserList(UserIndex).faccion.RepublicanosMatados
    
    If Matados < matadosArmada(UserList(UserIndex).faccion.Rango) Then
        Call WriteChatOverHead(UserIndex, "Mata " & matadosArmada(UserList(UserIndex).faccion.Rango) - Matados & " Criminales más para recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    UserList(UserIndex).faccion.Rango = UserList(UserIndex).faccion.Rango + 1
    If UserList(UserIndex).faccion.Rango >= 6 Then ' Segunda jeraquia xD
        Dim MiObj As obj
        MiObj.Amount = 1
        Dim bajos As Byte
        
        If UserList(UserIndex).Raza = Enano Or UserList(UserIndex).Raza = Gnomo Then
            bajos = 1
        End If
            
        Select Case UserList(UserIndex).Clase
            Case eClass.Clerigo
                MiObj.ObjIndex = 1566 + bajos
            Case eClass.Mago
                MiObj.ObjIndex = 1568 + bajos
            Case eClass.Guerrero
                MiObj.ObjIndex = 1570 + bajos
            Case eClass.Asesino
                MiObj.ObjIndex = 1572 + bajos
            Case eClass.Bardo
                MiObj.ObjIndex = 1574 + bajos
            Case eClass.Druida
                MiObj.ObjIndex = 1576 + bajos
            Case eClass.Gladiador
                MiObj.ObjIndex = 1578 + bajos
            Case eClass.Paladin
                MiObj.ObjIndex = 1580 + bajos
            Case eClass.Cazador
                MiObj.ObjIndex = 1582 + bajos
            Case eClass.Mercenario
                MiObj.ObjIndex = 1584 + bajos
            Case eClass.Nigromante
                MiObj.ObjIndex = 1586 + bajos
        End Select
        
        'Add Marius Le da la armadura faccionaria solo si ya no la tiene.
        If Not TieneObjetos(MiObj.ObjIndex, 1, UserIndex) Then
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            End If
        End If
    End If

End Sub
Public Sub RecompensaMilicia(ByVal UserIndex As Integer)
Dim Matados As Long
    
    If UserList(UserIndex).faccion.Rango = 7 Then
        Exit Sub
    End If
    
    Matados = (UserList(UserIndex).faccion.RenegadosMatados + UserList(UserIndex).faccion.ArmadaMatados + UserList(UserIndex).faccion.CiudadanosMatados)
    If Matados < matadosArmada(UserList(UserIndex).faccion.Rango) Then
        Call WriteChatOverHead(UserIndex, "Mata " & matadosArmada(UserList(UserIndex).faccion.Rango) - Matados & " Criminales más para recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    UserList(UserIndex).faccion.Rango = UserList(UserIndex).faccion.Rango + 1
    If UserList(UserIndex).faccion.Rango = 5 Then
        Dim MiObj As obj
        MiObj.Amount = 1
        Dim bajos As Byte
        
        If UserList(UserIndex).Raza = Enano Or UserList(UserIndex).Raza = Gnomo Then
            bajos = 1
        End If
            
        Select Case UserList(UserIndex).Clase
            Case eClass.Clerigo, eClass.Mago, eClass.Bardo, eClass.Druida, eClass.Nigromante
                MiObj.ObjIndex = 1592 + bajos
            Case eClass.Guerrero, eClass.Gladiador, eClass.Cazador, eClass.Mercenario, eClass.Paladin, eClass.Asesino
                MiObj.ObjIndex = 1590 + bajos
        End Select
        
        'Add Marius Le da la armadura faccionaria solo si ya no la tiene.
        If Not TieneObjetos(MiObj.ObjIndex, 1, UserIndex) Then
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            End If
        End If
    End If
End Sub
Public Sub RecompensaCaos(ByVal UserIndex As Integer)
Dim Matados As Long

    If UserList(UserIndex).faccion.Rango = 10 Then
        Exit Sub
    End If
    
    Matados = UserList(UserIndex).faccion.RenegadosMatados + UserList(UserIndex).faccion.CaosMatados + UserList(UserIndex).faccion.MilicianosMatados + UserList(UserIndex).faccion.RepublicanosMatados
    
    If Matados < matadosCaos(UserList(UserIndex).faccion.Rango) Then
        Call WriteChatOverHead(UserIndex, "Mata " & matadosCaos(UserList(UserIndex).faccion.Rango) - Matados & " enemigos más para recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    End If
    
    UserList(UserIndex).faccion.Rango = UserList(UserIndex).faccion.Rango + 1
    If UserList(UserIndex).faccion.Rango >= 6 Then ' Segunda jeraquia xD
        Dim MiObj As obj
        MiObj.Amount = 1
        Dim bajos As Byte
        
        If UserList(UserIndex).Raza = Enano Or UserList(UserIndex).Raza = Gnomo Then
            bajos = 1
        End If
            
        Select Case UserList(UserIndex).Clase
            Case eClass.Clerigo
                MiObj.ObjIndex = 1522 + bajos
            Case eClass.Mago
                MiObj.ObjIndex = 1524 + bajos
            Case eClass.Guerrero
                MiObj.ObjIndex = 1526 + bajos
            Case eClass.Asesino
                MiObj.ObjIndex = 1528 + bajos
            Case eClass.Bardo
                MiObj.ObjIndex = 1530 + bajos
            Case eClass.Druida
                MiObj.ObjIndex = 1532 + bajos
            Case eClass.Gladiador
                MiObj.ObjIndex = 1534 + bajos
            Case eClass.Paladin
                MiObj.ObjIndex = 1536 + bajos
            Case eClass.Cazador
                MiObj.ObjIndex = 1538 + bajos
            Case eClass.Mercenario
                MiObj.ObjIndex = 1540 + bajos
            Case eClass.Nigromante
                MiObj.ObjIndex = 1542 + bajos
        End Select
        
        'Add Marius Le da la armadura faccionaria solo si ya no la tiene.
        If Not TieneObjetos(MiObj.ObjIndex, 1, UserIndex) Then
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            End If
        End If
    End If
End Sub
Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer, Optional Expulsado As Boolean = True)
    
With UserList(UserIndex)
    Call ResetFacciones(UserIndex, False)
    .faccion.Renegado = 1

    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(.Char.CharIndex, UserTypeColor(UserIndex)))

    If Expulsado Then
        Call WriteConsoleMsg(1, UserIndex, "¡¡¡Has sido expulsado de las tropas reales!!!.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(1, UserIndex, "¡¡¡Te has retirado de las tropas reales!!!.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    If .Invent.ArmourEqpObjIndex Then
        'Desequipamos la armadura real si está equipada
        If ObjData(.Invent.ArmourEqpObjIndex).Real = 1 Then
            Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
        End If
    End If
    
    If .Invent.EscudoEqpObjIndex Then
        If ObjData(.Invent.EscudoEqpObjIndex).Real = 1 Then
            Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
        End If
    End If
    
    Call QuitarItemsFaccionarios(UserIndex)
    
    If .flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)

End With

End Sub
Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer, Optional ByVal Expulsar As Boolean)
    
With UserList(UserIndex)
    Call ResetFacciones(UserIndex, False)
    .faccion.Renegado = 1
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(.Char.CharIndex, UserTypeColor(UserIndex)))

    If Expulsar Then
        Call WriteConsoleMsg(1, UserIndex, "¡¡¡Has sido expulsado de las fuerza del caos!!!.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(1, UserIndex, "¡¡¡Te has retirado de las fuerza del caos!!!.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    If .Invent.ArmourEqpObjIndex Then
        'Desequipamos la armadura real si está equipada
        If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Caos = 1 Then
            Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
        End If
    End If
    
    If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
        If ObjData(.Invent.EscudoEqpObjIndex).Caos = 1 Then
            Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
        End If
    End If
    
    Call QuitarItemsFaccionarios(UserIndex)
    
    If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)

End With

End Sub
Public Sub ExpulsarFaccionMilicia(ByVal UserIndex As Integer, Optional ByVal Expulsar As Boolean)
    
With UserList(UserIndex)
    Call ResetFacciones(UserIndex, False)
    .faccion.Renegado = 1

    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(.Char.CharIndex, UserTypeColor(UserIndex)))

    If Expulsar Then
        Call WriteConsoleMsg(1, UserIndex, "¡¡¡Has sido expulsado de las tropas republicanas.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(1, UserIndex, "¡¡¡Te has retirado de las tropas republicanas.!!!.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    If UserList(UserIndex).Invent.ArmourEqpObjIndex Then
        'Desequipamos la armadura real si está equipada
        If ObjData(.Invent.ArmourEqpObjIndex).Milicia = 1 Then Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
    End If
    
    If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
        If ObjData(.Invent.EscudoEqpObjIndex).Milicia = 1 Then Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
    End If
    
    Call QuitarItemsFaccionarios(UserIndex)
    
    If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)

End With

End Sub
Public Function TituloCaos(ByVal UserIndex As Integer) As String
    Select Case UserList(UserIndex).faccion.Rango
        Case 1
            TituloCaos = "Miembro de las Hordas"
        Case 2
            TituloCaos = "Guerrero del Caos"
        Case 3
            TituloCaos = "Teniente del Caos"
        Case 4
            TituloCaos = "Comandante del Caos"
        Case 5
            TituloCaos = "General del Caos"
        Case 6
            TituloCaos = "Elite del Caos"
        Case 7
            TituloCaos = "Asolador de las Sombras"
        Case 8
            TituloCaos = "Caballero Negro"
        Case 9
            TituloCaos = "Emisario de las Sombras"
        Case 10
            TituloCaos = "Avatar del Apocalipsis"
        Case 200
            TituloCaos = "Lider Caótico"
    End Select
End Function
Public Function TituloReal(ByVal UserIndex As Integer) As String
    Select Case UserList(UserIndex).faccion.Rango
        Case 1
            TituloReal = "Legionario"
        Case 2
            TituloReal = "Soldado Real"
        Case 3
            TituloReal = "Teniente Real"
        Case 4
            TituloReal = "Comandante Real"
        Case 5
            TituloReal = "General Real"
        Case 6
            TituloReal = "Elite Real"
        Case 7
            TituloReal = "Guardian del Bien"
        Case 8
            TituloReal = "Caballero Imperial"
        Case 9
            TituloReal = "Justiciero"
        Case 10
            TituloReal = "Guardia Imperial"
        Case 200
            TituloReal = "Lider Imperial"
    End Select
End Function
Public Function TituloMilicia(ByVal UserIndex As Integer) As String
    Select Case UserList(UserIndex).faccion.Rango
        Case 1
            TituloMilicia = "Milicia de Reserva"
        Case 2
            TituloMilicia = "Miliciano"
        Case 3
            TituloMilicia = "Miliciano Elite"
        Case 4
            TituloMilicia = "Soldado de la República"
        Case 5
            TituloMilicia = "Soldado Raso"
        Case 6
            TituloMilicia = "Soldado Elite"
        Case 7
            TituloMilicia = "Comandante de la República"
        Case 200
            TituloMilicia = "Lider Republicano"
    End Select
End Function
Public Function matadosArmada(ByVal Rango As Byte) As Integer
    Select Case Rango
        Case 1
            matadosArmada = 350 / DIVFACCION
        Case 2
            matadosArmada = 450 / DIVFACCION
        Case 3
            matadosArmada = 550 / DIVFACCION
        Case 4
            matadosArmada = 650 / DIVFACCION
        Case 5
            matadosArmada = 750 / DIVFACCION
        Case 6
            matadosArmada = 850 / DIVFACCION
        Case 7
            matadosArmada = 950 / DIVFACCION
        Case 8
            matadosArmada = 1050 / DIVFACCION
        Case 9
            matadosArmada = 1150 / DIVFACCION
    End Select
End Function
Public Function matadosCaos(ByVal Rango As Byte) As Integer
    Select Case Rango
        Case 1
            matadosCaos = 500 / DIVFACCION
        Case 2
            matadosCaos = 600 / DIVFACCION
        Case 3
            matadosCaos = 700 / DIVFACCION
        Case 4
            matadosCaos = 800 / DIVFACCION
        Case 5
            matadosCaos = 900 / DIVFACCION
        Case 6
            matadosCaos = 1000 / DIVFACCION
        Case 7
            matadosCaos = 1100 / DIVFACCION
        Case 8
            matadosCaos = 1200 / DIVFACCION
        Case 9
            matadosCaos = 1300 / DIVFACCION
    End Select
End Function
Public Function matadosMilicia(ByVal Rango As Byte) As Integer
    Select Case Rango
        Case 1
            matadosMilicia = 300 / DIVFACCION
        Case 2
            matadosMilicia = 400 / DIVFACCION
        Case 3
            matadosMilicia = 500 / DIVFACCION
        Case 4
            matadosMilicia = 600 / DIVFACCION
        Case 5
            matadosMilicia = 700 / DIVFACCION
        Case 6
            matadosMilicia = 800 / DIVFACCION
    End Select
End Function
Public Sub QuitarItemsFaccionarios(ByVal UserIndex As Integer)
    Dim i As Byte
    Dim ObjIndex As Integer
    For i = 1 To MAX_INVENTORY_SLOTS
        ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
        If ObjIndex <> 0 Then
            If ObjData(ObjIndex).Shop = 0 And (ObjData(ObjIndex).Caos = 1 Or ObjData(ObjIndex).Real = 1 Or ObjData(ObjIndex).Milicia = 1) Then
                QuitarUserInvItem UserIndex, i, UserList(UserIndex).Invent.Object(i).Amount
                UpdateUserInv False, UserIndex, i
            End If
        End If
    Next i
            
End Sub

