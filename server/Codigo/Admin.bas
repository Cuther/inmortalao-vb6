Attribute VB_Name = "Admin"

Option Explicit

'ReAdd Marius
Public MotdMaxLines As Byte
Public MOTD() As String

Public PubMaxLines As Byte
Public PUBLICIDAD() As String

Public MsgEvento As String
'\ReAdd

Public Type tAPuestas
    Ganancias As Long
    Perdidas As Long
    Jugadas As Long
End Type
Public Apuestas As tAPuestas

Public tInicioServer As Long

'INTERVALOS
Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloSed As Integer
Public IntervaloHambre As Integer
Public IntervaloVeneno As Integer
Public IntervaloFuego As Integer
Public IntervaloParalizado As Integer
'Public IntervaloHechizosMagicos As Integer
Public IntervaloInvisible As Integer
Public IntervaloFrio As Integer
Public IntervaloWavFx As Integer
Public IntervaloNPCPuedeAtacar As Integer
Public IntervaloNPCAI As Integer
Public IntervaloInvocacion As Integer
Public IntervaloOculto As Integer '[Nacho]
Public IntervaloGolpeUsar As Long
Public IntervaloUserPuedeTrabajar As Long
Public IntervaloParaConexion As Long
Public IntervaloCerrarConexion As Long '[Gonzalo]

Public IntervaloUserPuedeUsar As Long
Public IntervaloUserPuedeAtacar As Long
Public IntervaloMagiaGolpe As Long
Public IntervaloGolpeMagia As Long
Public IntervaloFlechasCazadores As Long
Public IntervaloUserPuedeCastear As Long
Public IntervaloUserMove As Long



'BALANCE

Public PorcentajeRecuperoMana As Integer


Public Puerto As Integer

Public DeNoche As Boolean

Sub ReSpawnOrigPosNpcs()
    On Error Resume Next
    
    Dim i As Integer
    Dim MiNPC As npc
       
    For i = 1 To LastNPC
       If Npclist(i).flags.NPCActive Then
            If InMapBounds(Npclist(i).Orig.map, Npclist(i).Orig.x, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)
                Call ReSpawnNpc(MiNPC)
            End If
       End If
    Next i
End Sub



Public Sub PurgarPenas()
    Dim i As Long
    
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            If UserList(i).Counters.Pena > 0 Then
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, Libertad.map, Libertad.x, Libertad.Y, True)
                    Call WriteConsoleMsg(1, i, "Has sido liberado!", FontTypeNames.FONTTYPE_INFO)
                    
                    Call FlushBuffer(i)
                End If
            End If
        End If
    Next i
End Sub


Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal minutos As Long, Optional ByVal GmName As String = vbNullString, Optional ByVal razon As String = vbNullString)
    UserList(UserIndex).Counters.Pena = minutos
    
    
    'Add Marius Solucionamos un error con las arenas xD
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

    Call salir_listas_espera(UserIndex)
    '\Add
    
    If minutos = 0 Then
        Call WarpUserChar(UserIndex, Libertad.map, Libertad.x, Libertad.Y, True)
        Call WriteConsoleMsg(1, UserIndex, "Has sido liberado!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    
    Call WarpUserChar(UserIndex, Prision.map, Prision.x, Prision.Y, True)
    If Len(razon) = 0 Then
        Call WriteConsoleMsg(1, UserIndex, GmName & " te ha encarcelado por: " & razon & ", deberas permanecer en la carcel " & minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    ElseIf LenB(GmName) = 0 Then
        Call WriteConsoleMsg(1, UserIndex, "Has sido encarcelado, deberas permanecer en la carcel " & minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(1, UserIndex, GmName & " te ha encarcelado, deberas permanecer en la carcel " & minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub


Public Sub BanIpAgrega(ByVal ip As String)
    BanIps.Add ip
    
    Call BanIpGuardar
End Sub

Public Function BanIpBuscar(ByVal ip As String) As Long
Dim Dale As Boolean
Dim loopC As Long

    Dale = True
    loopC = 1
    Do While loopC <= BanIps.count And Dale
        Dale = (BanIps.Item(loopC) <> ip)
        loopC = loopC + 1
    Loop
    
    If Dale Then
        BanIpBuscar = 0
    Else
        BanIpBuscar = loopC - 1
    End If
End Function

Public Function BanIpQuita(ByVal ip As String) As Boolean
On Error Resume Next
Dim N As Long

    N = BanIpBuscar(ip)
    If N > 0 Then
        BanIps.Remove N
        BanIpGuardar
        BanIpQuita = True
    Else
        BanIpQuita = False
    End If

End Function

Public Sub BanIpGuardar()
Dim ArchivoBanIp As String
Dim ArchN As Long
Dim loopC As Long

    ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
    
    ArchN = FreeFile()
    Open ArchivoBanIp For Output As #ArchN
    
        For loopC = 1 To BanIps.count
            Print #ArchN, BanIps.Item(loopC)
        Next loopC
    
    Close #ArchN

End Sub

Public Sub BanIpCargar()
Dim ArchN As Long
Dim Tmp As String
Dim ArchivoBanIp As String

    ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
    
    Do While BanIps.count > 0
        BanIps.Remove 1
    Loop
    
    ArchN = FreeFile()
    Open ArchivoBanIp For Input As #ArchN
    
        Do While Not EOF(ArchN)
            Line Input #ArchN, Tmp
            BanIps.Add Tmp
        Loop
    
    Close #ArchN

End Sub



Public Function UserDarPrivilegioLevel(ByVal Name As String) As PlayerType
'***************************************************
'Author: Unknown
'Last Modification: 03/02/07
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'***************************************************
    If EsVIP(NameIndex(Name)) Then
        UserDarPrivilegioLevel = PlayerType.VIP
    ElseIf EsADMIN(NameIndex(Name)) Then
        UserDarPrivilegioLevel = PlayerType.Admin
    ElseIf EsDIOS(NameIndex(Name)) Then
        UserDarPrivilegioLevel = PlayerType.Dios
    ElseIf EsSEMI(NameIndex(Name)) Then
        UserDarPrivilegioLevel = PlayerType.Semi
    ElseIf EsCONSE(NameIndex(Name)) Then
        UserDarPrivilegioLevel = PlayerType.Conse
    
    Else
        UserDarPrivilegioLevel = PlayerType.User
    End If
End Function

