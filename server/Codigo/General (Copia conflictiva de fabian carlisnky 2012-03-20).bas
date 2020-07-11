Attribute VB_Name = "General"

Option Explicit

Global LeerNPCs As New clsIniReader

Sub DarCuerpoDesnudo(ByVal userindex As Integer, Optional ByVal Mimetizado As Boolean = False)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/14/07
'Da cuerpo desnudo a un usuario
'***************************************************
Dim CuerpoDesnudo As Integer
Select Case UserList(userindex).Genero
    Case eGenero.Hombre
        Select Case UserList(userindex).Raza
            Case eRaza.Humano
                CuerpoDesnudo = 21
            Case eRaza.Drow
                CuerpoDesnudo = 32
            Case eRaza.Elfo
                CuerpoDesnudo = 21
            Case eRaza.Gnomo
                CuerpoDesnudo = 53
            Case eRaza.Enano
                CuerpoDesnudo = 53
            Case eRaza.Orco
                CuerpoDesnudo = 248
        End Select
    Case eGenero.Mujer
        Select Case UserList(userindex).Raza
            Case eRaza.Humano
                CuerpoDesnudo = 39
            Case eRaza.Drow
                CuerpoDesnudo = 40
            Case eRaza.Elfo
                CuerpoDesnudo = 39
            Case eRaza.Gnomo
                CuerpoDesnudo = 60
            Case eRaza.Enano
                CuerpoDesnudo = 60
            Case eRaza.Orco
                CuerpoDesnudo = 249
        End Select
End Select

UserList(userindex).Char.body = CuerpoDesnudo

UserList(userindex).flags.Desnudo = 1

End Sub


Sub Bloquear(ByVal ToMap As Boolean, ByVal sndIndex As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal b As Boolean)
'b ahora es boolean,
'b=true bloquea el tile en (x,y)
'b=false desbloquea el tile en (x,y)
'ToMap = true -> Envia los datos a todo el mapa
'ToMap = false -> Envia los datos al user
'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s

If ToMap Then
    Call SendData(SendTarget.ToMap, sndIndex, PrepareMessageBlockPosition(x, Y, b))
Else
    Call WriteBlockPosition(sndIndex, x, Y, b)
End If

End Sub


Function HayAgua(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer) As Boolean

If map > 0 And map < NumMaps + 1 And x > 0 And x < 101 And Y > 0 And Y < 101 Then
    If ((MapData(map, x, Y).Graphic(1) >= 1505 And MapData(map, x, Y).Graphic(1) <= 1520) Or _
    (MapData(map, x, Y).Graphic(1) >= 5665 And MapData(map, x, Y).Graphic(1) <= 5680) Or _
    (MapData(map, x, Y).Graphic(1) >= 13547 And MapData(map, x, Y).Graphic(1) <= 13562)) And _
       MapData(map, x, Y).Graphic(2) = 0 Then
            HayAgua = True
    Else
            HayAgua = False
    End If
Else
  HayAgua = False
End If

End Function

Private Function HayLava(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/12/07
'***************************************************
If map > 0 And map < NumMaps + 1 And x > 0 And x < 101 And Y > 0 And Y < 101 Then
    If MapData(map, x, Y).Graphic(1) >= 5837 And MapData(map, x, Y).Graphic(1) <= 5852 Then
        HayLava = True
    Else
        HayLava = False
    End If
Else
  HayLava = False
End If

End Function


Sub EnviarSpawnList(ByVal userindex As Integer)
Dim k As Long
Dim npcNames() As String

ReDim npcNames(1 To UBound(SpawnList)) As String

For k = 1 To UBound(SpawnList)
    npcNames(k) = SpawnList(k).NpcName
Next k

Call WriteSpawnList(userindex, npcNames())

End Sub




Sub Main()
On Error Resume Next
Dim F As Date

ChDir App.Path
ChDrive App.Path

Call BanIpCargar

Prision.map = 47
Libertad.map = 47
Prision.x = 53
Libertad.x = 53

Prision.Y = 32
Libertad.Y = 68

Minutos = Format(Now, "Short Time")

IniPath = App.Path & "\"
DatPath = App.Path & "\Dat\"



LevelSkill(1).LevelValue = 3
LevelSkill(2).LevelValue = 5
LevelSkill(3).LevelValue = 7
LevelSkill(4).LevelValue = 10
LevelSkill(5).LevelValue = 13
LevelSkill(6).LevelValue = 15
LevelSkill(7).LevelValue = 17
LevelSkill(8).LevelValue = 20
LevelSkill(9).LevelValue = 23
LevelSkill(10).LevelValue = 25
LevelSkill(11).LevelValue = 27
LevelSkill(12).LevelValue = 30
LevelSkill(13).LevelValue = 33
LevelSkill(14).LevelValue = 35
LevelSkill(15).LevelValue = 37
LevelSkill(16).LevelValue = 40
LevelSkill(17).LevelValue = 43
LevelSkill(18).LevelValue = 45
LevelSkill(19).LevelValue = 47
LevelSkill(20).LevelValue = 50
LevelSkill(21).LevelValue = 53
LevelSkill(22).LevelValue = 55
LevelSkill(23).LevelValue = 57
LevelSkill(24).LevelValue = 60
LevelSkill(25).LevelValue = 63
LevelSkill(26).LevelValue = 65
LevelSkill(27).LevelValue = 67
LevelSkill(28).LevelValue = 70
LevelSkill(29).LevelValue = 73
LevelSkill(30).LevelValue = 75
LevelSkill(31).LevelValue = 77
LevelSkill(32).LevelValue = 80
LevelSkill(33).LevelValue = 83
LevelSkill(34).LevelValue = 85
LevelSkill(35).LevelValue = 87
LevelSkill(36).LevelValue = 90
LevelSkill(37).LevelValue = 93
LevelSkill(38).LevelValue = 95
LevelSkill(39).LevelValue = 97
LevelSkill(40).LevelValue = 100
LevelSkill(41).LevelValue = 100
LevelSkill(42).LevelValue = 100
LevelSkill(43).LevelValue = 100
LevelSkill(44).LevelValue = 100
LevelSkill(45).LevelValue = 100
LevelSkill(46).LevelValue = 100
LevelSkill(47).LevelValue = 100
LevelSkill(48).LevelValue = 100
LevelSkill(49).LevelValue = 100
LevelSkill(50).LevelValue = 100


ListaRazas(eRaza.Humano) = "Humano"
ListaRazas(eRaza.Elfo) = "Elfo"
ListaRazas(eRaza.Drow) = "Drow"
ListaRazas(eRaza.Gnomo) = "Gnomo"
ListaRazas(eRaza.Enano) = "Enano"
ListaRazas(eRaza.Orco) = "Orco"

ListaClases(eClass.Mago) = "Mago"
ListaClases(eClass.Clerigo) = "Clerigo"
ListaClases(eClass.Guerrero) = "Guerrero"
ListaClases(eClass.Asesino) = "Asesino"
ListaClases(eClass.Ladron) = "Ladron"
ListaClases(eClass.Bardo) = "Bardo"
ListaClases(eClass.Druida) = "Druida"
ListaClases(eClass.Paladin) = "Paladin"
ListaClases(eClass.Cazador) = "Cazador"
ListaClases(eClass.Pescador) = "Pescador"
ListaClases(eClass.Herrero) = "Herrero"
ListaClases(eClass.Leñador) = "Leñador"
ListaClases(eClass.Minero) = "Minero"
ListaClases(eClass.Carpintero) = "Carpintero"
ListaClases(eClass.Mercenario) = "Mercenario"
ListaClases(eClass.Nigromante) = "Nigromante"
ListaClases(eClass.Sastre) = "Sastre"
ListaClases(eClass.Gladiador) = "Gladiador"
    
SkillsNames(eSkill.Magia) = "Magia"
SkillsNames(eSkill.Robar) = "Robar"
SkillsNames(eSkill.Tacticas) = "Tacticas de combate"
SkillsNames(eSkill.armas) = "Combate con armas"
SkillsNames(eSkill.Meditar) = "Meditar"
SkillsNames(eSkill.Apuñalar) = "Apuñalar"
SkillsNames(eSkill.Ocultarse) = "Ocultarse"
SkillsNames(eSkill.Supervivencia) = "Supervivencia"
SkillsNames(eSkill.Talar) = "Talar arboles"
SkillsNames(eSkill.Comerciar) = "Comercio"
SkillsNames(eSkill.DefensaEscudos) = "Defensa con escudos"
SkillsNames(eSkill.Pesca) = "Pesca"
SkillsNames(eSkill.Mineria) = "Mineria"
SkillsNames(eSkill.Carpinteria) = "Carpinteria"
SkillsNames(eSkill.Herreria) = "Herreria"
SkillsNames(eSkill.Liderazgo) = "Liderazgo"
SkillsNames(eSkill.Domar) = "Domar animales"
SkillsNames(eSkill.Proyectiles) = "Armas de proyectiles"
SkillsNames(eSkill.artes) = "Artes Marciales"
SkillsNames(eSkill.Navegacion) = "Navegacion"
SkillsNames(eSkill.alquimia) = "Alquimia"
SkillsNames(eSkill.arrojadizas) = "Armas Arrojadizas"
SkillsNames(eSkill.botanica) = "Botanica"
SkillsNames(eSkill.Equitacion) = "Equitacion"
SkillsNames(eSkill.Musica) = "Musica"
SkillsNames(eSkill.Resistencia) = "Resistencia Magica"
SkillsNames(eSkill.Sastreria) = "Sastreria"

ListaAtributos(eAtributos.Fuerza) = "Fuerza"
ListaAtributos(eAtributos.Agilidad) = "Agilidad"
ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
ListaAtributos(eAtributos.Carisma) = "Carisma"
ListaAtributos(eAtributos.constitucion) = "Constitucion"


frmCargando.Show

IniPath = App.Path & "\"



'Bordes del mapa
MinXBorder = XMinMapSize + (XWindow \ 2)
MaxXBorder = XMaxMapSize - (XWindow \ 2)
MinYBorder = YMinMapSize + (YWindow \ 2)
MaxYBorder = YMaxMapSize - (YWindow \ 2)
DoEvents

frmCargando.Label1(2).Caption = "Iniciando Arrays..."
DoEvents

Call LoadGuildsDB
Call CargarSpawnList


'¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
frmCargando.Label1(2).Caption = "Cargando Server.ini"
DoEvents

MaxUsers = 0
Call LoadSini
Call CargaApuestas

'*************************************************
frmCargando.Label1(2).Caption = "Cargando NPCs.Dat"
Call CargaNpcsDat
DoEvents
'*************************************************

frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
Call LoadOBJData
    
frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
Call CargarHechizos
    
    
frmCargando.Label1(2).Caption = "Cargando Objetos de Herrería"
Call LoadArmasHerreria
Call LoadArmadurasHerreria

frmCargando.Label1(2).Caption = "Cargando Objetos de Carpintería"
Call LoadObjCarpintero

frmCargando.Label1(2).Caption = "Cargando Objetos de Alquimista"
Call LoadObjDruida

frmCargando.Label1(2).Caption = "Cargando Objetos de Sastre"

Call LoadObjSastre

frmCargando.Label1(2).Caption = "Cargando Balance.Dat"
Call LoadBalance    '4/01/08 Pablo ToxicWaste

' Jose Castelli  / Cuenta SQL
Call CargarDB
' Jose Castelli  / Cuenta SQL



frmCargando.Label1(2).Caption = "Cargando Mapas"
Call LoadMapData


'Comentado porque hay worldsave en ese mapa!
'Call CrearClanPretoriano(MAPA_PRETORIANO, ALCOBA2_X, ALCOBA2_Y)
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Dim LoopC As Integer

'Resetea las conexiones de los usuarios
For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
    Set UserList(LoopC).incomingData = New clsByteQueue
    Set UserList(LoopC).outgoingData = New clsByteQueue
Next LoopC

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

With frmMain
    .AutoSave.Enabled = True
    .tPiqueteC.Enabled = True
    .GameTimer.Enabled = True
    .Auditoria.Enabled = True
    .TIMER_AI.Enabled = True
    .npcataca.Enabled = True
End With

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Call SecurityIp.InitIpTables(1000)

Call IniciaWsApi(frmMain.hWnd)
SockListen = ListenForConnect(Puerto, hWndMsg, "")
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿


Unload frmCargando

Call frmMain.InitMain(1)

tInicioServer = GetTickCount() And &H7FFFFFFF

'Rem Nod Kopfnickend
'Add Nod Kokpfnickend (por las dudas seteamos todo por default)
Con.Execute "UPDATE `charflags` SET `Online` = 0"
Con.Execute "UPDATE `cuentas` SET `Online` = 0"

SendOnline
'/mod

End Sub

Function FileExist(ByVal file As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************
    FileExist = LenB(dir$(file, FileType)) <> 0
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'Gets a field from a delimited string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
    End If
End Function
Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        count = count + 1
    Loop While curPos <> 0
    
    FieldCount = count
End Function
Function MapaValido(ByVal map As Integer) As Boolean
MapaValido = map >= 1 And map <= NumMaps
End Function


Public Sub LogCriticEvent(desc As String)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & desc
Close #nfile

Exit Sub

Errhandler:

End Sub

Public Sub LogError(desc As String)
On Error GoTo Errhandler

Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\errores.log" For Append Shared As #nfile
        Print #nfile, Date & " " & time & " " & desc
    Close #nfile

Exit Sub

Errhandler:

End Sub



Public Sub LogHackAttemp(texto As String)
On Error GoTo Errhandler

Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
        Print #nfile, "----------------------------------------------------------"
        Print #nfile, Date & " " & time & " " & texto
        Print #nfile, "----------------------------------------------------------"
    Close #nfile

Exit Sub

Errhandler:

End Sub

Public Sub LogCheating(texto As String)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CH.log" For Append Shared As #nfile
    
Close #nfile

Exit Sub

Errhandler:

End Sub


Public Sub LogCriticalHackAttemp(texto As String)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

Errhandler:

End Sub

Public Sub LogAntiCheat(texto As String)
On Error GoTo Errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & texto
Print #nfile, ""
Close #nfile

Exit Sub

Errhandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
Dim arg As String
Dim i As Integer


For i = 1 To 33

arg = ReadField(i, cad, 44)

If LenB(arg) = 0 Then Exit Function

Next i

ValidInputNP = True

End Function


Sub Restart()


'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

Dim LoopC As Long

For LoopC = 1 To MaxUsers
    Call CloseSocket(LoopC)
Next

'Initialize statistics!!


For LoopC = 1 To UBound(UserList())
    Set UserList(LoopC).incomingData = Nothing
    Set UserList(LoopC).outgoingData = Nothing
Next LoopC

ReDim UserList(1 To MaxUsers) As User

For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
    Set UserList(LoopC).incomingData = New clsByteQueue
    Set UserList(LoopC).outgoingData = New clsByteQueue
Next LoopC

LastUser = 0
NumUsers = 0
SendOnline

Call FreeNPCs
Call FreeCharIndexes

Call LoadSini
Call LoadOBJData

Call LoadMapData

Call CargarHechizos

Call frmMain.InitMain(0)

End Sub



Public Sub TiempoInvocacion(ByVal userindex As Integer)
Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(userindex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
           Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia = _
           Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia - 1
           If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(userindex).MascotasIndex(i), 0)
        End If
    End If
Next i
End Sub

Public Sub EfectoFrio(ByVal userindex As Integer)
    
    Dim modifi As Integer
    
    If UserList(userindex).Counters.Frio < IntervaloFrio Then
        UserList(userindex).Counters.Frio = UserList(userindex).Counters.Frio + 1
    Else
        If MapInfo(UserList(userindex).Pos.map).Terreno = Nieve Then
            Call WriteConsoleMsg(1, userindex, "¡¡Estas muriendo de frio, abrigate o moriras!!.", FontTypeNames.FONTTYPE_INFO)
            modifi = Porcentaje(UserList(userindex).Stats.MaxHP, 5)
            
            UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - modifi
            
            If UserList(userindex).Stats.MinHP < 1 Then
                Call WriteConsoleMsg(1, userindex, "¡¡Has muerto de frio!!.", FontTypeNames.FONTTYPE_INFO)
                UserList(userindex).Stats.MinHP = 0
                Call UserDie(userindex)
            End If
            
            Call WriteUpdateHP(userindex)
        Else
            If UserList(userindex).Stats.MinSTA > 0 Then
                modifi = Porcentaje(UserList(userindex).Stats.MaxSTA, 5)
                Call QuitarSta(userindex, modifi)
                Call WriteUpdateSta(userindex)
            End If
        End If
        
        UserList(userindex).Counters.Frio = 0
    End If
End Sub

Public Sub EfectoInvisibilidad(ByVal userindex As Integer)

If UserList(userindex).Counters.Invisibilidad < IntervaloInvisible Then
    UserList(userindex).Counters.Invisibilidad = UserList(userindex).Counters.Invisibilidad + 1
Else
    UserList(userindex).Counters.Invisibilidad = 0
    UserList(userindex).flags.Invisible = 0
    If UserList(userindex).flags.Oculto = 0 Then
        Call WriteConsoleMsg(1, userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageSetInvisible(UserList(userindex).Char.CharIndex, False))
    End If
End If

End Sub


Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)

If Npclist(NpcIndex).Contadores.Paralisis > 0 Then
    Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 1
Else
    Npclist(NpcIndex).flags.Paralizado = 0
    Npclist(NpcIndex).flags.Inmovilizado = 0
End If

End Sub

Public Sub EfectoCegueEstu(ByVal userindex As Integer)

If UserList(userindex).Counters.Ceguera > 0 Then
    UserList(userindex).Counters.Ceguera = UserList(userindex).Counters.Ceguera - 1
Else
    If UserList(userindex).flags.Ceguera = 1 Then
        UserList(userindex).flags.Ceguera = 0
        Call WriteBlindNoMore(userindex)
    End If
    If UserList(userindex).flags.Estupidez = 1 Then
        UserList(userindex).flags.Estupidez = 0
        Call WriteDumbNoMore(userindex)
    End If

End If


End Sub


Public Sub EfectoParalisisUser(ByVal userindex As Integer)

If UserList(userindex).Counters.Paralisis > 0 Then
    UserList(userindex).Counters.Paralisis = UserList(userindex).Counters.Paralisis - 1
Else
    UserList(userindex).flags.Paralizado = 0
    UserList(userindex).flags.Inmovilizado = 0
    Call WriteParalizeOK(userindex)
End If

End Sub

Public Sub RecStamina(ByVal userindex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)

If UserList(userindex).flags.Trabajando Then Exit Sub
If UserList(userindex).flags.Desnudo Then
    If UserList(userindex).flags.Montando = 0 Then
        Exit Sub
    End If
End If

If MapData(UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.Y).Trigger = 1 And _
   MapData(UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.Y).Trigger = 2 And _
   MapData(UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.Y).Trigger = 4 Then Exit Sub


Dim massta As Integer
If UserList(userindex).Stats.MinSTA < UserList(userindex).Stats.MaxSTA And Not UserList(userindex).flags.Entrenando = 1 Then
    If UserList(userindex).Counters.STACounter < Intervalo Then
        UserList(userindex).Counters.STACounter = UserList(userindex).Counters.STACounter + 1
    Else
        EnviarStats = True
        UserList(userindex).Counters.STACounter = 0

        massta = RandomNumber(1, Porcentaje(UserList(userindex).Stats.MaxSTA, 5))
        UserList(userindex).Stats.MinSTA = UserList(userindex).Stats.MinSTA + massta
        If UserList(userindex).Stats.MinSTA > UserList(userindex).Stats.MaxSTA Then
            UserList(userindex).Stats.MinSTA = UserList(userindex).Stats.MaxSTA
        End If
    End If
End If

End Sub

Public Sub EfectoHechizoMagico(ByVal userindex As Integer)


    UserList(userindex).Stats.eCreateTipe = 0
    
    UserList(userindex).Stats.eMaxDef = 0
    UserList(userindex).Stats.eMinDef = 0
    UserList(userindex).Stats.eMaxHit = 0
    UserList(userindex).Stats.eMinHit = 0
    
    UserList(userindex).Stats.dMaxDef = 0
    UserList(userindex).Stats.dMinDef = 0


End Sub
Public Sub EfectoVeneno(ByVal userindex As Integer)
Dim N As Integer

If UserList(userindex).Counters.Veneno < IntervaloVeneno Then
    UserList(userindex).Counters.Veneno = UserList(userindex).Counters.Veneno + 1
Else
    Call WriteConsoleMsg(1, userindex, "Estás envenenado, si no te curas morirás.", FontTypeNames.FONTTYPE_VENENO)
    UserList(userindex).Counters.Veneno = 0
    N = UserList(userindex).flags.Envenenado - 2
    N = RandomNumber(1 * N, 5 * N)
    
    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - N
    
    If UserList(userindex).Stats.MinHP < 1 Then Call UserDie(userindex)
    Call WriteUpdateHP(userindex)
End If

End Sub
Public Sub EfectoIncineracion(ByVal userindex As Integer)
Dim N As Integer

If UserList(userindex).Counters.Fuego < IntervaloVeneno Then 'IntervaloFuego Then
  UserList(userindex).Counters.Fuego = UserList(userindex).Counters.Fuego + 1
Else
  Call WriteConsoleMsg(2, userindex, "Estás incendiado, si no te curas morirás.", FontTypeNames.FONTTYPE_VENENO)
  UserList(userindex).Counters.Fuego = 0
  N = RandomNumber(10, 20)
  
  UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - N
  
  Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_INCINERACION, UserList(userindex).Pos.x, UserList(userindex).Pos.Y))
  Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateCharParticle(UserList(userindex).Char.CharIndex, 96))
  
  If UserList(userindex).Stats.MinHP < 1 Then Call UserDie(userindex)
  Call WriteUpdateHP(userindex)
End If

End Sub
Public Function TieneSacri(ByVal userindex As Integer) As Byte
'****************************************************************
'Author: Leandro Mendoza (Mannakia)
'Desc:
'Last Modify: 21/10/10
'****************************************************************
On Error Resume Next
    Dim i As Long
    Dim ObjInd As Integer
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ObjInd = UserList(userindex).Invent.Object(i).ObjIndex
        If ObjInd > 0 Then
            If ObjData(ObjInd).EfectoMagico = eMagicType.Sacrificio Then
                TieneSacri = CByte(i)
                Exit Function
            End If
        End If
    Next i
    
    TieneSacri = 0

End Function
Public Sub DuracionPociones(ByVal userindex As Integer)

'Controla la duracion de las pociones
If UserList(userindex).flags.DuracionEfecto = 0 Then
    
    UserList(userindex).flags.TomoPocion = False
    UserList(userindex).flags.TipoPocion = 0
    'volvemos los atributos al estado normal
    Dim LoopX As Long
    For LoopX = 1 To NUMATRIBUTOS
          UserList(userindex).Stats.UserAtributos(LoopX) = UserList(userindex).Stats.UserAtributosBackUP(LoopX)
    Next
    Call WriteFuerza(userindex)
    Call WriteAgilidad(userindex)
    
    UserList(userindex).flags.DuracionEfecto = -1
    Exit Sub
End If

   UserList(userindex).flags.DuracionEfecto = UserList(userindex).flags.DuracionEfecto - 1


End Sub

Public Sub HambreYSed(ByVal userindex As Integer, ByRef fenviarAyS As Boolean)

If Not UserList(userindex).flags.Privilegios And PlayerType.User Then Exit Sub

'Sed
If UserList(userindex).Stats.MinAGU > 0 Then
    If UserList(userindex).Counters.AGUACounter < IntervaloSed Then
        UserList(userindex).Counters.AGUACounter = UserList(userindex).Counters.AGUACounter + 1
    Else
        UserList(userindex).Counters.AGUACounter = 0
        UserList(userindex).Stats.MinAGU = UserList(userindex).Stats.MinAGU - 10
        
        If UserList(userindex).Stats.MinAGU <= 0 Then
            UserList(userindex).Stats.MinAGU = 0
            UserList(userindex).flags.Sed = 1
        End If
        
        fenviarAyS = True
    End If
End If

'hambre
If UserList(userindex).Stats.MinHAM > 0 Then
   If UserList(userindex).Counters.COMCounter < IntervaloHambre Then
        UserList(userindex).Counters.COMCounter = UserList(userindex).Counters.COMCounter + 1
   Else
        UserList(userindex).Counters.COMCounter = 0
        UserList(userindex).Stats.MinHAM = UserList(userindex).Stats.MinHAM - 10
        If UserList(userindex).Stats.MinHAM <= 0 Then
               UserList(userindex).Stats.MinHAM = 0
               UserList(userindex).flags.Hambre = 1
        End If
        fenviarAyS = True
    End If
End If

End Sub

Public Sub Sanar(ByVal userindex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
'***************************************************************************
'Author: Leandro Mendoza(Mannakia)
'Desc:
'LastModify: 21/10/10
'***************************************************************************
If MapData(UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.Y).Trigger = 1 And _
   MapData(UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.Y).Trigger = 2 And _
   MapData(UserList(userindex).Pos.map, UserList(userindex).Pos.x, UserList(userindex).Pos.Y).Trigger = 4 Then Exit Sub

Dim mashit As Integer

'Mannakia
If UserList(userindex).Invent.MagicIndex <> 0 Then
    If ObjData(UserList(userindex).Invent.MagicIndex).EfectoMagico = eMagicType.AceleraVida Then
        Intervalo = Intervalo - Porcentaje(Intervalo, 40)  ' nose algo razonable
    End If
End If
'Mannakia

'con el paso del tiempo va sanando....pero muy lentamente ;-)
If UserList(userindex).Stats.MinHP < UserList(userindex).Stats.MaxHP Then
    If UserList(userindex).Counters.HPCounter < Intervalo Then
        UserList(userindex).Counters.HPCounter = UserList(userindex).Counters.HPCounter + 1
    Else
        mashit = RandomNumber(2, Porcentaje(UserList(userindex).Stats.MaxSTA, 5))
        
        UserList(userindex).Counters.HPCounter = 0
        UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP + mashit
        If UserList(userindex).Stats.MinHP > UserList(userindex).Stats.MaxHP Then UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
        Call WriteConsoleMsg(1, userindex, "Has sanado.", FontTypeNames.FONTTYPE_INFO)
        EnviarStats = True
    End If
End If

End Sub

Public Sub CargaNpcsDat()
    Dim npcfile As String
    
    npcfile = DatPath & "NPCs.dat"
    Call LeerNPCs.Initialize(npcfile)
End Sub

Sub PasarSegundo()
On Error GoTo Errhandler
    Dim i As Long
    
    
    
    tSeg = tSeg + 1
    If tSeg >= 15 Then
        tSeg = 0
        tMinuto = tMinuto + 1
        If tMinuto >= 60 Then
            tMinuto = 0
            tHora = tHora + 1
            If tHora = 24 Then tHora = 0
        End If
    End If
    
    If tSeg < 0 Then tSeg = 1
    
    Call PasarSegundotelep
    
    
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
        
        
        
            'Cerrar usuario
            If UserList(i).Counters.Saliendo Then
                UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
    
                If UserList(i).Counters.Salir <= 0 Then
                        
                        Call WriteMsg(i, 43)
                        Call WriteDisconnect(i)
                        Call FlushBuffer(i)
                        Call CloseSocket(i)
                Else
                
                    Call WriteMsg(i, 27, CStr(UserList(i).Counters.Salir))

                End If
            End If
            
            If UserList(i).Counters.Habla > 10 Then
                UserList(i).Counters.Silenciado = 500
            End If
            UserList(i).Counters.Habla = 0
      
            If UserList(i).flags.TimesWalk >= 20 Then
                If Not UserList(i).flags.Paralizado Then
                    UserList(i).flags.Paralizado = 1
                    UserList(i).Counters.Paralisis = IntervaloParalizado
                    Call Protocol.WriteParalizeOK(i)
                End If
            End If

            UserList(i).flags.TimesWalk = 0
            
        End If
    
    
    
    Next i
    
    
    
    
    
Exit Sub

Errhandler:
    Call LogError("Error en PasarSegundo. Err: " & err.description & " - " & err.Number & " - UserIndex: " & i)
   ' Resume Next
End Sub

Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)

    'Guardar Pjs
    Call GuardarUsuarios
    
    If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

    'Chauuu
    Unload frmMain

End Sub

 
Sub GuardarUsuarios()
    haciendoBK = True
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Servidor> Grabando Personajes", FontTypeNames.FONTTYPE_SERVER))
    
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUserSQL(i)
        End If
    Next i
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Servidor> Personajes Grabados", FontTypeNames.FONTTYPE_SERVER))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

    haciendoBK = False
End Sub




Public Sub FreeNPCs()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all NPC Indexes
'***************************************************
    Dim LoopC As Long
    
    ' Free all NPC indexes
    For LoopC = 1 To MAXNPCS
        Npclist(LoopC).flags.NPCActive = False
    Next LoopC
End Sub

Public Sub FreeCharIndexes()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all char indexes
'***************************************************
    ' Free all char indexes (set them all to 0)
    Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))
End Sub


Sub PasarSegundotelep()
Dim i As Integer, mapa As Integer, x As Integer, Y As Integer

For i = 1 To LastUser
    If UserList(i).Counters.CreoTeleport = True Then
        mapa = UserList(i).flags.DondeTiroMap
        x = UserList(i).flags.DondeTiroX
        Y = UserList(i).flags.DondeTiroY
    
        UserList(i).Counters.TimeTeleport = UserList(i).Counters.TimeTeleport + 1
    
        If UserList(i).Counters.TimeTeleport = 5 Then
            Call SendData(SendTarget.ToPCArea, i, PrepareMessageDestParticle(x, Y))
            Call SendData(SendTarget.ToPCArea, i, PrepareMessageCreateParticle(x, Y, 34))
            MapData(UserList(i).Pos.map, x, Y).TileExit.map = 49
            MapData(UserList(i).Pos.map, x, Y).TileExit.x = 50
            MapData(UserList(i).Pos.map, x, Y).TileExit.Y = 50
    
        Else
            If UserList(i).Counters.TimeTeleport = 18 Then
                Call SendData(SendTarget.ToPCArea, i, PrepareMessageDestParticle(x, Y))
                Call ControlarPortalLum(i)
            End If
        End If
   End If
Next i
 
End Sub



Sub ControlarPortalLum(ByVal userindex As Integer)
   
    If UserList(userindex).Counters.CreoTeleport = True Then
        MapData(UserList(userindex).flags.DondeTiroMap, UserList(userindex).flags.DondeTiroX, UserList(userindex).flags.DondeTiroY).TileExit.map = 0
        MapData(UserList(userindex).flags.DondeTiroMap, UserList(userindex).flags.DondeTiroX, UserList(userindex).flags.DondeTiroY).TileExit.x = 0
        MapData(UserList(userindex).flags.DondeTiroMap, UserList(userindex).flags.DondeTiroX, UserList(userindex).flags.DondeTiroY).TileExit.Y = 0
        UserList(userindex).flags.DondeTiroMap = 0
        UserList(userindex).flags.DondeTiroX = 0
        UserList(userindex).flags.DondeTiroY = 0
        UserList(userindex).flags.TiroPortalL = 0
        UserList(userindex).Counters.TimeTeleport = 0
        UserList(userindex).Counters.CreoTeleport = False
    End If

End Sub


Public Sub DoMetamorfosis(ByVal userindex As Integer, ByVal body As Integer, ByVal Head As Integer) 'Metamorfosis

Dim tUser As Integer
tUser = userindex

If UserList(userindex).flags.Muerto = 1 Then Exit Sub
 
UserList(tUser).Char.Head = Head
UserList(tUser).Char.body = body
UserList(tUser).Char.ShieldAnim = NingunEscudo
UserList(tUser).Char.WeaponAnim = NingunArma
UserList(tUser).Char.CascoAnim = NingunCasco

UserList(userindex).flags.Metamorfosis = 1
 
Call ChangeUserChar(userindex, UserList(tUser).Char.body, UserList(tUser).Char.Head, UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)

End Sub
 
Public Sub EfectoMetamorfosis(ByVal userindex As Integer) 'Metamorfosis
    If UserList(userindex).Counters.Metamorfosis < IntervaloInvisible * 2 Then
        UserList(userindex).Counters.Metamorfosis = UserList(userindex).Counters.Metamorfosis + 1
    Else
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
End Sub


Public Sub limpiaritemsmundo()


Dim Xnn, Ynn As Byte
Static limpma As Integer

If limpma > NumMaps Then
limpma = 0
End If

limpma = limpma + 1



If MapInfo(limpma).Pk = True Then
    For Ynn = YMinMapSize To YMaxMapSize
        For Xnn = XMinMapSize To XMaxMapSize
            If MapData(limpma, Xnn, Ynn).ObjInfo.ObjIndex > 0 And MapData(limpma, Xnn, Ynn).Blocked = 0 Then
                    Call EraseObj(10000, limpma, val(Xnn), val(Ynn))
            End If
        Next Xnn
    Next Ynn
End If




End Sub
