Attribute VB_Name = "General"

Option Explicit

Global LeerNPCs As New clsIniReader

Sub DarCuerpoDesnudo(ByVal UserIndex As Integer, Optional ByVal Mimetizado As Boolean = False)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/14/07
'Da cuerpo desnudo a un usuario
'***************************************************
Dim CuerpoDesnudo As Integer
Select Case UserList(UserIndex).Genero
    Case eGenero.Hombre
        Select Case UserList(UserIndex).Raza
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
        Select Case UserList(UserIndex).Raza
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

UserList(UserIndex).Char.body = CuerpoDesnudo

UserList(UserIndex).flags.Desnudo = 1

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

Sub LimpiarMundo()
'***************************************************
'Author: Unknow
'Last Modification: 04/15/2008
'01/14/2008: Marcos Martinez (ByVal) - La funcion FOR estaba mal. En ves de i habia un 1.
'04/15/2008: (NicoNZ) - La funcion FOR estaba mal, de la forma que se hacia tiraba error.
'***************************************************
On Error GoTo Errhandler

    Dim i As Integer
    Dim d As cGarbage
    Set d = New cGarbage
    
    For i = TrashCollector.count To 1 Step -1
        Set d = TrashCollector(i)
        Call EraseObj(1, d.map, d.x, d.Y)
        Call TrashCollector.Remove(i)
        Set d = Nothing
    Next i
    
    Call SecurityIp.IpSecurityMantenimientoLista
    
    Exit Sub

Errhandler:
    Call LogError("Error producido en el sub LimpiarMundo: " & err.description)
End Sub

Sub EnviarSpawnList(ByVal UserIndex As Integer)
Dim k As Long
Dim npcNames() As String

    ReDim npcNames(1 To UBound(SpawnList)) As String
    
    For k = 1 To UBound(SpawnList)
        npcNames(k) = SpawnList(k).NpcName
    Next k
    
    Call WriteSpawnList(UserIndex, npcNames())

End Sub




Sub Main()
On Error Resume Next
Dim F As Date

    ChDir App.Path
    ChDrive App.Path
    
    Call BanIpCargar
    
    Prision.map = 47
    Prision.x = 53
    Prision.Y = 32
    
    Libertad.map = 47
    Libertad.Y = 68
    Libertad.x = 53
    
    minutos = Format(Now, "Short Time")
    
    IniPath = App.Path & "\"
    DatPath = App.Path & "\Dat\"
    
    'Add Marius mas prolijidad en los log de errores
    LogError ("///////////////////////////////// Se levantó el servidor /////////////////////////////////")
    '\Add
    
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
    
    Call CargarSpawnList
    
    
    ' Jose Castelli  / Cuenta SQL
    frmCargando.Label1(2).Caption = "Conectando con la Base de Datos"
    DoEvents
    Call CargarDB
    ' Jose Castelli  / Cuenta SQL
    
    'Add Marius
    frmCargando.Label1(2).Caption = "Optimizando las tablas de la Base de Datos"
    DoEvents
    Call MySQL_Optimize
    '\Add
    
    'Add Marius
    frmCargando.Label1(2).Caption = "Cargando Clanes"
    DoEvents
    Call LoadGuildsDB
    '\Add
    
    '¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
    frmCargando.Label1(2).Caption = "Cargando Server.ini"
    DoEvents
    
    MaxUsers = 0
    Call LoadSini
    Call CargaApuestas
    
    '*************************************************
    frmCargando.Label1(2).Caption = "Cargando NPCs.Dat"
    Call CargaNpcsDat
    'Call insertarNpcsDB
    
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
        
        
    frmCargando.Label1(2).Caption = "Cargando Mapas"
    Call LoadMapData
    
    
    Dim loopC As Integer
    
    'Resetea las conexiones de los usuarios
    For loopC = 1 To MaxUsers
        UserList(loopC).ConnID = -1
        UserList(loopC).ConnIDValida = False
        Set UserList(loopC).incomingData = New clsByteQueue
        Set UserList(loopC).outgoingData = New clsByteQueue
    Next loopC
    
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    With frmMain
        .AutoSave.Enabled = True
        .tPiqueteC.Enabled = True
        .GameTimer.Enabled = True
        .Auditoria.Enabled = True
        .TIMER_AI.Enabled = True
        .npcataca.Enabled = True
        .minuto.Enabled = True
    End With
    
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    Call SecurityIp.InitIpTables(1000)
    
    Call IniciaWsApi(frmMain.hWnd)
    SockListen = ListenForConnect(Puerto, hWndMsg, "")
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    
    Unload frmCargando
    
    frmMain.Show
    'Call frmMain.InitMain(1)
    
    tInicioServer = GetTickCount() And &H7FFFFFFF
    
    'Add Nod Kokpfnickend (por las dudas seteamos todo por default)
    DB_Conn.Execute "UPDATE `charflags` SET `Online` = 0"
    DB_Conn.Execute "UPDATE `cuentas` SET `Online` = 0"
    Call extra_set("EventoX2", "0")
    SendOnline
    '/Add
    
    'Add Marius Iniciamos los valores de las arenas.
    Call arenas_iniciar
    '\Add
    
    
    'Add Nod kopfnickend
    'NPCs de venta de hechizos Fixeamos aca por que no tenemos world edit
        Dim Pos As WorldPos
        
        ' paladines en argal
        Pos.map = 151
        Pos.x = 67
        Pos.Y = 29
        Call SpawnNpc(156, Pos, True, False) 'Paladin
        
        'Bander
        Pos.map = 59
        Pos.x = 82
        Pos.Y = 82
        Call SpawnNpc(156, Pos, True, False) 'Paladin
        
        Pos.map = 59
        Pos.x = 82
        Pos.Y = 85
        Call SpawnNpc(157, Pos, True, False) 'Bardo
        
        Pos.map = 59
        Pos.x = 82
        Pos.Y = 88
        Call SpawnNpc(61, Pos, True, False) 'Mago
        
        'pos.map = 59
        'pos.x = 86
        'pos.Y = 88
        'Call SpawnNpc(NpcIndex, pos, True, False) '
        
        'pos.map = 59
        'pos.x = 90
        'pos.Y = 88
        'Call SpawnNpc(NpcIndex, pos, True, False) '
        
        Pos.map = 59
        Pos.x = 90
        Pos.Y = 82
        Call SpawnNpc(47, Pos, True, False) 'Mago
        
        Pos.map = 59
        Pos.x = 90
        Pos.Y = 85
        Call SpawnNpc(154, Pos, True, False) 'Clero
        
        Pos.map = 59
        Pos.x = 90
        Pos.Y = 88
        Call SpawnNpc(104, Pos, True, False) 'Druida
    
    '\Add
    
    'Add Nod Kopfnickend Fixeamos aca por que no hay worldedit (por ahora) (¬¬)
    Dim ET As obj
    ET.Amount = 1
    ET.ObjIndex = 378
        
    'Call MakeObj(ET, 237, 47, 55)
    With MapData(237, 47, 55)
        .TileExit.map = 237
        .TileExit.x = 75
        .TileExit.Y = 56
    End With

    'Call MakeObj(ET, 237, 46, 54)
    With MapData(237, 46, 54)
        .TileExit.map = 237
        .TileExit.x = 76
        .TileExit.Y = 56
    End With

    'Call MakeObj(ET, 237, 47, 53)
    With MapData(237, 47, 53)
        .TileExit.map = 237
        .TileExit.x = 78
        .TileExit.Y = 56
    End With

    'Call MakeObj(ET, 237, 48, 54)
    With MapData(237, 48, 54)
        .TileExit.map = 237
        .TileExit.x = 79
        .TileExit.Y = 56
    End With

    'Blockear para que no hagan trampa y vuelvan para atras
    MapData(237, 79, 57).Blocked = 1
    MapData(237, 78, 57).Blocked = 1
    MapData(237, 77, 57).Blocked = 1
    MapData(237, 76, 57).Blocked = 1
    MapData(237, 75, 57).Blocked = 1
    
    'Bloqueamos los tildes y los dejamos atrapados cuando los sumoneamos
    MapData(237, 44, 68).Blocked = 1
    MapData(237, 50, 68).Blocked = 1
    MapData(237, 44, 41).Blocked = 1
    MapData(237, 50, 41).Blocked = 1
    
    'Ponemos las sercas
    'ET.ObjIndex = 1190
    'Call MakeObj(ET, 237, 75, 57)
    'Call MakeObj(ET, 237, 79, 57)

    'Portales a Itermundia, (sin el objeto portal, asi son invisibles)
    'Pasan por la meta y van a intermundia derecho.
    With MapData(237, 75, 59)
        .TileExit.map = 49
        .TileExit.x = 52
        .TileExit.Y = 50
    End With

    With MapData(237, 76, 59)
        .TileExit.map = 49
        .TileExit.x = 51
        .TileExit.Y = 50
    End With

    With MapData(237, 77, 59)
        .TileExit.map = 49
        .TileExit.x = 50
        .TileExit.Y = 50
    End With

    With MapData(237, 78, 59)
        .TileExit.map = 49
        .TileExit.x = 49
        .TileExit.Y = 50
    End With

    With MapData(237, 79, 59)
        .TileExit.map = 49
        .TileExit.x = 48
        .TileExit.Y = 50
    End With
    '\Add
    
    'Add Marius Bloqueamos una pos para que no se roben objetos que vienen en el mapa
    MapData(682, 33, 32).Blocked = 1
    MapData(682, 33, 33).Blocked = 1
    MapData(682, 33, 34).Blocked = 1
    MapData(682, 33, 35).Blocked = 1
    '\Add
    
    'Add Marius Portal a sala de Aguite.
    Call MakeObj(ET, 34, 18, 68)
    With MapData(34, 18, 68)
        .TileExit.map = 238
        .TileExit.x = 50
        .TileExit.Y = 80
    End With
    '\Add
    
    'Add Marius Sillones en Jardin secreto
    ET.Amount = 1
    ET.ObjIndex = 1184
    
    Call MakeObj(ET, 248, 48, 50)
    Call MakeObj(ET, 248, 52, 50)
    Call MakeObj(ET, 248, 54, 52)
    Call MakeObj(ET, 248, 50, 52)
    Call MakeObj(ET, 248, 46, 52)
    '\Add
    
    MsgEvento = ""
    
    'Add Marius Seguimos cagando a los presos. (par que no trucheen presos) hay que tener mil ojos con estos putos!
    MapInfo(Prision.map).ResuSinEfecto = 1
    '\Add
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

Public Sub LogGM(Nombre As String, desc As String)
On Error GoTo Errhandler

Dim nfile As Integer
    
    If Nombre <> "Marius" Then
        nfile = FreeFile ' obtenemos un canal
        Open App.Path & "\logs\GMS\" & Nombre & ".log" For Append Shared As #nfile
            Print #nfile, Date & " " & time & " " & desc
        Close #nfile
    End If
Exit Sub

Errhandler:

End Sub

Public Sub LogBruteforce(desc As String)
On Error GoTo Errhandler

Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\BruteForce.log" For Append Shared As #nfile
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

Dim loopC As Long

    For loopC = 1 To MaxUsers
        Call CloseSocket(loopC)
    Next
    
    'Initialize statistics!!
    
    
    For loopC = 1 To UBound(UserList())
        Set UserList(loopC).incomingData = Nothing
        Set UserList(loopC).outgoingData = Nothing
    Next loopC
    
    ReDim UserList(1 To MaxUsers) As User
    
    For loopC = 1 To MaxUsers
        UserList(loopC).ConnID = -1
        UserList(loopC).ConnIDValida = False
        Set UserList(loopC).incomingData = New clsByteQueue
        Set UserList(loopC).outgoingData = New clsByteQueue
    Next loopC
    
    LastUser = 0
    NumUsers = 0
    SendOnline
    
    Call FreeNPCs
    Call FreeCharIndexes
    
    Call LoadSini
    Call LoadOBJData
    
    Call LoadMapData
    
    Call CargarHechizos
    
    'Call frmMain.InitMain(0)

End Sub



Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
           Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = _
           Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia - 1
           If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
        End If
    End If
Next i
End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)
    
    Dim modifi As Integer
    
    If UserList(UserIndex).Counters.Frio < IntervaloFrio Then
        UserList(UserIndex).Counters.Frio = UserList(UserIndex).Counters.Frio + 1
    Else
        If MapInfo(UserList(UserIndex).Pos.map).Terreno = Nieve Then
            Call WriteConsoleMsg(1, UserIndex, "¡¡Estas muriendo de frio, abrigate o moriras!!.", FontTypeNames.FONTTYPE_INFO)
            modifi = Porcentaje(UserList(UserIndex).Stats.MaxHP, 5)
            
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - modifi
            
            If UserList(UserIndex).Stats.MinHP < 1 Then
                Call WriteConsoleMsg(1, UserIndex, "¡¡Has muerto de frio!!.", FontTypeNames.FONTTYPE_INFO)
                UserList(UserIndex).Stats.MinHP = 0
                Call UserDie(UserIndex)
            End If
            
            Call WriteUpdateHP(UserIndex)
        Else
            If UserList(UserIndex).Stats.MinSTA > 0 Then
                modifi = Porcentaje(UserList(UserIndex).Stats.MaxSTA, 5)
                Call QuitarSta(UserIndex, modifi)
                Call WriteUpdateSta(UserIndex)
            End If
        End If
        
        UserList(UserIndex).Counters.Frio = 0
    End If
End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)

If UserList(UserIndex).Counters.Invisibilidad < IntervaloInvisible Then
    UserList(UserIndex).Counters.Invisibilidad = UserList(UserIndex).Counters.Invisibilidad + 1
Else
    UserList(UserIndex).Counters.Invisibilidad = 0
    UserList(UserIndex).flags.Invisible = 0
    If UserList(UserIndex).flags.Oculto = 0 Then
        Call WriteConsoleMsg(1, UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
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

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)

    If UserList(UserIndex).Counters.Ceguera > 0 Then
        UserList(UserIndex).Counters.Ceguera = UserList(UserIndex).Counters.Ceguera - 1
    Else
        If UserList(UserIndex).flags.Ceguera = 1 Then
            UserList(UserIndex).flags.Ceguera = 0
            Call WriteBlindNoMore(UserIndex)
        End If
        If UserList(UserIndex).flags.Estupidez = 1 Then
            UserList(UserIndex).flags.Estupidez = 0
            Call WriteDumbNoMore(UserIndex)
        End If
    
    End If

End Sub


Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)

    If UserList(UserIndex).Counters.Paralisis > 0 Then
        UserList(UserIndex).Counters.Paralisis = UserList(UserIndex).Counters.Paralisis - 1
    Else
        UserList(UserIndex).flags.Paralizado = 0
        UserList(UserIndex).flags.Inmovilizado = 0
        Call WriteParalizeOK(UserIndex)
    End If

End Sub

Public Sub RecStamina(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)

    If UserList(UserIndex).flags.Trabajando Then Exit Sub
    If UserList(UserIndex).flags.Desnudo Then
        If UserList(UserIndex).flags.Montando = 0 Then
            Exit Sub
        End If
    End If
    
    If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).Trigger = 1 And _
       MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).Trigger = 2 And _
       MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).Trigger = 4 Then Exit Sub
    
    
    Dim massta As Integer
    If UserList(UserIndex).Stats.MinSTA < UserList(UserIndex).Stats.MaxSTA And Not UserList(UserIndex).flags.Entrenando = 1 Then
        If UserList(UserIndex).Counters.STACounter < Intervalo Then
            UserList(UserIndex).Counters.STACounter = UserList(UserIndex).Counters.STACounter + 1
        Else
            EnviarStats = True
            UserList(UserIndex).Counters.STACounter = 0

            massta = RandomNumber(1, Porcentaje(UserList(UserIndex).Stats.MaxSTA + UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia), 5))
            UserList(UserIndex).Stats.MinSTA = UserList(UserIndex).Stats.MinSTA + massta
            If UserList(UserIndex).Stats.MinSTA > UserList(UserIndex).Stats.MaxSTA Then
                UserList(UserIndex).Stats.MinSTA = UserList(UserIndex).Stats.MaxSTA
            End If
        End If
    End If

End Sub

Public Sub EfectoHechizoMagico(ByVal UserIndex As Integer)


    UserList(UserIndex).Stats.eCreateTipe = 0
    
    UserList(UserIndex).Stats.eMaxDef = 0
    UserList(UserIndex).Stats.eMinDef = 0
    UserList(UserIndex).Stats.eMaxHit = 0
    UserList(UserIndex).Stats.eMinHit = 0
    
    UserList(UserIndex).Stats.dMaxDef = 0
    UserList(UserIndex).Stats.dMinDef = 0


End Sub
Public Sub EfectoVeneno(ByVal UserIndex As Integer)
Dim N As Integer

If UserList(UserIndex).Counters.Veneno < IntervaloVeneno Then
    UserList(UserIndex).Counters.Veneno = UserList(UserIndex).Counters.Veneno + 1
Else
    Call WriteConsoleMsg(1, UserIndex, "Estás envenenado, si no te curas morirás.", FontTypeNames.FONTTYPE_VENENO)
    UserList(UserIndex).Counters.Veneno = 0
    N = UserList(UserIndex).flags.Envenenado - 2
    N = RandomNumber(1 * N, 5 * N)
    
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - N
    
    If UserList(UserIndex).Stats.MinHP < 1 Then Call UserDie(UserIndex)
    Call WriteUpdateHP(UserIndex)
End If

End Sub
Public Sub EfectoIncineracion(ByVal UserIndex As Integer)
Dim N As Integer

If UserList(UserIndex).Counters.Fuego < IntervaloVeneno Then 'IntervaloFuego Then
  UserList(UserIndex).Counters.Fuego = UserList(UserIndex).Counters.Fuego + 1
Else
  Call WriteConsoleMsg(2, UserIndex, "Estás incendiado, si no te curas morirás.", FontTypeNames.FONTTYPE_VENENO)
  UserList(UserIndex).Counters.Fuego = 0
  N = RandomNumber(10, 20)
  
  UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - N
  
  Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_INCINERACION, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
  Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateCharParticle(UserList(UserIndex).Char.CharIndex, 96))
  
  If UserList(UserIndex).Stats.MinHP < 1 Then Call UserDie(UserIndex)
  Call WriteUpdateHP(UserIndex)
End If

End Sub
Public Function TieneSacri(ByVal UserIndex As Integer) As Byte
'****************************************************************
'Author: Leandro Mendoza (Mannakia)
'Desc:
'Last Modify: 21/10/10
'****************************************************************
On Error Resume Next
    Dim i As Long
    Dim ObjInd As Integer
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ObjInd = UserList(UserIndex).Invent.Object(i).ObjIndex
        If ObjInd > 0 Then
            If UserList(UserIndex).Invent.Object(i).Equipped = 1 And ObjData(ObjInd).EfectoMagico = eMagicType.Sacrificio Then
                TieneSacri = CByte(i)
                Exit Function
            End If
        End If
    Next i
    
    TieneSacri = 0

End Function
Public Sub DuracionPociones(ByVal UserIndex As Integer)

'Controla la duracion de las pociones
If UserList(UserIndex).flags.DuracionEfecto = 0 Then
    
    UserList(UserIndex).flags.TomoPocion = False
    UserList(UserIndex).flags.TipoPocion = 0
    'volvemos los atributos al estado normal
    Dim LoopX As Long
    For LoopX = 1 To NUMATRIBUTOS
          UserList(UserIndex).Stats.UserAtributos(LoopX) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopX)
    Next
    Call WriteFuerza(UserIndex)
    Call WriteAgilidad(UserIndex)
    
    UserList(UserIndex).flags.DuracionEfecto = -1
    Exit Sub
End If

   UserList(UserIndex).flags.DuracionEfecto = UserList(UserIndex).flags.DuracionEfecto - 1


End Sub

Public Sub HambreYSed(ByVal UserIndex As Integer, ByRef fenviarAyS As Boolean)

If Not UserList(UserIndex).flags.Privilegios And PlayerType.User Then Exit Sub

'Sed
If UserList(UserIndex).Stats.MinAGU > 0 Then
    If UserList(UserIndex).Counters.AGUACounter < IntervaloSed Then
        UserList(UserIndex).Counters.AGUACounter = UserList(UserIndex).Counters.AGUACounter + 1
    Else
        UserList(UserIndex).Counters.AGUACounter = 0
        UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU - 10
        
        If UserList(UserIndex).Stats.MinAGU <= 0 Then
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).flags.Sed = 1
        End If
        
        fenviarAyS = True
    End If
End If

'hambre
If UserList(UserIndex).Stats.MinHAM > 0 Then
   If UserList(UserIndex).Counters.COMCounter < IntervaloHambre Then
        UserList(UserIndex).Counters.COMCounter = UserList(UserIndex).Counters.COMCounter + 1
   Else
        UserList(UserIndex).Counters.COMCounter = 0
        UserList(UserIndex).Stats.MinHAM = UserList(UserIndex).Stats.MinHAM - 10
        If UserList(UserIndex).Stats.MinHAM <= 0 Then
               UserList(UserIndex).Stats.MinHAM = 0
               UserList(UserIndex).flags.Hambre = 1
        End If
        fenviarAyS = True
    End If
End If

End Sub

Public Sub Sanar(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
'***************************************************************************
'Author: Leandro Mendoza(Mannakia)
'Desc:
'LastModify: 21/10/10
'***************************************************************************
If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).Trigger = 1 And _
   MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).Trigger = 2 And _
   MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).Trigger = 4 Then Exit Sub

Dim mashit As Integer

'Mannakia
If UserList(UserIndex).Invent.MagicIndex <> 0 Then
    If ObjData(UserList(UserIndex).Invent.MagicIndex).EfectoMagico = eMagicType.AceleraVida Then
        Intervalo = Intervalo - Porcentaje(Intervalo, 40)  ' nose algo razonable
    End If
End If
'Mannakia

'con el paso del tiempo va sanando....pero muy lentamente ;-)
If UserList(UserIndex).Stats.MinHP < UserList(UserIndex).Stats.MaxHP Then
    If UserList(UserIndex).Counters.HPCounter < Intervalo Then
        UserList(UserIndex).Counters.HPCounter = UserList(UserIndex).Counters.HPCounter + 1
    Else
        mashit = RandomNumber(2, Porcentaje(UserList(UserIndex).Stats.MaxSTA, 5))
        
        UserList(UserIndex).Counters.HPCounter = 0
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + mashit
        If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
        Call WriteConsoleMsg(1, UserIndex, "Has sanado.", FontTypeNames.FONTTYPE_INFO)
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
                UserList(i).Counters.salir = UserList(i).Counters.salir - 1
    
                If UserList(i).Counters.salir < 1 Then
                        
                        Call WriteMsg(i, 43)
                        Call WriteDisconnect(i)
                        Call FlushBuffer(i)
                        Call CloseSocket(i)
                Else
                
                    Call WriteMsg(i, 27, CStr(UserList(i).Counters.salir))

                End If
            End If
            
            If UserList(i).Counters.Habla > 10 Then
                UserList(i).Counters.Silenciado = 500
            End If
            UserList(i).Counters.Habla = 0
      
            If UserList(i).flags.TimesWalk >= 20 Then
                If Not UserList(i).flags.Paralizado Then
                    UserList(i).flags.Paralizado = 1
                    UserList(i).Counters.Paralisis = (IntervaloParalizado * 3) 'Mod Marius le mandamos *3 a los putos cheteros
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
    'Des Nod Kopfnickend Por las dudas, falta la aplicacion para levantarla y parece poco practico
    'If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

    'Chauuu
    'Unload frmMain

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
    
    Call LimpiarMundo
        
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
    Dim loopC As Long
    
    ' Free all NPC indexes
    For loopC = 1 To MAXNPCS
        Npclist(loopC).flags.NPCActive = False
    Next loopC
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



Sub ControlarPortalLum(ByVal UserIndex As Integer)
   
    If UserList(UserIndex).Counters.CreoTeleport = True Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageDestParticle(UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY))
        
        MapData(UserList(UserIndex).flags.DondeTiroMap, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY).TileExit.map = 0
        MapData(UserList(UserIndex).flags.DondeTiroMap, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY).TileExit.x = 0
        MapData(UserList(UserIndex).flags.DondeTiroMap, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY).TileExit.Y = 0
        UserList(UserIndex).flags.DondeTiroMap = 0
        UserList(UserIndex).flags.DondeTiroX = 0
        UserList(UserIndex).flags.DondeTiroY = 0
        UserList(UserIndex).flags.TiroPortalL = 0
        UserList(UserIndex).Counters.TimeTeleport = 0
        UserList(UserIndex).Counters.CreoTeleport = False
    End If

End Sub


Public Sub DoMetamorfosis(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer) 'Metamorfosis

Dim tUser As Integer
tUser = UserIndex

    If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
     
    UserList(tUser).Char.Head = Head
    UserList(tUser).Char.body = body
    UserList(tUser).Char.ShieldAnim = NingunEscudo
    UserList(tUser).Char.WeaponAnim = NingunArma
    UserList(tUser).Char.CascoAnim = NingunCasco
    
    UserList(UserIndex).flags.Metamorfosis = 1
     
    Call ChangeUserChar(UserIndex, UserList(tUser).Char.body, UserList(tUser).Char.Head, UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)

End Sub
 
Public Sub EfectoMetamorfosis(ByVal UserIndex As Integer) 'Metamorfosis
    If UserList(UserIndex).Counters.Metamorfosis < IntervaloInvisible * 2 Then
        UserList(UserIndex).Counters.Metamorfosis = UserList(UserIndex).Counters.Metamorfosis + 1
    Else
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
    End If
End Sub


Public Sub limpiaritemsmundo()
Static limpma As Integer

    If limpma > NumMaps Then
        limpma = 0
    End If
    
    limpma = limpma + 1
    
    Call limpiamundo(limpma)
    
End Sub
'Add Marius
'Esto bugea las puertas dobles, no esta bueno.
'Faltaria comprobar si el objeto es del mapa o agarrable
Public Sub limpiamundo(map As Integer)
Dim Xnn, Ynn As Byte

    If map < 0 And map > NumMaps Then Exit Sub
    
    'If MapInfo(limpma).Pk = True Then
        For Ynn = YMinMapSize To YMaxMapSize
            For Xnn = XMinMapSize To XMaxMapSize
                 If MapData(map, Xnn, Ynn).ObjInfo.ObjIndex > 0 And MapData(map, Xnn, Ynn).Blocked = 0 Then
                    Call EraseObj(10000, map, val(Xnn), val(Ynn))
                End If
            Next Xnn
        Next Ynn
    'End If
    
End Sub
'\Add
