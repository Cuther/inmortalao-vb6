Attribute VB_Name = "modGuilds"

Option Explicit

'guilds nueva version. Hecho por el oso, eliminando los problemas
'de sincronizacion con los datos en el HD... entre varios otros
'º¬

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DECLARACIOENS PUBLICAS CONCERNIENTES AL JUEGO
'Y CONFIGURACION DEL SISTEMA DE CLANES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private GUILDINFOFILE   As String
'archivo .\guilds\guildinfo.ini o similar

Private Const MAX_GUILDS As Integer = 1000
'cantidad maxima de guilds en el servidor

Public CANTIDADDECLANES As Integer
'cantidad actual de clanes en el servidor

Private guilds(1 To MAX_GUILDS) As clsClan
'array global de guilds, se indexa por userlist().guildindex (Hasta antes que Marius haga el sistemas con BD)

Public Const CANTIDADMAXIMACODEX As Byte = 8
'cantidad maxima de codecs que se pueden definir

Public Const MAXASPIRANTES As Byte = 10
'cantidad maxima de aspirantes que puede tener un clan acumulados a la vez

Private Const MAXANTIFACCION As Byte = 5
'puntos maximos de antifaccion que un clan tolera antes de ser cambiada su alineacion

'Gemas para fundar clan
Private Const GEMA_LUNAR As Integer = 406
Private Const GEMA_AZUL As Integer = 407
Private Const GEMA_NARANJA As Integer = 408
Private Const GEMA_ROJA As Integer = 411
Private Const GEMA_GRIS As Integer = 409

Public Enum ALINEACION_GUILD
    ALINEACION_REPUBLICANO = 1
    ALINEACION_IMPERIAL = 2
    ALINEACION_CAOTICO = 3
    ALINEACION_RENEGADO = 4
End Enum
'alineaciones permitidas

Public Enum SONIDOS_GUILD
    SND_CREACIONCLAN = 44
    SND_ACEPTADOCLAN = 43
    SND_DECLAREWAR = 45
End Enum
'numero de .wav del cliente



'estado entre clanes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub LoadGuildsDB()
Dim TempStr     As String
Dim IndexGuild As Integer
Dim Alin        As ALINEACION_GUILD
     
    'Add Marius Sistema de clanes con Base de Datos
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Set RS = DB_Conn.Execute("SELECT * FROM `guildsinfo` LIMIT " & MAX_GUILDS)
    Dim i As Byte
    
    If Not (RS.BOF Or RS.EOF) Then
    
        CANTIDADDECLANES = RS.RecordCount
        
        For i = 1 To CANTIDADDECLANES
            Set guilds(i) = New clsClan
            TempStr = RS.Fields("GuildName")
            Alin = RS.Fields("Alineacion")
            IndexGuild = RS.Fields("GuildIndex")
            
            Debug.Print "ClanName> " & TempStr & " > " & Alin & " > " & IndexGuild & " i = " & i
            
            Call guilds(i).Inicializar(TempStr, i, Alin, IndexGuild)
            RS.MoveNext
        Next i
    Else
        CANTIDADDECLANES = 0
    End If
    Set RS = Nothing
    '\Add
    
    
    
    
    
End Sub

Public Function m_ConectarMiembroAClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As Boolean
Dim NuevoL  As Boolean
Dim NuevaA  As Boolean
Dim News    As String
    
    If GuildIndex > CANTIDADDECLANES Or GuildIndex <= 0 Then Exit Function 'x las dudas...
    
    If m_EstadoPermiteEntrar(UserIndex, GuildIndex) Then
        Call guilds(GuildIndex).ConectarMiembro(UserIndex)
        UserList(UserIndex).GuildIndex = GuildIndex
        m_ConectarMiembroAClan = True
        
    Else
        m_ConectarMiembroAClan = m_ValidarPermanencia(UserIndex, True, NuevaA, NuevoL)
        If NuevoL Then News = "El clan tiene nuevo líder."
        If NuevaA Then News = News & "El clan tiene nueva alineación."
    End If

End Function


Public Function m_ValidarPermanencia(ByVal UserIndex As Integer, ByVal SumaAntifaccion As Boolean, ByRef CambioAlineacion As Boolean, ByRef CambioLider As Boolean) As Boolean
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 25/03/2009
'25/03/2009: ZaMa - Desequipo los items faccionarios que tenga el funda al abandonar la faccion
'***************************************************

Dim GuildIndex  As Integer
Dim ML()        As String
Dim M           As String
Dim UI          As Integer
Dim Sale        As Boolean
Dim i           As Integer

    m_ValidarPermanencia = True
    GuildIndex = UserList(UserIndex).GuildIndex
    If GuildIndex > CANTIDADDECLANES And GuildIndex <= 0 Then Exit Function
    
    If Not m_EstadoPermiteEntrar(UserIndex, GuildIndex) Then

        m_ValidarPermanencia = False
        If SumaAntifaccion Then guilds(GuildIndex).PuntosAntifaccion = guilds(GuildIndex).PuntosAntifaccion + 1
        
        CambioAlineacion = (m_EsGuildFounder(UserList(UserIndex).Name, GuildIndex) Or guilds(GuildIndex).PuntosAntifaccion = MAXANTIFACCION)
        
        If CambioAlineacion Then
            'uso GetMemberList y no los iteradores pq voy a rajar gente y puedo altera
            CambioLider = False
            i = 1
            ML = guilds(GuildIndex).GetMemberList()
            M = ML(i)
            While LenB(M) <> 0
                'vamos a violar un poco de capas..
                UI = NameIndex(M)
                If UI > 0 Then
                    Sale = Not m_EstadoPermiteEntrar(UI, GuildIndex)
                Else
                    Sale = Not m_EstadoPermiteEntrarChar(M, GuildIndex)
                End If

                If Sale Then
                    If m_EsGuildFounder(M, GuildIndex) Then 'hay que sacarlo de las armadas
                     
                        If UI > 0 Then
                            If UserList(UI).faccion.ArmadaReal <> 0 Then
                                Call ExpulsarFaccionReal(UI)
                            ElseIf UserList(UI).faccion.FuerzasCaos <> 0 Then
                                Call ExpulsarFaccionCaos(UI)
                            End If
                        End If
                        m_ValidarPermanencia = True
                    Else    'sale si no es guildfounder
                        If m_EsGuildLeader(M, GuildIndex) Then
                            'pierde el liderazgo
                            CambioLider = True
                            Call guilds(GuildIndex).SetLeader(guilds(GuildIndex).Fundador)
                        End If

                        Call m_EcharMiembroDeClan(-1, M)
                    End If
                End If
                i = i + 1
                M = ML(i)
            Wend
        Else
            'no se va el fundador, el peor caso es que se vaya el lider
            
            'If m_EsGuildLeader(UserList(UserIndex).Name, GuildIndex) Then
            '    Call LogClanes("Se transfiere el liderazgo de: " & Guilds(GuildIndex).GuildName & " a " & Guilds(GuildIndex).Fundador)
            '    Call Guilds(GuildIndex).SetLeader(Guilds(GuildIndex).Fundador)  'transferimos el lideraztgo
            'End If
            Call m_EcharMiembroDeClan(-1, UserList(UserIndex).Name)   'y lo echamos
        End If
    End If
End Function

Public Sub m_DesconectarMiembroDelClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)
    If UserList(UserIndex).GuildIndex > CANTIDADDECLANES Then Exit Sub
    Call guilds(GuildIndex).DesConectarMiembro(UserIndex)
End Sub

Private Function m_EsGuildLeader(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
    m_EsGuildLeader = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).GetLeader)))
End Function

Private Function m_EsGuildFounder(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
    m_EsGuildFounder = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).Fundador)))
End Function

Public Function m_EcharMiembroDeClan(ByVal Expulsador As Integer, ByVal Expulsado As String) As Integer
'UI echa a Expulsado del clan de Expulsado
Dim UserIndex   As Integer
Dim GI          As Integer
    
    m_EcharMiembroDeClan = 0

    UserIndex = NameIndex(Expulsado)
    If UserIndex > 0 Then
        'pj online
        GI = UserList(UserIndex).GuildIndex
        If GI > 0 Then
            If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
                If m_EsGuildLeader(Expulsado, GI) Then guilds(GI).SetLeader (guilds(GI).Fundador)
                Call guilds(GI).DesConectarMiembro(UserIndex)
                Call guilds(GI).ExpulsarMiembro(Expulsado)
                UserList(UserIndex).GuildIndex = 0
                Call RefreshCharStatus(UserIndex)
                m_EcharMiembroDeClan = GI
            Else
                m_EcharMiembroDeClan = 0
            End If
        Else
            m_EcharMiembroDeClan = 0
        End If
    Else
        'pj offline
        GI = GetGuildIndexFromChar(Expulsado)
        If GI > 0 Then
            If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
                If m_EsGuildLeader(Expulsado, GI) Then guilds(GI).SetLeader (guilds(GI).Fundador)
                Call guilds(GI).ExpulsarMiembro(Expulsado)
                m_EcharMiembroDeClan = GI
            Else
                m_EcharMiembroDeClan = 0
            End If
        Else
            m_EcharMiembroDeClan = 0
        End If
    End If

End Function

Public Sub ActualizarWebSite(ByVal UserIndex As Integer, ByRef Web As String)
Dim GI As Integer

    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then Exit Sub
    
    Call guilds(GI).SetURL(Web)
    
End Sub


Public Sub ChangeCodexAndDesc(ByRef desc As String, ByRef codex() As String, ByVal GuildIndex As Integer)
    Dim i As Long
    
    If GuildIndex < 1 Or GuildIndex > CANTIDADDECLANES Then Exit Sub
    
    With guilds(GuildIndex)
        Call .SetDesc(desc)
        
        Call .SetCodex(codex)
    End With
End Sub

Public Sub ActualizarNoticias(ByVal UserIndex As Integer, ByRef Datos As String)
Dim GI              As Integer

    GI = UserList(UserIndex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then Exit Sub
    
    Call guilds(GI).SetNoticia(Datos)
    
End Sub

Public Function CrearNuevoClan(ByVal FundadorIndex As Integer, ByRef desc As String, ByRef GuildName As String, ByRef URL As String, ByRef codex() As String, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean
Dim CantCodex       As Integer
Dim i               As Long
Dim DummyString     As String

    CrearNuevoClan = False
    If Not PuedeFundarUnClan(FundadorIndex, Alineacion, DummyString) Then
        refError = DummyString
        Exit Function
    End If

    If GuildName = vbNullString Or Not GuildNameValido(GuildName) Then
        refError = "Nombre de clan inválido."
        Exit Function
    End If
    
    If YaExiste(GuildName) Then
        refError = "Ya existe un clan con ese nombre."
        Exit Function
    End If

    CantCodex = UBound(codex()) + 1

    'tenemos todo para fundar ya
    If CANTIDADDECLANES < UBound(guilds) Then
        CANTIDADDECLANES = CANTIDADDECLANES + 1
        'ReDim Preserve Guilds(1 To CANTIDADDECLANES) As clsClan

        'constructor custom de la clase clan
        Set guilds(CANTIDADDECLANES) = New clsClan
        Call guilds(CANTIDADDECLANES).Inicializar(GuildName, CANTIDADDECLANES, Alineacion)
        
        'Damos de alta al clan como nuevo inicializando sus archivos
        Call guilds(CANTIDADDECLANES).InicializarNuevoClan(UserList(FundadorIndex).Name)
        
        'seteamos codex y descripcion
        Call guilds(CANTIDADDECLANES).SetCodex(codex)
        
        Call guilds(CANTIDADDECLANES).SetDesc(desc)
        Call guilds(CANTIDADDECLANES).SetLeader(UserList(FundadorIndex).Name)
        Call guilds(CANTIDADDECLANES).SetURL(URL)
        
        '"conectamos" al nuevo miembro a la lista de la clase
        Call guilds(CANTIDADDECLANES).AceptarNuevoMiembro(UserList(FundadorIndex).Name)
        Call guilds(CANTIDADDECLANES).ConectarMiembro(FundadorIndex)
        UserList(FundadorIndex).GuildIndex = CANTIDADDECLANES
        Call RefreshCharStatus(FundadorIndex)
        
        For i = 1 To MAX_INVENTORY_SLOTS
            
            If UserList(FundadorIndex).Invent.Object(i).ObjIndex = GEMA_LUNAR Then
                QuitarUserInvItem FundadorIndex, i, 1
                UpdateUserInv False, FundadorIndex, i
            ElseIf UserList(FundadorIndex).Invent.Object(i).ObjIndex = GEMA_AZUL And esCiuda(FundadorIndex) Then
                QuitarUserInvItem FundadorIndex, i, 1
                UpdateUserInv False, FundadorIndex, i
            ElseIf UserList(FundadorIndex).Invent.Object(i).ObjIndex = GEMA_NARANJA And esRepu(FundadorIndex) Then
                QuitarUserInvItem FundadorIndex, i, 1
                UpdateUserInv False, FundadorIndex, i
            ElseIf UserList(FundadorIndex).Invent.Object(i).ObjIndex = GEMA_ROJA And esCaos(FundadorIndex) Then
                QuitarUserInvItem FundadorIndex, i, 1
                UpdateUserInv False, FundadorIndex, i
            ElseIf UserList(FundadorIndex).Invent.Object(i).ObjIndex = GEMA_GRIS And esRene(FundadorIndex) Then
                QuitarUserInvItem FundadorIndex, i, 1
                UpdateUserInv False, FundadorIndex, i
            End If
            
        Next i

    Else
        refError = "No hay mas slots para fundar clanes. Consulte a un administrador."
        Exit Function
    End If
    
    CrearNuevoClan = True
End Function


Public Function m_PuedeSalirDeClan(ByRef Nombre As String, ByVal GuildIndex As Integer, ByVal QuienLoEchaUI As Integer) As Boolean
'sale solo si no es fundador del clan.

    m_PuedeSalirDeClan = False
    If GuildIndex = 0 Then Exit Function
    
    'esto es un parche, si viene en -1 es porque la invoca la rutina de expulsion automatica de clanes x antifacciones
    If QuienLoEchaUI = -1 Then
        m_PuedeSalirDeClan = True
        Exit Function
    End If

    'cuando UI no puede echar a nombre?
    'si no es gm Y no es lider del clan del pj Y no es el mismo que se va voluntariamente
    If Not EsDIOS(QuienLoEchaUI) Then
        If Not m_EsGuildLeader(UCase$(UserList(QuienLoEchaUI).Name), GuildIndex) Then
            If UCase$(UserList(QuienLoEchaUI).Name) <> UCase$(Nombre) Then      'si no sale voluntariamente...
                Exit Function
            End If
        End If
    End If

    m_PuedeSalirDeClan = UCase$(guilds(GuildIndex).Fundador) <> UCase$(Nombre)

End Function

Public Function PuedeFundarUnClan(ByVal UserIndex As Integer, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean

    PuedeFundarUnClan = False
    
    If UserList(UserIndex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, no puedes fundar otro"
        Exit Function
    End If
    
    PuedeFundarUnClan = True
    refError = ""
    
    If UserList(UserIndex).Stats.ELV < 40 Then
        refError = refError & "Ser nivel 40 o superior. "
        PuedeFundarUnClan = False
    End If
    
    If UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) < 100 Then
        refError = refError & refError & "Tener 100 en liderazgo. "
        PuedeFundarUnClan = False
    End If
    
    If esCiuda(UserIndex) And TieneObjetos(GEMA_AZUL, 1, UserIndex) = False Then
        refError = refError & "Poseer 1 Gema Azul. "
        PuedeFundarUnClan = False
    ElseIf esRepu(UserIndex) And TieneObjetos(GEMA_NARANJA, 1, UserIndex) = False Then
        refError = refError & "Poseer 1 Gema Naranja, "
        PuedeFundarUnClan = False
    ElseIf esCaos(UserIndex) And TieneObjetos(GEMA_ROJA, 1, UserIndex) = False Then
        refError = refError & "Poseer 1 Gema Roja. "
        PuedeFundarUnClan = False
    ElseIf esRene(UserIndex) And TieneObjetos(GEMA_GRIS, 1, UserIndex) = False Then
        refError = refError & "Poseer 1 Gema Gris. "
        PuedeFundarUnClan = False
    End If
    
    If TieneObjetos(GEMA_LUNAR, 1, UserIndex) = False Then
        refError = refError & "Poseer 1 Gema Lunar. "
        PuedeFundarUnClan = False
    End If
    
    If Not PuedeFundarUnClan Then
        refError = "Para fundar un clan te falta: " & refError
        Exit Function
    End If
    
    PuedeFundarUnClan = True
    
End Function

Private Function m_EstadoPermiteEntrarChar(ByRef Personaje As String, ByVal GuildIndex As Integer) As Boolean
    Dim RS As New ADODB.Recordset
    
    Debug.Print "SELECT `charflags`.Nombre,`charfaccion`.* FROM `charflags`,`charfaccion` WHERE `charflags`.Nombre = '" & Personaje & "' AND `charflags`.IndexPJ = `charfaccion`.IndexPJ LIMIT 1"
    Set RS = DB_Conn.Execute("SELECT `charflags`.Nombre,`charfaccion`.* FROM `charflags`,`charfaccion` WHERE `charflags`.Nombre = '" & Personaje & "' AND `charflags`.IndexPJ = `charfaccion`.IndexPJ LIMIT 1")
    If RS.BOF Or RS.EOF Then Exit Function
    
    m_EstadoPermiteEntrarChar = False
    
    Select Case guilds(GuildIndex).Alineacion
        Case ALINEACION_GUILD.ALINEACION_IMPERIAL
            If RS!EjercitoReal = 1 Or RS!Ciudadano = 1 Then
                m_EstadoPermiteEntrarChar = True
            End If
        Case ALINEACION_GUILD.ALINEACION_CAOTICO
            If RS!EjercitoCaos = 1 Then
                m_EstadoPermiteEntrarChar = True
            End If
        Case ALINEACION_GUILD.ALINEACION_REPUBLICANO
            If RS!EjercitoMili = 1 Or RS!Republicano = 1 Then
                m_EstadoPermiteEntrarChar = True
            End If
        Case ALINEACION_GUILD.ALINEACION_RENEGADO
            If RS!Renegado = 1 Then
                m_EstadoPermiteEntrarChar = True
            End If
        
        Case Else
            m_EstadoPermiteEntrarChar = False
    End Select
    
End Function

Private Function m_EstadoPermiteEntrar(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As Boolean
    
    If UCase$(UserList(UserIndex).Name) = UCase$(modGuilds.GuildLeader(GuildIndex)) Then
        m_EstadoPermiteEntrar = True
        Exit Function
    End If
    
    Select Case guilds(GuildIndex).Alineacion
        Case ALINEACION_GUILD.ALINEACION_IMPERIAL
            m_EstadoPermiteEntrar = esArmada(UserIndex) Or esCiuda(UserIndex)
        Case ALINEACION_GUILD.ALINEACION_CAOTICO
            m_EstadoPermiteEntrar = esCaos(UserIndex)
        Case ALINEACION_GUILD.ALINEACION_REPUBLICANO
            m_EstadoPermiteEntrar = esMili(UserIndex) Or esRepu(UserIndex)
        Case ALINEACION_GUILD.ALINEACION_RENEGADO
            m_EstadoPermiteEntrar = esRene(UserIndex)
        Case Else
            m_EstadoPermiteEntrar = True
    End Select
    
End Function


Public Function String2Alineacion(ByRef S As String) As ALINEACION_GUILD
    Select Case S
        Case "Coático"
            String2Alineacion = ALINEACION_CAOTICO
        Case "Armada Real"
            String2Alineacion = ALINEACION_IMPERIAL
        Case "Milicia"
            String2Alineacion = ALINEACION_REPUBLICANO
        Case "Renegado"
            String2Alineacion = ALINEACION_RENEGADO
    End Select
End Function

Public Function Alineacion2String(ByVal Alineacion As ALINEACION_GUILD) As String
    Select Case Alineacion
        Case ALINEACION_GUILD.ALINEACION_CAOTICO
            Alineacion2String = "Caótico"
        Case ALINEACION_GUILD.ALINEACION_IMPERIAL
            Alineacion2String = "Imperial"
        Case ALINEACION_GUILD.ALINEACION_REPUBLICANO
            Alineacion2String = "Republicano"
        Case ALINEACION_GUILD.ALINEACION_RENEGADO
            Alineacion2String = "Renegado"
    End Select
End Function

Private Function GuildNameValido(ByVal cad As String) As Boolean
Dim car     As Byte
Dim i       As Integer

    'old function by morgo
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
    
        If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
            GuildNameValido = False
            Exit Function
        End If
        
    Next i
    
    GuildNameValido = True

End Function

Private Function YaExiste(ByVal GuildName As String) As Boolean
Dim i   As Integer

    YaExiste = False
    GuildName = UCase$(GuildName)
    
    For i = 1 To CANTIDADDECLANES
        YaExiste = (UCase$(guilds(i).GuildName) = GuildName)
        If YaExiste Then Exit Function
    Next i

End Function


Private Function GetGuildIndexFromChar(ByRef PlayerName As String) As Integer
Dim RS      As New ADODB.Recordset
Dim Indexpj As Integer

    Indexpj = GetIndexPJ(PlayerName)

    If Indexpj <> 0 Then
        Set RS = DB_Conn.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Function
        
        GetGuildIndexFromChar = RS!GuildIndex
    End If
    
End Function

Public Function GuildIndex(ByRef GuildName As String) As Integer
'me da el indice del guildname
Dim i As Integer

    GuildIndex = 0
    GuildName = UCase$(GuildName)
    For i = 1 To CANTIDADDECLANES
        If UCase$(guilds(i).GuildName) = GuildName Then
            GuildIndex = i
            Exit Function
        End If
    Next i
End Function

Public Function GuildIndexArray(ByRef GuildIndexDb As Integer) As Integer
'Nos da el ID de la DB y le damos el ID del array de clanes
Dim i As Integer

    GuildIndexArray = 0
    For i = 1 To CANTIDADDECLANES
        Debug.Print ">>> " & guilds(i).GuildIndex
        If guilds(i).GuildIndex() = GuildIndexDb Then
            GuildIndexArray = i
            Exit Function
        End If
    Next i
    
End Function


Public Function m_ListaDeMiembrosOnline(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As String
Dim i As Integer
    
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
        While i > 0
            'No mostramos dioses y admins
            If i <> UserIndex And Not EsDIOS(UserIndex) Then _
                m_ListaDeMiembrosOnline = m_ListaDeMiembrosOnline & UserList(i).Name & ","
            i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
        Wend
    End If
    If Len(m_ListaDeMiembrosOnline) > 0 Then
        m_ListaDeMiembrosOnline = Left$(m_ListaDeMiembrosOnline, Len(m_ListaDeMiembrosOnline) - 1)
    End If
End Function

Public Function PrepareGuildsList() As String()
    Dim tStr() As String
    Dim i As Long
    
    If CANTIDADDECLANES = 0 Then
        ReDim tStr(0) As String
    Else
        ReDim tStr(CANTIDADDECLANES - 1) As String
        
        For i = 1 To CANTIDADDECLANES
            tStr(i - 1) = guilds(i).GuildName
        Next i
    End If
    
    PrepareGuildsList = tStr
End Function


Public Sub SendGuildDetails(ByVal UserIndex As Integer, ByRef GuildName As String)
    Dim codex  As String
    Dim GI      As Integer
    Dim i       As Long

    GI = GuildIndex(GuildName)
    If GI = 0 Then Exit Sub
    
    With guilds(GI)
        codex = .GetCodex()
        
        Call Protocol.WriteGuildDetails(UserIndex, GuildName, .Fundador, .GetFechaFundacion, .GetLeader, _
            .GetURL, .CantidadDeMiembros, Alineacion2String(.Alineacion), _
            .PuntosAntifaccion & "/" & CStr(MAXANTIFACCION), _
            codex, .GetDesc)
    End With
End Sub

Public Sub SendGuildLeaderInfo(ByVal UserIndex As Integer)
'***************************************************
'Autor: Mariano Barrou (El Oso)
'Last Modification: 12/10/06
'Las Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'***************************************************
    Dim GI      As Integer
    Dim guildList() As String
    Dim MemberList() As String
    Dim aspirantsList() As String
    Dim noticia As String

    With UserList(UserIndex)
        GI = .GuildIndex
        
        guildList = PrepareGuildsList()
        
        If GI <= 0 Or GI > CANTIDADDECLANES Then
            'Send the guild list instead
            Call Protocol.WriteGuildList(UserIndex, guildList)
            Exit Sub
        End If
        
        If Not m_EsGuildLeader(.Name, GI) Then
            'Send the guild list instead
            Call Protocol.WriteGuildList(UserIndex, guildList)
            Exit Sub
        End If
        
        MemberList = guilds(GI).GetMemberList()
        noticia = guilds(GI).GetNoticia()
        aspirantsList = guilds(GI).GetAspirantes()
        
        Call WriteGuildLeaderInfo(UserIndex, guildList, MemberList, noticia, aspirantsList)
    End With
End Sub


Public Function m_Iterador_ProximoUserIndex(ByVal GuildIndex As Integer) As Integer
    'itera sobre los onlinemembers
    m_Iterador_ProximoUserIndex = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        m_Iterador_ProximoUserIndex = guilds(GuildIndex).m_Iterador_ProximoUserIndex()
    End If
End Function

Public Function Iterador_ProximoGM(ByVal GuildIndex As Integer) As Integer
    'itera sobre los gms escuchando este clan
    Iterador_ProximoGM = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        Iterador_ProximoGM = guilds(GuildIndex).Iterador_ProximoGM()
    End If
End Function


Public Function GMEscuchaClan(ByVal UserIndex As Integer, ByVal GuildName As String) As Integer
Dim GI As Integer

    'listen to no guild at all
    If LenB(GuildName) = 0 And UserList(UserIndex).EscucheClan <> 0 Then
        'Quit listening to previous guild!!
        Call WriteConsoleMsg(1, UserIndex, "Dejas de escuchar a : " & guilds(UserList(UserIndex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
        guilds(UserList(UserIndex).EscucheClan).DesconectarGM (UserIndex)
        Exit Function
    End If
    
'devuelve el guildindex
    GI = GuildIndex(GuildName)
    If GI > 0 Then
        If UserList(UserIndex).EscucheClan <> 0 Then
            If UserList(UserIndex).EscucheClan = GI Then
                'Already listening to them...
                Call WriteConsoleMsg(1, UserIndex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
                GMEscuchaClan = GI
                Exit Function
            Else
                'Quit listening to previous guild!!
                Call WriteConsoleMsg(1, UserIndex, "Dejas de escuchar a : " & guilds(UserList(UserIndex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
                guilds(UserList(UserIndex).EscucheClan).DesconectarGM (UserIndex)
            End If
        End If
        
        Call guilds(GI).ConectarGM(UserIndex)
        Call WriteConsoleMsg(1, UserIndex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
        GMEscuchaClan = GI
        UserList(UserIndex).EscucheClan = GI
    Else
        Call WriteConsoleMsg(1, UserIndex, "Error, el clan no existe", FontTypeNames.FONTTYPE_GUILD)
        GMEscuchaClan = 0
    End If
    
End Function

Public Sub GMDejaDeEscucharClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)
'el index lo tengo que tener de cuando me puse a escuchar
    UserList(UserIndex).EscucheClan = 0
    Call guilds(GuildIndex).DesconectarGM(UserIndex)
End Sub

Public Sub a_RechazarAspiranteChar(ByRef aspirante As String, ByVal guild As Integer, ByRef Detalles As String)
    If InStrB(aspirante, "\") <> 0 Then
        aspirante = Replace(aspirante, "\", "")
    End If
    If InStrB(aspirante, "/") <> 0 Then
        aspirante = Replace(aspirante, "/", "")
    End If
    If InStrB(aspirante, ".") <> 0 Then
        aspirante = Replace(aspirante, ".", "")
    End If
    Call guilds(guild).InformarRechazoEnChar(aspirante, Detalles)
End Sub

Public Function a_ObtenerRechazoDeChar(ByRef aspirante As String) As String
    Dim RS      As New ADODB.Recordset
    Dim Indexpj As Integer
    
    Indexpj = GetIndexPJ(aspirante)

    If Indexpj <> 0 Then
        Set RS = DB_Conn.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Function
    
        a_ObtenerRechazoDeChar = RS!MotivoRechazo

        Call DB_Conn.Execute("UPDATE `charguild` SET MotivoRechazo='' WHERE IndexPJ=" & Indexpj & " LIMIT 1")
   End If
End Function

Public Function a_RechazarAspirante(ByVal UserIndex As Integer, ByRef Nombre As String, ByRef refError As String) As Boolean
Dim GI              As Integer
Dim NroAspirante    As Integer

    a_RechazarAspirante = False
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No perteneces a ningún clan"
        Exit Function
    End If

    NroAspirante = guilds(GI).NumeroDeAspirante(Nombre)

    If NroAspirante = 0 Then
        refError = Nombre & " no es aspirante a tu clan"
        Exit Function
    End If

    Call guilds(GI).RetirarAspirante(Nombre, NroAspirante)
    refError = "Fue rechazada tu solicitud de ingreso a " & guilds(GI).GuildName
    a_RechazarAspirante = True

End Function

Public Function a_DetallesAspirante(ByVal UserIndex As Integer, ByRef Nombre As String) As String
Dim GI              As Integer
Dim NroAspirante    As Integer

    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        Exit Function
    End If
    
    NroAspirante = guilds(GI).NumeroDeAspirante(Nombre)
    If NroAspirante > 0 Then
        a_DetallesAspirante = guilds(GI).DetallesSolicitudAspirante(NroAspirante)
    End If
    
End Function

Public Sub SendDetallesPersonaje(ByVal UserIndex As Integer, ByVal Personaje As String)
    Dim GI          As Integer
    Dim NroAsp      As Integer
    Dim GuildName   As String
    Dim Miembro     As String
    Dim Pedidos     As String
    Dim GuildActual As Integer
    Dim list()      As String
    Dim i           As Long
    Dim RS      As New ADODB.Recordset
    Dim Indexpj As Integer
    Dim Raza As Byte, Clase As Byte, Genero As Byte
    Dim GLD As Long, Banco As Long, Nivel As Byte
    Dim iMatados As Integer, mMatados As Integer, cMatados As Integer
    Dim ciMatados As Integer, rMatados As Integer, reMatados As Integer
    Dim Real As Byte, Caos As Byte, Mili As Byte
    
    
    GI = UserList(UserIndex).GuildIndex
    
    Personaje = UCase$(Personaje)
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Call Protocol.WriteConsoleMsg(1, UserIndex, "No perteneces a ningún clan", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        Call Protocol.WriteConsoleMsg(1, UserIndex, "No eres el líder de tu clan", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If InStrB(Personaje, "\") <> 0 Then
        Personaje = Replace$(Personaje, "\", vbNullString)
    End If
    If InStrB(Personaje, "/") <> 0 Then
        Personaje = Replace$(Personaje, "/", vbNullString)
    End If
    If InStrB(Personaje, ".") <> 0 Then
        Personaje = Replace$(Personaje, ".", vbNullString)
    End If
    
    NroAsp = guilds(GI).NumeroDeAspirante(Personaje)
    
    If NroAsp = 0 Then
        list = guilds(GI).GetMemberList()
        
        For i = 0 To UBound(list())
            Debug.Print i & ">" & list(i) & " > " & UBound(list())
            If Personaje = list(i) Then Exit For
        Next i
        
        Debug.Print i & ">" & (i > UBound(list()))
        
        'If i > UBound(list()) Then
        '    Call Protocol.WriteConsoleMsg(1, UserIndex, "El personaje no es ni aspirante ni miembro del clan", FontTypeNames.FONTTYPE_INFO)
        '    Exit Sub
        'End If
    End If
    
    Indexpj = GetIndexPJ(Personaje)
    
    Set RS = DB_Conn.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
        
        GuildActual = RS!GuildIndex
        If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
            GuildName = "<" & guilds(GuildActual).GuildName & ">"
        Else
            GuildName = "Ninguno"
        End If
            
        'Get previous guilds
        Miembro = RS!Miembro
        If Len(Miembro) > 400 Then
            Miembro = ".." & Right$(Miembro, 400)
        End If
        
        Pedidos = RS!Pedidos
    Set RS = Nothing
    
    Set RS = DB_Conn.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
        Raza = RS!Raza
        Clase = RS!Clase
        Genero = RS!Genero
    Set RS = Nothing
    
    Set RS = DB_Conn.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
        Nivel = RS!ELV
        Banco = RS!Banco
        GLD = RS!GLD
    Set RS = Nothing

    Set RS = DB_Conn.Execute("SELECT * FROM `charfaccion` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
        iMatados = RS!ArmadaMatados
        mMatados = RS!MiliMatados
        cMatados = RS!CaosMatados
        ciMatados = RS!CiudMatados
        rMatados = RS!RepuMatados
        reMatados = RS!ReneMatados
        
        Real = RS!EjercitoReal
        Caos = RS!EjercitoCaos
        Mili = RS!EjercitoMili
    Set RS = Nothing
    
    Call Protocol.WriteCharacterInfo( _
        UserIndex, Personaje, Raza, Clase, Genero _
        , Nivel, GLD, Banco, _
        Pedidos, GuildName, Miembro, _
        iMatados, mMatados, cMatados, ciMatados, rMatados, reMatados, _
        Real, Mili, Caos)

End Sub

Public Function a_NuevoAspirante(ByVal UserIndex As Integer, ByRef clan As String, ByRef Solicitud As String, ByRef refError As String) As Boolean
Dim ViejoSolicitado     As String
Dim ViejoGuildINdex     As Integer
Dim ViejoNroAspirante   As Integer
Dim NuevoGuildIndex     As Integer

    a_NuevoAspirante = False

    If UserList(UserIndex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, debes salir del mismo antes de solicitar ingresar a otro"
        Exit Function
    End If
    
    If EsNewbie(UserIndex) Then
        refError = "Los newbies no tienen derecho a entrar a un clan."
        Exit Function
    End If

    NuevoGuildIndex = GuildIndex(clan)
    If NuevoGuildIndex = 0 Then
        refError = "Ese clan no existe! Avise a un administrador."
        Exit Function
    End If
    
    If Not m_EstadoPermiteEntrar(UserIndex, NuevoGuildIndex) Then
        refError = "Tu no puedes entrar a un clan de alineación " & Alineacion2String(guilds(NuevoGuildIndex).Alineacion)
        Exit Function
    End If

    If guilds(NuevoGuildIndex).CantidadAspirantes >= MAXASPIRANTES Then
        refError = "El clan tiene demasiados aspirantes. Contáctate con un miembro para que procese las solicitudes."
        Exit Function
    End If
        
    
    Dim Indexpj As Integer
    Indexpj = GetIndexPJ(UserList(UserIndex).Name)

    Dim RS      As New ADODB.Recordset
    Set RS = DB_Conn.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
    If RS.BOF Or RS.EOF Then Exit Function
        
    ViejoSolicitado = RS!AspiranteA

    If LenB(ViejoSolicitado) <> 0 Then
        'borramos la vieja solicitud
        ViejoGuildINdex = CInt(ViejoSolicitado)
        If ViejoGuildINdex <> 0 Then
            ViejoNroAspirante = guilds(ViejoGuildINdex).NumeroDeAspirante(UserList(UserIndex).Name)
            If ViejoNroAspirante > 0 Then
                Call guilds(ViejoGuildINdex).RetirarAspirante(UserList(UserIndex).Name, ViejoNroAspirante)
            End If
        Else
            'RefError = "Inconsistencia en los clanes, avise a un administrador"
            'Exit Function
        End If
    End If
    
    Call guilds(NuevoGuildIndex).NuevoAspirante(UserList(UserIndex).Name, Solicitud)
    a_NuevoAspirante = True
End Function

Public Function a_AceptarAspirante(ByVal UserIndex As Integer, ByRef aspirante As String, ByRef refError As String) As Boolean
Dim GI              As Integer
Dim NroAspirante    As Integer
Dim AspiranteUI     As Integer

    'un pj ingresa al clan :D

    a_AceptarAspirante = False
    
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No perteneces a ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    NroAspirante = guilds(GI).NumeroDeAspirante(aspirante)
    
    If NroAspirante = 0 Then
        refError = "El Pj no es aspirante al clan"
        Exit Function
    End If
    
    AspiranteUI = NameIndex(aspirante)
    If AspiranteUI > 0 Then
        'pj Online
        If Not m_EstadoPermiteEntrar(AspiranteUI, GI) Then
            refError = aspirante & " no puede entrar a un clan " & Alineacion2String(guilds(GI).Alineacion)
            Call guilds(GI).RetirarAspirante(aspirante, NroAspirante)
            Exit Function
        ElseIf Not UserList(AspiranteUI).GuildIndex = 0 Then
            refError = aspirante & " ya es parte de otro clan."
            Call guilds(GI).RetirarAspirante(aspirante, NroAspirante)
            Exit Function
        End If
    Else
        If Not m_EstadoPermiteEntrarChar(aspirante, GI) Then
            refError = aspirante & " no puede entrar a un clan " & Alineacion2String(guilds(GI).Alineacion)
            Call guilds(GI).RetirarAspirante(aspirante, NroAspirante)
            Exit Function
        ElseIf GetGuildIndexFromChar(aspirante) Then
            refError = aspirante & " ya es parte de otro clan."
            Call guilds(GI).RetirarAspirante(aspirante, NroAspirante)
            Exit Function
        End If
    End If
    'el pj es aspirante al clan y puede entrar
    
    Call guilds(GI).RetirarAspirante(aspirante, NroAspirante)
    Call guilds(GI).AceptarNuevoMiembro(aspirante)
    
    ' If player is online, update tag
    If AspiranteUI > 0 Then
        Call RefreshCharStatus(AspiranteUI)
    End If
    
    a_AceptarAspirante = True
End Function

Public Function GuildName(ByVal GuildIndex As Integer) As String
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
        Exit Function
    
    GuildName = guilds(GuildIndex).GuildName
End Function

Public Function GuildLeader(ByVal GuildIndex As Integer) As String
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
        Exit Function
    
    GuildLeader = guilds(GuildIndex).GetLeader
End Function

Public Function GuildAlignment(ByVal GuildIndex As Integer) As String
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
        Exit Function
    
    GuildAlignment = Alineacion2String(guilds(GuildIndex).Alineacion)
End Function

Public Function GuildFounder(ByVal GuildIndex As Integer) As String
'***************************************************
'Autor: ZaMa
'Returns the guild founder's name
'Last Modification: 25/03/2009
'***************************************************
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
        Exit Function
    
    GuildFounder = guilds(GuildIndex).Fundador
End Function
Public Function GuildDelete(ByVal GuildIndex As Integer) As Boolean
    
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
        Exit Function
        
    Dim ML() As String, uName As String
    Dim i As Long, UI As Integer, RS As ADODB.Recordset
    Dim ipj As Integer
    
    ML = guilds(GuildIndex).GetMemberList
    For i = 1 To UBound(ML)
        uName = ML(i)
        
        UI = NameIndex(uName)
        If UI > 0 Then 'Si esta conectado le informamos
            
            Call modGuilds.m_EcharMiembroDeClan(-1, UserList(UI).Name)
            UserList(UI).GuildIndex = 0
            WriteMsg UI, 20
            
        Else
            'Los offline no hace falta en .borrar se actualiza indistintamente.
        End If
    Next i
    
    'Listo, ahora un flags para que no aparezca en lista :P
    GuildDelete = guilds(GuildIndex).Borrar
    
End Function
