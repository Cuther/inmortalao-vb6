VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

''
' clase clan
'
' Es el "ADO" de los clanes. La interfaz entre el disco y
' el juego. Los datos no se guardan en memoria
' para evitar problemas de sincronizacion, y considerando
' que la performance de estas rutinas NO es critica.
' by el oso :p

Private p_GuildName                 As String
Private p_Alineacion                As ALINEACION_GUILD
Private p_OnlineMembers             As Collection   'Array de UserIndexes!
Private p_GMsOnline                 As Collection
Private p_IteradorRelaciones        As Integer
Private p_IteradorOnlineMembers     As Integer
Private p_IteradorPropuesta         As Integer
Private p_IteradorOnlineGMs         As Integer
Private p_GuildNumber               As Integer      'Numero de guild en el mundo

'Add Marius Index del Guild en la DB
Private p_IndexGuild As Integer

'Private GUILDINFOFILE               As String
'Private GUILDPATH                   As String       'aca pq me es mas comodo setearlo y pq en ningun disenio
'Private MEMBERSFILE                 As String       'decente la capa de arriba se entera donde estan
'Private SOLICITUDESFILE             As String       'los datos fisicamente

Private Const NEWSLENGTH = 1024
Private Const DESCLENGTH = 256
Private Const CODEXLENGTH = 256

Public Borrado As Boolean

Public Property Get GuildName() As String
    GuildName = p_GuildName
End Property

Public Property Get GuildIndex() As String
    GuildIndex = p_IndexGuild
End Property



'
'ALINEACION Y ANTIFACCION
'

Public Property Get Alineacion() As ALINEACION_GUILD
    Alineacion = p_Alineacion
End Property

Public Property Get PuntosAntifaccion() As Integer
    
    'Add Marius Sistema de clanes con base de datos
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    Set RS = DB_Conn.Execute("SELECT `Antifaccion` FROM  `guildsinfo` WHERE GuildIndex= " & p_IndexGuild & " LIMIT 1")
    If Not (RS.BOF Or RS.EOF) Then
        PuntosAntifaccion = RS!Antifaccion
    End If
    '\Add
    
End Property

Public Property Let PuntosAntifaccion(ByVal p As Integer)
    
    DB_Conn.Execute "UPDATE `guildsinfo` SET  Antifaccion=" & CStr(p) & " WHERE GuildIndex= " & p_IndexGuild & " LIMIT 1"

End Property

Public Sub CambiarAlineacion(ByVal NuevaAlineacion As ALINEACION_GUILD)
    p_Alineacion = NuevaAlineacion
    
    DB_Conn.Execute "UPDATE `guildsinfo` SET  Alineacion=" & CStr(p_Alineacion) & " WHERE GuildIndex= " & p_IndexGuild & " LIMIT 1"
    
End Sub

'
'INICIALIZADORES
'

Private Sub Class_Initialize()
    'GUILDPATH = App.Path & "\GUILDS\"
    'GUILDINFOFILE = GUILDPATH & "guildsinfo.inf"
End Sub

Private Sub Class_Terminate()
    Set p_OnlineMembers = Nothing
    Set p_GMsOnline = Nothing
End Sub



Public Sub Inicializar(ByVal GuildName As String, ByVal GuildNumber As Integer, ByVal Alineacion As ALINEACION_GUILD, Optional ByVal IndexGuild As Integer = 0)
Dim i As Integer

    p_GuildName = GuildName
    p_GuildNumber = GuildNumber
    p_Alineacion = Alineacion
    
    p_IndexGuild = IndexGuild
    
    Set p_OnlineMembers = New Collection
    Set p_GMsOnline = New Collection
    
    p_IteradorOnlineMembers = 0
    p_IteradorPropuesta = 0
    p_IteradorOnlineGMs = 0
    p_IteradorRelaciones = 0

End Sub

Public Function Borrar() As Boolean  ':O

    Set p_OnlineMembers = Nothing
    Set p_GMsOnline = Nothing
    
    'Sacamos del clan a todos  los miembros
    Call DB_Conn.Execute("UPDATE `charguild` SET GuildIndex=0 WHERE GuildIndex = " & p_IndexGuild)
    
    'Borramos el clan
    DB_Conn.Execute "DELETE FROM `guildsinfo` WHERE GuildIndex= " & p_IndexGuild & " LIMIT 1"
    
    'Borramos las solicitudes
    DB_Conn.Execute "DELETE FROM `guildsolicitudes` WHERE GuildIndex= " & p_IndexGuild & " LIMIT 1"
    
    'Reordenamos los ids autoincrement y ponemos todos numeros consecutivos.
    'Call DB_Conn.Execute("SET @count = 0")
    'Call DB_Conn.Execute("UPDATE `guildsinfo` SET `guildsinfo`.`GuildIndex` = @count:= @count + 1")
    'DB_Conn.Execute "ALTER TABLE `guildsinfo` AUTO_INCREMENT = @count:= @count + 1"
    
    p_IndexGuild = 0
    p_GuildName = ""
    p_GuildNumber = 0
    p_Alineacion = 0
    p_IteradorOnlineMembers = 0
    p_IteradorPropuesta = 0
    p_IteradorOnlineGMs = 0
    p_IteradorRelaciones = 0
    
    'Recargamos los clanes.
    'Call LoadGuildsDB
    
    Borrar = True
End Function

''
' esta TIENE QUE LLAMARSE LUEGO DE INICIALIZAR()
'
' @param Fundador Nombre del fundador del clan
'
Public Sub InicializarNuevoClan(ByRef Fundador As String)
    'Add Marius Sistema de clanes con base de datos
    Dim ipj As Integer
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    DB_Conn.Execute "SELECT @newNo := MAX(`GuildIndex`) + 1 FROM `guildsinfo`"
    Debug.Print "INSERT INTO `guildsinfo` (GuildIndex,Founder,GuildName,Date,Antifaccion,Alineacion,Leader) VALUES (@newNo,'" & Fundador & "','" & p_GuildName & "','" & Date & "','0','" & p_Alineacion & "','" & Fundador & "')"
    DB_Conn.Execute "INSERT INTO `guildsinfo` (GuildIndex,Founder,GuildName,Date,Antifaccion,Alineacion,Leader) VALUES (@newNo,'" & Fundador & "','" & p_GuildName & "','" & Date & "','0','" & p_Alineacion & "','" & Fundador & "')"
    
    Set RS = DB_Conn.Execute("SELECT @newNo as id FROM `guildsinfo`")
    If Not (RS.BOF Or RS.EOF) Then
        p_IndexGuild = RS!id
        Debug.Print "NewId Clan -> " & p_IndexGuild
    End If
    '\Add

End Sub

'
'MEMBRESIAS
'

Public Property Get Fundador() As String
    
    'Add Marius Sistema de clanes con base de datos
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    Set RS = DB_Conn.Execute("SELECT `Founder` FROM  `guildsinfo` WHERE GuildIndex= " & p_IndexGuild & " LIMIT 1")
    If Not (RS.BOF Or RS.EOF) Then
        Fundador = RS!founder
    End If
    '\Add
End Property

'Public Property Get JugadoresOnline() As String
'Dim i As Integer
'    'leve violacion de capas x aqui, je
'    For i = 1 To p_OnlineMembers.Count
'        JugadoresOnline = UserList(p_OnlineMembers.Item(i)).Name & "," & JugadoresOnline
'    Next i
'End Property

Public Property Get CantidadDeMiembros() As Integer
Dim OldQ As String
    'OldQ = GetVar(MEMBERSFILE, "INIT", "NroMembers")
    'CantidadDeMiembros = IIf(IsNumeric(OldQ), CInt(OldQ), 0)
    
    
    'Add Marius Sistema de clanes con base de datos
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    Set RS = DB_Conn.Execute("SELECT COUNT(*) AS Cantidad FROM `charguild` Where GuildIndex = " & p_IndexGuild)
    If Not (RS.BOF Or RS.EOF) Then
        CantidadDeMiembros = RS!Cantidad
    End If
    '\Add
End Property

Public Sub SetLeader(ByRef leader As String)
    
    'Add Marius Sistema de clanes con base de datos
    DB_Conn.Execute "UPDATE `guildsinfo` SET Leader='" & leader & "' WHERE GuildIndex= " & p_IndexGuild & " LIMIT 1"
    '\Add
    
End Sub

Public Function GetLeader() As String
    
    'Add Marius Sistema de clanes con base de datos
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    Set RS = DB_Conn.Execute("SELECT `Leader` FROM `guildsinfo` WHERE GuildIndex= " & p_IndexGuild & " LIMIT 1")
    If Not (RS.BOF Or RS.EOF) Then
        GetLeader = RS!leader
    End If
    '\Add
End Function

Public Function GetMemberList() As String()
    Dim list() As String

    'Add Marius Sistema de clanes con Base de Datos
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Set RS = DB_Conn.Execute("SELECT `charguild`.* , `charflags`.Nombre FROM `charguild`,`charflags` WHERE `charguild`.GuildIndex =" & p_IndexGuild & " AND `charflags`.IndexPJ = `charguild`.IndexPJ")
    
    Dim i As Integer
    Dim count  As Integer
    
    If Not (RS.BOF Or RS.EOF) Then
        
        count = RS.RecordCount
        ReDim list(count - 1) As String
        
        For i = 1 To count
            list(i - 1) = UCase$(RS.Fields("Nombre"))
            RS.MoveNext
        Next i
        
    End If
    Set RS = Nothing
    '\Add
    
    
    
    
    GetMemberList = list
End Function

Public Sub ConectarMiembro(ByVal UserIndex As Integer)
    p_OnlineMembers.Add UserIndex
End Sub

Public Sub DesConectarMiembro(ByVal UserIndex As Integer)
Dim i   As Integer
    For i = 1 To p_OnlineMembers.count
        If p_OnlineMembers.Item(i) = UserIndex Then
            p_OnlineMembers.Remove i
            Exit Sub
        End If
    Next i
End Sub

Public Sub AceptarNuevoMiembro(ByRef Nombre As String)
    Dim ruta    As String
    Dim RS      As New ADODB.Recordset
    Dim Indexpj As Integer
    Dim str     As String
    
    Indexpj = GetIndexPJ(Nombre)

    If Indexpj <> 0 Then
        Set RS = DB_Conn.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
        
        str = "UPDATE `charguild` SET IndexPJ=" & Indexpj
        str = str & ",GuildIndex=" & p_GuildNumber
        str = str & ",AspiranteA=" & 0
        str = str & " WHERE IndexPJ=" & Indexpj & " LIMIT 1"
        Call DB_Conn.Execute(str)
        
    End If
End Sub

Public Sub ExpulsarMiembro(ByRef Nombre As String)
Dim OldQ        As Integer
Dim Temps       As String
Dim i           As Integer
Dim EsMiembro   As Boolean
Dim MiembroDe   As String
Dim RS      As New ADODB.Recordset
Dim Indexpj As Integer
Dim str     As String

    Indexpj = GetIndexPJ(Nombre)
        
    If Indexpj <> 0 Then
        Set RS = DB_Conn.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & Indexpj & " AND GuildIndex=" & p_IndexGuild & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
        
        MiembroDe = RS!Miembro
        If Not InStr(1, MiembroDe, p_GuildName, vbTextCompare) > 0 Then
            If LenB(MiembroDe) <> 0 Then
                MiembroDe = MiembroDe & ","
            End If
            MiembroDe = MiembroDe & p_GuildName
        End If

        str = "UPDATE `charguild` SET IndexPJ=" & Indexpj
        str = str & ",GuildIndex=0"
        str = str & ",Miembro='" & MiembroDe & "'"
        str = str & " WHERE IndexPJ=" & Indexpj & " LIMIT 1"
        Call DB_Conn.Execute(str)

    End If

End Sub

'
'ASPIRANTES
'

Public Function GetAspirantes() As String()

    'Add Marius Sistema de clanes con Base de Datos
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    Debug.Print "SELECT Nombre FROM `guildsolicitudes` WHERE GuildIndex= " & p_IndexGuild & " LIMIT 10"
    Set RS = DB_Conn.Execute("SELECT Nombre FROM `guildsolicitudes` WHERE GuildIndex= " & p_IndexGuild & " LIMIT 10")
    Dim Cant As Integer
    Dim i As Integer
        
        
    If Not (RS.BOF Or RS.EOF) Then
    
        Cant = RS.RecordCount
        ReDim list(Cant - 1) As String
        
        For i = 1 To Cant
            list(i - 1) = RS.Fields("Nombre")
            RS.MoveNext
        Next i
        
    End If
    Set RS = Nothing
    '\Add

    GetAspirantes = list
End Function

Public Function CantidadAspirantes() As Integer
    
    'Add Marius Sistema de clanes con base de datos
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    Set RS = DB_Conn.Execute("SELECT COUNT(*) AS cant FROM  `guildsolicitudes` Where GuildIndex = " & p_IndexGuild & " LIMIT 10")
    If Not (RS.BOF Or RS.EOF) Then
        CantidadAspirantes = RS!Cant
    Else
        CantidadAspirantes = 0
    End If
    '\Add

End Function

Public Function DetallesSolicitudAspirante(ByVal NroAspirante As Integer) As String
    
    'Add Marius Sistema de clanes con base de datos
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    Set RS = DB_Conn.Execute("SELECT `Solicitud` FROM  `guildsolicitudes` WHERE IndexSol = " & NroAspirante & " AND GuildIndex= " & p_IndexGuild & " LIMIT 1")
    If Not (RS.BOF Or RS.EOF) Then
        DetallesSolicitudAspirante = RS!Solicitud
    End If
    '\Add
    
End Function

Public Function NumeroDeAspirante(ByRef Nombre As String) As Integer
   
    'Add Marius Sistema de clanes con base de datos
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    Set RS = DB_Conn.Execute("SELECT IndexSol FROM `guildsolicitudes` WHERE GuildIndex= " & p_IndexGuild & " AND Nombre = '" & Nombre & "' LIMIT 1")
    If Not (RS.BOF Or RS.EOF) Then
        NumeroDeAspirante = RS!IndexSol
    End If
    '\Add
    
End Function
Function User_AspiranteA(ByRef Nombre As String, ByVal aspirante As Integer)
    Dim RS      As New ADODB.Recordset
    Dim Indexpj As Integer
    Dim str     As String
    
    Indexpj = GetIndexPJ(Nombre)

    If Indexpj <> 0 Then
        Set RS = DB_Conn.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Function
    
        str = "UPDATE `charguild` SET"
        str = str & " AspiranteA=" & aspirante
        str = str & " WHERE IndexPJ=" & Indexpj & " LIMIT 1"
        Debug.Print str
        Call DB_Conn.Execute(str)
    End If
End Function
Public Sub NuevoAspirante(ByRef Nombre As String, ByRef Peticion As String)
    'Add Marius Sistema de clanes con base de datos
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    Debug.Print "SELECT * FROM `guildsolicitudes` WHERE GuildIndex = " & p_IndexGuild & " and Nombre = '" & Nombre & "' LIMIT 1"
    Set RS = DB_Conn.Execute("SELECT * FROM `guildsolicitudes` WHERE GuildIndex = " & p_IndexGuild & " and Nombre = '" & Nombre & "' LIMIT 1")
    If Not (RS.BOF Or RS.EOF) Then
        Debug.Print "Clan> Ya mandó solicitud"
    Else
        Debug.Print "INSERT INTO `guildsolicitudes` (GuildIndex,IndexPJ,Nombre,solicitud) VALUES (" & p_IndexGuild & ",0,'" & Nombre & "','" & IIf(Trim$(Peticion) = vbNullString, "Peticion vacia", Peticion) & "')"
        DB_Conn.Execute "INSERT INTO `guildsolicitudes` (GuildIndex,IndexPJ,Nombre,solicitud) VALUES (" & p_IndexGuild & ",0,'" & Nombre & "','" & IIf(Trim$(Peticion) = vbNullString, "Peticion vacia", Peticion) & "')"
        
        User_AspiranteA Nombre, p_GuildNumber
    End If
    '\Add
    
End Sub

Public Sub RetirarAspirante(ByRef Nombre As String, ByRef NroAspirante As Integer)
    Dim OldQ    As String
    Dim OldQI   As String
    Dim Pedidos As String
    Dim i       As Integer
    Dim RS      As New ADODB.Recordset
    Dim Indexpj As Integer
    Dim str     As String
    
    Indexpj = GetIndexPJ(Nombre)
    
    Set RS = DB_Conn.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
    If RS.BOF Or RS.EOF Then Exit Sub
    
    User_AspiranteA Nombre, 0
    Pedidos = RS!Pedidos
    If Not InStr(1, Pedidos, p_GuildName, vbTextCompare) > 0 Then
        If LenB(Pedidos) <> 0 Then
            Pedidos = Pedidos & ","
        End If
        Pedidos = Pedidos & p_GuildName
        
        str = "UPDATE `charguild` SET"
        str = str & " Pedidos='" & Pedidos & "'"
        str = str & " WHERE IndexPJ=" & Indexpj & " LIMIT 1"
        Call DB_Conn.Execute(str)
    End If
    
    DB_Conn.Execute "DELETE FROM `guildsolicitudes` WHERE GuildIndex= " & p_IndexGuild & " AND Nombre = '" & Nombre & "' LIMIT 1"
    
    Set RS = Nothing
    
End Sub

Public Sub InformarRechazoEnChar(ByRef Nombre As String, ByRef Detalles As String)
    Dim RS      As New ADODB.Recordset
    Dim Indexpj As Integer
    Dim str     As String
    
    Indexpj = GetIndexPJ(Nombre)

    If Indexpj <> 0 Then
        'Set RS = DB_Conn.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
        'If RS.BOF Or RS.EOF Then Exit Sub
    
        str = "UPDATE `charguild` SET"
        str = str & " MotivoRechazo='" & Detalles & "'"
        str = str & " WHERE IndexPJ=" & Indexpj & " LIMIT 1"
        Call DB_Conn.Execute(str)
    End If
End Sub

'
'DEFINICION DEL CLAN (CODEX Y NOTICIAS)
'

Public Function GetFechaFundacion() As String
    'Add Marius Sistema de clanes con base de datos
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    Set RS = DB_Conn.Execute("SELECT `Date` FROM  `guildsinfo` WHERE GuildIndex= " & p_IndexGuild & " LIMIT 1")
    If RS.BOF Or RS.EOF Then
        GetFechaFundacion = ""
    Else
        GetFechaFundacion = RS!Date
    End If
    '\Add
End Function

Public Sub SetCodex(ByRef codex() As String)
    Dim cad As String
    Dim i As Integer
    cad = ""
    
    For i = 0 To UBound(codex())
        Call ReplaceInvalidChars(codex(i))
        codex(i) = Left$(codex(i), CODEXLENGTH)
        cad = cad & codex(i) & "|"
    Next i
    
    For i = i To CANTIDADMAXIMACODEX
        cad = cad & "|"
    Next i
    
    cad = Left$(cad, Len(cad) - 1)
    
    DB_Conn.Execute "UPDATE `guildsinfo` SET Codexs='" & cad & "' WHERE GuildIndex= " & p_IndexGuild & " LIMIT 1"
End Sub

Public Function GetCodex() As String
    
    'Add Marius Sistema de clanes con base de datos
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    Set RS = DB_Conn.Execute("SELECT `Codexs` FROM  `guildsinfo` WHERE GuildIndex= " & p_IndexGuild & " LIMIT 1")
    If Not (RS.BOF Or RS.EOF) Then
        GetCodex = RS!Codexs
        
        If Len(GetCodex) < 8 Then GetCodex = "|||||||" 'Por las dudas si hay algun error y esta vacio, le mandamos 8 ORs para que no se cage el cliente.
        
        GetCodex = Replace(GetCodex, "|", vbNullChar)
    End If
    '\Add
    
End Function


Public Sub SetURL(ByRef URL As String)
    'Add Marius Sistema de clanes con base de datos
    DB_Conn.Execute "UPDATE `guildsinfo` SET  URL='" & Left$(URL, 40) & "' WHERE GuildIndex= " & p_IndexGuild & " LIMIT 1"
    '\Add
End Sub

Public Function GetURL() As String
    'Add Marius Sistema de clanes con base de datos
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    Set RS = DB_Conn.Execute("SELECT `URL` FROM  `guildsinfo` WHERE GuildIndex= " & p_IndexGuild & " LIMIT 1")
    If RS.BOF Or RS.EOF Then
        GetURL = ""
    Else
        GetURL = RS!URL
    End If
    '\Add
    
End Function

Public Sub SetDesc(ByRef desc As String)
    Call ReplaceInvalidChars(desc)
    desc = Left$(desc, DESCLENGTH)
    
    'Add Marius Sistema de clanes con base de datos
    DB_Conn.Execute "UPDATE `guildsinfo` SET `Desc`='" & desc & "' WHERE GuildIndex='" & p_IndexGuild & "' LIMIT 1"
    '\Add
    
End Sub

Public Function GetDesc() As String
    'Add Marius Sistema de clanes con base de datos
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    Set RS = DB_Conn.Execute("SELECT `Desc` FROM  `guildsinfo` WHERE GuildIndex= " & p_IndexGuild & " LIMIT 1")
    If RS.BOF Or RS.EOF Then
        GetDesc = ""
    Else
        GetDesc = RS!desc
    End If
    '\Add

End Function

Public Sub SetNoticia(ByRef noticia As String)
    Call ReplaceInvalidChars(noticia)
    noticia = Left$(noticia, 250)
    
    'Add Marius Sistema de clanes con base de datos
    DB_Conn.Execute "UPDATE `guildsinfo` SET `Noticia`='" & noticia & "' WHERE GuildIndex='" & p_IndexGuild & "' LIMIT 1"
    '\Add
    
End Sub

Public Function GetNoticia() As String
    'Add Marius Sistema de clanes con base de datos
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    Set RS = DB_Conn.Execute("SELECT `Noticia` FROM  `guildsinfo` WHERE GuildIndex= " & p_IndexGuild & " LIMIT 1")
    If RS.BOF Or RS.EOF Then
        GetNoticia = ""
    Else
        GetNoticia = RS!noticia
    End If
    '\Add

End Function

Public Function m_Iterador_ProximoUserIndex() As Integer

    If p_IteradorOnlineMembers < p_OnlineMembers.count Then
        p_IteradorOnlineMembers = p_IteradorOnlineMembers + 1
        m_Iterador_ProximoUserIndex = p_OnlineMembers.Item(p_IteradorOnlineMembers)
    Else
        p_IteradorOnlineMembers = 0
        m_Iterador_ProximoUserIndex = 0
    End If
End Function

Public Function Iterador_ProximoGM() As Integer
    
    If p_IteradorOnlineGMs < p_GMsOnline.count Then
        p_IteradorOnlineGMs = p_IteradorOnlineGMs + 1
        Iterador_ProximoGM = p_GMsOnline.Item(p_IteradorOnlineGMs)
    Else
        p_IteradorOnlineGMs = 0
        Iterador_ProximoGM = 0
    End If
End Function

Public Sub ConectarGM(ByVal UserIndex As Integer)
    p_GMsOnline.Add UserIndex
End Sub

Public Sub DesconectarGM(ByVal UserIndex As Integer)
Dim i As Integer
    For i = 1 To p_GMsOnline.count
        If p_GMsOnline.Item(i) = UserIndex Then
            p_GMsOnline.Remove (i)
        End If
    Next i
End Sub



'
'VARIAS, EXTRAS Y DEMASES
'

Private Sub ReplaceInvalidChars(ByRef S As String)
    If InStrB(S, Chr$(13)) <> 0 Then
        S = Replace(S, Chr$(13), vbNullString)
    End If
    If InStrB(S, Chr$(10)) <> 0 Then
        S = Replace(S, Chr$(10), vbNullString)
    End If
    If InStrB(S, "¬") <> 0 Then
        S = Replace(S, "¬", vbNullString)   'morgo usaba esto como "separador"
    End If
End Sub



