Attribute VB_Name = "Base"
Option Explicit
Public DB_Conn As ADODB.Connection

Public DB_User As String    'The database username - (default "root")
Public DB_Pass As String    'Password to your database for the corresponding username
Public DB_Name As String    'Name of the table in the database (default "vbgore")
Public DB_Host As String    'IP of the database server - use localhost if hosted locally! Only host remotely for multiple servers!
Public DB_Port As Integer   'Port of the database (default "3306")


Public Sub CargarDB()
    Dim ErrorString As String
    Dim DB_RS As New ADODB.Recordset
On Error GoTo Errhandler
    
    'Add Marius Me canse de anda cambiando las constantes
    DB_User = Trim$(GetVar(IniPath & "Server.ini", "MYSQL", "User"))
    DB_Pass = Trim$(GetVar(IniPath & "Server.ini", "MYSQL", "Password"))
    DB_Name = Trim$(GetVar(IniPath & "Server.ini", "MYSQL", "Database"))
    DB_Host = Trim$(GetVar(IniPath & "Server.ini", "MYSQL", "Host"))
    DB_Port = val(GetVar(IniPath & "Server.ini", "MYSQL", "Port"))
    '\Add
    
    Set DB_Conn = New ADODB.Connection
    DB_Conn.ConnectionString = "DRIVER={MySQL ODBC 5.3 ANSI Driver};" & _
                "SERVER=" & DB_Host & ";" & _
                "DATABASE=" & DB_Name & ";" & _
                "PORT=" & DB_Port & ";" & _
                "UID=" & DB_User & ";" & _
                "PWD=" & DB_Pass & "; OPTION=3"
    
    DB_Conn.CursorLocation = adUseClient
    DB_Conn.Open
    
    
    'Ejecutamos estas sentencias para asegurarnos que las tablas esten!
    DB_RS.Open "SELECT * FROM charatrib WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM charbanco WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM charcorreo WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM charfaccion WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM charflags WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM charguild WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM charhechizos WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM charinit WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM charinvent WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM charmascotafami WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM charskills WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM charstats WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM cuentas WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM extras WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close

    On Error GoTo 0
    
    Exit Sub
    
Errhandler:
    'MsgBox err.description
    'End
        
    'Add Marius
    
    'Refresh the errors
    DB_Conn.Errors.Refresh
    
    'Get the error string if there is one
    If DB_Conn.Errors.count > 0 Then ErrorString = DB_Conn.Errors.Item(0).description

    'Check for known errors
    If InStr(1, ErrorString, "Access denied for user ") Then
        MsgBox "Mysql: Usuario o Contraseña incorrecta. " & err.description & " ErrorString: " & ErrorString
    ElseIf InStr(1, ErrorString, "Can't connect to MySQL server on ") Then
        MsgBox "Mysql: No se pudo conectar con el server, verifique el host y el puerto. " & err.description & " ErrorString: " & ErrorString
    ElseIf InStr(1, ErrorString, "Unknown database ") Then
        MsgBox "Mysql: La base de datos no existe. " & err.description & " ErrorString: " & ErrorString
    ElseIf InStr(1, ErrorString, "Data source name not found and no default driver specified") Then
        MsgBox "Mysql: La base de datos invalida o no existe. " & err.description & " ErrorString: " & ErrorString
    ElseIf InStr(1, ErrorString, "Table '") & InStr(1, ErrorString, "' doesn't exist") Then
        MsgBox "Mysql: Alguna tabla no existe o tiene errores. " & err.description & " ErrorString: " & ErrorString
    Else
        MsgBox "Mysql: Error conectando con la base de datos... " & err.description & " ErrorString: " & ErrorString
    End If
    
    
    End
    '\Add
End Sub

'Add Marius
Public Sub MySQL_Optimize()
    
    'Reordenamos los ids del correo
    'DB_Conn.Execute "SET @correocount = 0"
    'DB_Conn.Execute "UPDATE `charcorreo` SET `charcorreo`.`Idmsj` = @correocount:= @correocount + 1"
    'DB_Conn.Execute "ALTER TABLE `charcorreo` AUTO_INCREMENT = MAX(Idmsj) + 1"
    
    'Reordenamos los ids de los clanes
    'DB_Conn.Execute "SET @guildcount = 0"
    'DB_Conn.Execute "UPDATE `guildsinfo` SET `guildsinfo`.`GuildIndex` = @guildcount:= @guildcount + 1"
    'DB_Conn.Execute "ALTER TABLE `guildsinfo` AUTO_INCREMENT = @guildcount:= @guildcount + 1"
    
    'Optimize the database tables
    DB_Conn.Execute "OPTIMIZE TABLE charatrib, charbanco, charcorreo, charfaccion, charflags, charguild, charhechizos, charinit, charinvent, charmascotafami, charskills, charstats, cuentas, extras, guildsinfo, guildsolicitudes"

End Sub
'\Add

Public Sub insertarObjsDB()


End Sub


Public Sub insertarNpcsDB()

    Dim Index As Integer
    Dim Name As Integer
    Dim desc As Integer
    Dim Movement As Integer
    Dim AguaValida As Integer
    Dim TierraInvalida As Integer
    Dim faccion As Integer
    Dim AtacaDoble As Integer
    Dim NPCtype As Integer
    Dim body As Integer
    Dim Head As Integer
    Dim heading As Integer
    Dim Attackable As Integer
    Dim Comercia As Integer
    Dim Hostile As Integer
    Dim GiveEXP As Integer
    Dim Veneno As Integer
    Dim Domable As Integer
    Dim GiveGLD As Integer
    Dim PoderAtaque As Integer
    Dim PoderEvasion As Integer
    Dim InvReSpawn As Integer
    Dim MaxHP As Integer
    Dim MinHP As Integer
    Dim MaxHit As Integer
    Dim MinHit As Integer
    Dim def As Integer
    Dim defM As Integer
    Dim Alineacion As Integer
    Dim NroItems As Integer
    Dim LanzaSpells As Integer
    Dim sp1 As Integer
    Dim sp2 As Integer
    Dim sp3 As Integer
    Dim sp4 As Integer
    Dim sp5 As Integer
    Dim NroCriaturas As Integer
    Dim ci1 As Integer
    Dim ci2 As Integer
    Dim ci3 As Integer
    Dim ci4 As Integer
    Dim ci5 As Integer
    Dim cn1 As Integer
    Dim cn2 As Integer
    Dim cn3 As Integer
    Dim cn4 As Integer
    Dim cn5 As Integer
    Dim respawn As Integer
    Dim BackUp As Integer
    Dim originpos As Integer
    Dim AfectaParalisis As Integer
    Dim Snd1 As Integer
    Dim Snd2 As Integer
    Dim Snd3 As Integer
    Dim nroexp As Integer
    Dim exp1 As Integer
    Dim exp2 As Integer
    Dim exp3 As Integer
    Dim TipoItems As Integer
    
    Dim objetosParams As String
    Dim objetosValues As String
    
    Dim spellsParams As String
    Dim spellsValues As String
    
    Dim ciParams As String
    Dim ciValues As String
    
    Dim cnParams As String
    Dim cnValues As String
    
    Dim expParams As String
    Dim expValues As String
    
    Dim loopC As Long
    Dim ln As String
    Dim aux As String
    
    
    Dim i As Integer
    
    On Error GoTo errh:
    Dim result As Integer
    
    Dim consulta As String
    
    
    
    For i = i To 610
        
        objetosParams = ""
        objetosValues = ""
        spellsParams = ""
        spellsValues = ""
        ciParams = ""
        ciValues = ""
        cnParams = ""
        cnValues = ""
        expParams = ""
        expValues = ""
        consulta = ""
                    
        result = OpenNPC(i)
        
        
        If result <> 0 Then
            
            With Npclist(result)
                
        
        If i = 610 Then
        MsgBox .Char.body
        End If
                
                If .Invent.NroItems > 0 Then
                    For loopC = 1 To .Invent.NroItems
                        objetosValues = objetosValues + "'" & .Invent.Object(loopC).ObjIndex & "-" & .Invent.Object(loopC).Amount & "-" & .Invent.Object(loopC).Prob & "',"
                        objetosParams = objetosParams + "`obj" & loopC & "`,"
                    Next loopC
                End If
                
                For loopC = 1 To .flags.LanzaSpells
                    spellsValues = spellsValues + "" & .Spells(loopC) & ","
                    spellsParams = spellsParams + "`sp" & loopC & "`,"
                Next loopC
                
                If .NPCtype = eNPCType.Entrenador Then
                    For loopC = 1 To .NroCriaturas
                
                        ciValues = ciValues + "" & .Criaturas(loopC).NpcIndex & ","
                        ciParams = ciParams + "`ci" & loopC & "`,"
                
                        cnValues = cnValues + "'" + .Criaturas(loopC).NpcName + "',"
                        cnParams = cnParams + "`cn" & loopC & "`,"
                    Next loopC
                End If
                
                For loopC = 1 To .NroExpresiones
                    expValues = expValues + .Expresiones(loopC) & ","
                    expParams = expParams + "`exp" & loopC & "`,"
                Next loopC
           

               consulta = "INSERT INTO `inmortalao`.`npcs`(`id`,`index`,`name`,`desc`,`movement`,`aguavalida`,`tierrainvalida`,`faccion`,`atacadoble`,`npctype`,`body`,`head`,`heading`,`Attackable`,`comercia`,`hostile`,`giveexp`,`veneno`,`domable`,`givegld`,`poderataque`,`poderevasion`,`invrespawn`,`maxhp`,`minhp`,`maxhit`,`minhit`,`def`,`defm`,`alineacion`,`nroitems`," + objetosParams + "`lanzaspells`," + spellsParams + "`nrocriaturas`," + ciParams + cnParams + "`respawn`,`backup`,`originpos`,`afectaparalisis`,`snd1`,`snd2`,`snd3`,`nroexp`," + expParams + "`tipoitems`) values" & _
                "(null," & i & ",'" + .Name + "','" + .desc + "'," & .Movement & "," & .flags.AguaValida & "," & .flags.TierraInvalida & "," & _
                "" & .flags.faccion & "," & .flags.AtacaDoble & "," & .NPCtype & "," & .Char.body & "," & .Char.Head & "," & .Char.heading & "," & .Attackable & "," & _
                "" & .Comercia & "," & .Hostile & "," & .GiveEXP & "," & .Veneno & "," & .flags.Domable & "," & .GiveGLD & "," & .PoderAtaque & "," & _
                "" & .PoderEvasion & "," & .InvReSpawn & "," & .Stats.MaxHP & "," & .Stats.MinHP & "," & .Stats.MaxHit & "," & .Stats.MinHit & "," & .Stats.def & "," & _
                "" & .Stats.defM & "," & .Stats.Alineacion & "," & .Invent.NroItems & "," + objetosValues & _
                "" & .flags.LanzaSpells & "," + spellsValues & .NroCriaturas & "," & _
                "" + ciValues + cnValues & _
                "" & .flags.respawn & "," & .flags.BackUp & "," & .flags.RespawnOrigPos & "," & .flags.AfectaParalisis & "," & .flags.Snd1 & "," & .flags.Snd2 & "," & .flags.Snd3 & "," & _
                "" & .NroExpresiones & "," + expValues & .TipoItems & ");"
            
                DB_Conn.Execute (consulta)
            End With
        End If
    Next i
    
    Exit Sub
    
errh:
    MsgBox consulta
    MsgBox err.description
    
    
End Sub

Public Function ChangeBan(ByVal Name As String, ByVal Baneado As Byte) As Boolean
    Dim Orden As String
    Dim tUser As Integer
    Dim RS As New ADODB.Recordset
    
    ChangeBan = False
    Set RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'" & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Function
        
        Orden = "UPDATE `charflags` SET "
        Orden = Orden & "Ban=" & CStr(Baneado)
        Orden = Orden & " WHERE IndexPJ=" & RS!Indexpj & " LIMIT 1"
        
        Call DB_Conn.Execute(Orden)
    
        tUser = NameIndex(Name)
        If tUser > 0 Then
            Call CloseSocket(tUser)
        End If
    
    Set RS = Nothing
    
    ChangeBan = True
End Function
'Add Marius
Public Function Pejotas(ByVal Name As String) As String
    Dim Orden As String
    Dim tUser As Integer
    Dim RS As New ADODB.Recordset
    
    Pejotas = ""
    Set RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'" & " LIMIT 1")
        If RS.BOF Or RS.EOF Then
            Pejotas = "El personaje no existe!"
            Exit Function
        End If
        
        'RS!Indexpj
        Set RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE id=" & RS!id & " LIMIT 10")
    
        If Not (RS.BOF Or RS.EOF) Then
        
            Dim ii As Byte
            For ii = 1 To RS.RecordCount
                Pejotas = Pejotas & RS.Fields("Nombre")
                
                If RS.Fields("Online") Then Pejotas = Pejotas & " (Online)"
                
                Pejotas = Pejotas & " | "
                
                RS.MoveNext
            Next ii
            If LenB(Pejotas) <> 0 Then Pejotas = Left$(Pejotas, Len(Pejotas) - 3)
        Else
            Pejotas = "Error al mostrar personajes"
        End If
    
    
    Set RS = Nothing
End Function
'\Add


Public Sub CerrarDB()
On Error GoTo ErrHandle
    DB_Conn.Close
    Set DB_Conn = Nothing
    Exit Sub
ErrHandle:
    Call LogError("CerrarDB " & err.description & " " & err.Number)
    End
    
End Sub
Public Sub SaveUserSQL(UserIndex As Integer, Optional ByVal account As String = "", Optional insertPj As Boolean = False)
        
    Dim ipj As Integer
    
    If Not account = "" Then
        Call AddUserInAccount(UserList(UserIndex).IndexAccount)
    End If
    
    If insertPj Then
        ipj = Insert_New_Table(UserList(UserIndex).Name, UserList(UserIndex).IndexAccount)
    Else
        ipj = GetIndexPJ(UserList(UserIndex).Name)
    End If

    SaveUserFlags UserIndex, ipj
    SaveUserStats UserIndex, ipj
    SaveUserInit UserIndex, ipj
    SaveUserInv UserIndex, ipj
    SaveUserBank UserIndex, ipj
    SaveUserHechi UserIndex, ipj
    SaveUserAtrib UserIndex, ipj
    SaveUserSkill UserIndex, ipj
    SaveUserFami UserIndex, ipj
    SaveUserFaccion UserIndex, ipj
    
    Exit Sub

End Sub

Sub SaveUserHechi(ByVal UserIndex As Integer, ByVal ipj As Integer)
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(UserIndex)
    
    If Len(mUser.Name) = 0 Then Exit Sub
    
    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charhechizos` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
     
    str = "UPDATE `charhechizos` SET"
    str = str & " IndexPJ=" & ipj
    For i = 1 To MAXUSERHECHIZOS
        str = str & ",H" & i & "=" & mUser.Stats.UserHechizos(i)
    Next i
    str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
    Call DB_Conn.Execute(str)
    '************************************************************************
    
    Exit Sub
ErrHandle:
    Resume Next
End Sub


Sub SaveUserFami(ByVal UserIndex As Integer, ByVal ipj As Integer)
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(UserIndex)
    
    If Len(mUser.Name) = 0 Then Exit Sub
    
    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charmascotafami` WHERE IndexPJ=" & ipj & " LIMIT 1")
    If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
     
    str = "UPDATE `charmascotafami` SET"
    str = str & " IndexPJ=" & ipj
    str = str & ",Tiene=" & UserList(UserIndex).masc.TieneFamiliar
    str = str & ",Nombre='" & UserList(UserIndex).masc.Nombre & "'"
    str = str & ",Tipo=" & UserList(UserIndex).masc.tipo
    str = str & ",Level=" & UserList(UserIndex).masc.ELV
    str = str & ",ELU=" & UserList(UserIndex).masc.ELU
    str = str & ",Exp=" & UserList(UserIndex).masc.Exp
    str = str & ",MinHP=" & UserList(UserIndex).masc.MinHP
    str = str & ",MaxHP=" & UserList(UserIndex).masc.MaxHP
    str = str & ",MinHIT=" & UserList(UserIndex).masc.MinHit
    str = str & ",MaxHIT=" & UserList(UserIndex).masc.MaxHit
    str = str & ",MAS1=0"
    str = str & ",MAS2=0"
    str = str & ",MAS3=0"
    str = str & ",NroMascotas=0"
     
    str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
    Call DB_Conn.Execute(str)
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Sub SaveUserFlags(ByVal UserIndex As Integer, ByVal ipj As Integer)
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim i As Byte
    Dim str As String
    
    If Len(UserList(UserIndex).Name) = 0 Then Exit Sub
    
        '************************************************************************
    Set RS = New ADODB.Recordset
    
    Set RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    Set RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE IndexPJ=" & ipj & " LIMIT 1")
    str = "UPDATE `charflags` SET "
    'str = str & " IndexPJ=" & ipj
    str = str & "Nombre='" & UserList(UserIndex).Name & "'"
    str = str & ",Navegando=" & UserList(UserIndex).flags.Navegando
    str = str & ",Envenenado=" & UserList(UserIndex).flags.Envenenado
    'Mod Marius Ahora funcionan xD
    str = str & ",Pena=" & UserList(UserIndex).Counters.Pena
    str = str & ",Paralizado=0"
    '\Mod
    str = str & ",Desnudo=" & UserList(UserIndex).flags.Desnudo
    str = str & ",Sed=" & UserList(UserIndex).flags.Sed
    str = str & ",Hambre=" & UserList(UserIndex).flags.Hambre
    str = str & ",Escondido=" & UserList(UserIndex).flags.Escondido
    str = str & ",Muerto=" & UserList(UserIndex).flags.Muerto
    'Add Nod Kopfnickend
    str = str & ",`desc`='" & UserList(UserIndex).desc & "'"
    '\Add
    str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
    Call DB_Conn.Execute(str)
    
    'Grabamos Estados
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub


Sub SaveUserFaccion(ByVal UserIndex As Integer, ByVal ipj As Integer)
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(UserIndex)
    
    If Len(mUser.Name) = 0 Then Exit Sub
    
    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charfaccion` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    str = "UPDATE `charfaccion` SET"
    
    'Graba Faccion
    str = str & " IndexPJ=" & ipj
    str = str & ",EjercitoReal=" & mUser.faccion.ArmadaReal
    str = str & ",EjercitoCaos=" & mUser.faccion.FuerzasCaos
    str = str & ",EjercitoMili=" & mUser.faccion.Milicia
    str = str & ",Republicano=" & mUser.faccion.Republicano
    str = str & ",Ciudadano=" & mUser.faccion.Ciudadano
    str = str & ",Rango=" & mUser.faccion.Rango
    str = str & ",Renegado=" & mUser.faccion.Renegado
    str = str & ",CiudMatados=" & mUser.faccion.CiudadanosMatados
    str = str & ",ReneMatados=" & mUser.faccion.RenegadosMatados
    str = str & ",RepuMatados=" & mUser.faccion.RepublicanosMatados
    str = str & ",CaosMatados=" & mUser.faccion.CaosMatados
    str = str & ",ArmadaMatados=" & mUser.faccion.ArmadaMatados
    str = str & ",MiliMatados=" & mUser.faccion.MilicianosMatados
    str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
    Call DB_Conn.Execute(str)
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Sub SaveUserInit(ByVal UserIndex As Integer, ByVal ipj As Integer)
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    Dim time As Long
    
    mUser = UserList(UserIndex)
    
    If Len(mUser.Name) = 0 Then Exit Sub
    
    
    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    str = "UPDATE `charinit` SET"
    str = str & " IndexPJ=" & ipj
    str = str & ",Genero=" & mUser.Genero
    str = str & ",Raza=" & mUser.Raza
    str = str & ",Hogar=" & mUser.Hogar
    str = str & ",Clase=" & mUser.Clase
    str = str & ",Heading=" & mUser.Char.heading
    str = str & ",Head=" & mUser.OrigChar.Head
    str = str & ",Body=" & mUser.Char.body
    str = str & ",Arma=" & mUser.Char.WeaponAnim
    str = str & ",Escudo=" & mUser.Char.ShieldAnim
    str = str & ",Casco=" & mUser.Char.CascoAnim
    str = str & ",LastIP='" & mUser.ip & "'"
    str = str & ",Mapa=" & mUser.Pos.map
    str = str & ",X=" & mUser.Pos.x
    str = str & ",Y=" & mUser.Pos.Y
    str = str & ",PAREJA='" & mUser.flags.miPareja & "'"
    
    'Add Marius Anti Frags
    For i = 1 To MAX_BUFFER_KILLEDS
        time = mUser.Matados_timer(i) - GetTickCount
        If time > 0 And mUser.Matados(i) <> 0 Then
            str = str & ",pj" & i & "=" & mUser.Matados(i)
            str = str & ",time" & i & "=" & time
        Else
            str = str & ",pj" & i & "= 0"
            str = str & ",time" & i & "= 0"
        End If
    Next i
    '\Add
    
    str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
    Call DB_Conn.Execute(str)
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Sub SaveUserInv(ByVal UserIndex As Integer, ByVal ipj As Integer)
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(UserIndex)
    
    If Len(mUser.Name) = 0 Then Exit Sub
    
    'Fix Marius
    mUser.Invent.MonturaSlot = 0
    '\Fix
    
    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charinvent` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
     
    str = "UPDATE `charinvent` SET"
    str = str & " IndexPJ=" & ipj
    For i = 1 To MAX_INVENTORY_SLOTS
        str = str & ",OBJ" & i & "=" & mUser.Invent.Object(i).ObjIndex
        str = str & ",CANT" & i & "=" & mUser.Invent.Object(i).Amount
    Next i
    str = str & ",CASCOSLOT=" & mUser.Invent.CascoEqpSlot
    str = str & ",ARMORSLOT=" & mUser.Invent.ArmourEqpSlot
    str = str & ",SHIELDSLOT=" & mUser.Invent.EscudoEqpSlot
    str = str & ",WEAPONSLOT=" & mUser.Invent.WeaponEqpSlot
    str = str & ",ANILLOSLOT=" & mUser.Invent.AnilloEqpSlot
    str = str & ",MUNICIONSLOT=" & mUser.Invent.MunicionEqpSlot
    str = str & ",BARCOSLOT=" & mUser.Invent.BarcoSlot
    str = str & ",NUDISLOT=" & mUser.Invent.NudiEqpSlot
    str = str & ",MONTUSLOT=" & mUser.Invent.MonturaSlot
     
    str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
    Call DB_Conn.Execute(str)
    '************************************************************************
    Exit Sub
    
ErrHandle:
    Resume Next
    
End Sub
Sub SaveUserBank(ByVal UserIndex As Integer, ByVal ipj As Integer)
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(UserIndex)
    
    If Len(mUser.Name) = 0 Then Exit Sub
    
    
    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charbanco` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    str = "UPDATE `charbanco` SET"
    str = str & " IndexPJ=" & ipj
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        str = str & ",OBJ" & i & "=" & mUser.BancoInvent.Object(i).ObjIndex
        str = str & ",CANT" & i & "=" & mUser.BancoInvent.Object(i).Amount
    Next i
    str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
    Call DB_Conn.Execute(str)
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Sub SaveUserStats(ByVal UserIndex As Integer, ByVal ipj As Integer)
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(UserIndex)
    
    If Len(mUser.Name) = 0 Then Exit Sub
    
    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
     
    str = "UPDATE `charstats` SET"
    str = str & " IndexPJ=" & ipj
    str = str & ",GLD=" & mUser.Stats.GLD
    str = str & ",BANCO=" & mUser.Stats.Banco
    str = str & ",MaxHP=" & mUser.Stats.MaxHP
    str = str & ",MinHP=" & mUser.Stats.MinHP
    str = str & ",MaxMAN=" & mUser.Stats.MaxMAN
    str = str & ",MinMAN=" & mUser.Stats.MinMAN
    str = str & ",MinSTA=" & mUser.Stats.MinSTA
    str = str & ",MaxSTA=" & mUser.Stats.MaxSTA
    str = str & ",MaxHIT=" & mUser.Stats.MaxHit
    str = str & ",MinHIT=" & mUser.Stats.MinHit
    str = str & ",MinAGU=" & mUser.Stats.MinAGU
    str = str & ",MinHAM=" & mUser.Stats.MinHAM
    str = str & ",SkillPtsLibres=" & mUser.Stats.SkillPts
    str = str & ",VecesMurioUsuario=" & mUser.Stats.VecesMuertos
    str = str & ",Exp=" & mUser.Stats.Exp
    str = str & ",ELV=" & mUser.Stats.ELV
    str = str & ",NpcsMuertes=" & mUser.Stats.NPCsMuertos
    str = str & ",ELU=" & mUser.Stats.ELU
    str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
    Call DB_Conn.Execute(str)
    '************************************************************************
    
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Sub SaveUserAtrib(ByVal UserIndex As Integer, ByVal ipj As Integer)

On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(UserIndex)
    
    If Len(mUser.Name) = 0 Then Exit Sub
    
    
    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charatrib` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    str = "UPDATE `charatrib` SET"
    str = str & " IndexPJ=" & ipj
    For i = 1 To NUMATRIBUTOS
        str = str & ",AT" & i & "=" & mUser.Stats.UserAtributosBackUP(i)
    Next i
    str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
    Call DB_Conn.Execute(str)
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Sub SaveUserSkill(ByVal UserIndex As Integer, ByVal ipj As Integer)
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(UserIndex)
    
    If Len(mUser.Name) = 0 Then Exit Sub
    
        '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charskills` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    str = "UPDATE `charskills` SET"
    str = str & " IndexPJ=" & ipj
    
    For i = 1 To NUMSKILLS
        str = str & ",SK" & i & "=" & mUser.Stats.UserSkills(i)
    Next i
    
    str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
    Call DB_Conn.Execute(str)
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Function LoadUserSQL(UserIndex As Integer, ByVal Name As String) As Boolean
On Error GoTo Errhandler
Dim i As Integer
Dim RS As New ADODB.Recordset
Dim ipj  As Integer
Dim priv As Integer


With UserList(UserIndex)

    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT (IndexPJ) FROM `charflags` WHERE Nombre='" & Name & "' LIMIT 1")
        If RS.BOF Or RS.EOF Then
            LoadUserSQL = False
            Call LogError("Error en LoadUserSQL/charflags: Error al cargar charflags. Personaje: " & Name)
            Exit Function
        End If
    
        ipj = RS!Indexpj
    Set RS = Nothing
    '************************************************************************
    
    .Indexpj = ipj
    
    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE IndexPJ=" & ipj & " LIMIT 1")
    
    If RS.BOF Or RS.EOF Then
        Call LogError("Error en LoadUserSQL/charflags: Error al cargar charflags Personaje: " & Name)
        LoadUserSQL = False
    End If

    .flags.ban = RS!ban
    .flags.Navegando = RS!Navegando
    .flags.Envenenado = RS!Envenenado
    .Counters.Pena = RS!Pena
    .flags.Paralizado = 0
    
    .flags.Desnudo = RS!Desnudo
    .flags.Sed = RS!Sed
    .flags.Hambre = RS!Hambre
    .flags.Escondido = RS!Escondido
    .flags.Muerto = RS!Muerto
    'Add Nod Kopfnickend Nunca se guardaba en la DB y por ende nunca se cargaba
    .desc = RS!desc
    '\Add
    
    priv = RS!Privilegio
    'Vemos que clase de user es (se lo usa para setear los privilegios al loguear el PJ)
    .flags.AdminPerseguible = True
    
    If priv = 11 Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.VIP
    ElseIf priv = 10 Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
        .flags.AdminPerseguible = False
    ElseIf priv = 9 Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
        .flags.AdminPerseguible = False
    ElseIf priv = 8 Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Semi
    ElseIf priv = 7 Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Conse
    
    'Add Lideres faccionarios
    ElseIf priv = 3 Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.FaccCaos
    ElseIf priv = 2 Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.FaccRepu
    ElseIf priv = 1 Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.FaccImpe
    
    Else ' Es un pobre diablo
        .flags.Privilegios = .flags.Privilegios Or PlayerType.User
    End If
    
    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charfaccion` WHERE IndexPJ=" & ipj & " LIMIT 1")
    
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Call LogError("Error en LoadUserSQL/charfaccion: Error al cargar charfaccion. Personaje: " & Name)
        Exit Function
    End If
    
    ' Carga Faccion
    .faccion.ArmadaReal = RS!EjercitoReal
    .faccion.FuerzasCaos = RS!EjercitoCaos
    .faccion.Milicia = RS!EjercitoMili
    .faccion.Republicano = RS!Republicano
    .faccion.Ciudadano = RS!Ciudadano
    .faccion.Rango = RS!Rango
    .faccion.Renegado = RS!Renegado
    .faccion.CiudadanosMatados = RS!CiudMatados
    .faccion.RenegadosMatados = RS!ReneMatados
    .faccion.RepublicanosMatados = RS!RepuMatados
    .faccion.CaosMatados = RS!CaosMatados
    .faccion.ArmadaMatados = RS!ArmadaMatados
    .faccion.MilicianosMatados = RS!MiliMatados
    ' Fin Carga Faccion
    
    'Add Marius Un fix por un error mio xD
    If .faccion.Renegado = 1 Then
        Call ResetFacciones(UserIndex, False)
        .faccion.Renegado = 1
    End If
    
    Set RS = Nothing
    '************************************************************************

    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charatrib` WHERE IndexPJ=" & ipj & " LIMIT 1")
    
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Call LogError("Error en LoadUserSQL/charatrib: Error al cargar charatrib. Personaje: " & Name)
        Exit Function
    End If
    
    For i = 1 To NUMATRIBUTOS
        .Stats.UserAtributos(i) = RS.Fields("AT" & i)
        .Stats.UserAtributosBackUP(i) = .Stats.UserAtributos(i)
    Next i
    
    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then
            LoadUserSQL = False
            Call LogError("Error en LoadUserSQL/charguild: Error al cargar charguild. Personaje: " & Name)
            Exit Function
        End If
        
        UserList(UserIndex).GuildIndex = RS!GuildIndex
    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charskills` WHERE IndexPJ=" & ipj & " LIMIT 1")
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Call LogError("Error en LoadUserSQL/charskills: Error al cargar charskill. Personaje: " & Name)
        Exit Function
    End If
    For i = 1 To NUMSKILLS
        .Stats.UserSkills(i) = RS.Fields("SK" & i)
    Next i
    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charbanco` WHERE IndexPJ=" & ipj & " LIMIT 1")
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Call LogError("Error en LoadUserSQL/charbanco: Error al cargar charbanco. Personaje: " & Name)
        Exit Function
    End If
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        .BancoInvent.Object(i).ObjIndex = RS.Fields("OBJ" & i)
        .BancoInvent.Object(i).Amount = RS.Fields("CANT" & i)
    Next i
    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charinvent` WHERE IndexPJ=" & ipj & " LIMIT 1")
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Call LogError("Error en LoadUserSQL/charinvent: Error al cargar charinvent. Personaje: " & Name)
        Exit Function
    End If
    For i = 1 To MAX_INVENTORY_SLOTS
        .Invent.Object(i).ObjIndex = RS.Fields("OBJ" & i)
        .Invent.Object(i).Amount = RS.Fields("CANT" & i)
    Next i
    .Invent.CascoEqpSlot = RS!CASCOSLOT
    .Invent.ArmourEqpSlot = RS!ARMORSLOT
    .Invent.EscudoEqpSlot = RS!SHIELDSLOT
    .Invent.WeaponEqpSlot = RS!WEAPONSLOT
    .Invent.AnilloEqpSlot = RS!ANILLOSLOT
    .Invent.MunicionEqpSlot = RS!MUNICIONSLOT
    .Invent.BarcoSlot = RS!BarcoSlot
    .Invent.NudiEqpSlot = RS!NUDISLOT
    .Invent.MonturaSlot = RS!MONTUSLOT
    Set RS = Nothing
    '************************************************************************

    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charmascotafami` WHERE IndexPJ=" & ipj)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Call LogError("Error en LoadUserSQL/charmascotafami: Error al cargar charmascotafami. Personaje: " & Name)
        Exit Function
    End If

    .masc.TieneFamiliar = RS!Tiene
    .masc.Nombre = RS!Nombre
    .masc.tipo = RS!tipo
    .masc.ELV = RS!level
    .masc.ELU = RS!ELU
    .masc.Exp = RS!Exp
    .masc.MinHP = RS!MinHP
    .masc.MaxHP = RS!MaxHP
    .masc.MinHit = RS!MinHit
    .masc.MaxHit = RS!MaxHit
    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charhechizos` WHERE IndexPJ=" & ipj & " LIMIT 1")
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Call LogError("Error en LoadUserSQL/charhechizos: Error al cargar charhechizo. Personaje: " & Name)
        Exit Function
    End If
    For i = 1 To MAXUSERHECHIZOS
        .Stats.UserHechizos(i) = RS.Fields("H" & i)
    Next i
    
    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & ipj & " LIMIT 1")
    
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Call LogError("Error en LoadUserSQL/charstats: Error al cargar charstats. Personaje: " & Name)
        Exit Function
    End If
    
    .Stats.GLD = RS!GLD
    .Stats.Banco = RS!Banco
    .Stats.MaxHP = RS!MaxHP
    .Stats.MinHP = RS!MinHP
    .Stats.MinSTA = RS!MinSTA
    .Stats.MaxSTA = RS!MaxSTA
    .Stats.MaxMAN = RS!MaxMAN
    .Stats.MinMAN = RS!MinMAN
    .Stats.MaxHit = RS!MaxHit
    .Stats.MinHit = RS!MinHit
    .Stats.MinAGU = RS!MinAGU
    .Stats.MinHAM = RS!MinHAM
    .Stats.MaxAGU = 100
    .Stats.MaxHAM = 100
    .Stats.SkillPts = RS!SkillPtsLibres
    .Stats.VecesMuertos = RS!VecesMurioUsuario
    .Stats.Exp = RS!Exp
    .Stats.ELV = RS!ELV
    .Stats.NPCsMuertos = RS!NpcsMuertes
    .Stats.ELU = RS!ELU
    
    Set RS = Nothing
    
    If .Stats.MinAGU < 1 Then .flags.Sed = 1
    If .Stats.MinHAM < 1 Then .flags.Hambre = 1
    If .Stats.MinHP < 1 Then .flags.Muerto = 1
    
    '************************************************************************
    
    '************************************************************************
    Set RS = DB_Conn.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & ipj & " LIMIT 1")
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Call LogError("Error en LoadUserSQL/charinit: Error al cargar charinit. Personaje: " & Name)
        Exit Function
    End If
    
    .Genero = RS!Genero
    .Raza = RS!Raza
    .Hogar = RS!Hogar
    .Clase = RS!Clase
    .OrigChar.heading = RS!heading
    .OrigChar.Head = RS!Head
    .OrigChar.body = RS!body
    .OrigChar.WeaponAnim = RS!Arma
    .OrigChar.ShieldAnim = RS!Escudo
    .OrigChar.CascoAnim = RS!casco
    .ip = RS!LastIP
    .Pos.map = RS!mapa
    .Pos.x = RS!x
    .Pos.Y = RS!Y
    .flags.miPareja = RS!PAREJA
    
    'Add Marius agregamos esto para que funcione los casamientos. (xD)
    If Len(.flags.miPareja) > 0 Then
        .flags.toyCasado = 1
    End If
    '\Add
    
    'Add Marius Anti Frags
    For i = 1 To MAX_BUFFER_KILLEDS
        .Matados(i) = RS.Fields("pj" & i)
        .Matados_timer(i) = RS.Fields("time" & i)
        
        If .Matados(i) > 0 And .Matados_timer(i) > 0 Then
            .Matados_timer(i) = GetTickCount + .Matados_timer(i)
        End If
        
    Next i
    '\Add
    
    If .flags.Muerto = 0 Then
        .Char = .OrigChar
        Call VerObjetosEquipados(UserIndex)
    Else
        .Char.body = iCuerpoMuerto
        .Char.Head = iCabezaMuerto
        .Char.WeaponAnim = NingunArma
        .Char.ShieldAnim = NingunEscudo
        .Char.CascoAnim = NingunCasco
    End If
    
    Set RS = Nothing

    '************************************************************************
    
    
    If Len(.desc) > 100 Then .desc = Left$(.desc, 100)

    .Stats.MaxAGU = 100
    .Stats.MaxHAM = 100
    
    
    
    '************************************************************************
         
         
    Set RS = DB_Conn.Execute("SELECT * FROM `charcorreo` WHERE IndexPJ=" & ipj & " LIMIT 20")
    
    If Not (RS.BOF Or RS.EOF) Then
    
        Dim ii As Byte
        
        .cant_mensajes = RS.RecordCount

        For ii = 1 To .cant_mensajes
            .Correos(ii).idmsj = RS.Fields("Idmsj")
            .Correos(ii).Mensaje = RS.Fields("Mensaje")
            .Correos(ii).De = RS.Fields("De")
            .Correos(ii).Cantidad = RS.Fields("Cantidad")
            .Correos(ii).Item = RS.Fields("Item")
            RS.MoveNext
        Next ii
    Else
        .cant_mensajes = 0
    End If
    Set RS = Nothing
    '************************************************************************
    
    
    
    LoadUserSQL = True

End With

Exit Function

Errhandler:
    Call LogError("Error en LoadUserSQL. N:" & Name & " - " & err.Number & "-" & err.description)
    Set RS = Nothing
    
End Function
Function Add_GLD_Subast(ByRef Name As String, ByVal oro As Integer)
    Dim RS As New ADODB.Recordset
    Dim Indexpj As Integer
    Dim str As String

    Indexpj = GetIndexPJ(Name)
    
    If Indexpj <> 0 Then
        Set RS = DB_Conn.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
        If (RS.BOF Or RS.EOF) = False Then
            str = "UPDATE `charstats` SET"
            str = str & " IndexPJ=" & Indexpj
            str = str & ",GLD=" & RS!GLD + oro
            str = str & ",BANCO=" & RS!Banco
            str = str & ",MaxHP=" & RS!MaxHP
            str = str & ",MinHP=" & RS!MinHP
            str = str & ",MaxMAN=" & RS!MaxMAN
            str = str & ",MinMAN=" & RS!MinMAN
            str = str & ",MinSTA=" & RS!MinSTA
            str = str & ",MaxSTA=" & RS!MaxSTA
            str = str & ",MaxHIT=" & RS!MaxHit
            str = str & ",MinHIT=" & RS!MinHit
            str = str & ",MinAGU=" & RS!MinAGU
            str = str & ",MinHAM=" & RS!MinHAM
            str = str & ",SkillPtsLibres=" & RS!SkillPtsLibres
            str = str & ",VecesMurioUsuario=" & RS!VecesMurioUsuario
            str = str & ",Exp=" & RS!Exp
            str = str & ",ELV=" & RS!ELV
            str = str & ",NpcsMuertes=" & RS!NpcsMuertes
            str = str & ",ELU=" & RS!ELU
            str = str & " WHERE IndexPJ=" & Indexpj & " LIMIT 1"
            
            Call DB_Conn.Execute(str)
            Set RS = Nothing
        End If
    End If
End Function
Function Add_Bank_Gold(ByRef Name As String, ByVal oro As Long) As Boolean
On Error GoTo LocalErr
    Dim RS As New ADODB.Recordset
    Dim Indexpj As Integer
    Dim str As String

    Indexpj = GetIndexPJ(Name)
    
    If Indexpj <> 0 Then
        Set RS = DB_Conn.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
        If (RS.BOF Or RS.EOF) = False Then
            str = "UPDATE `charstats` SET"
            str = str & " IndexPJ=" & Indexpj
            str = str & ",GLD=" & RS!GLD
            str = str & ",BANCO=" & RS!Banco + oro
            str = str & ",MaxHP=" & RS!MaxHP
            str = str & ",MinHP=" & RS!MinHP
            str = str & ",MaxMAN=" & RS!MaxMAN
            str = str & ",MinMAN=" & RS!MinMAN
            str = str & ",MinSTA=" & RS!MinSTA
            str = str & ",MaxSTA=" & RS!MaxSTA
            str = str & ",MaxHIT=" & RS!MaxHit
            str = str & ",MinHIT=" & RS!MinHit
            str = str & ",MinAGU=" & RS!MinAGU
            str = str & ",MinHAM=" & RS!MinHAM
            str = str & ",SkillPtsLibres=" & RS!SkillPtsLibres
            str = str & ",VecesMurioUsuario=" & RS!VecesMurioUsuario
            str = str & ",Exp=" & RS!Exp
            str = str & ",ELV=" & RS!ELV
            str = str & ",NpcsMuertes=" & RS!NpcsMuertes
            str = str & ",ELU=" & RS!ELU
            str = str & " WHERE IndexPJ=" & Indexpj & " LIMIT 1"
            
            Call DB_Conn.Execute(str)
            Set RS = Nothing
            
            Add_Bank_Gold = True
            Exit Function
        End If
    End If
    Exit Function
    
LocalErr:
    Add_Bank_Gold = False
    Exit Function
End Function
Function Add_Item_Subast(ByRef Name As String, ByVal obji As Integer, ByVal Cant As Integer) As Boolean
    Dim RS As New ADODB.Recordset
    Dim Indexpj As Integer
    Dim i As Integer, j As Integer, t As Integer
    Dim b As Boolean
    Dim str As String
    
    Indexpj = GetIndexPJ(UCase$(Name))

    If Indexpj <> 0 Then
        Set RS = DB_Conn.Execute("SELECT * FROM `charbanco` WHERE IndexPJ=" & Indexpj & " LIMIT 1")

        If (RS.BOF Or RS.EOF) = False Then
            str = "UPDATE `charbanco` SET"
            str = str & " IndexPJ=" & Indexpj
            'Buscamos en el banco
            For i = 1 To MAX_BANCOINVENTORY_SLOTS
                j = RS.Fields("OBJ" & i)
                If j = 0 And b = False Then
                    str = str & ",OBJ" & i & "=" & obji
                    str = str & ",CANT" & i & "=" & Cant
                    b = True
                Else
                    str = str & ",OBJ" & i & "=" & j
                    str = str & ",CANT" & i & "=" & RS.Fields("CANT" & i)
                End If
            Next i
            str = str & " WHERE IndexPJ=" & Indexpj & " LIMIT 1"
            
            If b = True Then
                Call DB_Conn.Execute(str)
                Set RS = Nothing
                Add_Item_Subast = True
                Exit Function
            End If
        End If
        
        str = ""
        b = False
        j = 0
        
        Set RS = Nothing
        
        '************************************************************************
        Set RS = DB_Conn.Execute("SELECT * FROM `charinvent` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
        If (RS.BOF Or RS.EOF) = False Then
        
            str = "UPDATE `charinvent` SET"
            str = str & " IndexPJ=" & Indexpj
            
            For i = 1 To MAX_INVENTORY_SLOTS
                j = RS.Fields("OBJ" & i)
                If j = 0 And b = False Then
                    str = str & ",OBJ" & i & "=" & obji
                    str = str & ",CANT" & i & "=" & Cant
                    b = True
                Else
                    str = str & ",OBJ" & i & "=" & j
                    str = str & ",CANT" & i & "=" & RS.Fields("CANT" & i)
                End If
            Next i
            str = str & ",CASCOSLOT=" & RS!CASCOSLOT
            str = str & ",ARMORSLOT=" & RS!ARMORSLOT
            str = str & ",SHIELDSLOT=" & RS!SHIELDSLOT
            str = str & ",WEAPONSLOT=" & RS!WEAPONSLOT
            str = str & ",ANILLOSLOT=" & RS!ANILLOSLOT
            str = str & ",MUNICIONSLOT=" & RS!MUNICIONSLOT
            str = str & ",BARCOSLOT=" & RS!BarcoSlot
            str = str & ",NUDISLOT=" & RS!NUDISLOT
            str = str & ",MONTUSLOT=" & RS!MONTUSLOT
             
            str = str & " WHERE IndexPJ=" & Indexpj & " LIMIT 1"
            
            If b = True Then
                Call DB_Conn.Execute(str)
                Add_Item_Subast = True
            Else
                Add_Item_Subast = False
            End If
                
            Set RS = Nothing
            Exit Function
        End If
    End If
    
End Function
Public Function BANCheckDB(ByVal Name As String) As Boolean
    Dim RS As New ADODB.Recordset
    Dim Baneado As Byte
    
    Set RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "' LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Function
    
        Baneado = RS!ban
        BANCheckDB = (Baneado >= 1)
    Set RS = Nothing

End Function

Function ExistePersonaje(Name As String) As Boolean
    Dim RS As New ADODB.Recordset
    
    Set RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "' LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Function
    Set RS = Nothing
    
    ExistePersonaje = True
End Function
Function GetIndexPJ(Name As String) As Integer
On Error GoTo err
    Dim RS As New ADODB.Recordset
    Dim Indexpj As Long
    Dim Index As Integer
    
    'Add Marius Si ya esta cargado para que buscarlo otra vez xD
    'Abria que testearlo ahora no tengo tiempo
    'index = NameIndex(Name)
    'If index > 0 And UserList(index).Indexpj <> 0 Then
    '    Indexpj = UserList(index).Indexpj
    '    Exit Function
    'End If
    '\Add
    
    Set RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")
        If RS.BOF Or RS.EOF Then
            GoTo err
        Else
            GetIndexPJ = RS!Indexpj
        End If
    Set RS = Nothing
    Exit Function
    
err:
    Set RS = Nothing
    GetIndexPJ = 0
    Exit Function
End Function

Public Sub SendOnline()
    Call extra_set("online", str(NumUsers))
    frmMain.lbl_online.Caption = "Online: " & NumUsers
End Sub

Public Sub VerObjetosEquipados(UserIndex As Integer)

With UserList(UserIndex).Invent
    If .CascoEqpSlot Then
        .Object(.CascoEqpSlot).Equipped = 1
        .CascoEqpObjIndex = .Object(.CascoEqpSlot).ObjIndex
        UserList(UserIndex).Char.CascoAnim = ObjData(.CascoEqpObjIndex).CascoAnim
    Else
        UserList(UserIndex).Char.CascoAnim = NingunCasco
    End If
    
    If .BarcoSlot Then .BarcoObjIndex = .Object(.BarcoSlot).ObjIndex
    
    If .ArmourEqpSlot Then
        .Object(.ArmourEqpSlot).Equipped = 1
        .ArmourEqpObjIndex = .Object(.ArmourEqpSlot).ObjIndex
        UserList(UserIndex).Char.body = ObjData(.ArmourEqpObjIndex).Ropaje
    Else
        Call DarCuerpoDesnudo(UserIndex)
    End If
    
    If .WeaponEqpSlot > 0 Then
        .Object(.WeaponEqpSlot).Equipped = 1
        .WeaponEqpObjIndex = .Object(.WeaponEqpSlot).ObjIndex
        If .Object(.WeaponEqpSlot).ObjIndex > 0 Then UserList(UserIndex).Char.WeaponAnim = ObjData(.WeaponEqpObjIndex).WeaponAnim
    Else
        UserList(UserIndex).Char.WeaponAnim = NingunArma
    End If
    
    If .EscudoEqpSlot > 0 Then
        .Object(.EscudoEqpSlot).Equipped = 1
        .EscudoEqpObjIndex = .Object(.EscudoEqpSlot).ObjIndex
        UserList(UserIndex).Char.ShieldAnim = ObjData(.EscudoEqpObjIndex).ShieldAnim
    Else
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    End If

    If .MunicionEqpSlot Then
        .Object(.MunicionEqpSlot).Equipped = 1
        .MunicionEqpObjIndex = .Object(.MunicionEqpSlot).ObjIndex
    End If
    
    If .NudiEqpSlot > 0 Then
        .Object(.NudiEqpSlot).Equipped = 1
        .NudiEqpIndex = .Object(.NudiEqpSlot).ObjIndex
        UserList(UserIndex).Char.WeaponAnim = ObjData(.NudiEqpIndex).WeaponAnim
    End If
    
    'Fix Marius
    UserList(UserIndex).Invent.MonturaSlot = 0
    '\Fix
    
    If UserList(UserIndex).Invent.MonturaSlot <> 0 Then
        UserList(UserIndex).Invent.MonturaObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MonturaSlot).ObjIndex
        UserList(UserIndex).Char.body = ObjData(UserList(UserIndex).Invent.MonturaObjIndex).Ropaje
        UserList(UserIndex).flags.Montando = 1
        Call WriteEquitateToggle(UserIndex)
    Else
        UserList(UserIndex).flags.Montando = 0
    End If

End With

End Sub
Public Function Insert_New_Table(ByRef Name As String, ByRef id As Long) As Integer
On Error GoTo Erro
    Dim ipj As Integer
    
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    DB_Conn.Execute "INSERT INTO `charflags` (id,Nombre) VALUES (" & id & ",'" & Name & "')"
    
    Set RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE Nombre='" & Name & "'")
        ipj = RS!Indexpj
    Set RS = Nothing

    DB_Conn.Execute "INSERT INTO `charatrib` (IndexPJ) VALUES (" & ipj & ")"
    DB_Conn.Execute "INSERT INTO `charbanco` (IndexPJ) VALUES (" & ipj & ")"
    'DB_Conn.Execute "INSERT INTO `charcorreo` (IndexPJ) VALUES (" & iPJ & ")"
    DB_Conn.Execute "INSERT INTO `charfaccion` (IndexPJ) VALUES (" & ipj & ")"
    DB_Conn.Execute "INSERT INTO `charguild` (IndexPJ) VALUES (" & ipj & ")"
    DB_Conn.Execute "UPDATE `charguild` SET GuildIndex=0 WHERE IndexPJ=" & ipj & " LIMIT 1"
    DB_Conn.Execute "UPDATE `charguild` SET AspiranteA=0 WHERE IndexPJ=" & ipj & " LIMIT 1"
    
    DB_Conn.Execute "INSERT INTO `charhechizos` (IndexPJ) VALUES (" & ipj & ")"
    DB_Conn.Execute "INSERT INTO `charinit` (IndexPJ) VALUES (" & ipj & ")"
    DB_Conn.Execute "INSERT INTO `charinvent` (IndexPJ) VALUES (" & ipj & ")"
    DB_Conn.Execute "INSERT INTO `charmascotafami` (IndexPJ) VALUES (" & ipj & ")"
    DB_Conn.Execute "INSERT INTO `charskills` (IndexPJ) VALUES (" & ipj & ")"
    DB_Conn.Execute "INSERT INTO `charstats` (IndexPJ) VALUES (" & ipj & ")"
    
    Insert_New_Table = ipj
    Exit Function
Erro:
    LogError "Insert_New_Table " & Name & " " & err.Number & " " & err.description
End Function


Public Sub Quitarcorreosql(ByVal idmsj As Long)
    Dim RS As New ADODB.Recordset
    
    Set RS = DB_Conn.Execute("DELETE FROM `charcorreo` WHERE Idmsj=" & idmsj & " LIMIT 1")

    Set RS = Nothing

End Sub


Public Function Cantidadmensajes(Indexpj As Integer) As Byte

    Dim RS As New ADODB.Recordset

    Set RS = DB_Conn.Execute("Select IndexPJ FROM `charcorreo` WHERE IndexPJ=" & Indexpj)
    Cantidadmensajes = RS.RecordCount
    Set RS = Nothing

End Function

Public Function EnviarCorreoSql(ByVal ipj As Integer, ByVal loopC As Byte, ByVal para As Integer) As Long
    Dim RS As New ADODB.Recordset
    Dim str As String
    With UserList(para)
        str = "INSERT INTO `charcorreo` SET"
        str = str & " IndexPj=" & ipj
        str = str & ",Mensaje='" & .Correos(loopC).Mensaje & "'"
        str = str & ",De='" & .Correos(loopC).De & "'"
        str = str & ",Cantidad=" & .Correos(loopC).Cantidad
        str = str & ",Item=" & .Correos(loopC).Item
    
        Call DB_Conn.Execute(str)
        
        Set RS = Nothing
        
        'Add Marius Sin esto se puede duplicar objetos
        Dim id As Long
        Set RS = DB_Conn.Execute("SELECT last_insert_id() as id")
        If RS.BOF Or RS.EOF Then
            'GoTo err
        Else
            EnviarCorreoSql = RS!id
        End If
        '\Add
        
    End With
    Set RS = Nothing
End Function

Public Sub onpj(ByVal UserIndex As Integer)
    DB_Conn.Execute "UPDATE `charflags` SET `Online` = '1' WHERE `IndexPJ` = " & UserList(UserIndex).Indexpj & " LIMIT 1"
End Sub

Public Sub offpj(ByVal UserIndex As Integer)
    DB_Conn.Execute "UPDATE `charflags` SET `Online` = '0' WHERE `IndexPJ` = " & UserList(UserIndex).Indexpj & " LIMIT 1"
End Sub
Public Sub torneo_contador(ByVal UserIndex As Integer)
    DB_Conn.Execute "UPDATE charflags SET Torneos = Torneos + 1 WHERE IndexPJ = " & UserIndex & " LIMIT 1"
End Sub

Public Sub extra_set(ByVal Nombre As String, ByVal valor As String)
    DB_Conn.Execute "UPDATE `extras` SET `valor` = '" & valor & "' WHERE `nombre` = '" & Nombre & "' LIMIT 1"
End Sub

Function extra_get(Nombre As String) As String
On Error GoTo err
    Dim RS As New ADODB.Recordset
    Dim valor As String

    Set RS = DB_Conn.Execute("SELECT * FROM `extras` WHERE Nombre='" & Nombre & "' LIMIT 1")
        If RS.BOF Or RS.EOF Then
            GoTo err
        Else
            extra_get = RS!valor
        End If
    Set RS = Nothing
    Exit Function
    
err:
    Set RS = Nothing
    extra_get = "0"
    Exit Function
End Function


