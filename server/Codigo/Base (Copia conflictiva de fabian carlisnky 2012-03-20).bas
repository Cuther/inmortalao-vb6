Attribute VB_Name = "Base"
Option Explicit
Public Con As ADODB.Connection

Public Const mySQLPass As String = "123456pharry" 'PC GAME
Public Const mySQLUser As String = "JOENOD" ' PC GAME

'Public Const mySQLPass As String = "" 'LOCAL
'Public Const mySQLUser As String = "root" 'LOCAL

Public Const mySQLHost As String = "localhost"
Public Const mySQLBase As String = "inmortalao"

Public Sub CargarDB()
On Error GoTo Errhandler

    Set Con = New ADODB.Connection
    Con.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};" & _
                "SERVER=" & mySQLHost & "; " & _
                "DATABASE=" & mySQLBase & ";" & _
                "UID=" & mySQLUser & ";" & _
                "PWD=" & mySQLPass & "; OPTION=3"
    
    Con.CursorLocation = adUseClient
    Con.Open
    Exit Sub
    
Errhandler:
    MsgBox err.description
    End
End Sub

Public Function ChangeBan(ByVal Name As String, ByVal Baneado As Byte) As Boolean
    Dim Orden As String
    Dim RS As New ADODB.Recordset
    
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'" & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Function
        
        Orden = "UPDATE `charflags` SET"
        Orden = Orden & " IndexPJ=" & RS!Indexpj
        Orden = Orden & ",Nombre='" & UCase$(Name) & "'"
        Orden = Orden & ",Ban=" & Baneado
        Orden = Orden & " WHERE IndexPJ=" & RS!Indexpj & " LIMIT 1"

        Call Con.Execute(Orden)
    Set RS = Nothing

End Function


Public Sub CerrarDB()
On Error GoTo ErrHandle
    Con.Close
    Set Con = Nothing
    Exit Sub
ErrHandle:
    Call LogError("CerrarDB " & err.description & " " & err.Number)
    End
    
End Sub
Public Sub SaveUserSQL(userindex As Integer, Optional ByVal account As String = "", Optional insertPj As Boolean = False)
        
    Dim ipj As Integer
    
    If Not account = "" Then
        Call AddUserInAccount(UserList(userindex).IndexAccount)
    End If
    
    If insertPj Then
        ipj = Insert_New_Table(UserList(userindex).Name, UserList(userindex).IndexAccount)
    Else
        ipj = GetIndexPJ(UserList(userindex).Name)
    End If

    SaveUserFlags userindex, ipj
    SaveUserStats userindex, ipj
    SaveUserInit userindex, ipj
    SaveUserInv userindex, ipj
    SaveUserBank userindex, ipj
    SaveUserHechi userindex, ipj
    SaveUserAtrib userindex, ipj
    SaveUserSkill userindex, ipj
    SaveUserFami userindex, ipj
    SaveUserFaccion userindex, ipj
    
    Exit Sub

End Sub

Sub SaveUserHechi(ByVal userindex As Integer, ByVal ipj As Integer)
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userindex)
    
    If Len(mUser.Name) = 0 Then Exit Sub
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charhechizos` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
     
    str = "UPDATE `charhechizos` SET"
    str = str & " IndexPJ=" & ipj
    For i = 1 To MAXUSERHECHIZOS
        str = str & ",H" & i & "=" & mUser.Stats.UserHechizos(i)
    Next i
    str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
    Call Con.Execute(str)
    '************************************************************************
    
    Exit Sub
ErrHandle:
    Resume Next
End Sub


Sub SaveUserFami(ByVal userindex As Integer, ByVal ipj As Integer)
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userindex)
    
    If Len(mUser.Name) = 0 Then Exit Sub
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charmascotafami` WHERE IndexPJ=" & ipj & " LIMIT 1")
    If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
     
    str = "UPDATE `charmascotafami` SET"
    str = str & " IndexPJ=" & ipj
    str = str & ",Tiene=" & UserList(userindex).masc.TieneFamiliar
    str = str & ",Nombre='" & UserList(userindex).masc.Nombre & "'"
    str = str & ",Tipo=" & UserList(userindex).masc.Tipo
    str = str & ",Level=" & UserList(userindex).masc.ELV
    str = str & ",ELU=" & UserList(userindex).masc.ELU
    str = str & ",Exp=" & UserList(userindex).masc.Exp
    str = str & ",MinHP=" & UserList(userindex).masc.MinHP
    str = str & ",MaxHP=" & UserList(userindex).masc.MaxHP
    str = str & ",MinHIT=" & UserList(userindex).masc.MinHit
    str = str & ",MaxHIT=" & UserList(userindex).masc.MaxHit
    str = str & ",MAS1=0"
    str = str & ",MAS2=0"
    str = str & ",MAS3=0"
    str = str & ",NroMascotas=0"
     
    str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
    Call Con.Execute(str)
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Sub SaveUserFlags(ByVal userindex As Integer, ByVal ipj As Integer)
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim i As Byte
    Dim str As String
    
    If Len(UserList(userindex).Name) = 0 Then Exit Sub
    
        '************************************************************************
    Set RS = New ADODB.Recordset
    
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    Dim Pena As Integer
    
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE IndexPJ=" & ipj & " LIMIT 1")
    str = "UPDATE `charflags` SET"
    str = str & " IndexPJ=" & ipj
    str = str & ",Nombre='" & UserList(userindex).Name & "'"
    str = str & ",Ban=" & UserList(userindex).flags.Ban
    str = str & ",Navegando=" & UserList(userindex).flags.Navegando
    str = str & ",Envenenado=" & UserList(userindex).flags.Envenenado
    str = str & ",Pena=" & Pena * 60
    str = str & ",Paralizado=" & UserList(userindex).flags.Paralizado
    str = str & ",Desnudo=" & UserList(userindex).flags.Desnudo
    str = str & ",Sed=" & UserList(userindex).flags.Sed
    str = str & ",Hambre=" & UserList(userindex).flags.Hambre
    str = str & ",Escondido=" & UserList(userindex).flags.Escondido
    str = str & ",Muerto=" & UserList(userindex).flags.Muerto
    str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
    Call Con.Execute(str)
    'Grabamos Estados
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub


Sub SaveUserFaccion(ByVal userindex As Integer, ByVal ipj As Integer)
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userindex)
    
    If Len(mUser.Name) = 0 Then Exit Sub
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charfaccion` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    str = "UPDATE `charfaccion` SET"
    
    'Graba Faccion
    str = str & " IndexPJ=" & ipj
    str = str & ",EjercitoReal=" & mUser.Faccion.ArmadaReal
    str = str & ",EjercitoCaos=" & mUser.Faccion.FuerzasCaos
    str = str & ",EjercitoMili=" & mUser.Faccion.Milicia
    str = str & ",Republicano=" & mUser.Faccion.Republicano
    str = str & ",Ciudadano=" & mUser.Faccion.Ciudadano
    str = str & ",Rango=" & mUser.Faccion.Rango
    str = str & ",Renegado=" & mUser.Faccion.Renegado
    str = str & ",CiudMatados=" & mUser.Faccion.CiudadanosMatados
    str = str & ",ReneMatados=" & mUser.Faccion.RenegadosMatados
    str = str & ",RepuMatados=" & mUser.Faccion.RepublicanosMatados
    str = str & ",CaosMatados=" & mUser.Faccion.CaosMatados
    str = str & ",ArmadaMatados=" & mUser.Faccion.ArmadaMatados
    str = str & ",MiliMatados=" & mUser.Faccion.MilicianosMatados
    str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
    Call Con.Execute(str)
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Sub SaveUserInit(ByVal userindex As Integer, ByVal ipj As Integer)
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userindex)
    
    If Len(mUser.Name) = 0 Then Exit Sub
    
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & ipj & " LIMIT 1")
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
    str = str & ",Mapa=" & mUser.pos.map
    str = str & ",X=" & mUser.pos.x
    str = str & ",Y=" & mUser.pos.Y
    str = str & ",PAREJA='" & mUser.flags.miPareja & "'"
    str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
    Call Con.Execute(str)
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Sub SaveUserInv(ByVal userindex As Integer, ByVal ipj As Integer)
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userindex)
    
    If Len(mUser.Name) = 0 Then Exit Sub
    
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charinvent` WHERE IndexPJ=" & ipj & " LIMIT 1")
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
    Call Con.Execute(str)
    '************************************************************************
    Exit Sub
    
ErrHandle:
    Resume Next
    
End Sub
Sub SaveUserBank(ByVal userindex As Integer, ByVal ipj As Integer)
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userindex)
    
    If Len(mUser.Name) = 0 Then Exit Sub
    
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charbanco` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    str = "UPDATE `charbanco` SET"
    str = str & " IndexPJ=" & ipj
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        str = str & ",OBJ" & i & "=" & mUser.BancoInvent.Object(i).ObjIndex
        str = str & ",CANT" & i & "=" & mUser.BancoInvent.Object(i).Amount
    Next i
    str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
    Call Con.Execute(str)
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Sub SaveUserStats(ByVal userindex As Integer, ByVal ipj As Integer)
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userindex)
    
    If Len(mUser.Name) = 0 Then Exit Sub
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & ipj & " LIMIT 1")
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
    Call Con.Execute(str)
    '************************************************************************
    
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Sub SaveUserAtrib(ByVal userindex As Integer, ByVal ipj As Integer)

On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userindex)
    
    If Len(mUser.Name) = 0 Then Exit Sub
    
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charatrib` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    str = "UPDATE `charatrib` SET"
    str = str & " IndexPJ=" & ipj
    For i = 1 To NUMATRIBUTOS
        str = str & ",AT" & i & "=" & mUser.Stats.UserAtributosBackUP(i)
    Next i
    str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
    Call Con.Execute(str)
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Sub SaveUserSkill(ByVal userindex As Integer, ByVal ipj As Integer)
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userindex)
    
    If Len(mUser.Name) = 0 Then Exit Sub
    
        '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charskills` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    str = "UPDATE `charskills` SET"
    str = str & " IndexPJ=" & ipj
    
    For i = 1 To NUMSKILLS
        str = str & ",SK" & i & "=" & mUser.Stats.UserSkills(i)
    Next i
    
    str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
    Call Con.Execute(str)
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Function LoadUserSQL(userindex As Integer, ByVal Name As String) As Boolean
On Error GoTo Errhandler
Dim i As Integer
Dim RS As New ADODB.Recordset
Dim ipj  As Integer

With UserList(userindex)

    '************************************************************************
    Set RS = Con.Execute("SELECT (IndexPJ) FROM `charflags` WHERE Nombre='" & Name & "' LIMIT 1")
        If RS.BOF Or RS.EOF Then
            LoadUserSQL = False
            Exit Function
        End If
    
        ipj = RS!Indexpj
    Set RS = Nothing
    '************************************************************************
    
    .Indexpj = ipj
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE IndexPJ=" & ipj & " LIMIT 1")
    
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If

    .flags.Ban = RS!Ban
    .flags.Navegando = RS!Navegando
    .flags.Envenenado = RS!Envenenado
    .Counters.Pena = RS!Pena * 60
    .flags.Paralizado = RS!Paralizado
    .flags.Desnudo = RS!Desnudo
    .flags.Sed = RS!Sed
    .flags.Hambre = RS!Hambre
    .flags.Escondido = RS!Escondido
    .flags.Muerto = RS!Muerto

    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charfaccion` WHERE IndexPJ=" & ipj & " LIMIT 1")
    
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    
    ' Carga Faccion
    .Faccion.ArmadaReal = RS!EjercitoReal
    .Faccion.FuerzasCaos = RS!EjercitoCaos
    .Faccion.Milicia = RS!EjercitoMili
    .Faccion.Republicano = RS!Republicano
    .Faccion.Ciudadano = RS!Ciudadano
    .Faccion.Rango = RS!Rango
    .Faccion.Renegado = RS!Renegado
    .Faccion.CiudadanosMatados = RS!CiudMatados
    .Faccion.RenegadosMatados = RS!ReneMatados
    .Faccion.RepublicanosMatados = RS!RepuMatados
    .Faccion.CaosMatados = RS!CaosMatados
    .Faccion.ArmadaMatados = RS!ArmadaMatados
    .Faccion.MilicianosMatados = RS!MiliMatados
    ' Fin Carga Faccion
    
    Set RS = Nothing
    '************************************************************************

    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charatrib` WHERE IndexPJ=" & ipj & " LIMIT 1")
    
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    
    For i = 1 To NUMATRIBUTOS
        .Stats.UserAtributos(i) = RS.Fields("AT" & i)
        .Stats.UserAtributosBackUP(i) = .Stats.UserAtributos(i)
    Next i
    
    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then
            LoadUserSQL = False
            Exit Function
        End If
        
        UserList(userindex).GuildIndex = RS!GuildIndex
    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charskills` WHERE IndexPJ=" & ipj & " LIMIT 1")
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    For i = 1 To NUMSKILLS
        .Stats.UserSkills(i) = RS.Fields("SK" & i)
    Next i
    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charbanco` WHERE IndexPJ=" & ipj & " LIMIT 1")
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        .BancoInvent.Object(i).ObjIndex = RS.Fields("OBJ" & i)
        .BancoInvent.Object(i).Amount = RS.Fields("CANT" & i)
    Next i
    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charinvent` WHERE IndexPJ=" & ipj & " LIMIT 1")
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
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
    Set RS = Con.Execute("SELECT * FROM `charmascotafami` WHERE IndexPJ=" & ipj)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If

    .masc.TieneFamiliar = RS!Tiene
    .masc.Nombre = RS!Nombre
    .masc.Tipo = RS!Tipo
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
    Set RS = Con.Execute("SELECT * FROM `charhechizos` WHERE IndexPJ=" & ipj & " LIMIT 1")
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    For i = 1 To MAXUSERHECHIZOS
        .Stats.UserHechizos(i) = RS.Fields("H" & i)
    Next i
    
    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & ipj & " LIMIT 1")
    
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
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
    Set RS = Con.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & ipj & " LIMIT 1")
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    
    .Genero = RS!Genero
    .Raza = RS!Raza
    .Hogar = RS!Hogar
    .Clase = RS!Clase
    .Char.heading = RS!heading
    .OrigChar.Head = RS!Head
    .Char.body = RS!body
    .Char.WeaponAnim = RS!Arma
    .Char.ShieldAnim = RS!Escudo
    .Char.CascoAnim = RS!casco
    .ip = RS!LastIP
    .pos.map = RS!mapa
    .pos.x = RS!x
    .pos.Y = RS!Y
    .flags.miPareja = RS!PAREJA
    
    If .flags.Muerto = 0 Then
        .Char = .OrigChar
        Call VerObjetosEquipados(userindex)
    Else
        .Char.body = iCuerpoMuerto
        .Char.Head = iCabezaMuerto
        .Char.WeaponAnim = NingunArma
        .Char.ShieldAnim = NingunEscudo
        .Char.CascoAnim = NingunCasco
    End If
    
    Set RS = Nothing

    '************************************************************************
    
    '************************************************************************
         
         
    Set RS = Con.Execute("SELECT * FROM `charcorreo` WHERE IndexPJ=" & ipj)

    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
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

    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    
    
    
    LoadUserSQL = True
    
    If Len(.desc) >= 80 Then .desc = Left$(.desc, 80)

    .Stats.MaxAGU = 100
    .Stats.MaxHAM = 100

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
        Set RS = Con.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
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
            
            Call Con.Execute(str)
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
        Set RS = Con.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
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
            
            Call Con.Execute(str)
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
        Set RS = Con.Execute("SELECT * FROM `charbanco` WHERE IndexPJ=" & Indexpj & " LIMIT 1")

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
                Call Con.Execute(str)
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
        Set RS = Con.Execute("SELECT * FROM `charinvent` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
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
                Call Con.Execute(str)
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
    
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")
        If RS.BOF Or RS.EOF Then Exit Function
    
        Baneado = RS!Ban
        BANCheckDB = (Baneado = 1)
    Set RS = Nothing

End Function

Function ExistePersonaje(Name As String) As Boolean
    Dim RS As New ADODB.Recordset
    
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")
        If RS.BOF Or RS.EOF Then Exit Function
    Set RS = Nothing
    
    ExistePersonaje = True
End Function
Function GetIndexPJ(Name As String) As Integer
On Error GoTo err
    Dim RS As New ADODB.Recordset
    Dim Indexpj As Long

    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")
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
    Con.Execute "UPDATE `extras` SET `valor` = '" & NumUsers & "' WHERE `nombre` LIKE 'online' LIMIT 1"
    frmMain.lbl_online.Caption = "Online: " & NumUsers
End Sub

Public Sub VerObjetosEquipados(userindex As Integer)

With UserList(userindex).Invent
    If .CascoEqpSlot Then
        .Object(.CascoEqpSlot).Equipped = 1
        .CascoEqpObjIndex = .Object(.CascoEqpSlot).ObjIndex
        UserList(userindex).Char.CascoAnim = ObjData(.CascoEqpObjIndex).CascoAnim
    Else
        UserList(userindex).Char.CascoAnim = NingunCasco
    End If
    
    If .BarcoSlot Then .BarcoObjIndex = .Object(.BarcoSlot).ObjIndex
    
    If .ArmourEqpSlot Then
        .Object(.ArmourEqpSlot).Equipped = 1
        .ArmourEqpObjIndex = .Object(.ArmourEqpSlot).ObjIndex
        UserList(userindex).Char.body = ObjData(.ArmourEqpObjIndex).Ropaje
    Else
        Call DarCuerpoDesnudo(userindex)
    End If
    
    If .WeaponEqpSlot > 0 Then
        .Object(.WeaponEqpSlot).Equipped = 1
        .WeaponEqpObjIndex = .Object(.WeaponEqpSlot).ObjIndex
        If .Object(.WeaponEqpSlot).ObjIndex > 0 Then UserList(userindex).Char.WeaponAnim = ObjData(.WeaponEqpObjIndex).WeaponAnim
    Else
        UserList(userindex).Char.WeaponAnim = NingunArma
    End If
    
    If .EscudoEqpSlot > 0 Then
        .Object(.EscudoEqpSlot).Equipped = 1
        .EscudoEqpObjIndex = .Object(.EscudoEqpSlot).ObjIndex
        UserList(userindex).Char.ShieldAnim = ObjData(.EscudoEqpObjIndex).ShieldAnim
    Else
        UserList(userindex).Char.ShieldAnim = NingunEscudo
    End If

    If .MunicionEqpSlot Then
        .Object(.MunicionEqpSlot).Equipped = 1
        .MunicionEqpObjIndex = .Object(.MunicionEqpSlot).ObjIndex
    End If
    
    If .NudiEqpSlot > 0 Then
        .Object(.NudiEqpSlot).Equipped = 1
        .NudiEqpIndex = .Object(.NudiEqpSlot).ObjIndex
        UserList(userindex).Char.WeaponAnim = ObjData(.NudiEqpIndex).WeaponAnim
    End If
    
    If UserList(userindex).Invent.MonturaSlot <> 0 Then
        UserList(userindex).Invent.MonturaObjIndex = UserList(userindex).Invent.Object(UserList(userindex).Invent.MonturaSlot).ObjIndex
        UserList(userindex).Char.body = ObjData(UserList(userindex).Invent.MonturaObjIndex).Ropaje
        UserList(userindex).flags.Montando = 1
        Call WriteEquitateToggle(userindex)
    Else
        UserList(userindex).flags.Montando = 0
    End If

End With

End Sub
Public Function Insert_New_Table(ByRef Name As String, ByRef id As Long) As Integer
On Error GoTo Erro
    Dim ipj As Integer
    
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    Con.Execute "INSERT INTO `charflags` (id,Nombre) VALUES (" & id & ",'" & Name & "')"
    
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & Name & "'")
        ipj = RS!Indexpj
    Set RS = Nothing

    Con.Execute "INSERT INTO `charatrib` (IndexPJ) VALUES (" & ipj & ")"
    Con.Execute "INSERT INTO `charbanco` (IndexPJ) VALUES (" & ipj & ")"
    'Con.Execute "INSERT INTO `charcorreo` (IndexPJ) VALUES (" & iPJ & ")"
    Con.Execute "INSERT INTO `charfaccion` (IndexPJ) VALUES (" & ipj & ")"
    Con.Execute "INSERT INTO `charguild` (IndexPJ) VALUES (" & ipj & ")"
    Con.Execute "UPDATE `charguild` SET GuildIndex=0 WHERE IndexPJ=" & ipj & " LIMIT 1"
    
    Con.Execute "INSERT INTO `charhechizos` (IndexPJ) VALUES (" & ipj & ")"
    Con.Execute "INSERT INTO `charinit` (IndexPJ) VALUES (" & ipj & ")"
    Con.Execute "INSERT INTO `charinvent` (IndexPJ) VALUES (" & ipj & ")"
    Con.Execute "INSERT INTO `charmascotafami` (IndexPJ) VALUES (" & ipj & ")"
    Con.Execute "INSERT INTO `charskills` (IndexPJ) VALUES (" & ipj & ")"
    Con.Execute "INSERT INTO `charstats` (IndexPJ) VALUES (" & ipj & ")"
    
    Insert_New_Table = ipj
    Exit Function
Erro:
    LogError "Insert_New_Table " & Name & " " & err.Number & " " & err.description
End Function


Public Sub Quitarcorreosql(ByVal idmsj As Long)
    Dim RS As New ADODB.Recordset
    
    Set RS = Con.Execute("DELETE FROM `charcorreo` WHERE Idmsj=" & idmsj & " LIMIT 1")

    Set RS = Nothing

End Sub


Public Function Cantidadmensajes(Indexpj As Integer) As Byte

    Dim RS As New ADODB.Recordset

    Set RS = Con.Execute("Select (IndexPJ)  FROM `charcorreo` WHERE IndexPJ=" & Indexpj)
    Cantidadmensajes = RS.RecordCount
    Set RS = Nothing

End Function


Public Sub EnviarCorreoSql(ByVal ipj As Integer, ByVal LoopC As Byte, ByVal para As Integer)
    Dim RS As New ADODB.Recordset
    Dim str As String
    With UserList(para)
    
    

    str = "INSERT INTO `charcorreo` SET"
    str = str & " IndexPj=" & ipj
    str = str & ",Mensaje='" & .Correos(LoopC).Mensaje & "'"
    str = str & ",De='" & .Correos(LoopC).De & "'"
    str = str & ",Cantidad=" & .Correos(LoopC).Cantidad
    str = str & ",Item=" & .Correos(LoopC).Item

    Call Con.Execute(str)

    End With
    Set RS = Nothing




End Sub


Public Sub onpj(ByVal userindex As Integer)
Con.Execute "UPDATE `charflags` SET `Online` = '1' WHERE `IndexPJ` = " & UserList(userindex).Indexpj
End Sub
Public Sub offpj(ByVal userindex As Integer)
Con.Execute "UPDATE `charflags` SET `Online` = '0' WHERE `IndexPJ` = " & UserList(userindex).Indexpj
End Sub


