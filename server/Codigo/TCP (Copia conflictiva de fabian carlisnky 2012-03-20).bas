Attribute VB_Name = "TCP"


Option Explicit
Enum lStat
    Incinerado = &H1
    Envenenado = &H2
    Comerciand = &H4
    Trabajando = &H8
    Transformado = &H10
    Ciego = &H20
    Inactivo = &H40
    Silenciado = &H80
End Enum

Enum lStatEx
    Paralizado = &H1
    Inmovilizado = &H2
    Hombre = &H4
    Mujer = &H8
End Enum
Sub DarCuerpo(ByVal userindex As Integer)
Dim NewBody As Integer
Dim UserRaza As Byte
Dim UserGenero As Byte
UserGenero = UserList(userindex).Genero
UserRaza = UserList(userindex).Raza
Select Case UserGenero
   Case eGenero.Hombre
        Select Case UserRaza
            Case eRaza.Humano
                NewBody = 1
            Case eRaza.Elfo
                NewBody = 2
            Case eRaza.Drow
                NewBody = 3
            Case eRaza.Enano
                NewBody = 95
            Case eRaza.Gnomo
                NewBody = 52
            Case eRaza.Orco
                NewBody = 251
        End Select
   Case eGenero.Mujer
        Select Case UserRaza
            Case eRaza.Humano
                NewBody = 351
            Case eRaza.Elfo
                NewBody = 352
            Case eRaza.Drow
                NewBody = 353
            Case eRaza.Gnomo
                NewBody = 138
            Case eRaza.Enano
                NewBody = 138
            Case eRaza.Orco
                NewBody = 252
        End Select
End Select
UserList(userindex).Char.body = NewBody
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 48 Or car > 57) Then
        Numeric = False
        Exit Function
    End If
    
Next i

Numeric = True

End Function




Function ValidateSkills(ByVal userindex As Integer) As Boolean

Dim LoopC As Integer

For LoopC = 1 To NUMSKILLS
    If UserList(userindex).Stats.UserSkills(LoopC) < 0 Then UserList(userindex).Stats.UserSkills(LoopC) = 0
    If UserList(userindex).Stats.UserSkills(LoopC) > 100 Then UserList(userindex).Stats.UserSkills(LoopC) = 100
Next LoopC

ValidateSkills = True
    
End Function

Sub ConnectNewUser(ByVal userindex As Integer, ByRef Name As String, ByRef account As String, ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, ByVal UserClase As eClass, _
                    ByRef skills() As Byte, ByRef UserEmail As String, ByVal Hogar As eCiudad, _
                    ByVal Fuerza As Byte, ByVal Agilidad As Byte, ByVal Inteligencia As Byte, _
                    ByVal Carisma As Byte, ByVal constitucion As Byte, ByVal Cabeza As Integer, _
                    ByVal petTipe As eMascota, ByRef petName As String)
'*************************************************
'Author: Unknown
'Last modified: 20/4/2007
'Conecta un nuevo Usuario
'23/01/2007 Pablo (ToxicWaste) - Agregué ResetFaccion al crear usuario
'24/01/2007 Pablo (ToxicWaste) - Agregué el nuevo mana inicial de los magos.
'12/02/2007 Pablo (ToxicWaste) - Puse + 1 de const al Elfo normal.
'20/04/2007 Pablo (ToxicWaste) - Puse -1 de fuerza al Elfo.
'09/01/2008 Pablo (ToxicWaste) - Ahora los modificadores de Raza se controlan desde Balance.dat
'*************************************************

'If ServerSoloGMs > 0 Then
'        Call WriteErrorMsg(Userindex, "Servidor restringido a administradores.")
'        Call FlushBuffer(Userindex)
        'Call CloseSocket(UserIndex)
'        Exit Sub
'End If

If Not AsciiValidos(Name) Or LenB(Name) = 0 Then
    Call WriteErrorMsg(userindex, "Nombre invalido.")
    Exit Sub
End If

If Len(Name) > 20 Then
    Call WriteErrorMsg(userindex, "El nombre es muy largo.")
    Exit Sub
End If

If UserList(userindex).flags.UserLogged Then
    Call LogCheating("El usuario " & UserList(userindex).Name & " ha intentado crear a " & Name & " desde la IP " & UserList(userindex).ip)
    Call CloseSocket(userindex)
    Exit Sub
End If

Dim LoopC As Long
Dim totalskpts As Long

'¿Existe el personaje?
If ExistePersonaje(Name) = True Then
    Call WriteErrorMsg(userindex, "Ya existe el personaje.")
    Exit Sub
End If


UserList(userindex).flags.Muerto = 0
UserList(userindex).flags.Escondido = 0

UserList(userindex).Name = Name
UserList(userindex).Clase = UserClase
UserList(userindex).Raza = UserRaza
UserList(userindex).Genero = UserSexo
UserList(userindex).email = UserEmail
UserList(userindex).Hogar = Hogar

UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = Fuerza + ModRaza(UserRaza).Fuerza
UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = Agilidad + ModRaza(UserRaza).Agilidad
UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = IIf(Inteligencia + ModRaza(UserRaza).Inteligencia < 0, 0, Inteligencia + ModRaza(UserRaza).Inteligencia)
UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = Carisma + ModRaza(UserRaza).Carisma
UserList(userindex).Stats.UserAtributos(eAtributos.constitucion) = constitucion + ModRaza(UserRaza).constitucion

UserList(userindex).Stats.UserAtributosBackUP(eAtributos.Fuerza) = Fuerza + ModRaza(UserRaza).Fuerza
UserList(userindex).Stats.UserAtributosBackUP(eAtributos.Agilidad) = Agilidad + ModRaza(UserRaza).Agilidad
UserList(userindex).Stats.UserAtributosBackUP(eAtributos.Inteligencia) = IIf(Inteligencia + ModRaza(UserRaza).Inteligencia < 0, 0, Inteligencia + ModRaza(UserRaza).Inteligencia)
UserList(userindex).Stats.UserAtributosBackUP(eAtributos.Carisma) = Carisma + ModRaza(UserRaza).Carisma
UserList(userindex).Stats.UserAtributosBackUP(eAtributos.constitucion) = constitucion + ModRaza(UserRaza).constitucion

If (Fuerza + Agilidad + Inteligencia + Carisma + constitucion) > 70 Or _
   (Fuerza < 6 Or Agilidad < 6 Or Inteligencia < 6 Or Carisma < 6 Or constitucion < 6) Or _
   (Fuerza > 18 Or Agilidad > 18 Or Inteligencia > 18 Or Carisma > 18 Or constitucion > 18) Then
    
    Call LogHackAttemp(UserList(userindex).Name & " intento hackear los atributos.")
    'Call BorrarUsuario(UserList(UserIndex).name)
    Call WriteErrorMsg(userindex, "Por favor vaya a molestar a otro servidor.")
    Call FlushBuffer(userindex)
    Call CloseSocket(userindex)
    Exit Sub
End If

For LoopC = 1 To NUMSKILLS
    If skills(LoopC - 1) >= 0 Then
        UserList(userindex).Stats.UserSkills(LoopC) = skills(LoopC - 1)
        totalskpts = totalskpts + Abs(UserList(userindex).Stats.UserSkills(LoopC))
    Else
        Call LogHackAttemp(UserList(userindex).Name & " intento hackear los skills.")
        'Call BorrarUsuario(UserList(UserIndex).name)
        Call CloseSocket(userindex)
        Exit Sub
    End If
Next LoopC

If totalskpts > 10 Then
    Call LogHackAttemp(UserList(userindex).Name & " intento hackear los skills.")
    'Call BorrarUsuario(UserList(UserIndex).name)
    Call CloseSocket(userindex)
    Exit Sub
End If
'%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%

UserList(userindex).Char.heading = eHeading.SOUTH

Call DarCuerpo(userindex)
UserList(userindex).Char.Head = Cabeza
UserList(userindex).OrigChar = UserList(userindex).Char

If UserClase = eClass.Mago Or _
   UserClase = eClass.Druida Or _
   UserClase = eClass.Cazador Then

    If Len(petName) > 30 Then
        Call WriteErrorMsg(userindex, "El nombre de la mascota no debe sobrepasar 30 letras.")
        Call FlushBuffer(userindex)
        Exit Sub
    ElseIf Len(petName) = 0 Then
        Call WriteErrorMsg(userindex, "Nombre de familiar o mascota invalido.")
        Call FlushBuffer(userindex)
        Exit Sub
    End If
    
    Call EntregarMascota(userindex, petTipe, petName)
    
Else
    UserList(userindex).masc.TieneFamiliar = 0
    UserList(userindex).masc.Tipo = 0
    UserList(userindex).masc.Nombre = ""
End If

UserList(userindex).Char.WeaponAnim = NingunArma
UserList(userindex).Char.ShieldAnim = NingunEscudo
UserList(userindex).Char.CascoAnim = NingunCasco

Dim MiInt As Long
MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributos(eAtributos.constitucion) \ 3)

UserList(userindex).Stats.MaxHP = 15 + MiInt
UserList(userindex).Stats.MinHP = 15 + MiInt

MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) \ 6)
If MiInt = 1 Then MiInt = 2

UserList(userindex).Stats.MaxSTA = 20 * MiInt
UserList(userindex).Stats.MinSTA = 20 * MiInt

UserList(userindex).Stats.MaxAGU = 100
UserList(userindex).Stats.MinAGU = 100

UserList(userindex).Stats.MaxHAM = 100
UserList(userindex).Stats.MinHAM = 100


'<-----------------MANA----------------------->
If UserClase = eClass.Mago Then  'Cambio en mana inicial (ToxicWaste)
    MiInt = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) * 3
    UserList(userindex).Stats.MaxMAN = MiInt
    UserList(userindex).Stats.MinMAN = MiInt
ElseIf UserClase = eClass.Clerigo Or UserClase = eClass.Druida _
    Or UserClase = eClass.Bardo Or UserClase = eClass.Asesino _
    Or UserClase = eClass.Nigromante Then
        UserList(userindex).Stats.MaxMAN = 50
        UserList(userindex).Stats.MinMAN = 50
Else
    UserList(userindex).Stats.MaxMAN = 0
    UserList(userindex).Stats.MinMAN = 0
End If

If UserClase = eClass.Mago Or UserClase = eClass.Clerigo Or _
   UserClase = eClass.Druida Or UserClase = eClass.Bardo Or _
   UserClase = eClass.Asesino Or UserClase = eClass.Nigromante Then
        UserList(userindex).Stats.UserHechizos(1) = 2
End If

If UserClase = eClass.Mago Or _
   UserClase = eClass.Druida Then
        UserList(userindex).Stats.UserHechizos(2) = 59
End If

If UserClase = eClass.Cazador Then
    UserList(userindex).Stats.UserHechizos(1) = 59
End If

'Castelli Casamiento
UserList(userindex).flags.miPareja = ""
'Castelli Casamiento

UserList(userindex).Stats.MaxHit = 2
UserList(userindex).Stats.MinHit = 1

UserList(userindex).Stats.GLD = 0

UserList(userindex).Stats.Exp = 0
UserList(userindex).Stats.ELU = 300
UserList(userindex).Stats.ELV = 1

'???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
UserList(userindex).Invent.NroItems = 4

UserList(userindex).Invent.Object(1).ObjIndex = 573
UserList(userindex).Invent.Object(1).Amount = 100

UserList(userindex).Invent.Object(2).ObjIndex = 572
UserList(userindex).Invent.Object(2).Amount = 100

'Esto depende de la clase
If UserList(userindex).Clase = eClass.Cazador Or _
    UserList(userindex).Clase = eClass.Druida Then
    
    UserList(userindex).Invent.Object(3).ObjIndex = 1355
    UserList(userindex).Invent.Object(3).Amount = 1
    UserList(userindex).Invent.Object(3).Equipped = 1
ElseIf UserList(userindex).Clase = eClass.Paladin Or _
    UserList(userindex).Clase = eClass.Guerrero Or _
    UserList(userindex).Clase = eClass.Herrero Or _
    UserList(userindex).Clase = eClass.Mercenario Or _
    UserList(userindex).Clase = eClass.Minero Or _
    UserList(userindex).Clase = eClass.Clerigo Or _
    UserList(userindex).Clase = eClass.Leñador Then
    
    UserList(userindex).Invent.Object(3).ObjIndex = 574
    UserList(userindex).Invent.Object(3).Amount = 1
    UserList(userindex).Invent.Object(3).Equipped = 1
ElseIf UserList(userindex).Clase = eClass.Gladiador Or _
    UserList(userindex).Clase = eClass.Bardo Then
    
    UserList(userindex).Invent.Object(3).ObjIndex = 1354
    UserList(userindex).Invent.Object(3).Amount = 1
    UserList(userindex).Invent.Object(3).Equipped = 1
ElseIf UserList(userindex).Clase = eClass.Asesino Or _
    UserList(userindex).Clase = eClass.Ladron Or _
    UserList(userindex).Clase = eClass.Sastre Or _
    UserList(userindex).Clase = eClass.Pescador Then
    
    UserList(userindex).Invent.Object(3).ObjIndex = 460
    UserList(userindex).Invent.Object(3).Amount = 1
    UserList(userindex).Invent.Object(3).Equipped = 1
ElseIf UserList(userindex).Clase = eClass.Mago Or _
    UserList(userindex).Clase = eClass.Nigromante Then
    
    UserList(userindex).Invent.Object(3).ObjIndex = 1356
    UserList(userindex).Invent.Object(3).Amount = 1
    UserList(userindex).Invent.Object(3).Equipped = 1
End If

Select Case UserRaza
    Case eRaza.Humano
        If UserList(userindex).Genero = eGenero.Hombre Then
            UserList(userindex).Invent.Object(4).ObjIndex = 463
        Else
            UserList(userindex).Invent.Object(4).ObjIndex = 1283
        End If
    Case eRaza.Elfo
        If UserList(userindex).Genero = eGenero.Hombre Then
            UserList(userindex).Invent.Object(4).ObjIndex = 464
        Else
            UserList(userindex).Invent.Object(4).ObjIndex = 1284
        End If
    Case eRaza.Drow
        If UserList(userindex).Genero = eGenero.Hombre Then
            UserList(userindex).Invent.Object(4).ObjIndex = 465
        Else
            UserList(userindex).Invent.Object(4).ObjIndex = 1285
        End If
    Case eRaza.Enano
        If UserList(userindex).Genero = eGenero.Hombre Then
            UserList(userindex).Invent.Object(4).ObjIndex = 562
        Else
            UserList(userindex).Invent.Object(4).ObjIndex = 563
        End If
    Case eRaza.Gnomo
        If UserList(userindex).Genero = eGenero.Hombre Then
            UserList(userindex).Invent.Object(4).ObjIndex = 466
        Else
            UserList(userindex).Invent.Object(4).ObjIndex = 563
        End If
    Case eRaza.Orco
        If UserList(userindex).Genero = eGenero.Hombre Then
            UserList(userindex).Invent.Object(4).ObjIndex = 988
        Else
            UserList(userindex).Invent.Object(4).ObjIndex = 1087
        End If
End Select

UserList(userindex).Invent.Object(4).Amount = 1
UserList(userindex).Invent.Object(4).Equipped = 1

UserList(userindex).Invent.Object(5).ObjIndex = 461
UserList(userindex).Invent.Object(5).Amount = 100

If UserList(userindex).Clase = eClass.Cazador Or _
    UserList(userindex).Clase = eClass.Druida Then
    
    UserList(userindex).Invent.Object(6).ObjIndex = 1357
    UserList(userindex).Invent.Object(6).Amount = 100
ElseIf UserList(userindex).Clase = eClass.Asesino Or _
    UserList(userindex).Clase = eClass.Ladron Then
    
    UserList(userindex).Invent.Object(6).ObjIndex = 576
    UserList(userindex).Invent.Object(6).Amount = 100
End If

UserList(userindex).Invent.Object(7).ObjIndex = 1601
UserList(userindex).Invent.Object(7).Amount = 1

Dim tmpObj As Obj
tmpObj.ObjIndex = 879: tmpObj.Amount = 1
Call MeterItemEnInventario(userindex, tmpObj) 'Mapa

UserList(userindex).Invent.ArmourEqpSlot = 4
UserList(userindex).Invent.ArmourEqpObjIndex = UserList(userindex).Invent.Object(4).ObjIndex

UserList(userindex).Invent.WeaponEqpObjIndex = UserList(userindex).Invent.Object(3).ObjIndex
UserList(userindex).Invent.WeaponEqpSlot = 3

'Valores Default de facciones al Activar nuevo usuario
Call ResetFacciones(userindex)

If UserList(userindex).Hogar = 0 Then
    UserList(userindex).Faccion.Ciudadano = 1
ElseIf UserList(userindex).Hogar = 1 Then
    UserList(userindex).Faccion.Republicano = 1
End If

Call SaveUserSQL(userindex, account, True)

'Open User
Call ConnectUser(userindex, Name, account)
  
End Sub

Sub CloseSocket(ByVal userindex As Integer)
Debug.Print "CLOSESOCKET!!!"

On Error GoTo Errhandler
  
       If userindex = LastUser Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser < 1 Then Exit Do
        Loop
    End If
    
    If UserList(userindex).ConnID <> -1 Then
        Call CloseSocketSL(userindex)
    End If
    
    If UserList(userindex).ComUsu.DestUsu > 0 Then
        If UserList(UserList(userindex).ComUsu.DestUsu).flags.UserLogged Then
            If UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.DestUsu = userindex Then
                Call WriteConsoleMsg(1, UserList(userindex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(userindex).ComUsu.DestUsu)
                Call FlushBuffer(UserList(userindex).ComUsu.DestUsu)
            End If
        End If
    End If
    
        'Empty buffer for reuse
    Call UserList(userindex).incomingData.ReadASCIIStringFixed(UserList(userindex).incomingData.length)


    If UserList(userindex).flags.UserLogged = True Then
        Call CloseUser(userindex)
    Else
        Call ResetUserSlot(userindex)
        End If

Exit Sub

Errhandler:
    Call ResetUserSlot(userindex)
    Call LogError("CloseSocket - Error = " & err.Number & " - Descripción = " & err.description & " - UserIndex = " & userindex)
End Sub


Public Function EnviarDatosASlot(ByVal userindex As Integer, ByRef Datos As String) As Long
    Dim Ret As Long
    
    Ret = WsApiEnviar(userindex, Datos)
    
    If Ret <> 0 And Ret <> WSAEWOULDBLOCK Then
        CloseSocket userindex
    End If
End Function
Function EstaPCarea(index As Integer, Index2 As Integer) As Boolean


Dim x As Integer, Y As Integer
For Y = UserList(index).pos.Y - MinYBorder + 1 To UserList(index).pos.Y + MinYBorder - 1
        For x = UserList(index).pos.x - MinXBorder + 1 To UserList(index).pos.x + MinXBorder - 1

            If MapData(UserList(index).pos.map, x, Y).userindex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
        Next x
Next Y
EstaPCarea = False
End Function

Function HayPCarea(pos As WorldPos) As Boolean


Dim x As Integer, Y As Integer
For Y = pos.Y - MinYBorder + 1 To pos.Y + MinYBorder - 1
        For x = pos.x - MinXBorder + 1 To pos.x + MinXBorder - 1
            If x > 0 And Y > 0 And x < 101 And Y < 101 Then
                If MapData(pos.map, x, Y).userindex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next x
Next Y
HayPCarea = False
End Function

Function HayOBJarea(pos As WorldPos, ObjIndex As Integer) As Boolean


Dim x As Integer, Y As Integer
For Y = pos.Y - MinYBorder + 1 To pos.Y + MinYBorder - 1
        For x = pos.x - MinXBorder + 1 To pos.x + MinXBorder - 1
            If MapData(pos.map, x, Y).ObjInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function
            End If
        
        Next x
Next Y
HayOBJarea = False
End Function
Function ValidateChr(ByVal userindex As Integer) As Boolean

ValidateChr = UserList(userindex).Char.Head <> 0 _
                And UserList(userindex).Char.body <> 0 _
                And ValidateSkills(userindex)

End Function

Sub ConnectUser(ByVal userindex As Integer, ByRef Name As String, ByRef account As String)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa - Agrego por default que el color de dialogo de los dioses, sea como el de su nick.
'***************************************************
Dim N As Integer
Dim tStr As String


With UserList(userindex)


'Reseteamos los FLAGS
.flags.Escondido = 0
.flags.TargetNPC = 0
.flags.TargetNpcTipo = eNPCType.Comun
.flags.TargetObj = 0
.flags.TargetUser = 0
.Char.FX = 0


'¿Existe el personaje?
If ExistePersonaje(Name) = False Then
    Call WriteErrorMsg(userindex, "El personaje no existe.")
    Call FlushBuffer(userindex)
    'Call CloseSocket(UserIndex)
    Exit Sub
End If

'¿Ya esta conectado el personaje?
If CheckForSameName(Name) Then
    If UserList(NameIndex(Name)).Counters.Saliendo Then
        Call WriteErrorMsg(userindex, "El usuario está saliendo.")
    Else
        Call WriteErrorMsg(userindex, "Usuario Conectado.")
    End If
    Call FlushBuffer(userindex)
    Exit Sub
End If

'Reseteamos los privilegios
.flags.Privilegios = 0

'Vemos que clase de user es (se lo usa para setear los privilegios al loguear el PJ)
If EsDios(Name) Then
    .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
    .flags.AdminPerseguible = False
ElseIf EsVIP(Name) Then
    .flags.Privilegios = .flags.Privilegios Or PlayerType.VIP
    .flags.AdminPerseguible = True
Else
    .flags.Privilegios = .flags.Privilegios Or PlayerType.User
    .flags.AdminPerseguible = True
End If


If .Stats.PuedeStaff = 0 Then
    If .flags.Privilegios And (PlayerType.Dios Or PlayerType.VIP) Then
        Call WriteErrorMsg(userindex, "No es tu personaje.")
        Call FlushBuffer(userindex)
        Call CloseSocket(userindex)
        Exit Sub
    End If
End If

Call LoadUserSQL(userindex, Name)

'Donador por jose ignacio castelli
  If Comprobar_Si_Donador(account) > 0 Then
    UserList(userindex).donador = 1
  End If

If Not ValidateChr(userindex) Then
    Call WriteErrorMsg(userindex, "Error en el personaje.")
    Call FlushBuffer(userindex)
    Call CloseSocket(userindex)
    Exit Sub
End If

If .Counters.IdleCount > 0 Then
.Counters.IdleCount = 0
End If

If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
If .Invent.WeaponEqpSlot = 0 And .Invent.NudiEqpSlot = 0 Then .Char.WeaponAnim = NingunArma

Call UpdateUserInv(True, userindex, 0)
Call UpdateUserHechizos(True, userindex, 0)

If .flags.Paralizado Then
    Call WriteParalizeOK(userindex, False)
End If

''
'TODO : Feo, esto tiene que ser parche cliente
If .flags.Estupidez = 0 Then
    Call WriteDumbNoMore(userindex)
End If

'Posicion de comienzo
If .pos.map = 0 Then
    Select Case .Hogar
        Case 0 ' Nix
            .pos.x = 40
            .pos.Y = 87
            .pos.map = 34
        Case 1 ' Illindor
            .pos.x = 50
            .pos.Y = 78
            .pos.map = 185
    End Select
Else
    If Not MapaValido(.pos.map) Then
        .pos.map = 1
    End If
End If


If .flags.Privilegios = PlayerType.Dios Or .flags.Privilegios = PlayerType.VIP Then 'PlayerType.Dios Or PlayerType.VIP Then
            .pos.x = 50
            .pos.Y = 50
            .pos.map = 248
End If








'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martín Sotuyo Dodero (Maraxus)
If MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex <> 0 Or MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).NpcIndex <> 0 Then
    Dim FoundPlace As Boolean
    Dim esAgua As Boolean
    Dim tX As Long
    Dim tY As Long
    
    FoundPlace = False
    esAgua = HayAgua(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y)
    
    For tY = UserList(userindex).pos.Y - 1 To UserList(userindex).pos.Y + 1
        For tX = UserList(userindex).pos.x - 1 To UserList(userindex).pos.x + 1
            If esAgua Then
                'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
                If LegalPos(UserList(userindex).pos.map, tX, tY, True, False) Then
                    FoundPlace = True
                    Exit For
                End If
            Else
                'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
                If LegalPos(UserList(userindex).pos.map, tX, tY, False, True) Then
                    FoundPlace = True
                    Exit For
                End If
            End If
        Next tX
        
        If FoundPlace Then _
            Exit For
    Next tY
    
    If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
        UserList(userindex).pos.x = tX
        UserList(userindex).pos.Y = tY
    Else
        'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
        If MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex <> 0 Then
            'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
            If UserList(MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex).ComUsu.DestUsu > 0 Then
                'Le avisamos al que estaba comerciando que se tuvo que ir.
                If UserList(UserList(MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex).ComUsu.DestUsu).flags.UserLogged Then
                    Call FinComerciarUsu(UserList(MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex).ComUsu.DestUsu)
                    Call WriteConsoleMsg(1, UserList(MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                    Call FlushBuffer(UserList(MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex).ComUsu.DestUsu)
                End If
                'Lo sacamos.
                If UserList(MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex).flags.UserLogged Then
                    Call FinComerciarUsu(MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex)
                    Call WriteErrorMsg(MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                    Call FlushBuffer(MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).userindex)
                End If
            End If

        End If
    End If
End If







'Nombre de sistema
.Name = Name

.showName = True 'Por default los nombres son visibles

'If in the water, and has a boat, equip it!
If .Invent.BarcoObjIndex > 0 And _
        (HayAgua(.pos.map, .pos.x, .pos.Y)) Then
    Dim Barco As ObjData
    Barco = ObjData(.Invent.BarcoObjIndex)
    .Char.Head = 0
    If .flags.Muerto <> 0 Then
        .Char.body = iFragataFantasmal
    End If
    .Char.body = 84
    .flags.Navegando = 1
End If


'Info
Call WriteUserIndexInServer(userindex) 'Enviamos el User index
Call WriteChangeMap(userindex, .pos.map, MapInfo(.pos.map).MapVersion) 'Carga el mapa
Call WritePlayMidi(userindex, val(ReadField(1, MapInfo(.pos.map).Music, 45)))


'Crea  el personaje del usuario
Call MakeUserChar(True, .pos.map, userindex, .pos.map, .pos.x, .pos.Y)

Call WriteUserCharIndexInServer(userindex)
''[/el oso]

Call WriteUpdateUserStats(userindex)

Call WriteUpdateHungerAndThirst(userindex)
If haciendoBK Then
    Call WritePauseToggle(userindex)
    Call WriteConsoleMsg(1, userindex, "Servidor> Por favor espera algunos segundos, WorldSave esta ejecutandose.", FontTypeNames.FONTTYPE_SERVER)
End If

If EnPausa Then
    Call WritePauseToggle(userindex)
    Call WriteConsoleMsg(1, userindex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER)
End If

.flags.UserLogged = True

MapInfo(.pos.map).NumUsers = MapInfo(.pos.map).NumUsers + 1

If .Stats.SkillPts > 0 Then
    Call WriteSendSkills(userindex)
End If

If .NroMascotas > 0 And MapInfo(.pos.map).Pk Then
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        If .MascotasType(i) > 0 Then
            .MascotasIndex(i) = SpawnNpc(.MascotasType(i), .pos, True, True)
            
            If .MascotasIndex(i) > 0 Then
                Npclist(.MascotasIndex(i)).MaestroUser = userindex
                Call FollowAmo(.MascotasIndex(i))
            Else
                .MascotasIndex(i) = 0
            End If
        End If
    Next i
End If

If .flags.Navegando = 1 Then
    Call WriteNavigateToggle(userindex)
End If

If esRene(userindex) Then
    Call WriteSafeModeOff(userindex)
    .flags.Seguro = False
Else
    .flags.Seguro = True
    Call WriteSafeModeOn(userindex)
End If

If .GuildIndex > 0 Then
    'welcome to the show baby...
If Not modGuilds.m_ConectarMiembroAClan(userindex, .GuildIndex) Then
        Call WriteConsoleMsg(1, userindex, "Tu estado no te permite entrar al clan.", FontTypeNames.FONTTYPE_GUILD)
    End If
End If

Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))

Call WriteLoggedMessage(userindex)

tStr = modGuilds.a_ObtenerRechazoDeChar(.Name)

If LenB(tStr) <> 0 Then
    Call WriteShowMessageBox(userindex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
End If


Call WriteFuerza(userindex)
Call WriteAgilidad(userindex)
WriteMensajeSigno userindex

WriteHora userindex

If NumUsers > RecordUsuarios Then
    Call SendData(SendTarget.toAll, 0, PrepareMessageConsoleMsg(1, "Record de usuarios conectados simultaneamente." & "Hay " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_BROWNI))
    RecordUsuarios = NumUsers
    
    Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(RecordUsuarios))
End If


If .flags.Privilegios <> PlayerType.Dios Then
NumUsers = NumUsers + 1
End If


'Add Nod Kopfnickend
'Pj online en la DB
Call onpj(userindex)

SendOnline

FlushBuffer userindex
DoEvents

End With

End Sub


Sub ResetFacciones(ByVal userindex As Integer)
    With UserList(userindex).Faccion
        .ArmadaReal = 0
        .FuerzasCaos = 0
        .Milicia = 0
        .Rango = 0
        
        .Ciudadano = 0
        .Republicano = 0
        .Renegado = 0
        
        .CaosMatados = 0
        .ArmadaMatados = 0
        .MilicianosMatados = 0
        
        .CiudadanosMatados = 0
        .RenegadosMatados = 0
        .RepublicanosMatados = 0
    End With
End Sub

Sub ResetContadores(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'05/20/2007 Integer - Agregue todas las variables que faltaban.
'*************************************************
    With UserList(userindex).Counters
        .AGUACounter = 0
        .AttackCounter = 0
        .Ceguera = 0
        .COMCounter = 0
        .Estupidez = 0
        .Frio = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Paralisis = 0
        .Pena = 0
        .PiqueteC = 0
        .STACounter = 0
        .Veneno = 0
        .Fuego = 0
        .Trabajando = 0
        .Ocultando = 0
        .Saliendo = False
        .Salir = 0
        .TiempoOculto = 0
        .TimerMagiaGolpe = 0
        .TimerGolpeMagia = 0
        .TimerLanzarSpell = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeUsarArco = 0
        .TimerPuedeTrabajar = 0
        .TimerUsar = 0
    End With
End Sub

Sub ResetCharInfo(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(userindex).Char
        .body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .loops = 0
        .heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
End Sub

Sub ResetBasicUserInfo(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(userindex)
        .Name = vbNullString
        .desc = vbNullString
        
        .pos.map = 0
        .pos.x = 0
        .pos.Y = 0
        .ip = vbNullString
        .Clase = 0
        .email = vbNullString
        .Genero = 0
        .Hogar = 0
        .Raza = 0

        .GrupoIndex = 0
        .GrupoSolicitud = 0
        
        With .Stats
            .Banco = 0
            .ELV = 0
            .ELU = 0
            .Exp = 0
            .def = 0
            '.CriminalesMatados = 0
            .NPCsMuertos = 0
            .VecesMuertos = 0
            
            .SkillPts = 0
            .GLD = 0
            .UserAtributos(1) = 0
            .UserAtributos(2) = 0
            .UserAtributos(3) = 0
            .UserAtributos(4) = 0
            .UserAtributos(5) = 0
            .UserAtributosBackUP(1) = 0
            .UserAtributosBackUP(2) = 0
            .UserAtributosBackUP(3) = 0
            .UserAtributosBackUP(4) = 0
            .UserAtributosBackUP(5) = 0
        End With
        
    End With
End Sub


Sub ResetGuildInfo(ByVal userindex As Integer)
    If UserList(userindex).EscucheClan > 0 Then
        Call modGuilds.GMDejaDeEscucharClan(userindex, UserList(userindex).EscucheClan)
        UserList(userindex).EscucheClan = 0
    End If
    If UserList(userindex).GuildIndex > 0 Then
        Call modGuilds.m_DesconectarMiembroDelClan(userindex, UserList(userindex).GuildIndex)
    End If
    UserList(userindex).GuildIndex = 0
End Sub

Sub ResetUserFlags(ByVal userindex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 06/28/2008
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'06/28/2008 NicoNZ - Agrego el flag Inmovilizado
'*************************************************
    With UserList(userindex).flags
        .Comerciando = False
        .Ban = 0
        .Escondido = 0
        .DuracionEfecto = 0
        .TargetNPC = 0
        .TargetNpcTipo = eNPCType.Comun
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TargetUser = 0
        .TipoPocion = 0
        .TomoPocion = False
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .ModoCombate = False
        .Navegando = 0
        .Montando = 0
        .Oculto = 0
        .Envenenado = 0
        .Metamorfosis = 0
        .Incinerado = 0
        .Invisible = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Meditando = 0
        .Trabajando = 0
        .Lingoteando = 0
        .Privilegios = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .Hechizo = 0
        .TimesWalk = 0
        .Silenciado = 0
        .AdminPerseguible = False
    End With
End Sub

Sub ResetUserSpells(ByVal userindex As Integer)
    Dim LoopC As Long
    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(userindex).Stats.UserHechizos(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserPets(ByVal userindex As Integer)
    Dim LoopC As Long
    
    UserList(userindex).NroMascotas = 0
        
    For LoopC = 1 To MAXMASCOTAS
        UserList(userindex).MascotasIndex(LoopC) = 0
        UserList(userindex).MascotasType(LoopC) = 0
    Next LoopC
    
End Sub

Sub ResetUserBanco(ByVal userindex As Integer)
    Dim LoopC As Long
    
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
          UserList(userindex).BancoInvent.Object(LoopC).Amount = 0
          UserList(userindex).BancoInvent.Object(LoopC).Equipped = 0
          UserList(userindex).BancoInvent.Object(LoopC).ObjIndex = 0
    Next LoopC
    
    UserList(userindex).BancoInvent.NroItems = 0
End Sub

Public Sub LimpiarComercioSeguro(ByVal userindex As Integer)
    With UserList(userindex).ComUsu
        If .DestUsu > 0 Then
            Call FinComerciarUsu(.DestUsu)
            Call FinComerciarUsu(userindex)
        End If
    End With
End Sub

Sub ResetUserSlot(ByVal userindex As Integer)

    'CASTELLI / Ubicamos esto aca ya que borra el sock cuando se va de la cuenta
    offcuenta (userindex)

ZeroMemory ByVal VarPtr(UserList(userindex)), LenB(UserList(userindex))


UserList(userindex).ConnIDValida = False
UserList(userindex).ConnID = -1
    

    
    Set UserList(userindex).incomingData = New clsByteQueue
    Set UserList(userindex).outgoingData = New clsByteQueue
End Sub


Sub ResetUserSlot2(ByVal userindex As Integer)

    If UserList(userindex).flags.accountlogged = False Then
        ResetUserSlot (userindex)
        Exit Sub
    End If

        
        Dim tUser As User
        tUser = UserList(userindex)
        

        
    
    
    ZeroMemory ByVal VarPtr(UserList(userindex)), LenB(UserList(userindex))
        

        UserList(userindex).account = tUser.account
        UserList(userindex).IndexAccount = tUser.IndexAccount 'Add Nod Kopfnickend
        UserList(userindex).ConnID = tUser.ConnID
        UserList(userindex).ConnIDValida = tUser.ConnIDValida
        UserList(userindex).Stats.PuedeStaff = tUser.Stats.PuedeStaff
        UserList(userindex).flags.accountlogged = True
        
    Set UserList(userindex).incomingData = New clsByteQueue
    Set UserList(userindex).outgoingData = New clsByteQueue
    
    
    
        
End Sub


Sub CloseUser(ByVal userindex As Integer)
On Error GoTo Errhandler

Dim N As Integer
Dim LoopC As Integer
Dim map As Integer
Dim Name As String
Dim i As Integer

Dim aN As Integer

If UserList(userindex).flags.automatico = True Then
Call Rondas_UsuarioDesconecta(userindex)
End If
If UserList(userindex).pos.map = 238 And UserList(userindex).flags.automatico = False Then
Call WarpUserChar(userindex, 34, 27, 72, True)
End If

If UserList(userindex).flags.Privilegios <> PlayerType.Dios Then
    If NumUsers > 0 Then NumUsers = NumUsers - 1
    End If
    
    SendOnline

aN = UserList(userindex).flags.AtacadoPorNpc
If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = 0
End If
aN = UserList(userindex).flags.NPCAtacado
If aN > 0 Then
    If Npclist(aN).flags.AttackedFirstBy = UserList(userindex).Name Then
        Npclist(aN).flags.AttackedFirstBy = vbNullString
    End If
End If
UserList(userindex).flags.AtacadoPorNpc = 0
UserList(userindex).flags.NPCAtacado = 0

Call ControlarPortalLum(userindex)
        'desequipamos items macigos
        If UserList(userindex).Invent.MagicIndex > 0 Then
            Call Desequipar(userindex, UserList(userindex).Invent.MagicSlot)
        End If

map = UserList(userindex).pos.map
Name = UCase$(UserList(userindex).Name)

UserList(userindex).Char.FX = 0
UserList(userindex).Char.loops = 0
Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(UserList(userindex).Char.CharIndex, 0, 0))


UserList(userindex).flags.UserLogged = False
UserList(userindex).Counters.Saliendo = False

    'Le devolvemos el body y head originales
    If UserList(userindex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(userindex)
    
    'si esta en Grupo le devolvemos la experiencia
    If UserList(userindex).GrupoIndex > 0 Then Call mdGrupo.SalirDeGrupo(userindex)

    If UserList(userindex).flags.inDuelo = 1 Then PerderDuelo userindex
    
    'castelli // desinvocamos fami
    If UserList(userindex).masc.invocado = True Then Call desinvocarfami(userindex)
    'castelli // desinvocamos fami

Call SaveUserSQL(userindex)

If MapInfo(map).NumUsers > 0 Then
    Call SendData(SendTarget.ToPCAreaButIndex, userindex, PrepareMessageRemoveCharDialog(UserList(userindex).Char.CharIndex))
End If

'Borrar el personaje
If UserList(userindex).Char.CharIndex > 0 Then
    Call EraseUserChar(userindex)
End If

'Borrar mascotas
For i = 1 To MAXMASCOTAS
    If UserList(userindex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(userindex).MascotasIndex(i)).flags.NPCActive Then _
            Call QuitarNPC(UserList(userindex).MascotasIndex(i))
    End If
Next i

'Update Map Users
MapInfo(map).NumUsers = MapInfo(map).NumUsers - 1

If MapInfo(map).NumUsers < 0 Then
    MapInfo(map).NumUsers = 0
End If

' Si el usuario habia dejado un msg en la gm's queue lo borramos
If Ayuda.Existe(UserList(userindex).Name) Then Call Ayuda.Quitar(UserList(userindex).Name)

Call offpj(userindex)

'CASTELLI ////  Reset info del clan, es un array dinamico doble, por eso_
'hay que resetearlo asi...
Call ResetGuildInfo(userindex)
'CASTELLI ////


Call ResetUserSlot(userindex)




Exit Sub

Errhandler:
Call LogError("Error en CloseUser. Número " & err.Number & " Descripción: " & err.description)

End Sub

Sub EntregarMascota(ByVal userindex As Integer, petTipe As eMascota, ByRef petName As String)
    With UserList(userindex)
        If .Clase = eClass.Mago Then
            If petTipe > 5 Or petTipe < 1 Then
                petTipe = eMascota.Fuego
            End If
        Else
            If petTipe < 6 Then
                petTipe = eMascota.Ent
            End If
        End If
        
        .masc.TieneFamiliar = 1
        .masc.Tipo = petTipe
        .masc.Nombre = petName
        
        .masc.ELV = 1
        .masc.ELU = 100
        .masc.MinHP = 10
        .masc.MaxHP = 10
        Select Case petTipe
            Case eMascota.Fuego, eMascota.Tierra
                .masc.MinHit = 5
                .masc.MaxHit = 15
                
            Case eMascota.Agua
                .masc.MinHit = 7
                .masc.MaxHit = 20
                
            Case eMascota.Ely
                .masc.MinHP = 15
                .masc.MaxHP = 15
                .masc.MinHit = 5
                .masc.MaxHit = 20
            Case eMascota.Fatuo
                .masc.MinHP = 7
                .masc.MaxHP = 7
                .masc.MinHit = 5
                .masc.MaxHit = 10
            
            'Caza o Druida
            Case eMascota.Tigre
                .masc.MinHP = 15
                .masc.MaxHP = 15
                .masc.MinHit = 10
                .masc.MaxHit = 20
            Case eMascota.Lobo
                .masc.MinHP = 20
                .masc.MaxHP = 20
                .masc.MinHit = 10
                .masc.MaxHit = 20
            Case eMascota.Oso
                .masc.MinHP = 20
                .masc.MaxHP = 20
                .masc.MinHit = 5
                .masc.MaxHit = 30
            Case eMascota.Oso
                .masc.MinHP = 17
                .masc.MaxHP = 17
                .masc.MinHit = 10
                .masc.MaxHit = 15
        End Select
        
    End With
End Sub
Public Sub EcharPjsNoPrivilegiados()
Dim LoopC As Long

For LoopC = 1 To LastUser
    If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
        If UserList(LoopC).flags.Privilegios And PlayerType.User Then
            Call SaveUserSQL(CInt(LoopC))
            Call CloseSocket(LoopC)
        End If
    End If
Next LoopC

End Sub
Function Generate_Char_Stat(ByVal userindex As Integer) As Byte
    With UserList(userindex)
        If .flags.Envenenado > 0 Then
            Generate_Char_Stat = Generate_Char_Stat Or lStat.Envenenado
        End If
    
        If .flags.Trabajando = 1 Then
            Generate_Char_Stat = Generate_Char_Stat Or lStat.Trabajando
        End If
    
        If .flags.Silenciado = 1 Then
            Generate_Char_Stat = Generate_Char_Stat Or lStat.Silenciado
        End If
    
        If .flags.Ceguera = 1 Then
            Generate_Char_Stat = Generate_Char_Stat Or lStat.Ciego
        End If
    
        If .flags.Incinerado = 1 Then
            Generate_Char_Stat = Generate_Char_Stat Or lStat.Incinerado
        End If
    
        If .flags.Metamorfosis = 1 Then
            Generate_Char_Stat = Generate_Char_Stat Or lStat.Transformado
        End If
    
        If .flags.Comerciando = 1 Then
            Generate_Char_Stat = Generate_Char_Stat Or lStat.Comerciand
        End If
    
        If .Counters.IdleCount > 1 Then
            Generate_Char_Stat = Generate_Char_Stat Or lStat.Inactivo
        End If
    End With
End Function
Function Generate_Char_StatEx(ByVal userindex As Integer) As Byte

With UserList(userindex)
    If .flags.Paralizado = 1 Then
       Generate_Char_StatEx = Generate_Char_StatEx Or lStatEx.Paralizado
    End If

    If .flags.Inmovilizado = 1 Then
        Generate_Char_StatEx = Generate_Char_StatEx Or lStatEx.Inmovilizado
    End If
    
    If .Genero = eGenero.Hombre Then
        Generate_Char_StatEx = Generate_Char_StatEx Or lStatEx.Hombre
    Else
        Generate_Char_StatEx = Generate_Char_StatEx Or lStatEx.Mujer
    End If
End With
End Function

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal userindex As Integer)

If UserList(userindex).ConnID <> -1 And UserList(userindex).ConnIDValida Then
    Call BorraSlotSock(UserList(userindex).ConnID)
    Call WSApiCloseSocket(UserList(userindex).ConnID)
    UserList(userindex).ConnIDValida = False
End If
End Sub
