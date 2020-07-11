Attribute VB_Name = "ES"

Option Explicit
'***************************
'Sinuhe - Map format .CSM
'***************************

Private Type tMapHeader
    NumeroBloqueados As Long
    NumeroLayers(2 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long
End Type

Private Type tDatosBloqueados
    x As Integer
    Y As Integer
End Type

Private Type tDatosGrh
    x As Integer
    Y As Integer
    GrhIndex As Long
End Type

Private Type tDatosTrigger
    x As Integer
    Y As Integer
    Trigger As Integer
End Type

Private Type tDatosLuces
    x As Integer
    Y As Integer
    color As Long
    Rango As Byte
End Type

Private Type tDatosParticulas
    x As Integer
    Y As Integer
    Particula As Long
End Type

Private Type tDatosNPC
    x As Integer
    Y As Integer
    NpcIndex As Integer
End Type

Private Type tDatosObjs
    x As Integer
    Y As Integer
    ObjIndex As Integer
    ObjAmmount As Integer
End Type

Private Type tDatosTE
    x As Integer
    Y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer
End Type

Private Type tMapSize
    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer
End Type

Private Type tMapDat
    map_name As String * 64
    battle_mode As Byte
    backup_mode As Byte
    restrict_mode As String * 4
    music_number As String * 16
    zone As String * 16
    terrain As String * 16
    ambient As String * 16
    base_light As Long
    letter_grh As Long
    extra1 As Long
    extra2 As Long
    extra3 As String * 32
End Type
Public Sub CargarSpawnList()
    Dim N As Integer, loopC As Integer
    N = val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(N) As tCriaturasEntrenador
    For loopC = 1 To N
        SpawnList(loopC).NpcIndex = val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & loopC))
        SpawnList(loopC).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & loopC)
    Next loopC
    
End Sub

Public Sub CargarHechizos()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'  ¡¡¡¡ NO USAR GetVar PARA LEER Hechizos.dat !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer Hechizos.dat se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

On Error GoTo Errhandler

frmCargando.Label1(2).Caption = "Cargando Hechizos ...."

Dim Hechizo As Integer
Dim leer As New clsIniReader

Call leer.Initialize(DatPath & "Hechizos.dat")

'obtiene el numero de hechizos
NumeroHechizos = val(leer.GetValue("INIT", "NumeroHechizos"))

ReDim Hechizos(1 To NumeroHechizos) As tHechizo

'frmCargando.cargar.min = 0
'frmCargando.cargar.max = NumeroHechizos
'frmCargando.cargar.value = 0

'Llena la lista
For Hechizo = 1 To NumeroHechizos
    
    Hechizos(Hechizo).Nombre = leer.GetValue("Hechizo" & Hechizo, "Nombre")
    Hechizos(Hechizo).desc = leer.GetValue("Hechizo" & Hechizo, "Desc")
    Hechizos(Hechizo).PalabrasMagicas = leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
    
    Hechizos(Hechizo).HechizeroMsg = leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
    Hechizos(Hechizo).TargetMsg = leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
    Hechizos(Hechizo).PropioMsg = leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
    
    Hechizos(Hechizo).tipo = val(leer.GetValue("Hechizo" & Hechizo, "Tipo"))
    
    Hechizos(Hechizo).WAV = val(leer.GetValue("Hechizo" & Hechizo, "WAV"))
    
    Hechizos(Hechizo).FXgrh = val(leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
    Hechizos(Hechizo).loops = val(leer.GetValue("Hechizo" & Hechizo, "Loops"))
    
    Hechizos(Hechizo).Particle = val(leer.GetValue("Hechizo" & Hechizo, "Particle"))
 
    Hechizos(Hechizo).SubeHP = val(leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
    Hechizos(Hechizo).MinHP = val(leer.GetValue("Hechizo" & Hechizo, "MinHP"))
    Hechizos(Hechizo).MaxHP = val(leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
    
    Hechizos(Hechizo).SubeAgilidad = val(leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
    Hechizos(Hechizo).MinAgilidad = val(leer.GetValue("Hechizo" & Hechizo, "MinAG"))
    Hechizos(Hechizo).MaxAgilidad = val(leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
    
    Hechizos(Hechizo).SubeFuerza = val(leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
    Hechizos(Hechizo).MinFuerza = val(leer.GetValue("Hechizo" & Hechizo, "MinFU"))
    Hechizos(Hechizo).MaxFuerza = val(leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
    
    Hechizos(Hechizo).Invisibilidad = val(leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
    Hechizos(Hechizo).Paraliza = val(leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
    Hechizos(Hechizo).Inmoviliza = val(leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
    Hechizos(Hechizo).RemoverParalisis = val(leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
    Hechizos(Hechizo).RemueveInvisibilidadParcial = val(leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
    
    Hechizos(Hechizo).CuraVeneno = val(leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
    Hechizos(Hechizo).Envenena = val(leer.GetValue("Hechizo" & Hechizo, "Envenena"))
    Hechizos(Hechizo).Incinera = val(leer.GetValue("Hechizo" & Hechizo, "Incinera"))
    Hechizos(Hechizo).Revivir = val(leer.GetValue("Hechizo" & Hechizo, "Revivir"))
    
    'Add Marius
    Hechizos(Hechizo).ExclusivoClase = UCase$(leer.GetValue("Hechizo" & Hechizo, "ExclusivoClase"))
    '\Add
    
    Hechizos(Hechizo).Resurreccion = val(leer.GetValue("Hechizo" & Hechizo, "Resurreccion"))
    Hechizos(Hechizo).ReviveFamiliar = val(leer.GetValue("Hechizo" & Hechizo, "ResucitaFamiliar"))
    Hechizos(Hechizo).Sanacion = val(leer.GetValue("Hechizo" & Hechizo, "Sanacion"))
    
    Hechizos(Hechizo).Ceguera = val(leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
    Hechizos(Hechizo).Estupidez = val(leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
    
    Hechizos(Hechizo).Invoca = val(leer.GetValue("Hechizo" & Hechizo, "Invoca"))
    Hechizos(Hechizo).NumNpc = val(leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
    Hechizos(Hechizo).Cant = val(leer.GetValue("Hechizo" & Hechizo, "Cant"))
    Hechizos(Hechizo).Mimetiza = val(leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
    
    Hechizos(Hechizo).AutoLanzar = val(leer.GetValue("hechizo" & Hechizo, "autolanzar"))
    Hechizos(Hechizo).Desencantar = val(leer.GetValue("hechizo" & Hechizo, "desencantar"))
    
    Hechizos(Hechizo).HechizoDeArea = val(leer.GetValue("hechizo" & Hechizo, "HechizoDeArea"))
    Hechizos(Hechizo).AreaEfecto = val(leer.GetValue("hechizo" & Hechizo, "AreaEfecto"))
    Hechizos(Hechizo).Afecta = val(leer.GetValue("hechizo" & Hechizo, "Afecta"))
    Hechizos(Hechizo).Metamorfosis = val(leer.GetValue("hechizo" & Hechizo, "metamorfosis"))
    
    Hechizos(Hechizo).Certero = val(leer.GetValue("hechizo" & Hechizo, "GolpeCertero"))
    
    If Hechizos(Hechizo).Metamorfosis = 1 Then
        Hechizos(Hechizo).Extrahit = val(leer.GetValue("hechizo" & Hechizo, "ExtraHIT"))
        Hechizos(Hechizo).Extradef = val(leer.GetValue("hechizo" & Hechizo, "ExtraDEF"))
        Hechizos(Hechizo).body = val(leer.GetValue("hechizo" & Hechizo, "body"))
        Hechizos(Hechizo).Head = val(leer.GetValue("hechizo" & Hechizo, "head"))
        Hechizos(Hechizo).MetaObj = val(leer.GetValue("hechizo" & Hechizo, "MetaObj"))
    End If
    
    If Hechizos(Hechizo).tipo = TipoHechizo.uCreateMagic Then
        Hechizos(Hechizo).CreaAlgo = val(leer.GetValue("Hechizo" & Hechizo, "CreaTipo"))
        Hechizos(Hechizo).MaxHit = val(leer.GetValue("Hechizo" & Hechizo, "MaxHit"))
        Hechizos(Hechizo).MinHit = val(leer.GetValue("Hechizo" & Hechizo, "MinHit"))
        
        Hechizos(Hechizo).MaxDef = val(leer.GetValue("Hechizo" & Hechizo, "MaxDef"))
        Hechizos(Hechizo).MinDef = val(leer.GetValue("Hechizo" & Hechizo, "MinDef"))
    End If
    
    Hechizos(Hechizo).MinSkill = val(leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
    Hechizos(Hechizo).ManaRequerido = val(leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))

    Hechizos(Hechizo).StaRequerido = val(leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
    
    Hechizos(Hechizo).Anillo = val(leer.GetValue("Hechizo" & Hechizo, "Anillo"))
    
    Hechizos(Hechizo).Target = val(leer.GetValue("Hechizo" & Hechizo, "Target"))
    
    'frmCargando.cargar.value = frmCargando.cargar.value + 1
    frmCargando.Label1(2).Caption = "Cargando hechizos " & Hechizo & "/" & NumeroHechizos
    DoEvents
Next Hechizo

Set leer = Nothing
Exit Sub

Errhandler:
 MsgBox "Error cargando hechizos.dat " & err.Number & ": " & err.description
 
End Sub

'ReAdd Marius
Sub LoadMotd()
    Dim i As Integer

    MotdMaxLines = val(GetVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines"))

    ReDim MOTD(1 To MotdMaxLines)
    For i = 1 To MotdMaxLines
        MOTD(i) = GetVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & i)
    Next i

End Sub
'\ReAdd

'Add Marius
Sub LoadPublicidad()
    Dim i As Integer

    PubMaxLines = val(GetVar(App.Path & "\Dat\publicidad.ini", "INIT", "NumLines"))

    ReDim PUBLICIDAD(1 To PubMaxLines)
    For i = 1 To PubMaxLines
        PUBLICIDAD(i) = GetVar(App.Path & "\Dat\publicidad.ini", "publicidad", "Line" & i)
    Next i

End Sub

Sub SendPublicidad()
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, PUBLICIDAD(RandomNumber(1, PubMaxLines)), FontTypeNames.FONTTYPE_GUILD))
End Sub
'\Add


Sub LoadArmasHerreria()

    Dim N As Integer, lc As Integer
    
    N = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))
    
    ReDim Preserve ArmasHerrero(1 To N) As Integer
    
    For lc = 1 To N
        ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
    Next lc

End Sub

Sub LoadArmadurasHerreria()

    Dim N As Integer, lc As Integer
    
    N = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))
    
    ReDim Preserve ArmadurasHerrero(1 To N) As Integer
    
    For lc = 1 To N
        ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
    Next lc

End Sub

Sub LoadBalance()
    Dim i As Long
    
    'Modificadores de Clase
    For i = 1 To NUMCLASES
        ModClase(i).Evasion = val(GetVar(DatPath & "Balance.dat", "MODEVASION", ListaClases(i)))
        ModClase(i).AtaqueArmas = val(GetVar(DatPath & "Balance.dat", "MODATAQUEARMAS", ListaClases(i)))
        ModClase(i).AtaqueProyectiles = val(GetVar(DatPath & "Balance.dat", "MODATAQUEPROYECTILES", ListaClases(i)))
        ModClase(i).DañoArmas = val(GetVar(DatPath & "Balance.dat", "MODDAÑOARMAS", ListaClases(i)))
        ModClase(i).DañoProyectiles = val(GetVar(DatPath & "Balance.dat", "MODDAÑOPROYECTILES", ListaClases(i)))
        ModClase(i).DañoWrestling = val(GetVar(DatPath & "Balance.dat", "MODDAÑOWRESTLING", ListaClases(i)))
        ModClase(i).Escudo = val(GetVar(DatPath & "Balance.dat", "MODESCUDO", ListaClases(i)))
    Next i
    
    'Modificadores de Raza
    For i = 1 To NUMRAZAS
        ModRaza(i).Fuerza = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Fuerza"))
        ModRaza(i).Agilidad = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Agilidad"))
        ModRaza(i).Inteligencia = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Inteligencia"))
        ModRaza(i).Carisma = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Carisma"))
        ModRaza(i).constitucion = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Constitucion"))
    Next i
    
    'Modificadores de Vida
    For i = 1 To NUMCLASES
        ModVida(i) = val(GetVar(DatPath & "Balance.dat", "MODVIDA", ListaClases(i)))
    Next i
    
    'Distribución de Vida
    For i = 1 To 5
        DistribucionEnteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "E" + CStr(i)))
    Next i
    For i = 1 To 4
        DistribucionSemienteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "S" + CStr(i)))
    Next i
    
    'Extra
    PorcentajeRecuperoMana = val(GetVar(DatPath & "Balance.dat", "EXTRA", "PorcentajeRecuperoMana"))

    'Grupo
    ExponenteNivelGrupo = val(GetVar(DatPath & "Balance.dat", "Grupo", "ExponenteNivelGrupo"))

    'Intervalos
    SanaIntervaloSinDescansar = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "SanaIntervaloSinDescansar"))
    StaminaIntervaloSinDescansar = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "StaminaIntervaloSinDescansar"))
    SanaIntervaloDescansar = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "SanaIntervaloDescansar"))
    StaminaIntervaloDescansar = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "StaminaIntervaloDescansar"))
    IntervaloSed = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloSed"))
    IntervaloHambre = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloHambre"))
    IntervaloVeneno = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloVeneno"))
    IntervaloParalizado = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloParalizado"))
    'IntervaloHechizosMagicos = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloHechizosMagicos"))
    IntervaloInvisible = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloInvisible"))
    IntervaloFrio = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloFrio"))
    IntervaloWavFx = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloWAVFX"))
    IntervaloInvocacion = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloInvocacion"))
    IntervaloParaConexion = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloParaConexion"))
    
    frmMain.TIMER_AI.Interval = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloNpcAI"))
    frmMain.npcataca.Interval = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloNpcPuedeAtacar"))
    IntervaloUserPuedeTrabajar = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloTrabajo"))

    IntervaloGolpeUsar = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloGolpeUsar"))
    
    IntervaloCerrarConexion = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloCerrarConexion"))
    IntervaloUserPuedeUsar = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloUserPuedeUsar"))
    IntervaloUserPuedeAtacar = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloUserPuedeAtacar"))
    
    IntervaloUserPuedeCastear = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloUserPuedeCastear"))
    IntervaloMagiaGolpe = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloMagiaGolpe"))
    IntervaloGolpeMagia = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloGolpeMagia"))
    IntervaloFlechasCazadores = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloFlechasCazadores"))
    IntervaloUserMove = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloMover"))

    IntervaloOculto = val(GetVar(DatPath & "Balance.dat", "INTERVALOS", "IntervaloOculto"))


End Sub

Sub LoadObjCarpintero()

Dim N As Integer, lc As Integer

    N = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))
    
    ReDim Preserve ObjCarpintero(1 To N) As Integer
    
    For lc = 1 To N
        ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
    Next lc

End Sub

Sub LoadObjDruida()

Dim N As Integer, lc As Integer

    N = val(GetVar(DatPath & "objdruida.dat", "INIT", "NumObjs"))
    
    ReDim Preserve ObjDruida(1 To N) As Integer
    
    For lc = 1 To N
        ObjDruida(lc) = val(GetVar(DatPath & "objdruida.dat", "Obj" & lc, "Index"))
    Next lc

End Sub

Sub LoadObjSastre()

Dim N As Integer, lc As Integer

    N = val(GetVar(DatPath & "objsastre.dat", "INIT", "NumObjs"))
    
    ReDim Preserve ObjSastre(1 To N) As Integer
    
    For lc = 1 To N
        ObjSastre(lc) = val(GetVar(DatPath & "objsastre.dat", "Obj" & lc, "Index"))
    Next lc

End Sub




Sub LoadOBJData()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'conmigo. Para leer desde el OBJ.DAT se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

'Call LogTarea("Sub LoadOBJData")

On Error GoTo Errhandler

frmCargando.Label1(2).Caption = "Cargando objetos...."

'*****************************************************************
'Carga la lista de objetos
'*****************************************************************
Dim Object As Integer
Dim leer As New clsIniReader

Dim consulta As String

Dim cpParams As String
Dim cpValues As String

Call leer.Initialize(DatPath & "Obj.dat")

'obtiene el numero de obj
NumObjDatas = val(leer.GetValue("INIT", "NumObjs"))

'frmCargando.cargar.min = 0
'frmCargando.cargar.max = NumObjDatas
'frmCargando.cargar.value = 0


ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
  
'Llena la lista
For Object = 1 To NumObjDatas
        
    cpParams = ""
    cpValues = ""
        
    ObjData(Object).Name = leer.GetValue("OBJ" & Object, "Name")
    
    'Pablo (ToxicWaste) Log de Objetos.
    ObjData(Object).Log = val(leer.GetValue("OBJ" & Object, "Log"))
    ObjData(Object).NoLog = val(leer.GetValue("OBJ" & Object, "NoLog"))
    '07/09/07
    
    ObjData(Object).GrhIndex = val(leer.GetValue("OBJ" & Object, "GrhIndex"))
    If ObjData(Object).GrhIndex = 0 Then
        ObjData(Object).GrhIndex = ObjData(Object).GrhIndex
    End If
    
    ObjData(Object).OBJType = val(leer.GetValue("OBJ" & Object, "ObjType"))
    
    ObjData(Object).Newbie = val(leer.GetValue("OBJ" & Object, "Newbie"))
    
    ObjData(Object).SubTipo = val(leer.GetValue("OBJ" & Object, "SubTipo"))

    ObjData(Object).EfectoMagico = val(leer.GetValue("OBJ" & Object, "efectomagico"))
    ObjData(Object).CuantoAumento = val(leer.GetValue("OBJ" & Object, "cuantoaumento"))
    
    ObjData(Object).Shop = 0
    ObjData(Object).Shop = val(leer.GetValue("OBJ" & Object, "Shop"))
    
    With ObjData(Object)
        Select Case ObjData(Object).OBJType
            Case eOBJType.otItemsMagicos
                Select Case ObjData(Object).EfectoMagico
                    Case 1, 2, 3, 6, 7, 12, 14, 18
                     '  ObjData(Object).CuantoAumento = val(leer.GetValue("OBJ" & Object, "cuantoaumento"))
                End Select
                
                If ObjData(Object).EfectoMagico = 2 Then _
                    ObjData(Object).QueAtributo = val(leer.GetValue("OBJ" & Object, "QueAtributo"))
                
                If ObjData(Object).EfectoMagico = 3 Then _
                ObjData(Object).QueSkill = val(leer.GetValue("OBJ" & Object, "QueSkill"))

            Case eOBJType.otArmadura
            
                If .SubTipo = 1 Then
                    ObjData(Object).CascoAnim = val(leer.GetValue("OBJ" & Object, "Anim"))
                ElseIf .SubTipo = 2 Then
                    ObjData(Object).ShieldAnim = val(leer.GetValue("OBJ" & Object, "Anim"))
                    ObjData(Object).DosManos = val(leer.GetValue("OBJ" & Object, "DosManos"))
                End If
                
                ObjData(Object).Real = val(leer.GetValue("OBJ" & Object, "Real"))
                ObjData(Object).Caos = val(leer.GetValue("OBJ" & Object, "Caos"))
                ObjData(Object).Milicia = val(leer.GetValue("OBJ" & Object, "Milicia"))
                
                ObjData(Object).LingH = val(leer.GetValue("OBJ" & Object, "LingH"))
                ObjData(Object).LingP = val(leer.GetValue("OBJ" & Object, "LingP"))
                ObjData(Object).LingO = val(leer.GetValue("OBJ" & Object, "LingO"))
                ObjData(Object).SkHerreria = val(leer.GetValue("OBJ" & Object, "SkHerreria"))
            
            Case eOBJType.otNudillos
                ObjData(Object).LingH = val(leer.GetValue("OBJ" & Object, "LingH"))
                ObjData(Object).LingP = val(leer.GetValue("OBJ" & Object, "LingP"))
                ObjData(Object).LingO = val(leer.GetValue("OBJ" & Object, "LingO"))
                ObjData(Object).SkHerreria = val(leer.GetValue("OBJ" & Object, "SkHerreria"))
                ObjData(Object).MaxHit = val(leer.GetValue("OBJ" & Object, "MaxHIT"))
                ObjData(Object).MinHit = val(leer.GetValue("OBJ" & Object, "MinHIT"))
                ObjData(Object).WeaponAnim = val(leer.GetValue("OBJ" & Object, "Anim"))
                
            Case eOBJType.otWeapon
                ObjData(Object).WeaponAnim = val(leer.GetValue("OBJ" & Object, "Anim"))
                ObjData(Object).Apuñala = val(leer.GetValue("OBJ" & Object, "Apuñala"))
                ObjData(Object).MaxHit = val(leer.GetValue("OBJ" & Object, "MaxHIT"))
                ObjData(Object).MinHit = val(leer.GetValue("OBJ" & Object, "MinHIT"))
                ObjData(Object).proyectil = val(leer.GetValue("OBJ" & Object, "Proyectil"))
                ObjData(Object).Municion = val(leer.GetValue("OBJ" & Object, "Municiones"))
                ObjData(Object).Refuerzo = val(leer.GetValue("OBJ" & Object, "Refuerzo"))
                
                ObjData(Object).LingH = val(leer.GetValue("OBJ" & Object, "LingH"))
                ObjData(Object).LingP = val(leer.GetValue("OBJ" & Object, "LingP"))
                ObjData(Object).LingO = val(leer.GetValue("OBJ" & Object, "LingO"))
                ObjData(Object).SkHerreria = val(leer.GetValue("OBJ" & Object, "SkHerreria"))
                ObjData(Object).Real = val(leer.GetValue("OBJ" & Object, "Real"))
                ObjData(Object).Caos = val(leer.GetValue("OBJ" & Object, "Caos"))
                ObjData(Object).DosManos = val(leer.GetValue("OBJ" & Object, "DosManos"))
            
            Case eOBJType.otInstrumentos
                ObjData(Object).Snd1 = val(leer.GetValue("OBJ" & Object, "SND1"))
                ObjData(Object).Snd2 = val(leer.GetValue("OBJ" & Object, "SND2"))
                ObjData(Object).Snd3 = val(leer.GetValue("OBJ" & Object, "SND3"))

                ObjData(Object).Real = val(leer.GetValue("OBJ" & Object, "Real"))
                ObjData(Object).Caos = val(leer.GetValue("OBJ" & Object, "Caos"))
                ObjData(Object).Milicia = val(leer.GetValue("OBJ" & Object, "Milicia"))
                
            Case eOBJType.otMinerales
                ObjData(Object).MinSkill = val(leer.GetValue("OBJ" & Object, "MinSkill"))
            
            Case eOBJType.otMonturas
                ObjData(Object).MinSkill = val(leer.GetValue("OBJ" & Object, "MinSkill"))
            
            Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
                ObjData(Object).IndexAbierta = val(leer.GetValue("OBJ" & Object, "IndexAbierta"))
                ObjData(Object).IndexCerrada = val(leer.GetValue("OBJ" & Object, "IndexCerrada"))
                ObjData(Object).IndexCerradaLlave = val(leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
            
            Case otPociones
                ObjData(Object).TipoPocion = val(leer.GetValue("OBJ" & Object, "TipoPocion"))
                ObjData(Object).MaxModificador = val(leer.GetValue("OBJ" & Object, "MaxModificador"))
                ObjData(Object).MinModificador = val(leer.GetValue("OBJ" & Object, "MinModificador"))
                ObjData(Object).DuracionEfecto = val(leer.GetValue("OBJ" & Object, "DuracionEfecto"))
            
            Case eOBJType.otBarcos
                ObjData(Object).MinSkill = val(leer.GetValue("OBJ" & Object, "MinSkill"))
                ObjData(Object).MaxHit = val(leer.GetValue("OBJ" & Object, "MaxHIT"))
                ObjData(Object).MinHit = val(leer.GetValue("OBJ" & Object, "MinHIT"))
            
            Case eOBJType.otFlechas
                ObjData(Object).MaxHit = val(leer.GetValue("OBJ" & Object, "MaxHIT"))
                ObjData(Object).MinHit = val(leer.GetValue("OBJ" & Object, "MinHIT"))
                ObjData(Object).SubTipo = val(leer.GetValue("OBJ" & Object, "SubTipo"))
    
            Case eOBJType.otPasajes
                ObjData(Object).DesdeMap = val(leer.GetValue("OBJ" & Object, "Desde"))
                ObjData(Object).HastaMap = val(leer.GetValue("OBJ" & Object, "Map"))
                ObjData(Object).HastaX = val(leer.GetValue("OBJ" & Object, "X"))
                ObjData(Object).HastaY = val(leer.GetValue("OBJ" & Object, "Y"))
                ObjData(Object).CantidadSkill = val(leer.GetValue("OBJ" & Object, "CantidadSkill"))
                
            Case eOBJType.otContenedores
                ObjData(Object).CuantoAgrega = val(leer.GetValue("OBJ" & Object, "CuantoAgrega"))
        
        End Select
    End With

    ObjData(Object).Ropaje = val(leer.GetValue("OBJ" & Object, "NumRopaje"))
    ObjData(Object).HechizoIndex = val(leer.GetValue("OBJ" & Object, "HechizoIndex"))
 
 
 
     If val(leer.GetValue("OBJ" & Object, "SICPO")) = 1 Then
        ObjData(Object).CPO = leer.GetValue("OBJ" & Object, "CPO")
    End If
    
    ObjData(Object).LingoteIndex = val(leer.GetValue("OBJ" & Object, "LingoteIndex"))
    
    ObjData(Object).MineralIndex = val(leer.GetValue("OBJ" & Object, "MineralIndex"))
    
    ObjData(Object).MaxHP = val(leer.GetValue("OBJ" & Object, "MaxHP"))
    ObjData(Object).MinHP = val(leer.GetValue("OBJ" & Object, "MinHP"))
    
    ObjData(Object).Mujer = val(leer.GetValue("OBJ" & Object, "Mujer"))
    ObjData(Object).Hombre = val(leer.GetValue("OBJ" & Object, "Hombre"))
    
    ObjData(Object).MinHAM = val(leer.GetValue("OBJ" & Object, "MinHam"))
    ObjData(Object).MinSed = val(leer.GetValue("OBJ" & Object, "MinAgu"))
    
    ObjData(Object).MinDef = val(leer.GetValue("OBJ" & Object, "MINDEF"))
    ObjData(Object).MaxDef = val(leer.GetValue("OBJ" & Object, "MAXDEF"))
    ObjData(Object).def = (ObjData(Object).MinDef + ObjData(Object).MaxDef) * 0.5
    
    ObjData(Object).RazaTipo = val(leer.GetValue("OBJ" & Object, "RazaTipo"))
    ObjData(Object).RazaEnana = val(leer.GetValue("OBJ" & Object, "RazaEnana"))
    ObjData(Object).MinELV = val(leer.GetValue("OBJ" & Object, "MinELV"))
    
    ObjData(Object).valor = val(leer.GetValue("OBJ" & Object, "Valor"))
    
    ObjData(Object).Crucial = val(leer.GetValue("OBJ" & Object, "Crucial"))
    
    ObjData(Object).Cerrada = val(leer.GetValue("OBJ" & Object, "abierta"))
    If ObjData(Object).Cerrada = 1 Then
        ObjData(Object).Llave = val(leer.GetValue("OBJ" & Object, "Llave"))
        ObjData(Object).clave = val(leer.GetValue("OBJ" & Object, "Clave"))
    End If
    
    'Puertas y llaves
    ObjData(Object).clave = val(leer.GetValue("OBJ" & Object, "Clave"))
    
    ObjData(Object).texto = leer.GetValue("OBJ" & Object, "Texto")
    ObjData(Object).GrhSecundario = val(leer.GetValue("OBJ" & Object, "VGrande"))
    
    ObjData(Object).Agarrable = val(leer.GetValue("OBJ" & Object, "Agarrable"))
    ObjData(Object).ForoID = leer.GetValue("OBJ" & Object, "ID")
    
    'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico
    Dim i As Integer
    Dim N As Integer
    Dim S As String
    i = 1: N = 1
    S = leer.GetValue("OBJ" & Object, "CP" & i)
    Do While Len(S) > 0
        If ClaseToEnum(S) > 0 Then ObjData(Object).ClaseProhibida(N) = ClaseToEnum(S)
        
        cpParams = cpParams + "`cp" + "" & i & "" + "`,"
        cpValues = cpValues + "'" + S + "',"
        
        If N = NUMCLASES Then Exit Do
        
        N = N + 1: i = i + 1
        S = leer.GetValue("OBJ" & Object, "CP" & i)

    Loop
    
    
    ObjData(Object).ClaseTipo = val(leer.GetValue("OBJ" & Object, "ClaseTipo"))
    
    ObjData(Object).DefensaMagicaMax = val(leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
    ObjData(Object).DefensaMagicaMin = val(leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
    


    
    ' Jose Castelli  / Resistencia Magica (RM)
    
    ObjData(Object).ResistenciaMagica = val(leer.GetValue("OBJ" & Object, "ResistenciaMagica"))
    
    ' Jose Castelli  / Resistencia Magica (RM)
    
    ObjData(Object).SkCarpinteria = val(leer.GetValue("OBJ" & Object, "SkCarpinteria"))
    
    If ObjData(Object).SkCarpinteria > 0 Then
        ObjData(Object).Madera = val(leer.GetValue("OBJ" & Object, "Madera"))
    End If
    
    
    ObjData(Object).SkPociones = val(leer.GetValue("OBJ" & Object, "SkPociones"))
    
    If ObjData(Object).SkPociones > 0 Then
        ObjData(Object).raies = val(leer.GetValue("OBJ" & Object, "Raices"))
    End If
    
    
    ObjData(Object).SkSastreria = val(leer.GetValue("OBJ" & Object, "SkSastreria"))
    
    If ObjData(Object).SkSastreria > 0 Then
        ObjData(Object).PielLobo = val(leer.GetValue("OBJ" & Object, "PielLobo"))
        ObjData(Object).PielLoboInvernal = val(leer.GetValue("OBJ" & Object, "PielOsoPolar"))
        ObjData(Object).PielOso = val(leer.GetValue("OBJ" & Object, "PielOsoPardo"))
    End If
    
    frmCargando.Label1(2).Caption = "Cargando objetos " & Object & "/" & NumObjDatas
    'frmCargando.cargar.value = frmCargando.cargar.value + 1
    
'Dim anim As Integer

'anim = val(leer.GetValue("OBJ" & Object, "Anim"))
    
'With ObjData(Object)


'consulta = "INSERT INTO `inmortalao`.`objs`(`name`,`grhIndex`,`objType`,`newbie`,`subTipo`,`efectoMagico`,`cuantoAumento`," & _
"`shop`,`queAtributo`,`queSkill`,`anim`,`dosManos`,`real`,`caos`,`milicia`,`lingH`,`lingP`,`lingO`,`skHerreria`,`maxHit`," & _
"`minHit`,`proyectil`,`municiones`,`refuerzo`,`snd1`,`snd2`,`snd3`,`minSkill`,`indexAbierta`,`indexCerrada`,`indexCerradaLlave`," & _
"`tipoPocion`,`maxModificador`,`minModificador`,`duracionEfecto`,`desde`,`map`,`x`,`y`,`cantidadSkill`,`cuantoAgrega`," & _
"`numRopaje`,`hechizoIndex`,`cpo`,`lingoteIndex`,`mineralIndex`,`maxHp`,`minHp`,`mujer`,`hombre`,`minHam`,`minAgu`," & _
"`minDef`,`maxDef`,`razaTipo`,`razaEnana`,`minElv`,`valor`,`crucial`,`abierta`,`texto`,`grhSecundario`,`agarrable`," & _
"`foroId`," + cpParams + "" & _
"`claseTipo`,`resistenciaMagica`,`skCarpinteria`,`madera`,`skpociones`,`raices`,`skSastreria`,`pielLobo`," & _
"`pielOsoPilar`,`pielOsoPardo`) values " & _
"('" + .name + "'," & .GrhIndex & "," & .OBJType & "," & .Newbie & "," & .SubTipo & "," & .EfectoMagico & "," & _
"" & .CuantoAumento & "," & .Shop & "," & .QueAtributo & "," & .QueSkill & "," & anim & "," & .DosManos & "," & .Real & "," & .Caos & "," & .Milicia & "," & .LingH & "," & _
"" & .LingP & "," & .LingO & "," & .SkHerreria & "," & .maxhit & "," & .minhit & "," & .proyectil & "," & .Municion & "," & .Refuerzo & "," & .snd1 & "," & .snd2 & "," & .snd3 & "," & .MinSkill & "," & .IndexAbierta & "," & .IndexCerrada & "," & .IndexCerradaLlave & "," & .TipoPocion & "," & _
"" & .MaxModificador & "," & .MinModificador & "," & .DuracionEfecto & "," & .DesdeMap & "," & .HastaMap & "," & .HastaX & "," & .HastaY & "," & .CantidadSkill & "," & .CuantoAgrega & "," & .Ropaje & "," & .HechizoIndex & ",'" + .CPO + "'," & .LingoteIndex & "," & .MineralIndex & "," & .maxhp & "," & .minhp & "," & _
"" & .Mujer & "," & .Hombre & "," & .MinHAM & "," & .MinSed & "," & .MinDef & "," & .MaxDef & "," & .RazaTipo & "," & .RazaEnana & "," & .MinELV & "," & .valor & "," & .Crucial & "," & .Cerrada & ",'" + .texto + "'," & .GrhSecundario & "," & .Agarrable & "," & _
"'" + .ForoID + "'," & _
"" + cpValues + "" & .ClaseTipo & "," & .ResistenciaMagica & "," & .SkCarpinteria & "," & _
"" & .Madera & "," & .SkPociones & "," & .raies & "," & .SkSastreria & "," & .PielLobo & "," & .PielLoboInvernal & "," & .PielOso & ");"

'End With
    
 '   DB_Conn.Execute (consulta)
    
    DoEvents
    
Next Object

Set leer = Nothing

Exit Sub

Errhandler:
    
    MsgBox "error cargando objetos " & err.Number & ": " & err.description


End Sub


Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
    szReturn = vbNullString
      
    sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
      
      
    GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, file
      
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function



Sub LoadMapData()

Dim map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

On Error GoTo man
    
    'NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps")) + 4
    NumMaps = 851
    Call InitAreas
    
  '  frmCargando.cargar.min = 0
  '  frmCargando.cargar.max = NumMaps
  '  frmCargando.cargar.value = 0
    
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo

    For map = 1 To NumMaps
        
        tFileName = App.Path & MapPath & "Mapa" & map
        Call CargarMapa(map, tFileName)
        
        'Add Marius Agregamos mas mapas seguros
        'Se agrego: intermundia (por decision de los users), lindos izquierda, lindos abajo, suramei abajo y bander derecha, tiama puerto
        If map = Banderbill.map Or map = Arghal.map Or _
           map = Rinkel.map Or map = Lindos.map Or _
           map = Suramei.map Or map = Illiandor.map Or _
           map = Orac.map Or map = Tiama.map Or _
           map = Ullathorpe.map Or map = Nix.map Or _
           map = NuevaEsperanza.map Or _
           map = NuevaEsperanzapuerto.map Or _
           map = 49 Or map = 364 Or map = 64 Or map = 63 Or map = 184 Or map = 217 Then
           
           MapInfo(map).Pk = False
        End If
        
        frmCargando.Label1(2).Caption = "Cargando Mapas " & map & "\" & NumMaps
        
        'frmCargando.cargar.value = frmCargando.cargar.value + 1
        DoEvents
    Next map

    'Add Marius Agregamos mas arenas =P Algun día tendremos WorldEdit
    Call CargarMapa(238, tFileName)
    Call CargarMapa(238, tFileName)
    Call CargarMapa(238, tFileName)
    Call CargarMapa(238, tFileName)
    '\Add
    
Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & map & " contiene errores")
    Call LogError(Date & " " & err.description & " " & err.HelpContext & " " & err.HelpFile & " " & err.source)

End Sub
Public Sub CargarMapa(ByVal map As Long, ByVal MAPFl As String)
On Error GoTo errh

Dim fh As Integer
Dim MH As tMapHeader
Dim Blqs() As tDatosBloqueados
Dim L1() As Long
Dim L2() As tDatosGrh
Dim L3() As tDatosGrh
Dim L4() As tDatosGrh
Dim Triggers() As tDatosTrigger
Dim Luces() As tDatosLuces
Dim Particulas() As tDatosParticulas
Dim Objetos() As tDatosObjs
Dim NPCs() As tDatosNPC
Dim TEs() As tDatosTE
Dim MapSize As tMapSize
Dim MapDat As tMapDat

Dim i As Long
Dim j As Long

    If Not FileExist(MAPFl & ".csm", vbNormal) Then _
        Exit Sub

    fh = FreeFile
    
    Dim fTxt As Integer
    fTxt = FreeFile
    
    
    Open MAPFl & ".csm" For Binary Access Read As fh
        Get #fh, , MH
        Get #fh, , MapSize
        Get #fh, , MapDat
        
        ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As Long
        
        Get #fh, , L1
        
        With MH
            If .NumeroBloqueados > 0 Then
                ReDim Blqs(1 To .NumeroBloqueados)
                Get #fh, , Blqs
            End If
            
            If .NumeroLayers(2) > 0 Then
                ReDim L2(1 To .NumeroLayers(2))
                Get #fh, , L2
            End If
            
            If .NumeroLayers(3) > 0 Then
                ReDim L3(1 To .NumeroLayers(3))
                Get #fh, , L3
            End If
            
            If .NumeroLayers(4) > 0 Then
                ReDim L4(1 To .NumeroLayers(4))
                Get #fh, , L4
            End If
            
            If .NumeroTriggers > 0 Then
                ReDim Triggers(1 To .NumeroTriggers)
                Get #fh, , Triggers
            End If
            
            If .NumeroParticulas > 0 Then
                ReDim Particulas(1 To .NumeroParticulas)
                Get #fh, , Particulas
            End If
            
            If .NumeroLuces > 0 Then
                ReDim Luces(1 To .NumeroLuces)
                Get #fh, , Luces
            End If
            
            If .NumeroOBJs > 0 Then
                ReDim Objetos(1 To .NumeroOBJs)
                Get #fh, , Objetos
            End If
                
            If .NumeroNPCs > 0 Then
                ReDim NPCs(1 To .NumeroNPCs)
                Get #fh, , NPCs
            End If
                
            If .NumeroTE > 0 Then
                ReDim TEs(1 To .NumeroTE)
                Get #fh, , TEs
            End If

        End With
    
    Close fh
    
    
    Open MAPFl & ".txt" For Output As fTxt
        Write #fTxt, MH.NumeroBloqueados & ";" & MH.NumeroLayers(2) & ";" & MH.NumeroLayers(3) & ";" & MH.NumeroLayers(4) & ";" & MH.NumeroLuces & ";" & MH.NumeroNPCs & ";" & MH.NumeroOBJs & ";" & MH.NumeroParticulas & ";" & MH.NumeroTE & ";" & MH.NumeroTriggers
        Write #fTxt, MapSize.XMax & ";" & MapSize.XMin & ";" & MapSize.YMax & ";" & MapSize.YMin
        Write #fTxt, MapDat.ambient & ";" & MapDat.backup_mode & ";" & MapDat.base_light & ";" & MapDat.battle_mode & ";" & MapDat.extra1 & ";" & MapDat.extra2 & ";" & MapDat.extra3 & ";" & MapDat.letter_grh & ";" & MapDat.map_name & ";" & MapDat.music_number & ";" & MapDat.restrict_mode & ";" & MapDat.terrain & ";" & MapDat.zone
        
        Dim x As Integer
        Dim Y As Integer
        
        For Y = MapSize.YMin To MapSize.YMax
            For x = MapSize.XMin To MapSize.XMax
                Write #fTxt, L1(x, Y)
            Next x
        Next Y
        
        With MH
        
            Dim vf As Long
            If .NumeroBloqueados > 0 Then
                For vf = 1 To .NumeroBloqueados
                    Write #fTxt, Blqs(vf).x & ";" & Blqs(vf).Y
                Next vf
            End If
            
            If .NumeroLayers(2) > 0 Then
                For vf = 1 To .NumeroLayers(2)
                    Write #fTxt, L2(vf).GrhIndex & ";" & L2(vf).x & ";" & L2(vf).Y
                Next vf
            End If
            
            If .NumeroLayers(3) > 0 Then
                For vf = 1 To .NumeroLayers(3)
                    Write #fTxt, L3(vf).GrhIndex & ";" & L3(vf).x & ";" & L3(vf).Y
                Next vf
            End If
            
            If .NumeroLayers(4) > 0 Then
                For vf = 1 To .NumeroLayers(4)
                    Write #fTxt, L4(vf).GrhIndex & ";" & L4(vf).x & ";" & L4(vf).Y
                Next vf
            End If
            
            If .NumeroTriggers > 0 Then
                For vf = 1 To .NumeroTriggers
                    Write #fTxt, Triggers(vf).Trigger & ";" & Triggers(vf).x & ";" & Triggers(vf).Y
                Next vf
            End If
            
            If .NumeroParticulas > 0 Then
                For vf = 1 To .NumeroParticulas
                    Write #fTxt, Particulas(vf).Particula & ";" & Particulas(vf).x & ";" & Particulas(vf).Y
                Next vf
            End If
            
            If .NumeroLuces > 0 Then
                For vf = 1 To .NumeroLuces
                    Write #fTxt, Luces(vf).color & ";" & Luces(vf).Rango & ";" & Luces(vf).x & ";" & Luces(vf).Y
                Next vf
            End If
            
            If .NumeroOBJs > 0 Then
                For vf = 1 To .NumeroOBJs
                    Write #fTxt, Objetos(vf).ObjAmmount & ";" & Objetos(vf).ObjIndex & ";" & Objetos(vf).x & ";" & Objetos(vf).Y
                Next vf
            End If
                
            If .NumeroNPCs > 0 Then
                For vf = 1 To .NumeroNPCs
                    Write #fTxt, NPCs(vf).NpcIndex & ";" & NPCs(vf).x & ";" & NPCs(vf).Y
                Next vf
            End If
                
            If .NumeroTE > 0 Then
                For vf = 1 To .NumeroTE
                    Write #fTxt, TEs(vf).DestM & ";" & TEs(vf).DestX & ";" & TEs(vf).DestY & ";" & TEs(vf).x & ";" & TEs(vf).Y
                Next vf
            End If
        
        End With
        
        
    Close fTxt
    
    
    With MH
        If .NumeroBloqueados > 0 Then
            For i = 1 To .NumeroBloqueados
                MapData(map, Blqs(i).x, Blqs(i).Y).Blocked = 1
            Next i
        End If
        
        If .NumeroLayers(2) > 0 Then
            For i = 1 To .NumeroLayers(2)
                MapData(map, L2(i).x, L2(i).Y).Graphic(2) = L2(i).GrhIndex
            Next i
        End If
        
        If .NumeroLayers(3) > 0 Then
            For i = 1 To .NumeroLayers(3)
                MapData(map, L3(i).x, L3(i).Y).Graphic(3) = L3(i).GrhIndex
            Next i
        End If
        
        If .NumeroLayers(4) > 0 Then
            For i = 1 To .NumeroLayers(4)
                MapData(map, L4(i).x, L4(i).Y).Graphic(4) = L4(i).GrhIndex
            Next i
        End If
        
        If .NumeroTriggers > 0 Then
            For i = 1 To .NumeroTriggers
                MapData(map, Triggers(i).x, Triggers(i).Y).Trigger = Triggers(i).Trigger
            Next i
        End If
        
        If .NumeroOBJs > 0 Then
            For i = 1 To .NumeroOBJs
                MapData(map, Objetos(i).x, Objetos(i).Y).ObjInfo.ObjIndex = Objetos(i).ObjIndex
                MapData(map, Objetos(i).x, Objetos(i).Y).ObjInfo.Amount = Objetos(i).ObjAmmount
                
                If ObjData(Objetos(i).ObjIndex).OBJType <> eOBJType.otPuertas Then
                    MapData(map, Objetos(i).x, Objetos(i).Y).ObjEsFijo = 1
                End If
            Next i
        End If
            
        If .NumeroNPCs > 0 Then
            For i = 1 To .NumeroNPCs
                MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex = NPCs(i).NpcIndex
                If MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex > 0 Then
                    Dim npcfile As String
                    
                    npcfile = DatPath & "NPCs.dat"
    
                    If val(GetVar(npcfile, "NPC" & MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex, "PosOrig")) = 1 Then
                        MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex = OpenNPC(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex)
                        Npclist(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex).Orig.map = map
                        Npclist(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex).Orig.x = NPCs(i).x
                        Npclist(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex).Orig.Y = NPCs(i).Y
                    Else
                        MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex = OpenNPC(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex)
                    End If
                    If Not MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex = 0 Then
                        Npclist(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex).Pos.map = map
                        Npclist(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex).Pos.x = NPCs(i).x
                        Npclist(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex).Pos.Y = NPCs(i).Y
                        
                        Npclist(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex).StartPos.map = map
                        Npclist(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex).StartPos.x = NPCs(i).x
                        Npclist(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex).StartPos.Y = NPCs(i).Y
                        
                        Call MakeNPCChar(True, 0, MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex, map, NPCs(i).x, NPCs(i).Y)
                    End If
                End If
            Next i
        End If
            
        If .NumeroTE > 0 Then
            For i = 1 To .NumeroTE
                MapData(map, TEs(i).x, TEs(i).Y).TileExit.map = TEs(i).DestM
                MapData(map, TEs(i).x, TEs(i).Y).TileExit.x = TEs(i).DestX
                MapData(map, TEs(i).x, TEs(i).Y).TileExit.Y = TEs(i).DestY
            Next i
        End If
    End With
    
    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax
            If L1(i, j) > 0 Then
                MapData(map, i, j).Graphic(1) = L1(i, j)
            End If
        Next i
    Next j
    
    MapDat.map_name = Trim$(MapDat.map_name)
    
    MapInfo(map).Name = MapDat.map_name
    MapInfo(map).Music = MapDat.music_number
    MapInfo(map).Seguro = MapDat.extra1
    
    If Not (Left$(MapDat.zone, 6) = "CIUDAD") Then
        MapInfo(map).Pk = True
    Else
        MapInfo(map).Pk = False
    End If
    
    MapInfo(map).Terreno = MapDat.terrain
    MapInfo(map).Zona = Trim$(MapDat.zone)
    MapInfo(map).Restringir = MapDat.restrict_mode
    MapInfo(map).BackUp = MapDat.backup_mode

    Exit Sub

errh:
    Call LogError("Error cargando mapa: " & map & " ." & err.description)
End Sub

Sub LoadSini()

Dim Temporal As Long
    
    'ReAdd Marius
    Call LoadMotd
    '\ReAdd
    
    'Add Marius
    Call LoadPublicidad
    '\Add
    
    frmCargando.Label1(2).Caption = "Cargando server.ini"
      
    Puerto = val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
    RondasAutomatico = val(GetVar(IniPath & "Server.ini", "INIT", "RondasAutomatico"))
    ModExpX = val(GetVar(IniPath & "Server.ini", "INIT", "ModExpX"))
    ModOroX = val(GetVar(IniPath & "Server.ini", "INIT", "ModOroX"))
    
    PuedenFundarClan = val(GetVar(IniPath & "Server.ini", "INIT", "PuedenFundarClan"))
    PuedeBorrarClan = val(GetVar(IniPath & "Server.ini", "INIT", "PuedeBorrarClan"))
    
    ModTrabajo = val(GetVar(IniPath & "Server.ini", "INIT", "ModTrabajo"))
    ModSkill = val(GetVar(IniPath & "Server.ini", "INIT", "ModSkill"))
    
    'AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
    'IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))
    
    'Lee la version correcta del cliente
    ULTIMAVERSION = GetVar(IniPath & "Server.ini", "INIT", "Version")
    PuedeCrearPersonajes = val(GetVar(IniPath & "Server.ini", "INIT", "PuedeCrearPersonajes"))
    ServerSoloGMs = val(GetVar(IniPath & "Server.ini", "init", "ServerSoloGMs"))
    
    RecordUsuarios = val(GetVar(IniPath & "Server.ini", "INIT", "Record"))
      
    'Max users
    Temporal = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))
    If MaxUsers = 0 Then
        MaxUsers = Temporal
        ReDim UserList(1 To MaxUsers) As User
    End If
    
    '&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    'Se agregó en LoadBalance y en el Balance.dat
    'PorcentajeRecuperoMana = val(GetVar(IniPath & "Server.ini", "BALANCE", "PorcentajeRecuperoMana"))
    
    ''&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    
    
    Nix.map = GetVar(DatPath & "Ciudades.dat", "NIX", "Mapa")
    Nix.x = GetVar(DatPath & "Ciudades.dat", "NIX", "X")
    Nix.Y = GetVar(DatPath & "Ciudades.dat", "NIX", "Y")
    
    Ullathorpe.map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
    Ullathorpe.x = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
    Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")
    
    Banderbill.map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
    Banderbill.x = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
    Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")
    
    Arghal.map = GetVar(DatPath & "Ciudades.dat", "Arghal", "Mapa")
    Arghal.x = GetVar(DatPath & "Ciudades.dat", "Arghal", "X")
    Arghal.Y = GetVar(DatPath & "Ciudades.dat", "Arghal", "Y")
    
    Illiandor.map = GetVar(DatPath & "Ciudades.dat", "Illiandor", "Mapa")
    Illiandor.x = GetVar(DatPath & "Ciudades.dat", "Illiandor", "X")
    Illiandor.Y = GetVar(DatPath & "Ciudades.dat", "Illiandor", "Y")
    
    Suramei.map = GetVar(DatPath & "Ciudades.dat", "Suramei", "Mapa")
    Suramei.x = GetVar(DatPath & "Ciudades.dat", "Suramei", "X")
    Suramei.Y = GetVar(DatPath & "Ciudades.dat", "Suramei", "Y")
    
    Lindos.map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
    Lindos.x = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
    Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")
    
    Orac.map = GetVar(DatPath & "Ciudades.dat", "Orac", "Mapa")
    Orac.x = GetVar(DatPath & "Ciudades.dat", "Orac", "X")
    Orac.Y = GetVar(DatPath & "Ciudades.dat", "Orac", "Y")
    
    Rinkel.map = GetVar(DatPath & "Ciudades.dat", "Rinkel", "Mapa")
    Rinkel.x = GetVar(DatPath & "Ciudades.dat", "Rinkel", "X")
    Rinkel.Y = GetVar(DatPath & "Ciudades.dat", "Rinkel", "Y")
    
    Tiama.map = GetVar(DatPath & "Ciudades.dat", "Tiama", "Mapa")
    Tiama.x = GetVar(DatPath & "Ciudades.dat", "Tiama", "X")
    Tiama.Y = GetVar(DatPath & "Ciudades.dat", "Tiama", "Y")
    
    NuevaEsperanza.map = GetVar(DatPath & "Ciudades.dat", "NuevaEsperanza", "Mapa")
    NuevaEsperanza.x = GetVar(DatPath & "Ciudades.dat", "NuevaEsperanza", "X")
    NuevaEsperanza.Y = GetVar(DatPath & "Ciudades.dat", "NuevaEsperanza", "Y")
    
    NuevaEsperanzapuerto.map = GetVar(DatPath & "Ciudades.dat", "Nuevaesperanzapuerto", "Mapa")
    NuevaEsperanzapuerto.x = GetVar(DatPath & "Ciudades.dat", "Nuevaesperanzapuerto", "X")
    NuevaEsperanzapuerto.Y = GetVar(DatPath & "Ciudades.dat", "Nuevaesperanzapuerto", "Y")

End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************

    writeprivateprofilestring Main, Var, value, file
    
End Sub



Sub BackUPnPc(NpcIndex As Integer)

Dim NpcNumero As Integer
Dim npcfile As String
Dim loopC As Integer


    NpcNumero = Npclist(NpcIndex).Numero
    
    npcfile = DatPath & "bkNPCs.dat"
    
    'General
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", Npclist(NpcIndex).Name)
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", Npclist(NpcIndex).desc)
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(Npclist(NpcIndex).Char.Head))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(Npclist(NpcIndex).Char.body))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(Npclist(NpcIndex).Char.heading))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(Npclist(NpcIndex).Movement))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(Npclist(NpcIndex).Attackable))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(Npclist(NpcIndex).Comercia))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(Npclist(NpcIndex).TipoItems))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(Npclist(NpcIndex).GiveEXP))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", val(Npclist(NpcIndex).GiveGLD))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(Npclist(NpcIndex).InvReSpawn))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(Npclist(NpcIndex).NPCtype))
    
    
    'Stats
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(Npclist(NpcIndex).Stats.Alineacion))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.def))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(Npclist(NpcIndex).Stats.MaxHit))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(Npclist(NpcIndex).Stats.MaxHP))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(Npclist(NpcIndex).Stats.MinHit))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(Npclist(NpcIndex).Stats.MinHP))
    
    
    
    
    'Flags
    Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(Npclist(NpcIndex).flags.respawn))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(Npclist(NpcIndex).flags.BackUp))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(Npclist(NpcIndex).flags.Domable))
    
    'Inventario
    Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(Npclist(NpcIndex).Invent.NroItems))
    If Npclist(NpcIndex).Invent.NroItems > 0 Then
       For loopC = 1 To MAX_INVENTORY_SLOTS
            Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & loopC, Npclist(NpcIndex).Invent.Object(loopC).ObjIndex & "-" & Npclist(NpcIndex).Invent.Object(loopC).Amount)
       Next
    End If


End Sub




Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal motivo As String)

    Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "Reason", motivo)
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
        Print #mifile, UserList(BannedIndex).Name
    Close #mifile

End Sub


Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal motivo As String)

    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).Name)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile

End Sub


Sub ban(ByVal BannedName As String, ByVal Baneador As String, ByVal motivo As String)

    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)
    
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile

End Sub

Public Sub CargaApuestas()

    Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

End Sub

