Attribute VB_Name = "modDeclaraciones"
'************************************************************************
'************************************************************************
'Inmortal AO v0.1 Beta
'Type object: Module
'Names: modDeclaraciones.bas
'Description: The all declaration is on this module
'Author: Jose Ignacio Castelli
'************************************************************************
'************************************************************************

Option Explicit

Public Const iFragataFantasmal = 87
Public Const iBarca = 84
Public Const iGalera = 85
Public Const iGaleon = 86

Public serverList As String

Public Security As New clsSecurity
Public perm As Boolean
Public ALaMierda As Boolean
Public thFPSAndHour As Long

Public FramesPerSecCounter As Integer
Public FPS As Integer

Public Const CANT_GRH_INDEX As Long = 32016

Public CurServerIp As String
Public CurServerPort As Integer

Public mueve As Byte 'Drag and Drop ¿Only in 800x600?

Public Type tCurrentUser
    SendingType As Byte
    sndPrivateTo As String
End Type
Public CurrentUser As tCurrentUser

Public IsLeader As Boolean
Public GrupoIndex As Integer


Public MIDI_ACTIVATE As Byte
    
'Meteorologia
Public tHora As Byte
Public tMinuto As Byte
Public tSeg As Byte
Public tCartel As Integer

'Public TileEngine As clsTileEngineX
Public Audio As clsAudio
Public Inventario As clsGrapchicalInventory
Public CustomKeys As clsCustomKeys
Public MainTimer As clsTimer
Public incomingData As clsByteQueue
Public outgoingData As clsByteQueue
Public UserIndex As Integer

Public BloqMov As Boolean
Public BloqDir As E_Heading

'Sonidos
Public Const SND_CLICK As Integer = 190
Public Const SND_NAVEGANDO As Integer = 50
Public Const SND_OVER As Integer = 0
Public Const SND_DICE As Integer = 188
Public Const SND_FUEGO As Integer = 79

Public Const SND_LLUVIAIN As Integer = 191
Public Const SND_LLUVIAOUT As Integer = 194
Public Const SND_AMBIENTE_NOCHE As Integer = 107
Public Const SND_AMBIENTE_NOCHE_CIU As Integer = 141

Public Const SND_NIEVEIN As Integer = 191
Public Const SND_NIEVEOUT As Integer = 194

Public Const SND_RESUCITAR As Integer = 104
Public Const SND_CURAR As Integer = 240
Public Const SND_WARHORN As Integer = 252
Public Const SND_RETIRARORO As Integer = 172
Public Const SND_MENSAJE As Integer = 227
Public Const SND_BODA As Integer = 161


' Head index of the casper. Used to know if a char is killed

' Constantes de intervalo
Public Const INT_MACRO_HECHIS As Integer = 2788

Public Const INT_ATTACK As Integer = 1500
Public Const INT_ARROWS As Integer = 1400
Public Const INT_CAST_SPELL As Integer = 1000
Public Const INT_CAST_ATTACK As Integer = 1000
Public Const INT_WORK As Integer = 1000
Public Const INT_USEITEMU As Integer = 270
Public Const INT_USEITEMDCK As Integer = 125
Public Const INT_SENTRPU As Integer = 2000

Public Const CASPER_HEAD As Integer = 500
Public Const FRAGATA_FANTASMAL As Integer = 87

Public Const NUMATRIBUTES As Byte = 5

Public RenderInv As Boolean
Public Default_RGB(0 To 3) As Long

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String

Public UserCiego As Boolean
Public UserEstupido As Boolean

Public FogataBufferIndex As Long

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

Public Const PrimerBodyBarco = 84
Public Const UltimoBodyBarco = 87

Public NumEscudosAnims As Integer

Public ArmasHerrero(1 To 22) As Integer
Public ArmadurasHerrero(1 To 45) As Integer
Public ObjCarpintero(1 To 22) As Integer
Public ObjSastre(1 To 9) As Integer
Public ObjAlquimia(1 To 9) As Integer

Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory

'Direcciones
Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tIndiceCuerpo
    body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type

'Objetos
Public Const MAX_INVENTORY_OBJS As Integer = 10000
Public Const MAX_INVENTORY_SLOTS As Byte = 25
Public Const MAX_NPC_INVENTORY_SLOTS As Byte = 50
Public Const MAXHECHI As Byte = 35

Public Const MAXSKILLPOINTS As Byte = 100

Public Const FLAGORO As Integer = MAX_INVENTORY_SLOTS + 1

Public Const Fogata As Integer = 1521

Enum eMascota
    fuego = 1
    Tierra
    Agua
    Ely
    Fatuo
    Tigre
    Lobo
    Oso
    Ent
End Enum

Type tFamiliar
    TieneFamiliar As Byte
    nombre As String
    ELV As Byte
    MinHP As Integer
    MaxHP As Integer
    ELU As Long
    EXP As Long
    MinHit As Integer
    MaxHit As Integer
    Habilidad As String
    tipo As String
End Type

Public UserPet As tFamiliar
Public PetPercExp As Long

Public Enum eClass
    Clerigo = 1
    Mago = 2
    Guerrero = 3
    Asesino = 4
    Ladron = 5
    Bardo = 6
    Druida = 7
    Gladiador = 8
    Paladin = 9
    Cazador = 10
    Pescador = 11
    Herrero = 12
    Leñador = 13
    Minero = 14
    Carpintero = 15
    Sastre = 16
    Mercenario = 17
    Nigromante = 18
End Enum

Public Enum eCiudad
    cUllathorpe = 1 'Imperial
    cNix            'Imperial
    cBanderbill     'Imperial
    cArghal         'Imperial
    
    cIlliandor      'Republicana
    cLindos         'Republicana
    cSuramei        'Republicana
    
    cOrac           'Coatica
                        
    cNuevaEsperanza 'Neutral
    cRinkel         'Neutral
    cTiama          'Neutral
End Enum

Enum eRaza
    HUMANO = 1
    ELFO
    ElfoOscuro
    Gnomo
    Enano
    Orco
End Enum

Public Enum eSkill
    Tacticas = 1
    Armas = 2
    Artes = 3
    Apuñalar = 4
    Arrojadizas = 5
    Proyectiles = 6
    DefensaEscudos = 7
    Magia = 8
    Resistencia = 9
    Meditar = 10
    Ocultarse = 11
    Domar = 12
    Musica = 13
    Robar = 14
    Comercio = 15
    Supervivencia = 16
    Liderazgo = 17
    Pesca = 18
    Mineria = 19
    Talar = 20
    Botanica = 21
    Herreria = 22
    Carpinteria = 23
    Alquimia = 24
    Sastreria = 25
    Navegacion = 26
    Equitacion = 27
End Enum

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Enum eGenero
    Hombre = 1
    Mujer
End Enum

Public Enum PlayerType
    User = &H1
    Conse = &H2
    Semi = &H3
    Dios = &H4
    Admin = &H5
    VIP = &H6
End Enum

Public Enum eObjType
    otUseOnce = 1           '  1 Comidas
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otLibros = 12
    otBebidas = 13
    otLeña = 14
    otFuego = 15
    otESCUDO = 16
    otCASCO = 17
    otHerramientas = 18
    otTeleport = 19
    otMuebles = 20
    otItemsMagicos = 21
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otLingotes = 29
    otPieles = 30
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManchas = 35          'No se usa
    otPasajes = 36
    otMapas = 38
    otBolsas = 39 ' Blosas de Oro  (contienen más de 10k de oro)
    otPozos = 40 'Pozos Mágicos
    otEsposas = 41
    otRaíces = 42
    otCadáveres = 43
    otMonturas = 44
    otPuestos = 45 ' Puestos de Entrenamiento
    otNudillos = 46
    otAnillos = 47
    otCorreo = 48
    otCualquiera = 1000
End Enum

Type tHeadRange
    mStart As Integer
    mEnd As Integer
    fStart As Integer
    fEnd As Integer
End Type
Public Head_Range() As tHeadRange

Public Const FundirMetal As Integer = 88

Type tListaFamiliares
    name As String
    Desc As String
    Imagen As String
    
    tipe As eMascota
End Type

Public ListaFamiliares() As tListaFamiliares

Public Const MENSAJE_CRIATURA_FALLA_GOLPE As String = "¡¡¡La criatura fallo el golpe!!!"
Public Const MENSAJE_CRIATURA_MATADO As String = "¡¡¡La criatura te ha matado!!!"
Public Const MENSAJE_RECHAZO_ATAQUE_ESCUDO As String = "¡¡¡Has rechazado el ataque con el escudo!!!"
Public Const MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO  As String = "¡¡¡El usuario rechazo el ataque con su escudo!!!"
Public Const MENSAJE_FALLADO_GOLPE As String = "¡¡¡Has fallado el golpe!!!"
Public Const MENSAJE_SEGURO_ACTIVADO As String = "Seguro activado"
Public Const MENSAJE_SEGURO_DESACTIVADO As String = "Seguro desactivado"
Public Const MENSAJE_PIERDE_NOBLEZA As String = "¡¡Has perdido puntaje de nobleza y ganado puntaje de criminalidad!! Si sigues ayudando a criminales te convertirás en uno de ellos y serás perseguido por las tropas de las ciudades."
Public Const MENSAJE_USAR_MEDITANDO As String = "¡Estás meditando! Debes dejar de meditar para usar objetos."

Public Const MENSAJE_GOLPE_CABEZA As String = "¡¡La criatura te ha pegado en la cabeza por "
Public Const MENSAJE_GOLPE_BRAZO_IZQ As String = "¡¡La criatura te ha pegado el brazo izquierdo por "
Public Const MENSAJE_GOLPE_BRAZO_DER As String = "¡¡La criatura te ha pegado el brazo derecho por "
Public Const MENSAJE_GOLPE_PIERNA_IZQ As String = "¡¡La criatura te ha pegado la pierna izquierda por "
Public Const MENSAJE_GOLPE_PIERNA_DER As String = "¡¡La criatura te ha pegado la pierna derecha por "
Public Const MENSAJE_GOLPE_TORSO  As String = "¡¡La criatura te ha pegado en el torso por "

Public Const MENSAJE_1 As String = "¡¡"
Public Const MENSAJE_2 As String = "!!"

Public Const MENSAJE_GOLPE_CRIATURA_1 As String = "¡¡Le has pegado a la criatura por "

Public Const MENSAJE_ATAQUE_FALLO As String = " te ataco y fallo!!"

Public Const MENSAJE_RECIVE_IMPACTO_CABEZA As String = " te ha pegado en la cabeza por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ As String = " te ha pegado el brazo izquierdo por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_DER As String = " te ha pegado el brazo derecho por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ As String = " te ha pegado la pierna izquierda por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_DER As String = " te ha pegado la pierna derecha por "
Public Const MENSAJE_RECIVE_IMPACTO_TORSO As String = " te ha pegado en el torso por "

Public Const MENSAJE_PRODUCE_IMPACTO_1 As String = "¡¡Le has pegado a "
Public Const MENSAJE_PRODUCE_IMPACTO_CABEZA As String = " en la cabeza por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ As String = " en el brazo izquierdo por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_DER As String = " en el brazo derecho por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ As String = " en la pierna izquierda por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_DER As String = " en la pierna derecha por "
Public Const MENSAJE_PRODUCE_IMPACTO_TORSO As String = " en el torso por "

Public Const MENSAJE_TRABAJO As String = "Haz click sobre el objetivo..."

Public LlegaronEstadisticas As Boolean

'Inventario
Type Inventory
    OBJIndex As Integer
    name As String
    grhindex As Integer
    Amount As Long
    Equipped As Byte
    Valor As Single
    OBJType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
    PuedeUsar As Byte
End Type

Type NpCinV
    OBJIndex As Integer
    name As String
    grhindex As Integer
    Amount As Integer
    Valor As Single
    OBJType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
End Type

Type tEstadisticasUsu
    CiudadanosMatados As Long
    RenegadosMatados As Long
    RepublicanosMatados As Long
    ArmadaMatados As Long
    MiliciaMatados As Long
    CaosMatados As Long
    UsuariosMatados As Long
    NpcMatados As Long
    Clase As Byte
    Raza As Byte
    Genero As Byte
End Type

Public Nombres As Boolean

Public OtroInventario(1 To MAX_INVENTORY_SLOTS) As Inventory

Public UserHechizos(1 To MAXHECHI) As Integer
Public UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV
Public UserMeditar As Boolean
Public UserName As String
Public UserAccount As String
Public UserPassword As String
Public UserMaxHP As Integer
Public UserMinHP As Integer
Public UserMaxMAN As Integer
Public UserMinMAN As Integer
Public UserMaxSTA As Integer
Public UserMinSTA As Integer
Public UserMaxAGU As Byte
Public UserMinAGU As Byte
Public UserMaxHAM As Byte
Public UserMinHAM As Byte
Public UserGLD As Long
Public UserLVL As Integer
Public UserEstado As Byte
Public UserPasarNivel
Public UserExp
Public UserEstadisticas As tEstadisticasUsu
Public UserDescansar As Boolean
Public UserStat As Byte
Public FPSFLAG As Boolean
Public Pausa As Boolean
Public IScombate As Boolean
Public UserParalizado As Boolean
Public UserNavegando As Boolean
Public UserHogar As Byte
Public UserMontando As Boolean
Public Comerciando As Boolean
Public UserClase As eClass
Public UserSexo As eGenero
Public UserRaza As eRaza
Public UserEmail As String

Public PetName As String
Public PetType As eMascota

Public Const NUMCIUDADES As Byte = 11
Public Const NUMSKILLS As Integer = 27
Public Const NUMATRIBUTOS As Integer = 5
Public Const NUMCLASES As Integer = 18
Public Const NUMRAZAS As Integer = 6

Public UserSkills(1 To NUMSKILLS) As Byte
Public SkillsOrig(1 To NUMSKILLS) As Byte
Public SkillsNames(1 To NUMSKILLS) As String

Public UserAtributos(1 To NUMATRIBUTOS) As Byte
Public AtributosNames(1 To NUMATRIBUTOS) As String

Public Ciudades(1 To NUMCIUDADES) As String

Public ListaRazas(1 To NUMRAZAS) As String
Public ListaClases(1 To NUMCLASES) As String

Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer
Public Logged As Boolean

Public UsingSkill As Integer

Public pingTime As Long

Public Enum E_MODO
    Normal = 1
    CrearNuevoPj = 2
    Dados = 3
    CrearNuevaCuenta = 4
    ConectarCuenta = 5
End Enum

Public EstadoLogin As E_MODO

Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
    eo_Part
End Enum

Public Mensajes As Boolean

''
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param trigger_2 ???
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'
Public Enum eTrigger
    NADA = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

'Server stuff
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public Connected As Boolean 'True when connected to server
Public UserMap As Integer

'Control
Public prgRun As Boolean 'When true the program ends


'********** FUNCIONES API ***********
Public Declare Function GetTickCount Lib "kernel32" () As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetActiveWindow Lib "user32" () As Long

'Para ejecutar el Internet Explorer para el manual
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public trueno As Byte
Public TalkMode As Byte

'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'********************************************
'*************Configuracion******************
'********************************************
Public Sound As Byte
Public Music As Byte
Public AmbientSound As Byte
Public RepitMusic As Byte
Public InvertCanal As Byte
Public VolumeSound As Integer
Public VolumeMusic As Integer
Public TileBufferSize As Byte

Public Window As Byte
Public Sinc As Byte

Public BitPixel As Byte

'New
Public CursoresStandar As Byte
Public Cursores As Byte

Public ChatGlobal As Byte
Public ChatFaccionario As Byte
'********************************************
'*************/Configuracion*****************
'********************************************

'Particle Groups
Public TotalStreams As Integer
Public StreamData() As Stream

'RGB Type
Public Type RGB
    r As Long
    g As Long
    b As Long
End Type

Public Type Stream
    name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
    
    speed As Single
    life_counter As Long
    
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

Public meteo_particle As Integer

'****************************************************************
'****************************************************************
'**********************SISTEMA DE CUENTAS************************
'****************************************************************
'****************************************************************
Public Type PjCuenta
    nombre      As String
    Head        As Integer
    body        As Integer
    Shield      As Byte
    Casco       As Byte
    Weapon      As Byte
    Nivel       As Byte
    Mapa        As Integer
    Clase       As Byte
    color       As Byte
    
    tipPet      As Byte
End Type

Public cPJ(0 To 9) As PjCuenta
'****************************************************************
'****************************************************************
'****************************************************************
'****************************************************************
'****************************************************************


Type tServerInfo
    port As Integer
    Ip As String
    name As String
    id As Byte
End Type
Public lServer(1 To 10) As tServerInfo

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Type tSubasta
    active As Boolean 'No se usa este slot en el array ?
    
    OBJIndex As String 'Objeto qe se subasta
    cant     As Integer 'Cantidad de objetos
    
    nckVndedor As String 'Nick del vendedor
    nckCmprdor As String 'Nick del que oferto en caso de ser oferta
    
    actOfert As Long 'Actual oferta en Oro
    fnlOfert As Long 'Oferta final en caso de comprar de una
    
    hsDura As Byte 'Duracion
    mnDura As Byte 'Duracion
    
    grhindex As Long
End Type

Public lstSubastas(1 To 100) As tSubasta

Type tBoton
    TipoAccion As Integer
    SendString As String
    hlist As Integer
    invslot As Integer
End Type

Public MacroKeys() As tBoton
Public BotonElegido As Integer

'Renderizacion sin DirectX 8
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

'Calcular lapso a lapso
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'Timer con proc
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

'Cargar interfaces desde memoria
Public Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Public Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Public Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

'Sensibilidad del mouse
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Const SPI_SETMOUSESPEED = 113
Public Const SPI_GETMOUSESPEED = 112


'Copia archivos
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Public MouseS As Long

'***************************************************************
'***************************************************************
'***************************************************************
'***************************************************************
'***************************************************************
    
    'Posicion en un mapa
    Public Type Position
        X As Long
        Y As Long
    End Type
    
    Public UserPos As Position
    
    'Map sizes in tiles
    Public Const XMinMapSize As Byte = 1
    Public Const XMaxMapSize As Byte = 100
    Public Const YMinMapSize As Byte = 1
    Public Const YMaxMapSize As Byte = 100
    
    'Sets a Grh animation to loop indefinitely.
    Public Const INFINITE_LOOPS As Integer = -1
    
    Public Type Grh
        grhindex As Long
        FrameCounter As Single
        speed As Single
        Started As Byte
        Loops As Integer
        angle As Single
    End Type
    
    'Posicion en el Mundo
    Public Type WorldPos
        map As Integer
        X As Integer
        Y As Integer
    End Type
    
    Public Type GrhData
        sX As Integer
        sY As Integer
        FileNum As Integer
        pixelWidth As Integer
        pixelHeight As Integer
        TileWidth As Single
        TileHeight As Single
        NumFrames As Integer
        Frames() As Integer
        speed As Single
        mini_map_color As Long
        
        tu(3) As Single
        tv(3) As Single
        hardcor As Byte
    End Type
    
    'Lista de cuerpos
    Public Type BodyData
        Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
        HeadOffset As Position
    End Type
    
    'Lista de cabezas
    Public Type HeadData
        Head(E_Heading.NORTH To E_Heading.WEST) As Grh
    End Type
    
    'Lista de las animaciones de las armas
    Public Type WeaponAnimData
        WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
    End Type
    
    'Lista de las animaciones de los escudos
    Public Type ShieldAnimData
        ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
    End Type
    
    Public Type tAura
        Grh As Grh
        color As Long
    End Type
    
    'Apariencia del personaje
    Public Type Char
        active As Byte
        heading As E_Heading
        Pos As Position
        donador As Boolean
        
        Bandera As Byte
        
        label_color(3) As Long
        state As Byte
        
        iBody As Integer
        body As BodyData
        
        iHead As Integer
        Head As HeadData
        
        Casco As HeadData
        
        Arma As WeaponAnimData
        UsandoArma As Boolean
        
        dl As Boolean
        dialog() As String
        dialogColor(3) As Long
        dialogLife As Long
        dialogStart As Long
        dialogHeight As Byte
        dialogIndex As Byte
        
        Escudo As ShieldAnimData
        ShieldOffSetY As Integer
        
        plusGrh(2) As tAura
        
        FX As Grh
        fxIndex As Integer
        
        AlphaX As Double
        last_tick As Long
        
        nombre As String
        clan As String
        offNameX As Byte
        offClanX As Byte
    
        particle_count As Integer
        particle_group() As Long
        
        pie As Boolean
        Muerto As Boolean
        invisible As Boolean
        Priv As Byte
        
        Moving As Byte
        
        MoveOffSetX As Single
        MoveOffSetY As Single
        
        ScrollDirectionX As Integer
        ScrollDirectionY As Integer
    End Type
    
    'Info de un objeto
    Public Type Obj
        OBJIndex As Integer
        Amount As Integer
        tipe As Byte
        EsFijo As Byte
    End Type
    
    'Tipo de las celdas del mapa
    Public Type MapBlock
        Graphic(1 To 4) As Grh
        CharIndex As Integer
        ObjGrh As Grh
        
        light_value(0 To 3) As Long
        
        luz As Integer
        color(3) As Long
        
        particle_group_index As Integer
        effectIndex As Integer
        
        NPCIndex As Integer
        OBJInfo As Obj
        TileExit As WorldPos
        Blocked As Byte
        
        FX As Integer
        fXGrh As Grh
        
        Trigger As Integer
    End Type
    
    Public MapName As String
    
    'Status del user
    Public engineBaseSpeed As Single
    
    Public NumHeads As Integer
    Public NumFxs As Integer
    
    Public NumChars As Integer
    Public LastChar As Integer
    
    Public NumWeaponAnims As Integer
    Public NumShieldAnims As Integer
    
    Public GrhData() As GrhData
    Public BodyData() As BodyData
    Public HeadData() As HeadData
    Public FxData() As tIndiceFx
    Public WeaponAnimData() As WeaponAnimData
    Public ShieldAnimData() As ShieldAnimData
    Public CascoAnimData() As HeadData
    
    Public MapData() As MapBlock
    Public charlist(1 To 10000) As Char
    
    Public char_current As Integer
    
    Public bTecho As Boolean

'***************************************************************
'***************************************************************
'***************************************************************
'***************************************************************
'***************************************************************

'*****************************
'Resolucion*******************
'*****************************
'Constantes


Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_TEST = &H4

'Estructura
Public Type DevMode
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Public Declare Function EnumDisplaySettings Lib "user32" _
    Alias "EnumDisplaySettingsA" ( _
    ByVal lpszDeviceName As Long, _
    ByVal iModeNum As Long, _
    lpDevMode As Any) As Boolean

Public Declare Function ChangeDisplaySettings Lib "user32" _
    Alias "ChangeDisplaySettingsA" ( _
    lpDevMode As Any, _
    ByVal dwFlags As Long) As Long

Public Declare Function GetDeviceCaps Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal nIndex As Long) As Long

Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" ( _
    ByVal lpDriverName As String, _
    ByVal lpDeviceName As String, _
    ByVal lpOutput As String, _
    ByVal lpInitData As Any) As Long

Public OldX As Long
Public OldY As Long
Public OldBit As Long

Public ChangeResolution As Boolean

Type tCorreo
    mensaje As String
    De As String
    
    Cantidad As Integer
    item As Integer
End Type
    
Public Correos(1 To 20) As tCorreo
    
