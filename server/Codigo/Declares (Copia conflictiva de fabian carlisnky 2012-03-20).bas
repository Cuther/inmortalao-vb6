Attribute VB_Name = "Declaraciones"
Option Explicit

Public Const MENSAJES_TOPE_CORREO As Byte = 20 ' Cantidad de mensajes de correo



Public Security As New clsSecurity
Public RondasAutomatico As Byte
Public ModExpX As Long
Public ModOroX As Long
Public ModTrabajo As Long
Public ModSkill As Long

Public tHora As Byte
Public tMinuto As Byte
Public tSeg As Byte

Public TrashCollector As New Collection

Public Const MAXSPAWNATTEMPS = 60
Public Const INFINITE_LOOPS As Integer = -1
Public Const FXSANGRE = 14

Public Const NO_3D_SOUND As Byte = 0

Public Const iFragataFantasmal = 87
Public Const iBarca = 84
Public Const iGalera = 85
Public Const iGaleon = 86

Public Enum iMinerales
    HierroCrudo = 192
    PlataCruda = 193
    OroCrudo = 194
    LingoteDeHierro = 386
    LingoteDePlata = 387
    LingoteDeOro = 388
End Enum

Public Enum PlayerType
    User = &H1
    Dios = &H2
    VIP = &H4
End Enum

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

Public Enum eRaza
    Humano = 1
    Elfo
    Drow
    Gnomo
    Enano
    Orco
End Enum

Enum eGenero
    Hombre = 1
    Mujer
End Enum

Public Const LimiteNewbie As Byte = 13

Public Type tCabecera 'Cabecera de los con
    desc As String * 255
    crc As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

'Barrin 3/10/03
'Cambiado a 2 segundos el 30/11/07
Public Const TIEMPO_INICIOMEDITAR As Integer = 2000

Public Const NingunEscudo As Integer = 2
Public Const NingunCasco As Integer = 2
Public Const NingunArma As Integer = 2

Public Const EspadaMataDragonesIndex As Integer = 402

Public Const MAXMASCOTASENTRENADOR As Byte = 7

Public Enum FXIDs
    FXWARP = 1
    FXMEDITARCHICO = 4
    FXMEDITARMEDIANO = 5
    FXMEDITARGRANDE = 6
    FXMEDITARXGRANDE = 16
    FXMEDITARXXGRANDE = 34
End Enum

Public Const TIEMPO_CARCEL_PIQUETE As Long = 10

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
    Nada = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

''
' constantes para el trigger 6
'
' @see eTrigger
' @param TRIGGER6_PERMITE TRIGGER6_PERMITE
' @param TRIGGER6_PROHIBE TRIGGER6_PROHIBE
' @param TRIGGER6_AUSENTE El trigger no aparece
'
Public Enum eTrigger6
    TRIGGER6_PERMITE = 1
    TRIGGER6_PROHIBE = 2
    TRIGGER6_AUSENTE = 3
End Enum

'TODO : Reemplazar por un enum
Public Const Bosque As String = "BOSQUE"
Public Const Nieve As String = "NIEVE"
Public Const Desierto As String = "DESIERTO"
Public Const Ciudad As String = "CIUDAD"
Public Const Campo As String = "CAMPO"
Public Const Dungeon As String = "DUNGEON"

' <<<<<< Targets >>>>>>
Public Enum TargetType
    uUsuarios = 1
    uNPC = 2
    uUsuariosYnpc = 3
    uTerreno = 4
End Enum

' <<<<<< Acciona sobre >>>>>>
Public Enum TipoHechizo
    uPropiedades = 1
    uEstado = 2
    uInvocacion = 4
    uCreateTelep = 5
    uFamiliar = 6
    uMaterializa = 7
    uPropEsta = 8
    uCalmacion = 9
    uCreateMagic = 10
    uEquipamiento = 11
    uDetectarInvis = 12
End Enum

Public Const MAX_MENSAJES_FORO As Byte = 35

Public Const MAXUSERHECHIZOS As Byte = 35


' TODO: Y ESTO ? LO CONOCE GD ?
Public Const EsfuerzoTalarGeneral As Byte = 4
Public Const EsfuerzoTalarLeñador As Byte = 2

Public Const EsfuerzoBotanicaGeneral As Byte = 4
Public Const EsfuerzoBotanicaDruida As Byte = 2

Public Const EsfuerzoPescarPescador As Byte = 1
Public Const EsfuerzoPescarGeneral As Byte = 3

Public Const EsfuerzoExcavarMinero As Byte = 2
Public Const EsfuerzoExcavarGeneral As Byte = 5

Public Const FX_TELEPORT_INDEX As Integer = 1

' La utilidad de esto es casi nula, sólo se revisa si fue a la cabeza...
Public Enum PartesCuerpo
    bCabeza = 1
    bPiernaIzquierda = 2
    bPiernaDerecha = 3
    bBrazoDerecho = 4
    bBrazoIzquierdo = 5
    bTorso = 6
End Enum

Public Const Guardias As Integer = 6

Public Const MAXREP As Long = 6000000
Public Const MAXORO As Long = 90000000
Public Const MAXEXP As Long = 199999999

Public Const MAXUSERMATADOS As Long = 65000

Public Const MAXATRIBUTOS As Byte = 35
Public Const MINATRIBUTOS As Byte = 6

Public Const LingoteHierro As Integer = 386
Public Const LingotePlata As Integer = 387
Public Const LingoteOro As Integer = 388
Public Const Leña As Integer = 58
Public Const Raiz As Integer = 888

Public Const PielLobo As Integer = 414
Public Const PielOso As Integer = 415
Public Const PielLoboInvernal As Integer = 1145


Public Const MAXNPCS As Integer = 10000
Public Const MAXCHARS As Integer = 10000

Public Const DAGA As Integer = 15
Public Const FOGATA_APAG As Integer = 136
Public Const FOGATA As Integer = 63

Public Const ObjArboles As Integer = 4

Public Const MARTILLO_HERRERO As Integer = 389
Public Const SERRUCHO_CARPINTERO As Integer = 198
Public Const HACHA_LEÑADOR As Integer = 127
Public Const PIQUETE_MINERO As Integer = 187
Public Const RED_PESCA As Integer = 138
Public Const CAÑA_PESCA As Integer = 881
Public Const TIJERAS As Integer = 885
Public Const COSTURERO As Integer = 886
Public Const OLLA As Integer = 887

Public Enum eNPCType
    Comun = 0
    Revividor = 1
    GuardiaReal = 2
    Entrenador = 3
    Banquero = 4
    Noble = 5
    DRAGON = 6
    Timbero = 7
    Guardiascaos = 8
    ResucitadorNewbie = 9
    Pirata = 10
    Bot = 11
End Enum
'  0 NPCs Comunes
'  1 Resucitadores
'  2 Guardias
'  3 Entrenadores
'  4 Banqueros
'  5 Facciones
'  6 (nada)
'  7 Transportadores
'  8 Carceleros
'  9 (nada)
' 10 Paga recompensas
' 11 Veterinarias
' 12 Apuestas
' 13 Presentadores de Quest
' 14 Objetivos de Quest
' 15 Centinelas
' 16 Subastadores
Public Const MIN_APUÑALAR As Byte = 10

'********** CONSTANTANTES ***********

''
' Cantidad de skills
Public Const NUMSKILLS As Byte = 27

''
' Cantidad de Atributos
Public Const NUMATRIBUTOS As Byte = 5

''
' Cantidad de Clases
Public Const NUMCLASES As Byte = 18

''
' Cantidad de Razas
Public Const NUMRAZAS As Byte = 6

''
' Valor maximo de cada skill
Public Const MAXSKILLPOINTS As Byte = 100

''
'Direccion
'
' @param NORTH Norte
' @param EAST Este
' @param SOUTH Sur
' @param WEST Oeste
'
Public Enum eHeading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

''
' Cantidad maxima de mascotas
Public Const MAXMASCOTAS As Byte = 3

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const vlASALTO As Integer = 100
Public Const vlASESINO As Integer = 1000
Public Const vlCAZADOR As Integer = 5
Public Const vlNoble As Integer = 5
Public Const vlLadron As Integer = 25
Public Const vlProleta As Integer = 2

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const iCuerpoMuerto As Integer = 8
Public Const iCabezaMuerto As Integer = 500


Public Const iORO As Byte = 12
Public Const Pescado As Byte = 139

Public Enum PECES_POSIBLES
    PESCADO1 = 139
    PESCADO2 = 544
    PESCADO3 = 545
    PESCADO4 = 546
End Enum

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Enum eSkill
    Tacticas = 1
    armas = 2
    artes = 3
    Apuñalar = 4
    arrojadizas = 5
    Proyectiles = 6
    DefensaEscudos = 7
    Magia = 8
    Resistencia = 9
    Meditar = 10
    Ocultarse = 11
    Domar = 12
    Musica = 13
    Robar = 14
    Comerciar = 15
    Supervivencia = 16
    Liderazgo = 17
    Pesca = 18
    Mineria = 19
    Talar = 20
    botanica = 21
    Herreria = 22
    Carpinteria = 23
    alquimia = 24
    Sastreria = 25
    Navegacion = 26
    Equitacion = 27
End Enum

Public Const FundirMetal = 88

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    constitucion = 5
End Enum

Public Const AdicionalHPGuerrero As Byte = 2 'HP adicionales cuando sube de nivel
Public Const AdicionalHPCazador As Byte = 1 'HP adicionales cuando sube de nivel

Public Const AumentoSTDef As Byte = 15
Public Const AumentoSTLadron As Byte = AumentoSTDef + 3
Public Const AumentoSTMago As Byte = AumentoSTDef - 1
Public Const AumentoSTLeñador As Byte = AumentoSTDef + 23
Public Const AumentoSTPescador As Byte = AumentoSTDef + 20
Public Const AumentoSTMinero As Byte = AumentoSTDef + 25

'Tamaño del mapa
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

'Tamaño del tileset
Public Const TileSizeX As Byte = 32
Public Const TileSizeY As Byte = 32

'Tamaño en Tiles de la pantalla de visualizacion
Public Const XWindow As Byte = 17
Public Const YWindow As Byte = 13

'Sonidos
Public Const SND_SWING As Byte = 2
Public Const SND_TALAR As Byte = 13
Public Const SND_PESCAR As Byte = 14
Public Const SND_MINERO As Byte = 15
Public Const SND_WARP As Byte = 3
Public Const SND_PUERTA As Byte = 5
Public Const SND_NIVEL As Byte = 128

Public Const SND_USERMUERTE As Byte = 11
Public Const SND_IMPACTO As Byte = 10
Public Const SND_IMPACTO2 As Byte = 12
Public Const SND_LEÑADOR As Byte = 13
Public Const SND_FOGATA As Byte = 14
Public Const SND_AVE As Byte = 21
Public Const SND_AVE2 As Byte = 22
Public Const SND_AVE3 As Byte = 34
Public Const SND_GRILLO As Byte = 28
Public Const SND_GRILLO2 As Byte = 29
Public Const SND_SACARARMA As Byte = 25
Public Const SND_ESCUDO As Byte = 37
Public Const SND_BEBER As Byte = 135
Public Const SND_REMO As Byte = 255
Public Const SND_NEWCLAN As Byte = 44
Public Const SND_NEWMEMBER As Byte = 43
Public Const SND_OUT As Byte = 45
Public Const SND_INCINERACION As Byte = 123
Public Const SND_DROPITEM As Byte = 132
Public Const SND_REVIVE As Byte = 204

Public Const MARTILLOHERRERO As Byte = 41
Public Const LABUROCARPINTERO As Byte = 42

''
' Cantidad maxima de objetos por slot de inventario
Public Const MAX_INVENTORY_OBJS As Integer = 10000

''
' Cantidad de "slots" en el inventario
Public Const MAX_INVENTORY_SLOTS As Byte = 25

''
' Constante para indicar que se esta usando ORO
Public Const FLAGORO As Integer = MAX_INVENTORY_SLOTS + 1

' CATEGORIAS PRINCIPALES
Public Enum eOBJType
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
    otruna = 49
    otCualquiera = 1000
End Enum

'Texto
'Public Const FONTTYPE_TALK As String = "~255~255~255~0~0"
'Public Const FONTTYPE_FIGHT As String = "~255~0~0~1~0"
'Public Const FONTTYPE_WARNING As String = "~32~51~223~1~1"
'Public Const FONTTYPE_INFO As String = "~65~190~156~0~0"
'Public Const FONTTYPE_INFOBOLD As String = "~65~190~156~1~0"
'Public Const FONTTYPE_EJECUCION As String = "~130~130~130~1~0"
'Public Const FONTTYPE_Grupo As String = "~255~180~255~0~0"
'Public Const FONTTYPE_VENENO As String = "~0~255~0~0~0"
'Public Const FONTTYPE_GUILD As String = "~255~255~255~1~0"
'Public Const FONTTYPE_SERVER As String = "~0~185~0~0~0"
'Public Const FONTTYPE_GUILDMSG As String = "~228~199~27~0~0"
'Public Const FONTTYPE_CONSEJO As String = "~130~130~255~1~0"
'Public Const FONTTYPE_CONSEJOCAOS As String = "~255~60~00~1~0"
'Public Const FONTTYPE_CONSEJOVesA As String = "~0~200~255~1~0"
'Public Const FONTTYPE_CONSEJOCAOSVesA As String = "~255~50~0~1~0"
'Public Const FONTTYPE_CENTINELA As String = "~0~255~0~1~0"

'Estadisticas
Public Const STAT_MAXELV As Byte = 50
Public Const STAT_MAXHP As Integer = 999
Public Const STAT_MAXSTA As Integer = 999
Public Const STAT_MAXMAN As Integer = 9999
Public Const STAT_MAXHIT_UNDER36 As Byte = 99
Public Const STAT_MAXHIT_OVER36 As Integer = 999
Public Const STAT_MAXDEF As Byte = 99



' **************************************************************
' **************************************************************
' ************************ TIPOS *******************************
' **************************************************************
' **************************************************************

Public Type tHechizo
    Nombre As String
    desc As String
    PalabrasMagicas As String
    
    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String
    
    HechizoDeArea As Byte
    AreaEfecto As Byte
    Afecta As Byte

    Tipo As TipoHechizo
    
    WAV As Integer
    FXgrh As Integer
    loops As Byte
    Particle As Integer
    
    SubeHP As Byte
    MinHP As Integer
    MaxHP As Integer
    
    
    SubeAgilidad As Byte
    MinAgilidad As Integer
    MaxAgilidad As Integer
    
    SubeFuerza As Byte
    MinFuerza As Integer
    MaxFuerza As Integer

    Invisibilidad As Byte
    Paraliza As Byte
    Inmoviliza As Byte
    RemoverParalisis As Byte
    CuraVeneno As Byte
    Envenena As Byte
    Incinera As Byte
    Estupidez As Byte
    Ceguera As Byte
    Revivir As Byte
    Resurreccion As Byte
    ReviveFamiliar As Byte
    
    'Jose Castelli
    CreaTelep As Byte
    AutoLanzar As Byte
    'Jose Castelli
    
    'Mannakia
    Desencantar As Byte
    Sanacion As Byte
    Certero As Byte
    
    CreaAlgo As Byte
    MinDef As Integer
    MaxDef As Integer
    
    MinHit As Byte
    MaxHit As Byte
    'Mannakia
    
    'By jose castelli // Metamorfosis
    Metamorfosis As Byte
    Asusta As Byte
    Extrahit As Byte
    Extradef As Byte
    body As Integer
    Head As Integer
    'By jose castelli  // Metamorfosis
    
    Mimetiza As Byte
    RemueveInvisibilidadParcial As Byte
    
    Invoca As Byte
    NumNpc As Integer
    Cant As Integer

    MinSkill As Integer
    ManaRequerido As Integer

    StaRequerido As Integer

    Target As TargetType
    
    Anillo As Byte
        '1 Espectral
        '2 Penumbra
End Type

Public Type LevelSkill
    LevelValue As Integer
End Type

Public Type UserObj
    ObjIndex As Integer
    Amount As Integer
    Equipped As Byte
    Prob As Byte
End Type

Public Type Inventario
    Object(1 To MAX_INVENTORY_SLOTS) As UserObj
    WeaponEqpObjIndex As Integer
    WeaponEqpSlot As Byte
    NudiEqpSlot As Byte
    NudiEqpIndex As Integer
    ArmourEqpObjIndex As Integer
    ArmourEqpSlot As Byte
    EscudoEqpObjIndex As Integer
    EscudoEqpSlot As Byte
    CascoEqpObjIndex As Integer
    CascoEqpSlot As Byte
    MunicionEqpObjIndex As Integer
    MunicionEqpSlot As Byte
    MonturaObjIndex As Integer
    MonturaSlot As Byte
    AnilloEqpObjIndex As Integer
    AnilloEqpSlot As Byte
    BarcoObjIndex As Integer
    BarcoSlot As Byte
    MagicIndex As Integer
    MagicSlot As Integer
    NroItems As Integer
End Type

Public Type tGrupoData
    PIndex As Integer
    RemXP As Double 'La exp. en el server se cuenta con Doubles
    TargetUser As Integer 'Para las invitaciones
End Type

Public Type Position
    x As Integer
    Y As Integer
End Type

Public Type WorldPos
    map As Integer
    x As Integer
    Y As Integer
End Type

Public Type FXdata
    Nombre As String
    GrhIndex As Integer
    Delay As Integer
End Type

Public Enum eMagicType
    ResistenciaMagica = 1
    ModificaAtributo = 2
    ModificaSkill = 3
    AceleraVida = 4
    AceleraMana = 5
    AumentaGolpe = 6
    DisminuyeGolpe = 7
    Nada = 8
    MagicasNoAtacan = 9
    Incinera = 10
    Paraliza = 11
    CarroMinerales = 12
    CaminaOculto = 13
    DañoMagico = 14
    Sacrificio = 15
    Silencio = 16
    NadieDetecta = 17
    Experto = 18
    Envenena = 19
End Enum

'Datos de user o npc
Public Type Char
    CharIndex As Integer
    Head As Integer
    body As Integer

    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
    
    FX As Integer
    loops As Integer
    
    heading As eHeading
End Type

'Tipos de objetos
Public Type ObjData
    Name As String 'Nombre del obj
    
    OBJType As eOBJType 'Tipo enum que determina cuales son las caract del obj
    
    GrhIndex As Integer ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    
    'Solo contenedores
    MaxItems As Integer
    Conte As Inventario
    Apuñala As Byte
    
    HechizoIndex As Integer
    
    ForoID As String
    
    MinHP As Integer ' Minimo puntos de vida
    MaxHP As Integer ' Maximo puntos de vida
    
    SubTipo As Byte
    
    MineralIndex As Integer
    LingoteInex As Integer
    
    
    proyectil As Integer
    Municion As Integer
    
    Crucial As Byte
    Newbie As Integer
    
    DesdeMap As Long
    HastaMap As Long
    HastaY As Byte
    HastaX As Byte
    CantidadSkill As Byte
    
    'Pociones
    TipoPocion As Byte
    MaxModificador As Integer
    MinModificador As Integer
    DuracionEfecto As Long
    MinSkill As Integer
    LingoteIndex As Integer
    
    MinHit As Integer 'Minimo golpe
    MaxHit As Integer 'Maximo golpe
    
    MinHAM As Integer
    MinSed As Integer
    
    def As Integer
    MinDef As Integer ' Armaduras
    MaxDef As Integer ' Armaduras
    
    Ropaje As Integer 'Indice del grafico del ropaje
    
    WeaponAnim As Integer ' Apunta a una anim de armas
    ShieldAnim As Integer ' Apunta a una anim de escudo
    CascoAnim As Integer
    
    valor As Long     ' Precio
    
    Cerrada As Integer
    Llave As Byte
    clave As Long 'si clave=llave la puerta se abre o cierra
    
    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer
    
    RazaTipo As Byte '1 Altas 2 Bajas 3 Orcas
    RazaEnana As Byte
    MinELV As Byte
    
    QueAtributo As Byte
    EfectoMagico As eMagicType
    CuantoAumento As Byte
    QueSkill As Byte
    
    CuantoAgrega As Integer ' Para los contenedores
    
    Mujer As Byte
    Hombre As Byte

    Agarrable As Byte
    
    LingH As Integer
    LingO As Integer
    LingP As Integer
    Madera As Integer
    
    SkPociones As Integer
    raies As Integer
    PielLobo As Integer
    PielOso As Integer
    PielLoboInvernal As Integer
    
    SkSastreria As Integer
    SkHerreria As Integer
    SkCarpinteria As Integer
    
    texto As String
    
    'Clases que no tienen permitido usar este obj
    ClaseProhibida(1 To NUMCLASES) As eClass
    ClaseTipo As Byte
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    
    Real As Integer
    Caos As Integer
    Milicia As Integer

    DefensaMagicaMax As Integer
    DefensaMagicaMin As Integer
    
    CPO As String

    
    Refuerzo As Byte
    ResistenciaMagica As Integer
    
    Log As Byte 'es un objeto que queremos loguear? Pablo (ToxicWaste) 07/09/07
    NoLog As Byte 'es un objeto que esta prohibido loguear?
    
    DosManos As Byte
End Type

Public Type Obj
    ObjIndex As Integer
    Amount As Integer
End Type

'[Pablo ToxicWaste]
Public Type ModClase
    Evasion As Double
    AtaqueArmas As Double
    AtaqueProyectiles As Double
    DañoArmas As Double
    DañoProyectiles As Double
    DañoWrestling As Double
    Escudo As Double
End Type

Public Type ModRaza
    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Carisma As Single
    constitucion As Single
End Type
'[/Pablo ToxicWaste]

'[KEVIN]
'Banco Objs
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
'[/KEVIN]

'[KEVIN]
Public Type BancoInventario
    Object(1 To MAX_BANCOINVENTORY_SLOTS) As UserObj
    NroItems As Integer
End Type
'[/KEVIN]


'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'******* T I P O S   D E    U S U A R I O S **************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

'Estadisticas de los usuarios
Public Type UserStats
    GLD As Long 'Dinero
    Banco As Long
    
    MaxHP As Integer
    MinHP As Integer
    
    MaxSTA As Integer
    MinSTA As Integer
    MaxMAN As Integer
    MinMAN As Integer
    MaxHit As Integer
    MinHit As Integer
    
    MaxHAM As Integer
    MinHAM As Integer
    
    MaxAGU As Integer
    MinAGU As Integer
        
    def As Integer
    Exp As Double
    ELV As Byte
    ELU As Long
    UserSkills(1 To NUMSKILLS) As Byte
    UserAtributos(1 To NUMATRIBUTOS) As Byte
    UserAtributosBackUP(1 To NUMATRIBUTOS) As Byte
    UserHechizos(1 To MAXUSERHECHIZOS) As Integer
    
    eMinDef As Byte
    eMaxDef As Byte
    eMinHit As Byte
    eMaxHit As Byte
    eCreateTipe As Byte
    
    dMaxDef As Integer
    dMinDef As Integer
    
    NPCsMuertos As Integer
    
    VecesMuertos As Long
    
    SkillPts As Integer
    
    PuedeStaff As Byte
End Type

'Flags
Public Type UserFlags

    'Torneos automaticos
    automatico As Boolean
    'Torneos automaticos
    
    NoFalla As Byte

   'Castelli Casamientos
    toyCasado As Byte
    yaOfreci As Byte
    miPareja As String
   'Castelli Casamientos

    'Mannakia Duelos 1vs1
    vicDuelo As Integer 'Victima
    inDuelo As Byte 'ESTA DUELEANDO ?
    solDuelo As Integer 'Solicito ?
    
    'Sistema de Grupo
    Solicito As Integer
    Invito As Integer
    'Mannakia
    
    ' Castelli Metamorfossis
    Metamorfosis As Byte
    ' Castelli Metamorfossis

    DondeTiroMap As Integer
    DondeTiroX As Integer
    DondeTiroY As Integer
    TiroPortalL As Integer

    Muerto As Byte '¿Esta muerto?
    Escondido As Byte '¿Esta escondido?
    Comerciando As Boolean '¿Esta comerciando?
    UserLogged As Boolean '¿Esta online?
    accountlogged As Boolean ' ¿Cuenta online?
    Meditando As Boolean
    ModoCombate As Boolean

    Hambre As Byte
    Sed As Byte
    
    Entrenando As Byte

    Resucitando As Byte
    Envenenado As Byte
    Incinerado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    Estupidez As Byte
    Ceguera As Byte
    Invisible As Byte
    Oculto As Byte
    Desnudo As Byte
    Descansar As Boolean
    Hechizo As Integer
    TomoPocion As Boolean
    TipoPocion As Byte
    
    Trabajando As Boolean
    Lingoteando As Byte

    Navegando As Byte
    Montando As Byte
    
    Seguro As Boolean

    DuracionEfecto As Long
    TargetNPC As Integer ' Npc señalado por el usuario
    TargetNpcTipo As eNPCType ' Tipo del npc señalado
    NpcInv As Integer
    
    Ban As Byte

    TargetUser As Integer ' Usuario señalado
    
    TargetObj As Integer ' Obj señalado
    TargetObjMap As Integer
    TargetObjX As Integer
    TargetObjY As Integer
    
    TargetMap As Integer
    TargetX As Integer
    TargetY As Integer
    
    TargetObjInvIndex As Integer
    TargetObjInvSlot As Integer
    
    AtacadoPorNpc As Integer
    AtacadoPorUser As Integer
    NPCAtacado As Integer
    Privilegios As PlayerType

    OldBody As Integer
    OldHead As Integer
    AdminInvisible As Byte
    AdminPerseguible As Boolean

    TimesWalk As Long
    
    UltimoMensaje As Byte

    Silenciado As Byte
    
End Type

Public Type UserCounters
    Silenciado As Integer
    Habla As Integer
    
    CreoTeleport As Boolean
    TimeTeleport As Integer

    Metamorfosis As Integer

    IdleCount As Long
    AttackCounter As Integer
    HPCounter As Integer
    STACounter As Integer
    Frio As Integer
    COMCounter As Integer
    AGUACounter As Integer
    Veneno As Integer
    Fuego As Integer
    Paralisis As Integer
    Ceguera As Integer
    Estupidez As Integer
    
    Invisibilidad As Integer
    TiempoOculto As Integer
    
    PiqueteC As Long
    Pena As Long

    Saliendo As Boolean
    Salir As Integer

    IntervaloRevive As Long
    
    TimerLanzarSpell As Long
    TimerPuedeAtacar As Long
    TimerPuedeUsarArco As Long
    TimerPuedeTrabajar As Long
    TimerUsar As Long
    TimerMagiaGolpe As Long
    TimerGolpeMagia As Long
    TimerGolpeUsar As Long
    
    Trabajando As Long  ' Para el centinela
    Ocultando As Long   ' Unico trabajo no revisado por el centinela
    
    failedUsageAttempts As Long
End Type

'Cosas faccionarias.
Public Type tFacciones
    ArmadaReal As Byte
    Ciudadano As Byte
    
    FuerzasCaos As Byte
    Renegado As Byte
    
    Milicia As Byte
    Republicano As Byte
    
    CiudadanosMatados As Integer
    RenegadosMatados As Integer
    RepublicanosMatados As Integer
    
    MilicianosMatados As Integer
    ArmadaMatados As Integer
    CaosMatados As Integer
    
    Rango As Byte

End Type

Public Enum eMascota
    Fuego = 1
    Tierra
    Agua
    Ely
    Fatuo
    
    Tigre
    Lobo
    Oso
    Ent
End Enum

Type Mascota
    TieneFamiliar As Byte
    invocado As Boolean
    
    Nombre As String
    
    MinHP As Integer
    MaxHP As Integer
    
    ELV As Byte
    ELU As Long
    Exp As Long
    
    Tipo As eMascota
    
    MinHit As Integer
    MaxHit As Integer
    
    gDesarma As Byte
    gEntorpece As Byte
    gEnseguece As Byte
    gParaliza As Byte
    gEnvenena As Byte
    
    
    Curar As Byte
    Desencanta As Byte
    Descargas As Byte
    Paraliza As Byte
    Inmoviliza As Byte
    Tormentas As Byte
    Misil As Byte
    DetecInvi As Byte
    
    NpcIndex As Integer
End Type


Type tCorreo
    De As String
    Mensaje As String
    Cantidad As Integer
    Item As Integer
    idmsj As Long
End Type

'Tipo de los Usuarios
Public Type User
    Redundance As Byte
    Name As String
    account As String
    IndexAccount As Long 'Add Nod Kopfnickend
    Indexpj As Long
    
    donador As Integer

    showName As Boolean 'Permite que los GMs oculten su nick con el comando /SHOWNAME
    
    
    ' Sistema de correo Por Castelli
    cVer As Byte
    Correos(1 To 20) As tCorreo
    cant_mensajes As Byte ' Obvio que tope en 20
    ' Sistema de correo Por Castelli
    
    Char As Char
    OrigChar As Char
    
    desc As String ' Descripcion

    masc As Mascota
    
    Clase As eClass
    Raza As eRaza
    Genero As eGenero
    email As String
    Hogar As Byte
        
    Invent As Inventario
    
    pos As WorldPos
    AuxPos As WorldPos
    
    ConnIDValida As Boolean
    ConnID As Long 'ID
    
    BancoInvent As BancoInventario

    Counters As UserCounters
    
    MascotasIndex(1 To MAXMASCOTAS) As Integer
    MascotasType(1 To MAXMASCOTAS) As Integer
    NroMascotas As Integer
    
    Stats As UserStats
    flags As UserFlags

    Faccion As tFacciones
    
    ip As String

    ComUsu As tCOmercioUsuario
    
    GuildIndex As Integer   'puntero al array global de guilds
    FundandoGuildAlineacion As ALINEACION_GUILD     'esto esta aca hasta que se parchee el cliente y se pongan cadenas de datos distintas para cada alineacion
    EscucheClan As Integer
    
    GrupoIndex As Integer
    GrupoSolicitud As Integer
    
    
    AreasInfo As AreaInfo
    
    'Outgoing and incoming messages
    outgoingData As clsByteQueue
    incomingData As clsByteQueue
End Type


'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'**  T I P O S   D E    N P C S **************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Public Type NPCStats
    Alineacion As Integer
    MaxHP As Long
    MinHP As Long
    MaxHit As Integer
    MinHit As Integer
    def As Integer
    defM As Integer
End Type

Public Type NpcCounters
    Paralisis As Integer
    TiempoExistencia As Long
End Type

Public Type NPCFlags
    AfectaParalisis As Byte
    Domable As Integer
    Respawn As Byte
    NPCActive As Boolean '¿Esta vivo?
    Follow As Boolean
    Faccion As Byte
    AtacaDoble As Byte
    LanzaSpells As Byte
    
    ExpCount As Long
    
    OldMovement As TipoAI
    OldHostil As Byte
    
    AguaValida As Byte
    TierraInvalida As Byte
    
    Sound As Integer
    AttackedBy As Integer
    AttackedFirstBy As String
    BackUp As Byte
    RespawnOrigPos As Byte
    
    Envenenado As Byte
    Incinerado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    Invisible As Byte
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
End Type

Public Type tCriaturasEntrenador
    NpcIndex As Integer
    NpcName As String
    tmpIndex As Integer
End Type

Type tBot
    UpMana As Integer
    ManaMax As Integer
    
    UpVida As Integer
    VidaMax As Integer
    
    TargetUser As Integer
    TargetNPC As Integer
    
    IntervaloAtaque As Long
    IntervaloHechizo As Long
    
    RandomDire As Byte
End Type


Public Type npc
    Name As String
    Char As Char 'Define como se vera
    desc As String

    NPCtype As eNPCType
    Numero As Integer
    Faccion As Byte
    InvReSpawn As Byte

    Comercia As Integer
    Target As Long
    TargetNPC As Long
    TipoItems As Integer

    Veneno As Byte
    Fuego As Byte
    
    pos As WorldPos 'Posicion
    oldPos As WorldPos
    Orig As WorldPos
    StartPos As WorldPos
    lastHeading  As eHeading
    
    SkillDomar As Integer

    Movement As TipoAI
    Attackable As Byte
    Hostile As Byte
    PoderAtaque As Long
    PoderEvasion As Long

    GiveEXP As Long
    GiveGLD As Long

    Stats As NPCStats
    flags As NPCFlags
    Contadores As NpcCounters
    
    Invent As Inventario
    CanAttack As Byte
    
    NroExpresiones As Byte
    Expresiones() As String ' le da vida ;)
    
    NroSpells As Byte
    Spells() As Integer  ' le da vida ;)
    
    '<<<<Entrenadores>>>>>
    NroCriaturas As Integer
    Criaturas() As tCriaturasEntrenador
    MaestroUser As Integer
    MaestroNpc As Integer
    Mascotas As Integer
    
    IsFamiliar As Boolean
    
    AreasInfo As AreaInfo
End Type

'**********************************************************
'**********************************************************
'******************** Tipos del mapa **********************
'**********************************************************
'**********************************************************
'Tile
Public Type MapBlock
    Blocked As Byte
    Graphic(1 To 4) As Integer
    userindex As Integer
    NpcIndex As Integer
    BotIndex As Integer
    
    ObjInfo As Obj
    ObjEsFijo As Byte
    
    TileExit As WorldPos
    Trigger As eTrigger
End Type

'Info del mapa
Type MapInfo
    NumUsers As Integer
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    Seguro As Byte
    Pk As Boolean
    MagiaSinEfecto As Byte
    NoEncriptarMP As Byte
    InviSinEfecto As Byte
    ResuSinEfecto As Byte
    
    Terreno As String
    Zona As String
    Restringir As String
    BackUp As Byte
End Type

'********** V A R I A B L E S     P U B L I C A S ***********

Public ULTIMAVERSION As String

Public ListaRazas(1 To NUMRAZAS) As String
Public SkillsNames(1 To NUMSKILLS) As String
Public ListaClases(1 To NUMCLASES) As String
Public ListaAtributos(1 To NUMATRIBUTOS) As String


Public RecordUsuarios As Long

'
'Directorios
'

''
'Ruta base del server, en donde esta el "server.ini"
Public IniPath As String


''
'Ruta base para los archivos de mapas
Public MapPath As String

''
'Ruta base para los DATs
Public DatPath As String

''
'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

''
'Numero de usuarios actual
Public NumUsers As Integer
Public LastUser As Integer
Public LastChar As Integer
Public NumChars As Integer
Public LastNPC As Integer
Public NumNPCs As Integer
Public NumFX As Integer
Public NumMaps As Integer
Public NumObjDatas As Integer
Public NumeroHechizos As Integer
Public AllowMultiLogins As Byte
Public IdleLimit As Integer
Public MaxUsers As Integer
Public Minutos As String
Public haciendoBK As Boolean
Public PuedeCrearPersonajes As Integer
Public ServerSoloGMs As Integer


Public EnPausa As Boolean

'*****************ARRAYS PUBLICOS*************************
Public UserList() As User 'USUARIOS
Public Npclist(1 To MAXNPCS) As npc 'NPCS
Public MapData() As MapBlock
Public MapInfo() As MapInfo
Public Hechizos() As tHechizo
Public CharList(1 To MAXCHARS) As Integer
Public ObjData() As ObjData
Public FX() As FXdata
Public SpawnList() As tCriaturasEntrenador
Public LevelSkill(1 To 50) As LevelSkill
Public ArmasHerrero() As Integer
Public ArmadurasHerrero() As Integer
Public ObjCarpintero() As Integer
Public ObjDruida() As Integer
Public ObjSastre() As Integer
Public BanIps As New Collection
Public Parties(1 To MAX_PARTIES) As clsGrupo
Public ModClase(1 To NUMCLASES) As ModClase
Public ModRaza(1 To NUMRAZAS) As ModRaza
Public ModVida(1 To NUMCLASES) As Double
Public DistribucionEnteraVida(1 To 5) As Integer
Public DistribucionSemienteraVida(1 To 4) As Integer
'*********************************************************

Public Nix As WorldPos
Public Ullathorpe As WorldPos
Public Banderbill As WorldPos
Public Arghal As WorldPos

Public Lindos As WorldPos
Public Illiandor As WorldPos
Public Suramei As WorldPos

Public Orac As WorldPos

Public Rinkel As WorldPos
Public Tiama As WorldPos
Public NuevaEsperanza As WorldPos

Public NuevaEsperanzapuerto As WorldPos

Public Prision As WorldPos
Public Libertad As WorldPos

Public Ayuda As New cCola

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef destination As Any, ByVal length As Long)
