Attribute VB_Name = "Protocol"
'************************************************************************
'************************************************************************
'Inmmortal AO v0.1 Beta
'Type object:
'Names:
'Description:
'Author: Jose Ignacio Castelli (Fedudok)
'************************************************************************
'************************************************************************
Option Explicit

''
' TODO : /BANIP y /UNBANIP ya no trabajan con nicks. Esto lo puede mentir en forma local el cliente con un paquete a NickToIp

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

Private Enum Stat
    Incinerado = &H1
    Envenenado = &H2
    Comerciand = &H4
    Trabajando = &H8
    Transformado = &H10
    Ciego = &H20
    Inactivo = &H40
    Silenciado = &H80
End Enum

Private Enum StatEx
    Paralizado = &H1
    Inmovilizado = &H2
    Hombre = &H4
    Mujer = &H8
End Enum

Private Type tFont
    red As Byte
    green As Byte
    blue As Byte
    bold As Boolean
    italic As Boolean
End Type

Private Enum ServerPacketID
    Logged                  ' 0
    RemoveDialogs           ' 1
    RemoveCharDialog        ' 2
    NavigateToggle          ' 3
    EquitateToggle          ' 4
    Disconnect              ' 5
    CommerceEnd             ' 6
    BankEnd                 ' 7
    CommerceInit            ' 8
    BankInit                ' 9
    UserCommerceInit        ' 10
    UserCommerceEnd         ' 11
    ShowBlacksmithForm      ' 12
    ShowCarpenterForm       ' 13
    ShowAlquimiaForm        ' 14
    ShowSastreForm          ' 15
    NPCSwing                ' 16
    NPCKillUser             ' 17
    BlockedWithShieldUser   ' 18
    BlockedWithShieldOther  ' 19
    UserSwing               ' 20
    SafeModeOn              ' 21
    SafeModeOff             ' 22
    NobilityLost            ' 23
    CantUseWhileMeditating  ' 24
    UpdateSta               ' 25
    UpdateMana              ' 26
    UpdateHP                ' 27
    UpdateGold              ' 28
    UpdateExp               ' 29
    ChangeMap               ' 30
    posUpdate               ' 31
    NPCHitUser              ' 32
    UserHitNPC              ' 33
    UserAttackedSwing       ' 34
    UserHittedByUser        ' 35
    UserHittedUser          ' 36
    ChatOverHead            ' 37
    ConsoleMsg              ' 38
    GuildChat               ' 39
    ShowMessageBox          ' 40
    UserIndexInServer       ' 41
    UserCharIndexInServer   ' 42
    CharacterCreate         ' 43
    CharacterRemove         ' 44
    CharacterMove           ' 45
    ForceCharMove           ' 46
    CharacterChange         ' 47
    CharStatus              ' 48
    ObjectCreate            ' 49
    ObjectDelete            ' 50
    BlockPosition           ' 51
    PlayMidi                ' 52
    PlayWave                ' 53
    guildList               ' 54
    AreaChanged             ' 55
    PauseToggle             ' 56
    CreateFX                ' 57
    CreateFXMap             ' 58
    UpdateUserStats         ' 59
    WorkRequestTarget       ' 60
    ChangeInventorySlot     ' 61
    ChangeBankSlot          ' 62
    ChangeSpellSlot         ' 63
    atributes               ' 64
    BlacksmithWeapons       ' 65
    BlacksmithArmors        ' 66
    CarpenterObjects        ' 67
    SastreObjects           ' 68
    AlquimiaObjects         ' 69
    RestOK                  ' 70
    ErrorMsg                ' 71
    Blind                   ' 72
    Dumb                    ' 73
    ShowSignal              ' 74
    ChangeNPCInventorySlot  ' 75
    ShowGuildFundationForm  ' 76
    ParalizeOK              ' 77
    ShowUserRequest         ' 78
    TradeOK                 ' 79
    BankOK                  ' 80
    ChangeUserTradeSlot     ' 81
    UpdateHungerAndThirst   ' 82
    MiniStats               ' 83
    AddForumMsg             ' 84
    ShowForumForm           ' 85
    SetInvisible            ' 86
    MeditateToggle          ' 87
    BlindNoMore             ' 88
    DumbNoMore              ' 89
    SendSkills              ' 90
    TrainerCreatureList     ' 91
    Pong                    ' 92
    UpdateTagAndStatus      ' 93
    SpawnList               ' 94
    ShowSOSForm             ' 95
    ShowGMPanelForm         ' 96
    UserNameList            ' 97
    AddPJ                   ' 98
    ShowAccount             ' 99
    CharacterInfo           ' 100
    GuildLeaderInfo         ' 101
    GuildDetails            ' 102
    Fuerza                  ' 103
    Agilidad                ' 104
    Subasta                 ' 105
    ParticleCreate          ' 106
    CharParticleCreate      ' 107
    DestParticle            ' 108
    DestCharParticle        ' 109
    hora                    ' 110
    Grupo                   ' 111
    ShowGrupoForm           ' 112
    Messages                ' 113
    showCorreoForm          ' 114
    AddCorreoMsg            ' 115
    ShowFamiliarForm        ' 116
    CharMsgStatus           ' 117
    MensajeSigno            ' 118
    Disconnect2             ' 119
End Enum

Private Enum ClientPacketID
    ConnectAccount          '0
    CreateNewAccount        '1
    LoginExistingChar       '2
    LoginNewChar            '3
    Talk                    '4
    Whisper                 '5
    Walk                    '6
    RequestPositionUpdate   '7
    Attack                  '8
    PickUp                  '9
    CombatModeToggle        '10
    SafeToggle              '11
    RequestGuildLeaderInfo  '12
    RequestEstadisticas     '13
    CommerceEnd             '14
    UserCommerceEnd         '15
    BankEnd                 '16
    UserCommerceOk          '17
    UserCommerceReject      '18
    Drop                    '19
    CastSpell               '20
    LeftClick               '21
    DoubleClick             '22
    Work                    '23
    UseItem                 '24
    CraftBlacksmith         '25
    CraftCarpenter          '26
    Craftalquimia           '27
    CraftSastre             '28
    WorkLeftClick           '29
    CreateNewGuild          '30
    SpellInfo               '31
    EquipItem               '32
    ChangeHeading           '33
    ModifySkills            '34
    Train                   '35
    CommerceBuy             '36
    BankExtractItem         '37
    CommerceSell            '38
    BankDeposit             '39
    ForumPost               '40
    MoveSpell               '41
    ClanCodexUpdate         '42
    UserCommerceOffer       '43
    GuildRequestJoinerInfo  '44
    GuildNewWebsite         '45
    GuildAcceptNewMember    '46
    GuildRejectNewMember    '47
    GuildKickMember         '48
    GuildUpdateNews         '49
    GuildMemberInfo         '50
    GuildRequestMembership  '51
    GuildRequestDetails     '52
    Online                  '53
    Quit                    '54
    GuildLeave              '55
    RequestAccountState     '56
    PetStand                '57
    PetFollow               '58
    TrainList               '59
    Rest                    '60
    Meditate                '61
    Resucitate              '62
    Heal                    '63
    Help                    '64
    CommerceStart           '65
    BankStart               '66
    Enlist                  '67
    Information             '68
    Reward                  '69
    UpTime                  '70
    GrupoLeave              '71
    GrupoKick               '72
    GuildMessage            '73
    GrupoMessage            '74
    CentinelReport          '75
    GuildOnline             '76
    RoleMasterRequest       '77
    GMRequest               '78
    bugReport               '79
    ChangeDescription       '80
    Gamble                  '81
    LeaveFaction            '82
    BankExtractGold         '83
    BankTransferGold        '84
    BankDepositGold         '85
    Denounce                '86
    GuildFundate            '87
    Ping                    '88
    Casamiento              '89
    Acepto                  '90
    Divorcio                '91
    MessagesGM              '92
    Subasta                 '93
    RequestGrupo            '94
    Duelo                   '95
    BorrarMensaje           '96
    ExtraerItem             '97
    EnviarMensaje           '98
    AdoptarMascota          '99
    DelClan                 '100
    ChatFaccion             '101
    DragAndDrop             '102
    Hogar                   '103
    Participar              '104
    Pena                    '105
    RequestStats            '106 /EST
    Friends                 '107 /FADD /FDEL /FLIST
End Enum

Public Enum gMessages
    GMMessage               '/GMSG
    showName                '/SHOWNAME
    OnlineArmada            '/ONLINEREAL
    OnlineCaos              '/ONLINECAOS
    OnlineMilicia           '/ONLINEMI
    GoNearby                '/IRCERCA
    comment                 '/REM
    serverTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
    WarpMeToTarget          '/TELEPLOC
    WarpChar                '/TELEP
    Silence                 '/SILENCIAR
    SOSShowList             '/SHOW SOS
    SOSRemove               'SOSDONE
    GoToChar                '/IRA
    invisible               '/INVISIBLE
    GMPanel                 '/PANELGM
    RequestUserList         'LISTUSU
    Working                 '/TRABAJANDO
    Hiding                  '/OCULTANDO
    Jail                    '/CARCEL
    KillNPC                 '/RMATA
    WarnUser                '/ADVERTENCIA
    EditChar                '/MOD
    ReviveChar              '/REVIVIR
    OnlineGM                '/ONLINEGM
    OnlineMap               '/ONLINEMAP
    Kick                    '/ECHAR
    Execute                 '/EJECUTAR
    BanChar                 '/BAN
    UnbanChar               '/UNBAN
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNPCInventory       '/RESETINV
    CleanWorld              '/LIMPIAR
    ServerMessage           '/RMSG
    NickToIP                '/NICK2IP
    IPToNick                '/IP2NICK
    GuildOnlineMembers      '/ONCLAN
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    SetCharDescription      '/SETDESC
    ForceMIDIToMap          '/FORCEMIDIMAP
    ForceWAVEToMap          '/FORCEWAVMAP
    TalkAsNPC               '/TALKAS
    DestroyAllItemsInArea   '/MASSDEST
    ItemsInTheFloor         '/PISO
    MakeDumb                '/ESTUPIDO
    MakeDumbNoMore          '/NOESTUPIDO
    DumpIPTables            '/DUMPSECURITY
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no args
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    GuildMemberList         '/MIEMBROSCLAN
    ShowGuildMessages       '/SHOWCMSG
    GuildBan                '/BANCLAN
    BanIP                   '/BANIP
    UnbanIP                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    ChaosLegionKick         '/NOCAOS
    RoyalArmyKick           '/NOREAL
    MiliciaKick             '/NOMILI
    ForceMIDIAll            '/FORCEMIDI
    ForceWAVEAll            '/FORCEWAV
    TileBlockedToggle       '/BLOQ
    KillNPCNoRespawn        '/MATA
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    SystemMessage           '/SMSG
    CreateNPC               '/ACC
    CreateNPCWithRespawn    '/RACC
    NavigateToggle          '/NAVE
    ServerOpenToUsersToggle '/HABILITAR
    TurnOffServer           '/APAGAR
    TurnCriminal            '/CONDEN
    ResetFactions           '/RAJAR
    RemoveCharFromGuild     '/RAJARCLAN
    ToggleCentinelActivated '/CENTINELAACTIVADO
    DoBackUp                '/DOBACKUP
    Ignored                 '/IGNORADO
    CheckSlot               '/SLOT
    KickAllChars            '/ECHARTODOSPJS
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    Restart                 '/REINICIAR
    SaveMap                 '/GUARDAMAPA
    ChangeMapInfoPK         '/MODMAPINFO PK
    ChangeMapInfoBackup     '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand       '/MODMAPINFO TERRENO
    ChangeMapInfoZone       '/MODMAPINFO ZONA
    SaveChars               '/GRABAR
    CleanSOS                '/BORRAR SOS
    CancelTorneo            '/CANCELTORNEO
    CrearTorneo             '/CREARTORNEO
    Pejotas                 '/PEJOTAS
    SlashSlash              '// <comando>
End Enum


Public Enum FontTypeNames
    FONTTYPE_TALK '255~255~255~0~0
    FONTTYPE_FIGHT '255~0~0~1~0
    FONTTYPE_WARNING '32~51~223~1~1
    FONTTYPE_INFO '65~190~156~0~0
    FONTTYPE_VENENO '0~255~0~0~0
    FONTTYPE_GUILD '255~255~255~1~0
    FONTTYPE_TALKITALIC '255~255~255~0~1
    FONTTYPE_SERVER '0~185~0~0~0
    FONTTYPE_CLAN '228~199~27~0~0
    FONTTYPE_RED '255~0~0~0~0
    FONTTYPE_BROWNB '204~193~115~1~0
    FONTTYPE_BROWNI '204~193~115~0~1
    FONTTYPE_PRIVADO '182~226~29~0~0
    FONTTYPE_GLOBAL '139~248~244~0~1
    FONTTYPE_GRUPO '0~128~128~0~0
    FONTTYPE_FACCION '228~199~27~0~0
    
    FONTTYPE_FACCION_IMPE '0~80~200~1~1
    FONTTYPE_FACCION_REPU '243~147~1~1~1
    FONTTYPE_FACCION_CAOS '197~0~5~1~1
End Enum

Public FontTypes(21) As tFont


                                                 
                          
''
' Initializes the fonts array

Public Sub InitFonts()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
        
    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        .red = 255
        .green = 255
        .blue = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
        .red = 255
        .bold = 1
    End With
        
    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
        .red = 32
        .green = 51
        .blue = 223
        .bold = 1
        .italic = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        .red = 65
        .green = 190
        .blue = 156
    End With
    

    With FontTypes(FontTypeNames.FONTTYPE_VENENO)
        .green = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILD)
        .red = 255
        .green = 255
        .blue = 255
        .bold = 1
    End With

    With FontTypes(FontTypeNames.FONTTYPE_TALKITALIC)
        .red = 255
        .green = 255
        .blue = 255
        .bold = 0
        .italic = 1
    End With
        
    With FontTypes(FontTypeNames.FONTTYPE_SERVER)
        .green = 185
    End With

    With FontTypes(FontTypeNames.FONTTYPE_CLAN)
        .red = 228
        .green = 199
        .blue = 27
    End With
        
    With FontTypes(FontTypeNames.FONTTYPE_RED)
        .red = 255
    End With

    With FontTypes(FontTypeNames.FONTTYPE_BROWNB)
        .red = 204
        .green = 193
        .blue = 115
        .bold = 1
        .italic = 0
    End With

    With FontTypes(FontTypeNames.FONTTYPE_BROWNI)
        .red = 204
        .green = 193
        .blue = 115
        .italic = 1
    End With

    With FontTypes(FontTypeNames.FONTTYPE_PRIVADO)
        .red = 182
        .green = 226
        .blue = 29
    End With
   
    With FontTypes(FontTypeNames.FONTTYPE_GLOBAL)
        .red = 139
        .green = 248
        .blue = 244
        .italic = 1
    End With
        
    With FontTypes(FontTypeNames.FONTTYPE_GRUPO)
        .green = 128
        .blue = 128
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FACCION)
        .red = 228
        .green = 199
        .blue = 27
        .bold = 0
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FACCION_IMPE)
        .red = 0
        .green = 80
        .blue = 200
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FACCION_REPU)
        .red = 243
        .green = 147
        .blue = 1
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FACCION_CAOS)
        .red = 197
        .green = 0
        .blue = 5
        .bold = 1
    End With
End Sub

''
' Handles incoming data.

Public Sub HandleIncomingData()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

On Error Resume Next

    Dim package As Byte
    
    package = incomingData.PeekByte()


    Select Case package
    
    
    
        Case ServerPacketID.Logged                  ' LOGGED
            Call HandleLogged
        
        Case ServerPacketID.RemoveDialogs           ' QTDL
            Call HandleRemoveDialogs
        
        Case ServerPacketID.RemoveCharDialog        ' QDL
            Call HandleRemoveCharDialog
        
        Case ServerPacketID.NavigateToggle          ' NAVEG
            Call HandleNavigateToggle
            
        Case ServerPacketID.EquitateToggle
            Call HandleEquitateToggle
        
        Case ServerPacketID.Disconnect              ' FINOK
            Call HandleDisconnect
        
        Case ServerPacketID.CommerceEnd             ' FINCOMOK
            Call HandleCommerceEnd
        
        Case ServerPacketID.BankEnd                 ' FINBANOK
            Call HandleBankEnd
        
        Case ServerPacketID.CommerceInit            ' INITCOM
            Call HandleCommerceInit
        
        Case ServerPacketID.BankInit                ' INITBANCO
            Call HandleBankInit
        
        Case ServerPacketID.UserCommerceInit        ' INITCOMUSU
            Call HandleUserCommerceInit
        
        Case ServerPacketID.UserCommerceEnd         ' FINCOMUSUOK
            Call HandleUserCommerceEnd
        
        Case ServerPacketID.ShowBlacksmithForm      ' SFH
            Call HandleShowBlacksmithForm
        
        Case ServerPacketID.ShowCarpenterForm       ' SFC
            Call HandleShowCarpenterForm
        
        Case ServerPacketID.ShowSastreForm
            Call HandleShowSastreForm
        
        Case ServerPacketID.ShowAlquimiaForm
            Call HandleShowalquimiaForm
        
        Case ServerPacketID.NPCSwing                ' N1
            Call HandleNPCSwing
        
        Case ServerPacketID.NPCKillUser             ' 6
            Call HandleNPCKillUser
        
        Case ServerPacketID.BlockedWithShieldUser   ' 7
            Call HandleBlockedWithShieldUser
        
        Case ServerPacketID.BlockedWithShieldOther  ' 8
            Call HandleBlockedWithShieldOther
        
        Case ServerPacketID.UserSwing               ' U1
            Call HandleUserSwing
            
        Case ServerPacketID.SafeModeOn              ' SEGON
            Call HandleSafeModeOn
        
        Case ServerPacketID.SafeModeOff             ' SEGOFF
            Call HandleSafeModeOff
        
        Case ServerPacketID.NobilityLost            ' PN
            Call HandleNobilityLost
        
        Case ServerPacketID.CantUseWhileMeditating  ' M!
            Call HandleCantUseWhileMeditating
        
        Case ServerPacketID.UpdateSta               ' ASS
            Call HandleUpdateSta
        
        Case ServerPacketID.UpdateMana              ' ASM
            Call HandleUpdateMana
        
        Case ServerPacketID.UpdateHP                ' ASH
            Call HandleUpdateHP
        
        Case ServerPacketID.UpdateGold              ' ASG
            Call HandleUpdateGold
        
        Case ServerPacketID.UpdateExp               ' ASE
            Call HandleUpdateExp
        
        Case ServerPacketID.ChangeMap               ' CM
            Call HandleChangeMap
        
        Case ServerPacketID.posUpdate               ' PU
            Call HandlePosUpdate
        
        Case ServerPacketID.NPCHitUser              ' N2
            Call HandleNPCHitUser
        
        Case ServerPacketID.UserHitNPC              ' U2
            Call HandleUserHitNPC
        
        Case ServerPacketID.UserAttackedSwing       ' U3
            Call HandleUserAttackedSwing
        
        Case ServerPacketID.UserHittedByUser        ' N4
            Call HandleUserHittedByUser
        
        Case ServerPacketID.UserHittedUser          ' N5
            Call HandleUserHittedUser
        
        Case ServerPacketID.ChatOverHead            ' ||
            Call HandleChatOverHead
        
        Case ServerPacketID.ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
            Call HandleConsoleMessage
        
        Case ServerPacketID.GuildChat               ' |+
            Call HandleGuildChat
        
        Case ServerPacketID.ShowMessageBox          ' !!
            Call HandleShowMessageBox
        
        Case ServerPacketID.UserIndexInServer       ' IU
            Call HandleUserIndexInServer
        
        Case ServerPacketID.UserCharIndexInServer   ' IP
            Call Handlechar_currentInServer
        
        Case ServerPacketID.CharacterCreate         ' CC
            Call HandleCharacterCreate
        
        Case ServerPacketID.CharacterRemove         ' BP
            Call HandleCharacterRemove
        
        Case ServerPacketID.CharacterMove           ' MP, +, * and _ '
            Call HandleCharacterMove
            
        Case ServerPacketID.ForceCharMove
            Call HandleForceCharMove
        
        Case ServerPacketID.CharacterChange         ' CP
            Call HandleCharacterChange
        
        Case ServerPacketID.CharStatus
            Call HandleCharStatus
        
        Case ServerPacketID.ObjectCreate            ' HO
            Call HandleObjectCreate
        
        Case ServerPacketID.ObjectDelete            ' BO
            Call HandleObjectDelete
        
        Case ServerPacketID.BlockPosition           ' BQ
            Call HandleBlockPosition
            
        Case ServerPacketID.PlayMidi                ' TM
            Call HandlePlayMusic
        
        Case ServerPacketID.PlayWave                ' TW
            Call HandlePlayWave
        
        Case ServerPacketID.guildList               ' GL
            Call HandleGuildList
        
        Case ServerPacketID.AreaChanged             ' CA
            Call HandleAreaChanged
        
        Case ServerPacketID.PauseToggle             ' BKW
            Call HandlePauseToggle
        
        Case ServerPacketID.CreateFX                ' CFX
            Call HandleCreateFX
        
        Case ServerPacketID.CreateFXMap
            Call HandleCreateFXMap
            
        Case ServerPacketID.UpdateUserStats         ' EST
            Call HandleUpdateUserStats
        
        Case ServerPacketID.WorkRequestTarget       ' T01
            Call HandleWorkRequestTarget
        
        Case ServerPacketID.ChangeInventorySlot     ' CSI
            Call HandleChangeInventorySlot
        
        Case ServerPacketID.ChangeBankSlot          ' SBO
            Call HandleChangeBankSlot
        
        Case ServerPacketID.ChangeSpellSlot         ' SHS
            Call HandleChangeSpellSlot
        
        Case ServerPacketID.atributes               ' ATR
            Call HandleAtributes
        
        Case ServerPacketID.BlacksmithWeapons       ' LAH
            Call HandleBlacksmithWeapons
        
        Case ServerPacketID.BlacksmithArmors        ' LAR
            Call HandleBlacksmithArmors
        
        Case ServerPacketID.CarpenterObjects        ' OBR
            Call HandleCarpenterObjects
            
        Case ServerPacketID.AlquimiaObjects
            Call HandleAlquimiaObjects
            
        Case ServerPacketID.SastreObjects
            Call HandleSastreObjects
        
        
        Case ServerPacketID.RestOK                  ' DOK
            Call HandleRestOK
        
        Case ServerPacketID.ErrorMsg                ' ERR
            Call HandleErrorMessage
        
        Case ServerPacketID.Blind                   ' CEGU
            Call HandleBlind
        
        Case ServerPacketID.Dumb                    ' DUMB
            Call HandleDumb
        
        Case ServerPacketID.ShowSignal              ' MCAR
            Call HandleShowSignal
        
        Case ServerPacketID.ChangeNPCInventorySlot  ' NPCI
            Call HandleChangeNPCInventorySlot
        
        Case ServerPacketID.UpdateHungerAndThirst   ' EHYS
            Call HandleUpdateHungerAndThirst
        
        Case ServerPacketID.MiniStats               ' MEST
            Call HandleMiniStats
            
        Case ServerPacketID.AddForumMsg             ' FMSG
            Call HandleAddForumMessage
        
        Case ServerPacketID.AddCorreoMsg
            Call HandleAddCorreoMessage
        
        Case ServerPacketID.ShowForumForm           ' MFOR
            Call HandleShowForumForm
            
        Case ServerPacketID.showCorreoForm
            Call HandleShowCorreoForm
        
        Case ServerPacketID.SetInvisible            ' NOVER
            Call HandleSetInvisible
        
        Case ServerPacketID.MeditateToggle          ' MEDOK
            Call HandleMeditateToggle
        
        Case ServerPacketID.BlindNoMore             ' NSEGUE
            Call HandleBlindNoMore
        
        Case ServerPacketID.DumbNoMore              ' NESTUP
            Call HandleDumbNoMore
        
        Case ServerPacketID.SendSkills              ' SKILLS
            Call HandleSendSkills
        
        Case ServerPacketID.TrainerCreatureList     ' LSTCRI
            Call HandleTrainerCreatureList
        
        Case ServerPacketID.CharacterInfo           ' CHRINFO
            Call HandleCharacterInfo
        
        Case ServerPacketID.GuildLeaderInfo         ' LEADERI
            Call HandleGuildLeaderInfo
        
        Case ServerPacketID.GuildDetails            ' CLANDET
            Call HandleGuildDetails
        
        Case ServerPacketID.ShowGuildFundationForm  ' SHOWFUN
            Call HandleShowGuildFundationForm
        
        Case ServerPacketID.ParalizeOK              ' PARADOK
            Call HandleParalizeOK
        
        Case ServerPacketID.ShowUserRequest         ' PETICIO
            Call HandleShowUserRequest
        
        Case ServerPacketID.TradeOK                 ' TRANSOK
            Call HandleTradeOK
        
        Case ServerPacketID.BankOK                  ' BANCOOK
            Call HandleBankOK
        
        Case ServerPacketID.ChangeUserTradeSlot     ' COMUSUINV
            Call HandleChangeUserTradeSlot
            
        Case ServerPacketID.Pong
            Call HandlePong
        
        Case ServerPacketID.UpdateTagAndStatus
            Call HandleUpdateTagAndStatus

        
        '*******************
        'GM messages
        '*******************
        Case ServerPacketID.SpawnList               ' SPL
            Call HandleSpawnList
        
        Case ServerPacketID.ShowSOSForm             ' RSOS and MSOS
            Call HandleShowSOSForm
        
        Case ServerPacketID.ShowGMPanelForm         ' ABPANEL
            Call HandleShowGMPanelForm
        
        Case ServerPacketID.UserNameList            ' LISTUSU
            Call HandleUserNameList
            
        Case ServerPacketID.ParticleCreate          ' PC
            Call HandleParticle
        
        Case ServerPacketID.CharParticleCreate      ' CPC
            Call HandleCharParticle
            
        Case ServerPacketID.DestParticle            ' DP
            Call HandleDestParticle
        
        Case ServerPacketID.DestCharParticle        ' DPC
            Call HandleDestCharParticle
            
        Case ServerPacketID.AddPJ
            Call HandleAddPj
            
        Case ServerPacketID.ShowAccount
            Call HandleShowAccount
            
        Case ServerPacketID.Fuerza                  ' PF
            Call HandleFuerza
            
        Case ServerPacketID.Agilidad                ' PG
            Call HandleAgilidad
            
        Case ServerPacketID.Subasta
            Call HandleSubastRequest
            
        Case ServerPacketID.hora
            Call HandleHora
        
        Case ServerPacketID.Grupo
            Call HandleGrupo
            
        Case ServerPacketID.ShowGrupoForm
            Call HandleGrupoForm
        
        Case ServerPacketID.Messages
            Call HandleMessages
            
        Case ServerPacketID.ShowFamiliarForm
            Call HandleShorFamiliarForm
            
        Case ServerPacketID.CharMsgStatus
            Call HandleCharMsgStatus
            
        Case ServerPacketID.MensajeSigno
            Call HandleMensajeSigno
            
        Case ServerPacketID.Disconnect2
            Call HandleDisconnect2
            
            
        Case Else
            'ERROR : Abort!
            Exit Sub

    End Select
    
    'Done with this packet, move on to next one
    If incomingData.length > 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then
        Err.Clear
        Call HandleIncomingData
    End If
End Sub

''
' Handles the Logged message.

Private Sub HandleLogged()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Security.Redundance = incomingData.ReadByte()
    
    ' Variable initialization
    Nombres = True
    
    'Set connected state
    Call SetConnected
    

End Sub

''
' Handles the RemoveDialogs message.

Private Sub HandleRemoveDialogs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call TileEngine.Char_Dialog_Remove_All
End Sub

''
' Handles the RemoveCharDialog message.

Private Sub HandleRemoveCharDialog()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check if the packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call TileEngine.Char_Dialog_Remove(incomingData.ReadInteger())
End Sub

''
' Handles the NavigateToggle message.

Private Sub HandleNavigateToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserNavegando = Not UserNavegando
End Sub
Private Sub HandleEquitateToggle()
    Call incomingData.ReadByte
    
    UserMontando = Not UserMontando
    
    If UserMontando = True Then
        TileEngine.Engine_Scroll_Pixels 6.4
    Else
        TileEngine.Engine_Scroll_Pixels 5.2
    End If
End Sub

''
' Handles the Disconnect message.

Private Sub HandleDisconnect()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte

    CleanClient
    
    'Unload all forms except frmMain and frmConnect
    Dim frm As Form
    
    For Each frm In Forms
        If frm.name <> frmMain.name And frm.name <> frmConnect.name Then
            Unload frm
        End If
    Next
    
        frmMain.Socket1.Disconnect
        EstadoLogin = E_MODO.ConectarCuenta
        frmMain.Socket1.Connect
        DoEvents

    
End Sub

Private Sub HandleDisconnect2()
'***************************************************
'Author: Jose Ignacio Castelli
'Last Modification: 22/2/11
'Explicacion: Es para desconectar definitivamente al haber aplicado Sistema de cuentas
'***************************************************

    
    'Remove packet ID
    Call incomingData.ReadByte
    Unload frmConnect
    perm = True
   
    frmMain.Socket1.Disconnect
    frmConnect.Visible = True
    frmPanelAccount.Visible = False
    If frmMain.Visible = True Then
    Unload frmMain
    End If

    
End Sub

''
' Handles the CommerceEnd message.

Private Sub HandleCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Clear item's list
    frmComerciar.List1(0).Clear
    frmComerciar.List1(1).Clear
    
    'Reset vars
    Comerciando = False
    
    'Hide form
    Unload frmComerciar
End Sub

''
' Handles the BankEnd message.

Private Sub HandleBankEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmBancoObj.List1(0).Clear
    frmBancoObj.List1(1).Clear
    
    Unload frmBancoObj
    Comerciando = False
End Sub

''
' Handles the CommerceInit message.

Private Sub HandleCommerceInit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Fill our inventory list
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.OBJIndex(i) <> 0 Then
            frmComerciar.List1(1).AddItem Inventario.ItemName(i)
        Else
            frmComerciar.List1(1).AddItem "(Nada)"
        End If
    Next i
    
    'Set state and show form
    Comerciando = True
    frmComerciar.Show , frmMain
End Sub

''
' Handles the BankInit message.

Private Sub HandleBankInit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    If incomingData.ReadByte Then
        Call frmGoliath.ParseBancoInfo(incomingData.ReadLong, incomingData.ReadByte)
    Else
        Call frmBancoObj.List1(1).Clear
        
        'Fill the inventory list
        For i = 1 To MAX_INVENTORY_SLOTS
            If Inventario.OBJIndex(i) <> 0 Then
                frmBancoObj.List1(1).AddItem Inventario.ItemName(i)
            Else
                frmBancoObj.List1(1).AddItem "(Nada)"
            End If
        Next i
        
        Call frmBancoObj.List1(0).Clear
        
        'Fill the bank list
        For i = 1 To MAX_BANCOINVENTORY_SLOTS
            If UserBancoInventory(i).OBJIndex <> 0 Then
                frmBancoObj.List1(0).AddItem UserBancoInventory(i).name
            Else
                frmBancoObj.List1(0).AddItem "(Nada)"
            End If
        Next i

        'Set state and show form
        Comerciando = True
        frmBancoObj.Show , frmMain
    End If
End Sub

''
' Handles the UserCommerceInit message.

Private Sub HandleUserCommerceInit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Clears lists if necessary
    If frmComerciarUsu.List1.ListCount > 0 Then frmComerciarUsu.List1.Clear
    If frmComerciarUsu.List2.ListCount > 0 Then frmComerciarUsu.List2.Clear
    
    'Fill inventory list
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.OBJIndex(i) <> 0 Then
            frmComerciarUsu.List1.AddItem Inventario.ItemName(i)
            frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = Inventario.Amount(i)
        Else
            frmComerciarUsu.List1.AddItem ""
            frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = 0
        End If
    Next i
    
    'Set state and show form
    Comerciando = True
    frmComerciarUsu.Show , frmMain
End Sub

''
' Handles the UserCommerceEnd message.

Private Sub HandleUserCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Clear the lists
    frmComerciarUsu.List1.Clear
    frmComerciarUsu.List2.Clear
    
    'Destroy the form and reset the state
    Unload frmComerciarUsu
    Comerciando = False
End Sub

''
' Handles the ShowBlacksmithForm message.

Private Sub HandleShowBlacksmithForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmHerrero.Show , frmMain

End Sub

''
' Handles the ShowCarpenterForm message.

Private Sub HandleShowCarpenterForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmCarp.Show , frmMain
    
End Sub


Private Sub HandleShowSastreForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte

    frmSastre.Show , frmMain
    
End Sub

Private Sub HandleShowalquimiaForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte

    frmAlquimia.Show , frmMain
    
End Sub


''
' Handles the NPCSwing message.

Private Sub HandleNPCSwing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, False)
End Sub

''
' Handles the NPCKillUser message.

Private Sub HandleNPCKillUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)
End Sub

''
' Handles the BlockedWithShieldUser message.

Private Sub HandleBlockedWithShieldUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
End Sub

''
' Handles the BlockedWithShieldOther message.

Private Sub HandleBlockedWithShieldOther()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
End Sub

''
' Handles the UserSwing message.

Private Sub HandleUserSwing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, False)
End Sub


''
' Handles the SafeModeOn message.

Private Sub HandleSafeModeOn()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmMain.DibujarSeguro
    Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_SEGURO_ACTIVADO, 255, 0, 0, True, False, False)
End Sub

''
' Handles the SafeModeOff message.

Private Sub HandleSafeModeOff()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmMain.DesDibujarSeguro
    Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_SEGURO_DESACTIVADO, 255, 0, 0, True, False, False)
End Sub

''
' Handles the NobilityLost message.

Private Sub HandleNobilityLost()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecChat, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, False)
End Sub

''
' Handles the CantUseWhileMeditating message.

Private Sub HandleCantUseWhileMeditating()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecChat, MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, False)
End Sub

''
' Handles the UpdateSta message.

Private Sub HandleUpdateSta()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinSTA = incomingData.ReadInteger()
    frmMain.STAShp.width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 91)
    frmMain.lblST.Caption = UserMinSTA & "/" & UserMaxSTA
End Sub

''
' Handles the UpdateMana message.

Private Sub HandleUpdateMana()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinMAN = incomingData.ReadInteger()
    
    If UserMaxMAN > 0 Then
        frmMain.MANShp.width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 91)
        frmMain.lblMP.Caption = UserMinMAN & "/" & UserMaxMAN
    Else
        frmMain.lblMP.Caption = ""
        frmMain.MANShp.width = 0
    End If
End Sub

' Handles the UpdateHungerAndThirst message.

Private Sub HandleUpdateHungerAndThirst()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMaxAGU = 100
    UserMaxHAM = 100
    UserMinAGU = incomingData.ReadByte()
    UserMinHAM = incomingData.ReadByte()
    
    frmMain.lblHAM.Caption = UserMinHAM & "/" & UserMaxHAM
    frmMain.lblSED.Caption = UserMinAGU & "/" & UserMaxAGU
    
    frmMain.AGUAsp.width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 91)
    frmMain.COMIDAsp.width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 91)
End Sub
''
' Handles the UpdateHP message.

Private Sub HandleUpdateHP()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinHP = incomingData.ReadInteger()
    frmMain.Hpshp.width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 91)
    frmMain.lblHP.Caption = UserMinHP & "/" & UserMaxHP
    
    'Is the user alive??
    If UserMinHP = 0 Then
        UserEstado = 1
    Else
        UserEstado = 0
    End If
End Sub

''
' Handles the UpdateGold message.

Private Sub HandleUpdateGold()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'- 08/14/07: Added GldLbl color variation depending on User Gold and Level
'***************************************************
    'Check packet is complete
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserGLD = incomingData.ReadLong()
    
    frmMain.GldLbl.Caption = UserGLD
End Sub

''
' Handles the UpdateExp message.

Private Sub HandleUpdateExp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserExp = incomingData.ReadLong()
    
    If Not UserExp = 0 And Not UserPasarNivel = 0 Then
        frmMain.lblExp.Caption = CStr(Round(UserExp * 100 / UserPasarNivel) & "%") & " (" & UserExp & "/" & UserPasarNivel & ")"
    Else
        frmMain.lblExp.Caption = "¡Nivel maximo!"
    End If
        
    If Not UserExp = 0 And Not UserPasarNivel = 0 Then
        frmMain.ExpShp.width = Round(((UserExp / 100) / (UserPasarNivel / 100)) * 120)
    Else
        frmMain.ExpShp.Visible = False
        'frmMain.ExpShp.width = 120
    End If
End Sub

''
' Handles the ChangeMap message.

Private Sub HandleChangeMap()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMap = incomingData.ReadInteger()

    Call incomingData.ReadInteger

    Call TileEngine.Map_Load(UserMap)
End Sub

''
' Handles the PosUpdate message.

Private Sub HandlePosUpdate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Remove char from old position
    If MapData(UserPos.X, UserPos.Y).CharIndex = char_current Then
        MapData(UserPos.X, UserPos.Y).CharIndex = 0
    End If
    
    'Set new pos
    UserPos.X = incomingData.ReadByte()
    UserPos.Y = incomingData.ReadByte()
    
    'Set char
    MapData(UserPos.X, UserPos.Y).CharIndex = char_current
    charlist(char_current).Pos = UserPos
    
    'Are we under a roof?
    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4 Or _
                MapData(UserPos.X, UserPos.Y).Trigger >= 20, True, False)
                
    
    Call DibujarMiniMapPos
End Sub

''
' Handles the NPCHitUser message.

Private Sub HandleNPCHitUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Select Case incomingData.ReadByte()
        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_GOLPE_CABEZA & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, False)
        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_GOLPE_BRAZO_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, False)
        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_GOLPE_BRAZO_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, False)
        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_GOLPE_PIERNA_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, False)
        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_GOLPE_PIERNA_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, False)
        Case bTorso
            Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_GOLPE_TORSO & CStr(incomingData.ReadInteger() & "!!"), 255, 0, 0, True, False, False)
    End Select
End Sub
''
' Handles the UserHitNPC message.

Private Sub HandleUserHitNPC()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_GOLPE_CRIATURA_1 & CStr(incomingData.ReadLong()) & MENSAJE_2, 255, 0, 0, True, False, False)
End Sub

''
' Handles the UserAttackedSwing message.

Private Sub HandleUserAttackedSwing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_1 & charlist(incomingData.ReadInteger()).nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)
End Sub

''
' Handles the UserHittingByUser message.

Private Sub HandleUserHittedByUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim attacker As String
    
    attacker = charlist(incomingData.ReadInteger()).nombre
    
    Select Case incomingData.ReadByte
        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_CABEZA & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bTorso
            Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_TORSO & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
    End Select
End Sub

''
' Handles the UserHittedUser message.

Private Sub HandleUserHittedUser()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim victim As String
    
    victim = charlist(incomingData.ReadInteger()).nombre
    
    Select Case incomingData.ReadByte
        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_CABEZA & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
        Case bTorso
            Call AddtoRichTextBox(frmMain.RecCombat, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_TORSO & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)
    End Select
End Sub

''
' Handles the ChatOverHead message.

Private Sub HandleChatOverHead()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim chat As String
    Dim CharIndex As Integer
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    Dim name As String, Pos As Integer
    
    chat = buffer.ReadASCIIString()
    CharIndex = buffer.ReadInteger()
    
    r = buffer.ReadByte()
    g = buffer.ReadByte()
    b = buffer.ReadByte()
    
    Pos = InStr(charlist(CharIndex).nombre, "<"): If Pos = 0 Then Pos = Len(charlist(CharIndex).nombre) + 2
    name = Left$(charlist(CharIndex).nombre, Pos - 2)
    
    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If charlist(CharIndex).active Then _
        Call TileEngine.Char_Dialog_Create(CharIndex, chat, D3DColorXRGB(r, g, b))
    
    If buffer.ReadByte = 1 Then
        If Trim(chat) <> "" Then

            If charlist(CharIndex).Priv = 1 Then 'Rene
                Call AddtoRichTextBox(frmMain.RecChat, "[" & name & "] " & chat, 114, 115, 108, 0, 1)
            ElseIf charlist(CharIndex).Priv = 2 Then 'Impe
                Call AddtoRichTextBox(frmMain.RecChat, "[" & name & "] " & chat, 0, 80, 200, 0, 1)
            ElseIf charlist(CharIndex).Priv = 4 Then 'Repu
                Call AddtoRichTextBox(frmMain.RecChat, "[" & name & "] " & chat, 243, 147, 1, 0, 1)
            ElseIf charlist(CharIndex).Priv = 3 Then 'Caos
                Call AddtoRichTextBox(frmMain.RecChat, "[" & name & "] " & chat, 197, 0, 5, 0, 1)
                
            'Add Marius Lideres faccionarios
            ElseIf charlist(CharIndex).Priv = 9 Then 'Lider Impe
                Call AddtoRichTextBox(frmMain.RecChat, "[" & name & "] " & chat, 0, 80, 200, 1, 1)
            ElseIf charlist(CharIndex).Priv = 10 Then 'Lider Repu
                Call AddtoRichTextBox(frmMain.RecChat, "[" & name & "] " & chat, 243, 147, 1, 1, 1)
            ElseIf charlist(CharIndex).Priv = 11 Then 'Lider Caos
                Call AddtoRichTextBox(frmMain.RecChat, "[" & name & "] " & chat, 197, 0, 5, 1, 1)
            '\Add
            ElseIf charlist(CharIndex).Priv = 5 Then ' conse
                Call AddtoRichTextBox(frmMain.RecChat, "[" & name & "] " & chat, 0, 160, 0, 1, 1)
            ElseIf charlist(CharIndex).Priv = 6 Then 'semi
                Call AddtoRichTextBox(frmMain.RecChat, "[" & name & "] " & chat, 0, 160, 0, 1, 1)
            ElseIf charlist(CharIndex).Priv = 7 Then 'dios
                Call AddtoRichTextBox(frmMain.RecChat, "[" & name & "] " & chat, 0, 160, 0, 1, 1)
            ElseIf charlist(CharIndex).Priv = 8 Then ' admin
                Call AddtoRichTextBox(frmMain.RecChat, "[" & name & "] " & chat, 0, 160, 0, 1, 1)
            Else
                Call AddtoRichTextBox(frmMain.RecChat, "[" & name & "] " & chat, 255, 255, 255, 0, 1)
            End If
        End If
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleConsoleMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim chat As String
    Dim fontIndex As Integer
    Dim str As String
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    Dim consola As Byte
    
    chat = buffer.ReadASCIIString()
    consola = buffer.ReadByte()
    fontIndex = buffer.ReadByte()
    
    If consola = 3 Then
        If ChatGlobal = 0 Then
            Exit Sub
        End If
    End If
    
    If fontIndex = FontTypeNames.FONTTYPE_FACCION Then
        If ChatFaccionario = 0 Then
            Exit Sub
        End If
    End If
    
    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)
            If val(str) > 255 Then
                r = 255
            Else
                r = val(str)
            End If
            
            str = ReadField(3, chat, 126)
            If val(str) > 255 Then
                g = 255
            Else
                g = val(str)
            End If
            
            str = ReadField(4, chat, 126)
            If val(str) > 255 Then
                b = 255
            Else
                b = val(str)
            End If
        If consola = 1 Then
            Call AddtoRichTextBox(frmMain.RecChat, Left$(chat, InStr(1, chat, "~") - 1), r, g, b, val(ReadField(5, chat, 126)) <> 0, val(ReadField(6, chat, 126)) <> 0)
        ElseIf consola = 2 Then
            Call AddtoRichTextBox(frmMain.RecCombat, Left$(chat, InStr(1, chat, "~") - 1), r, g, b, val(ReadField(5, chat, 126)) <> 0, val(ReadField(6, chat, 126)) <> 0)
        ElseIf consola = 3 Then
            Call AddtoRichTextBox(frmMain.RecGlobal, Left$(chat, InStr(1, chat, "~") - 1), r, g, b, val(ReadField(5, chat, 126)) <> 0, val(ReadField(6, chat, 126)) <> 0)
        End If
    Else
        With FontTypes(fontIndex)
            If consola = 1 Then
                Call AddtoRichTextBox(frmMain.RecChat, chat, .red, .green, .blue, .bold, .italic)
            ElseIf consola = 2 Then
                Call AddtoRichTextBox(frmMain.RecCombat, chat, .red, .green, .blue, .bold, .italic)
            ElseIf consola = 3 Then
                Call AddtoRichTextBox(frmMain.RecGlobal, chat, .red, .green, .blue, .bold, .italic)
            End If
        End With
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the GuildChat message.

Private Sub HandleGuildChat()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/07/08 (NicoNZ)
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim chat As String
    Dim str As String
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    Dim tmp As Integer
    Dim Cont As Integer
    
    
    chat = buffer.ReadASCIIString()
    
    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)
        If val(str) > 255 Then
            r = 255
        Else
            r = val(str)
        End If
        
        str = ReadField(3, chat, 126)
        If val(str) > 255 Then
            g = 255
        Else
            g = val(str)
        End If
        
        str = ReadField(4, chat, 126)
        If val(str) > 255 Then
            b = 255
        Else
            b = val(str)
        End If
        
        Call AddtoRichTextBox(frmMain.RecChat, Left$(chat, InStr(1, chat, "~") - 1), r, g, b, val(ReadField(5, chat, 126)) <> 0, val(ReadField(6, chat, 126)) <> 0)
    Else
        With FontTypes(FontTypeNames.FONTTYPE_CLAN)
            Call AddtoRichTextBox(frmMain.RecChat, chat, .red, .green, .blue, .bold, .italic)
        End With
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ShowMessageBox message.

Private Sub HandleShowMessageBox()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Dim msg As String
    
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    msg = buffer.ReadASCIIString()
    If buffer.ReadByte = 0 Then
        frmMensaje.msg.Caption = msg
        frmMensaje.Show
    Else
        frmPregunta.Show
        frmPregunta.SetAccion buffer.ReadByte, msg
    End If
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the UserIndexInServer message.

Private Sub HandleUserIndexInServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserIndex = incomingData.ReadInteger()
End Sub

''
' Handles the char_currentInServer message.

Private Sub Handlechar_currentInServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    char_current = incomingData.ReadInteger()
    UserPos = charlist(char_current).Pos
    
    'Are we under a roof?
    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4 Or _
                MapData(UserPos.X, UserPos.Y).Trigger >= 20, True, False)

    Call TileEngine.Char_Refresh(char_current)

    Call DibujarMiniMapPos
End Sub

''
' Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 23 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim CharIndex As Integer
    Dim body As Integer
    Dim Head As Integer
    Dim heading As E_Heading
    Dim X As Byte
    Dim Y As Byte
    Dim Weapon As Integer
    Dim Shield As Integer
    Dim helmet As Integer
    Dim privs As Integer
    Dim Pos As Integer
    
    CharIndex = buffer.ReadInteger()
    body = buffer.ReadInteger()
    Head = buffer.ReadInteger()
    heading = buffer.ReadByte()
    X = buffer.ReadByte()
    Y = buffer.ReadByte()
    Weapon = buffer.ReadInteger()
    Shield = buffer.ReadInteger()
    helmet = buffer.ReadInteger()
    
    
    With charlist(CharIndex)
        Call TileEngine.Char_SetFx(CharIndex, buffer.ReadInteger(), buffer.ReadInteger())
        
        .nombre = buffer.ReadASCIIString()
        Pos = InStr(.nombre, "<")
        If Pos = 0 Then
            .clan = ""
            .offClanX = 0
            
        Else
            .clan = mid$(.nombre, Pos)
            .offClanX = TileEngine.Text_Width(.clan, 1) / 2
            
            .nombre = Left$(.nombre, Pos - 2)
        End If
        
        .offNameX = TileEngine.Text_Width(.nombre, 1) / 2
        
        .Priv = buffer.ReadByte()
        
        .donador = buffer.ReadBoolean()
        
        .Bandera = buffer.ReadByte()
    
    End With
    
    Call TileEngine.Char_Create(CharIndex, body, Head, heading, X, Y, Weapon, Shield, helmet)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the CharacterRemove message.

Private Sub HandleCharacterRemove()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    
    Call TileEngine.Char_Remove(CharIndex)

End Sub

''
' Handles the CharacterMove message.

Private Sub HandleCharacterMove()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim X As Byte
    Dim Y As Byte
    
    CharIndex = incomingData.ReadInteger() Xor (13246 Xor 789)
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    With charlist(CharIndex)
        If .fxIndex >= 40 And .fxIndex <= 49 Then   'If it's meditating, we remove the FX
            .fxIndex = 0
        End If
        
        ' Play steps sounds if the user is not an admin of any kind
        'Mod Marius
        'If .Priv = 1 Or .Priv = 2 Or .Priv = 3 Or .Priv = 4 Or .Priv = 9 Or .Priv = 10 Or .Priv = 11 Then
        'If .Priv <> 1 And .Priv <> 2 And .Priv <> 3 And .Priv <> 4 Then
            Call TileEngine.Char_Pasos_Render(CharIndex)
        'End If
    End With
    
    Call TileEngine.Char_Move_Pos(CharIndex, X, Y)
End Sub
Private Sub HandleForceCharMove()
    
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Direccion As Byte
    
    Direccion = incomingData.ReadByte()

    Call TileEngine.Char_Move_Head(char_current, Direccion)
    Call TileEngine.Engine_MoveScreen(Direccion)

End Sub
Sub HandleCharStatus()
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim Priv As Byte
    
    CharIndex = incomingData.ReadInteger
    Priv = incomingData.ReadInteger
    
    charlist(CharIndex).Priv = Priv

    Select Case Priv
        Case 1 'Gris
            TileEngine.Engine_Long_To_RGB_List charlist(CharIndex).label_color, D3DColorXRGB(175, 175, 175)
        Case 2 'Azul
            TileEngine.Engine_Long_To_RGB_List charlist(CharIndex).label_color, D3DColorXRGB(39, 131, 243)
        Case 3 'Rojo
            TileEngine.Engine_Long_To_RGB_List charlist(CharIndex).label_color, D3DColorXRGB(217, 0, 5)
        Case 4 'Naranja
            TileEngine.Engine_Long_To_RGB_List charlist(CharIndex).label_color, D3DColorXRGB(243, 147, 1)
        
        Case 9 'Lider Impe
            TileEngine.Engine_Long_To_RGB_List charlist(CharIndex).label_color, D3DColorXRGB(39, 131, 243)
        Case 10 'Lider Repu
            TileEngine.Engine_Long_To_RGB_List charlist(CharIndex).label_color, D3DColorXRGB(243, 147, 1)
        Case 11 'Lider Caos
            TileEngine.Engine_Long_To_RGB_List charlist(CharIndex).label_color, D3DColorXRGB(217, 0, 5)
            
        Case 5 'Verde - conse
            TileEngine.Engine_Long_To_RGB_List charlist(CharIndex).label_color, D3DColorXRGB(0, 142, 72)
        Case 6 'Verde - semi
            TileEngine.Engine_Long_To_RGB_List charlist(CharIndex).label_color, D3DColorXRGB(0, 142, 72)
        Case 7 'Verde - dios
            TileEngine.Engine_Long_To_RGB_List charlist(CharIndex).label_color, D3DColorXRGB(0, 142, 72)
        Case 8 'Verde - admin
            TileEngine.Engine_Long_To_RGB_List charlist(CharIndex).label_color, D3DColorXRGB(10, 10, 10)
    End Select
    
End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 18 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim tempInt As Integer
    Dim headIndex As Integer
    Dim Escudo As Integer, Arma As Integer, body As Integer
    
    CharIndex = incomingData.ReadInteger()
    
    With charlist(CharIndex)
        tempInt = incomingData.ReadInteger()
        body = tempInt
        If tempInt < LBound(BodyData()) Or tempInt > UBound(BodyData()) Then
            .body = BodyData(0)
            .iBody = 0
        Else
            .body = BodyData(tempInt)
            .iBody = tempInt
        End If
        
        headIndex = incomingData.ReadInteger()
        
        If tempInt < LBound(HeadData()) Or tempInt > UBound(HeadData()) Then
            .Head = HeadData(0)
            .iHead = 0
        Else
            .Head = HeadData(headIndex)
            .iHead = headIndex
        End If
        
        Dim oldMuerto As Boolean
        oldMuerto = .Muerto
        
        .Muerto = (headIndex = CASPER_HEAD)
        If .Muerto = False And oldMuerto = True Then
            Call TileEngine.Char_Particle_Group_Remove(CharIndex, 22)
        End If
        
        .heading = incomingData.ReadByte()
        
        Arma = incomingData.ReadInteger
        Escudo = incomingData.ReadInteger
        If Arma <> 0 Then .Arma = WeaponAnimData(Arma)
        If Escudo <> 0 Then .Escudo = ShieldAnimData(Escudo)
        
        TileEngine.Char_Set_Aura CharIndex, Escudo, Arma, body
        
        tempInt = incomingData.ReadInteger()
        If tempInt <> 0 Then .Casco = CascoAnimData(tempInt)
        
        Call TileEngine.Char_SetFx(CharIndex, incomingData.ReadInteger(), incomingData.ReadInteger())
    End With
End Sub

''
' Handles the ObjectCreate message.

Private Sub HandleObjectCreate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    Dim Obj As Integer
    Dim tipe As Byte
    Dim Amount As Integer
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    Obj = incomingData.ReadInteger()
    tipe = incomingData.ReadByte()
    Amount = incomingData.ReadInteger
    
    If Obj = 378 Then
        Call TileEngine.General_Particle_Create(34, X, Y)
    Else
        TileEngine.Map_Obj_Create X, Y, objs(Obj).Grh, Obj, tipe, Amount
    End If
End Sub

''
' Handles the ObjectDelete message.

Private Sub HandleObjectDelete()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    If MapData(X, Y).particle_group_index > 0 Then
        If TileEngine.Particle_Get_Type(MapData(X, Y).particle_group_index) = 16 Then
            Call TileEngine.Particle_Group_Remove(MapData(X, Y).particle_group_index)
        End If
    End If
    
    TileEngine.Map_Obj_Delete X, Y

End Sub

''
' Handles the BlockPosition message.

Private Sub HandleBlockPosition()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    If incomingData.ReadBoolean() Then
        MapData(X, Y).Blocked = 1
    Else
        MapData(X, Y).Blocked = 0
    End If
End Sub

''
' Handles the PlayMusic message.

Private Sub HandlePlayMusic()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim currentMidi As Byte
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    currentMidi = incomingData.ReadByte()
    
    If currentMidi Then
        Call Audio.PlayMusic(CStr(currentMidi) & "", incomingData.ReadInteger())
    Else
        'Remove the bytes to prevent errors
        Call incomingData.ReadInteger
    End If
End Sub

''
' Handles the PlayWave message.

Private Sub HandlePlayWave()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Last Modified by: Rapsodius
'Added support for 3D Sounds.
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
        
    Dim wave As Integer
    Dim SrcX As Byte
    Dim SrcY As Byte
    
    wave = incomingData.ReadInteger()
    SrcX = incomingData.ReadByte()
    SrcY = incomingData.ReadByte()
    
    If wave = 105 Then 'Trueno
        trueno = 20
        Exit Sub
    End If
    
    Call Audio.PlayWave(wave, SrcX, SrcY)
        
End Sub

''
' Handles the GuildList message.

Private Sub HandleGuildList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    'Clear guild's list
    frmGuildAdm.guildslist.Clear
    
    Dim guilds() As String
    guilds = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    Dim i As Long
    For i = 0 To UBound(guilds())
        Call frmGuildAdm.guildslist.AddItem(guilds(i))
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    frmGuildAdm.Show vbModeless, frmMain
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the AreaChanged message.

Private Sub HandleAreaChanged()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
        
    Call TileEngine.Map_Change_Area(X, Y)
End Sub

''
' Handles the PauseToggle message.

Private Sub HandlePauseToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Pausa = Not Pausa
End Sub
Private Sub HandleCreateFXMap()
    With incomingData
        Call .ReadByte
    
        Dim X As Byte, Y As Byte, FX As Integer, Loops As Integer
        X = .ReadByte
        Y = .ReadByte
        FX = .ReadInteger
        Loops = .ReadInteger
        
        TileEngine.Map_FX_Create X, Y, FX, Loops
    End With
End Sub
''
' Handles the CreateFX message.

Private Sub HandleCreateFX()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim FX As Integer
    Dim Loops As Integer
    
    CharIndex = incomingData.ReadInteger()
    FX = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    Call TileEngine.Char_SetFx(CharIndex, FX, Loops)
End Sub

''
' Handles the UpdateUserStats message.

Private Sub HandleUpdateUserStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 26 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMaxHP = incomingData.ReadInteger()
    UserMinHP = incomingData.ReadInteger()
    UserMaxMAN = incomingData.ReadInteger()
    UserMinMAN = incomingData.ReadInteger()
    UserMaxSTA = incomingData.ReadInteger()
    UserMinSTA = incomingData.ReadInteger()
    UserGLD = incomingData.ReadLong()
    UserLVL = incomingData.ReadByte()
    UserPasarNivel = incomingData.ReadLong()
    UserExp = incomingData.ReadLong()
    
    
    If Not UserExp = 0 And Not UserPasarNivel = 0 Then
        frmMain.lblExp.Caption = CStr(Round(UserExp * 100 / UserPasarNivel) & "%") & " (" & UserExp & "/" & UserPasarNivel & ")"
    Else
        frmMain.lblExp.Caption = "¡Nivel maximo!"
    End If

    
    If Not UserExp = 0 And Not UserPasarNivel = 0 Then
        frmMain.ExpShp.width = Round(((UserExp / 100) / (UserPasarNivel / 100)) * 120)
    Else
        frmMain.ExpShp.Visible = False
        'frmMain.ExpShp.width = 120
    End If
    
    If Not UserMinHP < 0 Then frmMain.Hpshp.width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 91)
    frmMain.lblHP.Caption = UserMinHP & "/" & UserMaxHP
    
    If UserMaxMAN > 0 Then
        frmMain.MANShp.width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 91)
        frmMain.lblMP.Caption = UserMinMAN & "/" & UserMaxMAN
    Else
        frmMain.lblMP.Caption = ""
        frmMain.MANShp.width = 0
    End If
    
    frmMain.STAShp.width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 91)
    frmMain.lblST.Caption = UserMinSTA & "/" & UserMaxSTA
    
    frmMain.GldLbl.Caption = UserGLD
    frmMain.LvlLbl.Caption = UserLVL
    
    If UserMinHP = 0 Then
        UserEstado = 1
    Else
        UserEstado = 0
    End If
    
End Sub

''
' Handles the WorkRequestTarget message.

Private Sub HandleWorkRequestTarget()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UsingSkill = incomingData.ReadByte()
    
    Select Case UsingSkill
        Case Magia
            Call AddtoRichTextBox(frmMain.RecChat, MENSAJE_TRABAJO, 100, 100, 120, 0, 0)
            cCursores.Parse_Form frmMain, E_CAST
        Case Pesca
            Call AddtoRichTextBox(frmMain.RecChat, MENSAJE_TRABAJO, 100, 100, 120, 0, 0)
            cCursores.Parse_Form frmMain, E_SHOOT
        Case Robar
            Call AddtoRichTextBox(frmMain.RecChat, MENSAJE_TRABAJO, 100, 100, 120, 0, 0)
            cCursores.Parse_Form frmMain, E_SHOOT
        Case Talar
            Call AddtoRichTextBox(frmMain.RecChat, MENSAJE_TRABAJO, 100, 100, 120, 0, 0)
            cCursores.Parse_Form frmMain, E_SHOOT
        Case Mineria
            Call AddtoRichTextBox(frmMain.RecChat, MENSAJE_TRABAJO, 100, 100, 120, 0, 0)
            cCursores.Parse_Form frmMain, E_SHOOT
        Case FundirMetal
            Call AddtoRichTextBox(frmMain.RecChat, MENSAJE_TRABAJO, 100, 100, 120, 0, 0)
            cCursores.Parse_Form frmMain, E_SHOOT
        Case Proyectiles
            Call AddtoRichTextBox(frmMain.RecChat, MENSAJE_TRABAJO, 100, 100, 120, 0, 0)
            cCursores.Parse_Form frmMain, E_ARROW
        Case Arrojadizas
            Call AddtoRichTextBox(frmMain.RecChat, MENSAJE_TRABAJO, 100, 100, 120, 0, 0)
            cCursores.Parse_Form frmMain, E_ATTACK
        Case Else
            cCursores.Parse_Form frmMain
    End Select
    
    Select Case UsingSkill
        Case Pesca, Talar, Mineria, FundirMetal
            UsingSkill = 0
    End Select
End Sub

''
' Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim slot As Byte
    Dim OBJIndex As Integer
    Dim name As String
    Dim Amount As Integer
    Dim Equipped As Boolean
    Dim grhindex As Integer
    Dim OBJType As Byte
    Dim MaxHit As Integer
    Dim MinHit As Integer
    Dim defense As Integer
    Dim value As Single
    Dim Puede As Byte
    
    slot = buffer.ReadByte()
    OBJIndex = buffer.ReadInteger()
    name = General_Locale_Name_Obj(OBJIndex)
    Amount = buffer.ReadInteger()
    Equipped = buffer.ReadBoolean()
    grhindex = buffer.ReadInteger()
    OBJType = buffer.ReadByte()
    MaxHit = buffer.ReadInteger()
    MinHit = buffer.ReadInteger()
    defense = buffer.ReadInteger()
    value = buffer.ReadSingle()
    Puede = buffer.ReadByte
    
    Call Inventario.SetItem(slot, OBJIndex, Amount, Equipped, grhindex, OBJType, MaxHit, MinHit, defense, value, name, Puede)
    
    RenderInv = True
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ChangeBankSlot message.

Private Sub HandleChangeBankSlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim slot As Byte
    slot = buffer.ReadByte()
    
    With UserBancoInventory(slot)
        .OBJIndex = buffer.ReadInteger()
        .name = General_Locale_Name_Obj(.OBJIndex)
        .Amount = buffer.ReadInteger()
        .grhindex = buffer.ReadInteger()
        .OBJType = buffer.ReadByte()
        .MaxHit = buffer.ReadInteger()
        .MinHit = buffer.ReadInteger()
        .Def = buffer.ReadInteger()
        .Valor = buffer.ReadLong()
    End With
    
    If frmBancoObj.List1(0).ListCount >= slot Then _
        Call frmBancoObj.List1(0).RemoveItem(slot - 1)
    
    Call frmBancoObj.List1(0).AddItem(IIf(UserBancoInventory(slot).name <> "", UserBancoInventory(slot).name, "(Nada)"), slot - 1)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ChangeSpellSlot message.

Private Sub HandleChangeSpellSlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim slot As Byte
    slot = buffer.ReadByte()
    
    UserHechizos(slot) = buffer.ReadInteger()
    
    If slot <= frmMain.hlst.ListCount Then
        frmMain.hlst.List(slot - 1) = General_Locale_Name_Spell(UserHechizos(slot))
    Else
        Call frmMain.hlst.AddItem(General_Locale_Name_Spell(UserHechizos(slot)))
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the Attributes message.

Private Sub HandleAtributes()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 1 + NUMATRIBUTES Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim i As Long
    
    For i = 1 To NUMATRIBUTES
        UserAtributos(i) = incomingData.ReadByte()
    Next i
    
    'Show them in character creation
    If EstadoLogin = E_MODO.Dados Then
        With frmCrearPersonaje
            If .Visible Then
                .lbAtt(0).Caption = UserAtributos(1)
                .lbAtt(1).Caption = UserAtributos(2)
                .lbAtt(2).Caption = UserAtributos(3)
                .lbAtt(3).Caption = UserAtributos(4)
                .lbAtt(4).Caption = UserAtributos(5)
            End If
        End With
    Else
        LlegaronEstadisticas = True
    End If
End Sub

''
' Handles the BlacksmithWeapons message.

Private Sub HandleBlacksmithWeapons()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim count As Integer
    Dim i As Long
    Dim tmp As String
    
    count = buffer.ReadInteger()
    
    Call frmHerrero.lstArmas.Clear
    
    For i = 1 To count
        tmp = buffer.ReadASCIIString() & " ("           'Get the object's name
        tmp = tmp & CStr(buffer.ReadInteger()) & ","    'The iron needed
        tmp = tmp & CStr(buffer.ReadInteger()) & ","    'The silver needed
        tmp = tmp & CStr(buffer.ReadInteger()) & ")"    'The gold needed
        
        Call frmHerrero.lstArmas.AddItem(tmp)
        ArmasHerrero(i) = buffer.ReadInteger()
    Next i
    
    For i = i To UBound(ArmasHerrero())
        ArmasHerrero(i) = 0
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the BlacksmithArmors message.

Private Sub HandleBlacksmithArmors()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim count As Integer
    Dim i As Long
    Dim tmp As String
    
    count = buffer.ReadInteger()
    
    Call frmHerrero.lstArmaduras.Clear
    
    For i = 1 To count
        tmp = buffer.ReadASCIIString() & " ("           'Get the object's name
        tmp = tmp & CStr(buffer.ReadInteger()) & ","    'The iron needed
        tmp = tmp & CStr(buffer.ReadInteger()) & ","    'The silver needed
        tmp = tmp & CStr(buffer.ReadInteger()) & ")"    'The gold needed
        
        Call frmHerrero.lstArmaduras.AddItem(tmp)
        ArmadurasHerrero(i) = buffer.ReadInteger()
    Next i
    
    For i = i To UBound(ArmadurasHerrero())
        ArmadurasHerrero(i) = 0
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the CarpenterObjects message.

Private Sub HandleCarpenterObjects()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim count As Integer
    Dim i As Long
    Dim tmp As String
    
    count = buffer.ReadInteger()
    
    Call frmCarp.lstArmas.Clear
    
    For i = 1 To count
        tmp = buffer.ReadASCIIString() & " ("          'Get the object's name
        tmp = tmp & CStr(buffer.ReadInteger()) & ")"    'The wood needed
        
        Call frmCarp.lstArmas.AddItem(tmp)
        ObjCarpintero(i) = buffer.ReadInteger()
    Next i
    
    For i = i To UBound(ObjCarpintero())
        ObjCarpintero(i) = 0
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub



Private Sub HandleAlquimiaObjects()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim count As Integer
    Dim i As Long
    Dim tmp As String
    
    count = buffer.ReadInteger()
    
    Call frmAlquimia.lstPociones.Clear
    
    For i = 1 To count
        tmp = buffer.ReadASCIIString() & " ("          'Get the object's name
        tmp = tmp & CStr(buffer.ReadInteger()) & ")"    'The wood needed
        
        Call frmAlquimia.lstPociones.AddItem(tmp)
        ObjAlquimia(i) = buffer.ReadInteger()
    Next i
    
    For i = i To UBound(ObjAlquimia())
        ObjAlquimia(i) = 0
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub





' Handles the CarpenterObjects message.

Private Sub HandleSastreObjects()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
      
    
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim count As Integer
    Dim i As Long
    Dim tmp As String

    
    count = buffer.ReadInteger()

    Call frmSastre.lstRopas.Clear
    

    For i = 1 To count
        tmp = buffer.ReadASCIIString() & " ("
        tmp = tmp & CStr(buffer.ReadInteger()) & "/" & _
        CStr(buffer.ReadInteger()) & "/" & _
        CStr(buffer.ReadInteger()) & ")"
        
        Call frmSastre.lstRopas.AddItem(tmp)
        ObjSastre(i) = buffer.ReadInteger()
    Next i
  
    For i = i To UBound(ObjSastre())
        ObjSastre(i) = 0
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    
      
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub


''
' Handles the RestOK message.

Private Sub HandleRestOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserDescansar = Not UserDescansar
End Sub

''
' Handles the ErrorMessage message.

Private Sub HandleErrorMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    frmMensaje.msg.Caption = buffer.ReadASCIIString()
    frmMensaje.Show , frmMain
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the Blind message.

Private Sub HandleBlind()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCiego = True
End Sub

''
' Handles the Dumb message.

Private Sub HandleDumb()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEstupido = True
End Sub

''
' Handles the ShowSignal message.

Private Sub HandleShowSignal()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim tmp As String
    tmp = buffer.ReadASCIIString()
    
    Rem Mannakia 'Esto hay que sacarlo por dios'
    'Call InitCartel(tmp, Buffer.ReadInteger())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ChangeNPCInventorySlot message.

Private Sub HandleChangeNPCInventorySlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
  
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim slot As Byte
    slot = buffer.ReadByte()
    
    With NPCInventory(slot)
        .Amount = buffer.ReadInteger()
        .Valor = buffer.ReadSingle()
        .grhindex = buffer.ReadInteger()
        .OBJIndex = buffer.ReadInteger()
        .OBJType = buffer.ReadByte()
        .MaxHit = buffer.ReadInteger()
        .MinHit = buffer.ReadInteger()
        .Def = buffer.ReadInteger()
        .name = General_Locale_Name_Obj(.OBJIndex)
    End With
    
    If frmComerciar.List1(0).ListCount >= slot Then _
        Call frmComerciar.List1(0).RemoveItem(slot - 1)
    
    Call frmComerciar.List1(0).AddItem(IIf(NPCInventory(slot).name <> "", NPCInventory(slot).name, "(Nada)"), slot - 1)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the MiniStats message.

Private Sub HandleMiniStats()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 20 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    With UserEstadisticas
        .CiudadanosMatados = incomingData.ReadLong()
        .RenegadosMatados = incomingData.ReadLong()
        .RepublicanosMatados = incomingData.ReadLong()
        
        .ArmadaMatados = incomingData.ReadLong()
        .CaosMatados = incomingData.ReadLong()
        .MiliciaMatados = incomingData.ReadLong()
        
        .UsuariosMatados = incomingData.ReadLong()
        
        .NpcMatados = incomingData.ReadInteger()
        .Clase = incomingData.ReadByte
        .Raza = incomingData.ReadByte
        .Genero = incomingData.ReadByte
        
        UserStat = incomingData.ReadByte
        
        SkillPoints = incomingData.ReadInteger
    End With
    
    With UserPet
        .TieneFamiliar = incomingData.ReadByte
        If .TieneFamiliar = 1 Then
            .nombre = incomingData.ReadASCIIString
            
            .ELV = incomingData.ReadByte
            .ELU = incomingData.ReadLong
            .EXP = incomingData.ReadLong
            
            .MinHP = incomingData.ReadInteger
            .MaxHP = incomingData.ReadInteger
            
            .MinHit = incomingData.ReadInteger
            .MaxHit = incomingData.ReadInteger
            
            .tipo = incomingData.ReadByte
        
            If .ELV >= 10 Then
                If .tipo = eMascota.Ely Then
                    .Habilidad = "Cura al amo"
                ElseIf .tipo = eMascota.Fatuo Then
                    .Habilidad = "Lanza Misiles"
                End If
            End If
            
            If .ELV >= 15 Then
                If .tipo = eMascota.Ely Or .tipo = eMascota.Fatuo Then
                    .Habilidad = .Habilidad & " - Inmoviliza"
                ElseIf .tipo = eMascota.Lobo Or .tipo = eMascota.Tigre Then
                    .Habilidad = "Golpe Entorpece"
                ElseIf .tipo = eMascota.Ent Then
                    .Habilidad = "Golpe Envenena"
                End If
            End If
            
            If .ELV >= 20 Then
                If .tipo = eMascota.Tigre Or .tipo = eMascota.Ent Or .tipo = eMascota.Lobo Then
                    .Habilidad = .Habilidad & " - Golpe Paraliza"
                ElseIf .tipo = eMascota.Ely Or .tipo = eMascota.Fatuo Then
                    .Habilidad = .Habilidad & " - Lanza Descargas"
                ElseIf .tipo = eMascota.fuego Then
                    .Habilidad = "Lanza Tormentas"
                ElseIf .tipo = eMascota.Agua Then
                    .Habilidad = "Paraliza"
                ElseIf .tipo = eMascota.Tierra Then
                    .Habilidad = "Inmoviliza"
                End If
            End If
            
            If .ELV >= 30 Then
                If .tipo = eMascota.Ely Then
                    .Habilidad = .Habilidad & " - Desencanta al amo"
                ElseIf .tipo = eMascota.Fatuo Then
                    .Habilidad = .Habilidad & " - Detecta Invisibilidad"
                ElseIf .tipo = eMascota.Oso Then
                    .Habilidad = "Golpe Desarma"
                ElseIf .tipo = eMascota.Tigre Or .tipo = eMascota.Lobo Then
                    .Habilidad = .Habilidad & " - Golpe Enceguece"
                ElseIf .tipo = eMascota.Ent Then
                    .Habilidad = .Habilidad & " - Golpe Desarma"
                End If
            End If
        End If
    End With
End Sub



''
' Handles the AddForumMessage message.

Private Sub HandleAddForumMessage()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim Title As String
    Dim Message As String
    
    Title = buffer.ReadASCIIString()
    Message = buffer.ReadASCIIString()
    
    Call frmForo.List.AddItem(Title)
    frmForo.Text(frmForo.List.ListCount - 1).Text = Message
    Call Load(frmForo.Text(frmForo.List.ListCount))
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub



Private Sub HandleAddCorreoMessage()
    If incomingData.length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim De As String
    Dim Message As String
    Dim item As Integer
    Dim Index As Byte
    Dim cant As Integer
    
    Index = buffer.ReadByte
    Message = buffer.ReadASCIIString()
    De = buffer.ReadASCIIString()
    item = buffer.ReadInteger
    cant = buffer.ReadInteger
    
    Correos(Index).mensaje = Message
    Correos(Index).De = De
    Correos(Index).item = item
    Correos(Index).Cantidad = cant
    
    If frmCorreo.Visible Then
        frmCorreo.lstMsg.List(Index - 1) = IIf(De = "", "(Nada)", De)
    End If
    
    frmCorreo.ActualizarCorreo
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub






''
' Handles the ShowForumForm message.

Private Sub HandleShowForumForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If Not frmForo.Visible Then
        frmForo.Show , frmMain
    End If
End Sub

Private Sub HandleShowCorreoForm()
'***************************************************
'Author: Jose Castelli
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If Not frmCorreo.Visible Then
        frmCorreo.Show , frmMain
    End If
    
    Dim i As Long
    For i = 1 To 20
        If Correos(i).De <> "" Then
            frmCorreo.lstMsg.AddItem Correos(i).De
        Else
            frmCorreo.lstMsg.AddItem "(Nada)"
        End If
    Next i
End Sub


''
' Handles the SetInvisible message.

Private Sub HandleSetInvisible()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    charlist(CharIndex).invisible = incomingData.ReadBoolean()

End Sub

''
' Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMeditar = Not UserMeditar

End Sub

''
' Handles the BlindNoMore message.

Private Sub HandleBlindNoMore()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCiego = False
End Sub

''
' Handles the DumbNoMore message.

Private Sub HandleDumbNoMore()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEstupido = False
End Sub

''
' Handles the SendSkills message.

Private Sub HandleSendSkills()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 1 + NUMSKILLS Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim i As Long
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = incomingData.ReadByte()
        SkillsOrig(i) = UserSkills(i)
    Next i
   
End Sub

''
' Handles the TrainerCreatureList message.

Private Sub HandleTrainerCreatureList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim creatures() As String
    Dim i As Long
    
    creatures = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(creatures())
        Call frmEntrenador.lstCriaturas.AddItem(creatures(i))
    Next i
    frmEntrenador.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub


' Handles the CharacterInfo message.

Private Sub HandleCharacterInfo()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 34 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    With frmCharInfo
        If .frmType = CharInfoFrmType.frmMembers Then
            .Rechazar.Visible = False
            .Aceptar.Visible = False
            .Echar.Visible = True
            .Desc.Visible = False
        Else
            .Rechazar.Visible = True
            .Aceptar.Visible = True
            .Echar.Visible = False
            .Desc.Visible = True
        End If
        
        .nombre.Caption = "Nombre: " & buffer.ReadASCIIString()
        .Raza.Caption = "Raza: " & ListaRazas(buffer.ReadByte())
        .Clase.Caption = "Clase: " & ListaClases(buffer.ReadByte())
        
        If buffer.ReadByte() = 1 Then
            .Genero.Caption = "Genero: Hombre"
        Else
            .Genero.Caption = "Genero: Mujer"
        End If
        
        .Nivel.Caption = "Nivel: " & buffer.ReadByte()
        .Oro.Caption = "Oro: " & buffer.ReadLong()
        .Banco.Caption = "Banco: " & buffer.ReadLong()
        
        .txtPeticiones.Text = buffer.ReadASCIIString()
        '.guildactual.Caption = "Clan: " & Buffer.ReadASCIIString()
        buffer.ReadASCIIString
        .txtMiembro.Text = buffer.ReadASCIIString()
        
        Dim faccion As Byte
        faccion = buffer.ReadByte
        If faccion = 1 Then
            .ejercito.Caption = "Facción: Ejercito Real"
        ElseIf faccion = 2 Then
            .ejercito.Caption = "Facción: Milicia Republicana"
        ElseIf faccion = 3 Then
            .ejercito.Caption = "Facción: Fuerzas del caos"
        Else
            .ejercito.Caption = "Facción: Ninguna"
        End If
        
        .Armadas.Caption = "Armadas matados: " & buffer.ReadInteger
        .Milicianos.Caption = "Milicianos matados: " & buffer.ReadInteger
        .Caoticos.Caption = "Caoticos matados: " & buffer.ReadInteger
        .Imperiales.Caption = "Imperiales matados: " & buffer.ReadInteger
        .Republicanos.Caption = "Republicanos matados: " & buffer.ReadInteger
        .Renegados.Caption = "Renegados matados: " & buffer.ReadInteger
        
        Call .Show(vbModeless, frmMain)
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the GuildLeaderInfo message.

Private Sub HandleGuildLeaderInfo()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 9 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim List() As String
    Dim i As Long
    
    With frmGuildLeader
        'Get list of existing guilds
        List = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .guildslist.Clear
        
        For i = 0 To UBound(List())
            Call .guildslist.AddItem(List(i))
        Next i
        
        'Get list of guild's members
        List = Split(buffer.ReadASCIIString(), SEPARATOR)
        .Miembros.Caption = "El clan cuenta con " & CStr(UBound(List()) + 1) & " miembros."
        
        'Empty the list
        Call .members.Clear
        
        For i = 0 To UBound(List())
            Call .members.AddItem(List(i))
        Next i
        
        .txtguildnews = buffer.ReadASCIIString()
        
        'Get list of join requests
        List = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .solicitudes.Clear
        
        For i = 0 To UBound(List())
            Call .solicitudes.AddItem(List(i))
        Next i
        
        .Show , frmMain
    End With

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the GuildDetails message.

Private Sub HandleGuildDetails()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 26 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    With frmGuildBrief
        .nombre.Caption = "Nombre:" & buffer.ReadASCIIString()
        .fundador.Caption = "Fundador:" & buffer.ReadASCIIString()
        .creacion.Caption = "Fecha de creacion:" & buffer.ReadASCIIString()
        .lider.Caption = "Líder:" & buffer.ReadASCIIString()
        .web.Caption = "Web site:" & buffer.ReadASCIIString()
        .Miembros.Caption = "Miembros:" & buffer.ReadInteger()
        
        .lblAlineacion.Caption = "Alineación: " & buffer.ReadASCIIString()
        .antifaccion.Caption = "Puntos Antifaccion: " & buffer.ReadASCIIString()
        
        Dim codexStr() As String
        Dim i As Long
        
        codexStr = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        For i = 0 To 7
            .Codex(i).Caption = codexStr(i)
        Next i
        
        .Desc.Text = buffer.ReadASCIIString()
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    frmGuildBrief.Show vbModeless, frmMain
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ShowGuildFundationForm message.

Private Sub HandleShowGuildFundationForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    CreandoClan = True
    frmGuildFoundation.Show , frmMain
End Sub

''
' Handles the ParalizeOK message.

Private Sub HandleParalizeOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserParalizado = Not UserParalizado
End Sub

''
' Handles the ShowUserRequest message.

Private Sub HandleShowUserRequest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(buffer.ReadASCIIString())
    Call frmUserRequest.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the TradeOK message.

Private Sub HandleTradeOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If frmComerciar.Visible Then
        Dim i As Long
        
        Call frmComerciar.List1(1).Clear
        
        For i = 1 To MAX_INVENTORY_SLOTS
            If Inventario.OBJIndex(i) <> 0 Then
                Call frmComerciar.List1(1).AddItem(Inventario.ItemName(i))
            Else
                Call frmComerciar.List1(1).AddItem("(Nada)")
            End If
        Next i
        
        'Alter order according to if we bought or sold so the labels and grh remain the same
        If frmComerciar.LasActionBuy Then
            frmComerciar.List1(1).ListIndex = frmComerciar.LastIndex2
            frmComerciar.List1(0).ListIndex = frmComerciar.LastIndex1
        Else
            frmComerciar.List1(0).ListIndex = frmComerciar.LastIndex1
            frmComerciar.List1(1).ListIndex = frmComerciar.LastIndex2
        End If
    End If
End Sub

''
' Handles the BankOK message.

Private Sub HandleBankOK()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim i As Long
    
    If frmBancoObj.Visible Then
        
        Call frmBancoObj.List1(1).Clear
        
        For i = 1 To MAX_INVENTORY_SLOTS
            If Inventario.OBJIndex(i) <> 0 Then
                Call frmBancoObj.List1(1).AddItem(Inventario.ItemName(i))
            Else
                Call frmBancoObj.List1(1).AddItem("(Nada)")
            End If
        Next i
        
        'Alter order according to if we bought or sold so the labels and grh remain the same
        If frmBancoObj.LasActionBuy Then
            frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
            frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
        Else
            frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
            frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
        End If
    End If
End Sub

''
' Handles the ChangeUserTradeSlot message.

Private Sub HandleChangeUserTradeSlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    With OtroInventario(1)
        .OBJIndex = buffer.ReadInteger()
        .name = General_Locale_Name_Obj(.OBJIndex)
        .Amount = buffer.ReadLong()
        .grhindex = buffer.ReadInteger()
        .OBJType = buffer.ReadByte()
        .MaxHit = buffer.ReadInteger()
        .MinHit = buffer.ReadInteger()
        .Def = buffer.ReadInteger()
        .Valor = buffer.ReadLong()
        
        frmComerciarUsu.List2.Clear
        
        Call frmComerciarUsu.List2.AddItem(.name)
        frmComerciarUsu.List2.ItemData(frmComerciarUsu.List2.NewIndex) = .Amount
        
        frmComerciarUsu.lblEstadoResp.Visible = False
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub


''
' Handles the SpawnList message.

Private Sub HandleSpawnList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim creatureList() As String
    Dim i As Long
    
    creatureList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(creatureList())
        Call frmSpawnList.lstCriaturas.AddItem(creatureList(i))
    Next i
    frmSpawnList.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub
''
' Handles the SpawnList message.


Private Sub HandleShowSOSForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim sosList() As String
    Dim i As Long
    
    sosList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(sosList())
        Call frmMSG.List1.AddItem(sosList(i))
    Next i
    
    frmMSG.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the ShowGMPanelForm message.

Private Sub HandleShowGMPanelForm()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmPanelGm.Show vbModeless, frmMain
    
End Sub

''
' Handles the UserNameList message.

Private Sub HandleUserNameList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim userList() As String
    Dim i As Long
    
    userList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    If frmPanelGm.Visible Then
        frmPanelGm.cboListaUsus.Clear
        For i = 0 To UBound(userList())
            Call frmPanelGm.cboListaUsus.AddItem(userList(i))
        Next i
        If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the Pong message.

Private Sub HandlePong()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecChat, "El ping es " & CLng(GetElapsedTime * 10) & " ms.", 255, 0, 0, True, False, False)
    
    pingTime = 0
End Sub

''
' Handles the UpdateTag message.

Private Sub HandleUpdateTagAndStatus()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim CharIndex As Integer
    Dim userTipe As Byte
    Dim userTag As String
    Dim Pos As Integer
    
    CharIndex = buffer.ReadInteger()
    userTipe = buffer.ReadByte
    userTag = buffer.ReadASCIIString()
    
    'Update char status adn tag!
    With charlist(CharIndex)
        .Priv = userTipe
        Select Case .Priv
            Case 1 'Gris
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(128, 128, 128)
            Case 2 'Azul
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(0, 80, 200)
            Case 3 'Rojo
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(197, 0, 5)
            Case 4 'Naranja
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(243, 147, 1)
            
            Case 9 'Lider Impe
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(0, 80, 200)
            Case 10 'Lider Repu
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(243, 147, 1)
            Case 11 'Lider Caos
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(197, 0, 5)
            
            
            Case 5 'Verde - conse
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(0, 128, 0)
            Case 6 'Verde - semi
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(0, 128, 0)
            Case 7 'Verde - dios
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(0, 128, 0)
            Case 8 'Verde - admin
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(10, 10, 10)

        End Select
        
        Pos = InStr(userTag, "<")
        If Pos = 0 Then
            .clan = ""
            .offClanX = 0
            
            .nombre = userTag
        Else
            .clan = mid$(userTag, Pos)
            .offClanX = TileEngine.Text_Width(.clan, 1) / 2
            
            .nombre = Left$(userTag, Pos - 2)
        End If
        
        .offNameX = TileEngine.Text_Width(.nombre, 1) / 2
        
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If error <> 0 Then _
        Err.Raise error
End Sub



Public Sub WriteTalk(ByVal chat As String, ByVal mode As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.Talk)
        Call .WriteASCIIString(chat)
        Call .WriteByte(mode)
    End With
End Sub

''
' Writes the "Whisper" message to the outgoing data buffer.
'
' @param    charIndex The index of the char to whom to whisper.
' @param    chat The chat text to be sent to the user.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhisper(ByVal name As String, ByVal chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Whisper" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Whisper)
        
        Call .WriteASCIIString(name)
        
        Call .WriteASCIIString(chat)
    End With
End Sub

''
' Writes the "Walk" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWalk(ByVal heading As E_Heading)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Walk" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Walk)
        
        Call .WriteByte(heading)
    End With
End Sub

''
' Writes the "RequestPositionUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestPositionUpdate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestPositionUpdate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestPositionUpdate)
End Sub

''
' Writes the "Attack" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttack()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Attack" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Attack)
End Sub

''
' Writes the "PickUp" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePickUp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PickUp" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PickUp)
End Sub

''
' Writes the "CombatModeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCombatModeToggle()

'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CombatModeToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CombatModeToggle)
End Sub

''
' Writes the "SafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SafeToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.SafeToggle)
End Sub

''
' Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestGuildLeaderInfo()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestGuildLeaderInfo)
End Sub


''
' Writes the "RequestSkills" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestEstadisticas()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestSkills" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestEstadisticas)
End Sub


''
' Writes the "CommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceEnd" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CommerceEnd)
End Sub

''
' Writes the "UserCommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceEnd" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceEnd)
End Sub

''
' Writes the "BankEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankEnd" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.BankEnd)
End Sub

''
' Writes the "UserCommerceOk" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOk()
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/10/07
'Writes the "UserCommerceOk" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceOk)
End Sub

''
' Writes the "UserCommerceReject" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceReject()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceReject" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceReject)
End Sub

''
' Writes the "Drop" message to the outgoing data buffer.
'
' @param    slot Inventory slot where the item to drop is.
' @param    amount Number of items to drop.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDrop(ByVal slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Drop" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Drop)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

''
' Writes the "CastSpell" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell to cast is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCastSpell(ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CastSpell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CastSpell)
        
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "LeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeftClick(ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LeftClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.LeftClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "DoubleClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoubleClick(ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DoubleClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.DoubleClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "Work" message to the outgoing data buffer.
'
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWork(ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Work" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Work)
        
        Call .WriteByte(Skill)
    End With
End Sub



''
' Writes the "UseItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to use is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseItem(ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UseItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UseItem)
        
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "CraftBlacksmith" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftBlacksmith(ByVal item As Integer, cant As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CraftBlacksmith" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftBlacksmith)
        
        Call .WriteInteger(item)
        Call .WriteInteger(cant)
    End With
End Sub

''
' Writes the "CraftCarpenter" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftCarpenter(ByVal item As Integer, ByVal cant As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CraftCarpenter" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftCarpenter)
        
        Call .WriteInteger(item)
        Call .WriteInteger(cant)
    End With
End Sub


Public Sub WriteCraftSastre(ByVal item As Integer, ByVal cant As Integer)

    With outgoingData
        Call .WriteByte(ClientPacketID.CraftSastre)
        
        Call .WriteInteger(item)
        Call .WriteInteger(cant)
    End With
End Sub
Public Sub WriteCraftalquimia(ByVal item As Integer, ByVal cant As Integer)


    With outgoingData
        Call .WriteByte(ClientPacketID.Craftalquimia)
        
        Call .WriteInteger(item)
        Call .WriteInteger(cant)
    End With
End Sub

''
' Writes the "WorkLeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkLeftClick(ByVal X As Byte, ByVal Y As Byte, ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WorkLeftClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.WorkLeftClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        Call .WriteByte(Skill)
    End With
End Sub

''
' Writes the "CreateNewGuild" message to the outgoing data buffer.
'
' @param    desc    The guild's description
' @param    name    The guild's name
' @param    site    The guild's website
' @param    codex   Array of all rules of the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNewGuild(ByVal Desc As String, ByVal name As String, ByVal Site As String, ByRef Codex() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateNewGuild" message to the outgoing data buffer
'***************************************************
    Dim temp As String
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNewGuild)
        
        Call .WriteASCIIString(Desc)
        Call .WriteASCIIString(name)
        Call .WriteASCIIString(Site)
        
        For i = LBound(Codex()) To UBound(Codex())
            temp = temp & Codex(i) & SEPARATOR
        Next i
        
        If Len(temp) Then _
            temp = Left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)
    End With
End Sub

''
' Writes the "SpellInfo" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpellInfo(ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpellInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SpellInfo)
        
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "EquipItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to equip is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquipItem(ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "EquipItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.EquipItem)
        
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "ChangeHeading" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeHeading(ByVal heading As E_Heading)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeHeading" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeHeading)
        
        Call .WriteByte(heading)
    End With
End Sub

''
' Writes the "ModifySkills" message to the outgoing data buffer.
'
' @param    skillEdt a-based array containing for each skill the number of points to add to it.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteModifySkills(ByRef skillEdt() As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ModifySkills" message to the outgoing data buffer
'***************************************************
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ModifySkills)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(skillEdt(i))
        Next i
    End With
End Sub

''
' Writes the "Train" message to the outgoing data buffer.
'
' @param    creature Position within the list provided by the server of the creature to train against.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrain(ByVal creature As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Train" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Train)
        
        Call .WriteByte(creature)
    End With
End Sub

''
' Writes the "CommerceBuy" message to the outgoing data buffer.
'
' @param    slot Position within the NPC's inventory in which the desired item is.
' @param    amount Number of items to buy.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceBuy(ByVal slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceBuy" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceBuy)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

''
' Writes the "BankExtractItem" message to the outgoing data buffer.
'
' @param    slot Position within the bank in which the desired item is.
' @param    amount Number of items to extract.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractItem(ByVal slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankExtractItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractItem)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

''
' Writes the "CommerceSell" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to sell.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceSell(ByVal slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceSell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceSell)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

''
' Writes the "BankDeposit" message to the outgoing data buffer.
'
' @param    slot Position within the user inventory in which the desired item is.
' @param    amount Number of items to deposit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDeposit(ByVal slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankDeposit" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDeposit)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

''
' Writes the "ForumPost" message to the outgoing data buffer.
'
' @param    title The message's title.
' @param    message The body of the message.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForumPost(ByVal Title As String, ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForumPost" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ForumPost)
        
        Call .WriteASCIIString(Title)
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "MoveSpell" message to the outgoing data buffer.
'
' @param    upwards True if the spell will be moved up in the list, False if it will be moved downwards.
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveSpell(ByVal upwards As Boolean, ByVal slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MoveSpell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveSpell)
        
        Call .WriteBoolean(upwards)
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "ClanCodexUpdate" message to the outgoing data buffer.
'
' @param    desc New description of the clan.
' @param    codex New codex of the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteClanCodexUpdate(ByVal Desc As String, ByRef Codex() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ClanCodexUpdate" message to the outgoing data buffer
'***************************************************
    Dim temp As String
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ClanCodexUpdate)
        
        Call .WriteASCIIString(Desc)
        
        For i = LBound(Codex()) To UBound(Codex())
            temp = temp & Codex(i) & SEPARATOR
        Next i
        
        If Len(temp) Then _
            temp = Left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)
    End With
End Sub

''
' Writes the "UserCommerceOffer" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to offer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOffer(ByVal slot As Byte, ByVal Amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceOffer" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UserCommerceOffer)
        
        Call .WriteByte(slot)
        Call .WriteLong(Amount)
    End With
End Sub


''
' Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer.
'
' @param    username The user who wants to join the guild whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestJoinerInfo(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestJoinerInfo)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GuildNewWebsite" message to the outgoing data buffer.
'
' @param    url The guild's new website's URL.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNewWebsite(ByVal url As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildNewWebsite" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildNewWebsite)
        
        Call .WriteASCIIString(url)
    End With
End Sub

''
' Writes the "GuildAcceptNewMember" message to the outgoing data buffer.
'
' @param    username The name of the accepted player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptNewMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAcceptNewMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptNewMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GuildRejectNewMember" message to the outgoing data buffer.
'
' @param    username The name of the rejected player.
' @param    reason The reason for which the player was rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectNewMember(ByVal UserName As String, ByVal reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRejectNewMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectNewMember)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
    End With
End Sub

''
' Writes the "GuildKickMember" message to the outgoing data buffer.
'
' @param    username The name of the kicked player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildKickMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildKickMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildKickMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GuildUpdateNews" message to the outgoing data buffer.
'
' @param    news The news to be posted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildUpdateNews(ByVal news As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildUpdateNews" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildUpdateNews)
        
        Call .WriteASCIIString(news)
    End With
End Sub

''
' Writes the "GuildMemberInfo" message to the outgoing data buffer.
'
' @param    username The user whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildMemberInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMemberInfo)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub


''
' Writes the "GuildRequestMembership" message to the outgoing data buffer.
'
' @param    guild The guild to which to request membership.
' @param    application The user's application sheet.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestMembership(ByVal guild As String, ByVal Application As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestMembership" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestMembership)
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(Application)
    End With
End Sub

''
' Writes the "GuildRequestDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestDetails(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestDetails" message to the outgoing data buffer
'***************************************************
    With outgoingData
        
        If 1 < Len(guild) Then
            Call .WriteByte(ClientPacketID.GuildRequestDetails)
            Call .WriteASCIIString(guild)
        End If
        
    End With
End Sub

''
' Writes the "Online" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnline()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Online" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Online)
End Sub
'Add Marius
Public Sub WritePena()
    Call outgoingData.WriteByte(ClientPacketID.Pena)
End Sub
Public Sub WriteRequestStats()
    Call outgoingData.WriteByte(ClientPacketID.RequestStats)
End Sub
'\Add

''
' Writes the "Quit" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteQuit()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/16/08
'Writes the "Quit" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Quit)
    
    
    
End Sub

''
' Writes the "GuildLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeave()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildLeave" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildLeave)
End Sub

''
' Writes the "RequestAccountState" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAccountState()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestAccountState" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestAccountState)
End Sub

''
' Writes the "PetStand" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetStand()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PetStand" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PetStand)
End Sub

''
' Writes the "PetFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetFollow()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PetFollow" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PetFollow)
End Sub

''
' Writes the "TrainList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TrainList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.TrainList)
End Sub

''
' Writes the "Rest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Rest" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Rest)
End Sub


Public Sub WriteCasament(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.Casamiento)
        Call .WriteASCIIString(UserName)
    End With
End Sub
 
Public Sub WriteAcepto(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.Acepto)
        Call .WriteASCIIString(UserName)
    End With
End Sub
 
Public Sub WriteDivorciate(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.Divorcio)
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "Meditate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Meditate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Meditate)
End Sub

''
' Writes the "Resucitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResucitate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Resucitate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Resucitate)
End Sub

''
' Writes the "Heal" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHeal()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Heal" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Heal)
End Sub

''
' Writes the "Help" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHelp()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Help" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Help)
End Sub


''
' Writes the "CommerceStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceStart()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceStart" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CommerceStart)
End Sub

''
' Writes the "BankStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankStart()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankStart" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.BankStart)
End Sub

''
' Writes the "Enlist" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEnlist()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Enlist" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Enlist)
End Sub

''
' Writes the "Information" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInformation()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Information" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Information)
End Sub

''
' Writes the "Reward" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReward()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Reward" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Reward)
End Sub

''
' Writes the "UpTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpTime()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpTime" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UpTime)
End Sub

''
' Writes the "PartyLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyLeave()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyLeave" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GrupoLeave)
End Sub


''
' Writes the "GuildMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestDetails" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "PartyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GrupoMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "CentinelReport" message to the outgoing data buffer.
'
' @param    number The number to report to the centinel.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCentinelReport(ByVal Number As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CentinelReport" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CentinelReport)
        
        Call .WriteInteger(Number)
    End With
End Sub

''
' Writes the "GuildOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnline()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOnline" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildOnline)
End Sub


''
' Writes the "RoleMasterRequest" message to the outgoing data buffer.
'
' @param    message The message to send to the role masters.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoleMasterRequest(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RoleMasterRequest" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RoleMasterRequest)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "GMRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMRequest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMRequest" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMRequest)
End Sub

''
' Writes the "BugReport" message to the outgoing data buffer.
'
' @param    message The message explaining the reported bug.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBugReport(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BugReport" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.bugReport)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "ChangeDescription" message to the outgoing data buffer.
'
' @param    desc The new description of the user's character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeDescription(ByVal Desc As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeDescription" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeDescription)
        
        Call .WriteASCIIString(Desc)
    End With
End Sub

''
' Writes the "Gamble" message to the outgoing data buffer.
'
' @param    amount The amount to gamble.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGamble(ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Gamble" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Gamble)
        
        Call .WriteInteger(Amount)
    End With
End Sub


''
' Writes the "LeaveFaction" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeaveFaction(Optional action As Byte = 0)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LeaveFaction" message to the outgoing data buffer
'***************************************************

    Call outgoingData.WriteByte(ClientPacketID.LeaveFaction)
    Call outgoingData.WriteByte(action)
End Sub
Public Sub WriteHogar()
'***************************************************
'Author: Jose Ignacio Castelli (fedudok)
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Hogar)
    End With
End Sub

Public Sub WriteParticipar(ByVal Message As String)
'***************************************************
'Author: Jose Ignacio Castelli (fedudok)
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Participar)
        Call .WriteASCIIString(Message)
    End With
End Sub


''
' Writes the "BankExtractGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to extract from the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractGold(ByVal Amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankExtractGold" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractGold)
        
        Call .WriteLong(Amount)
    End With
End Sub
''
' Writes the "BankExtractGold" message to the outgoing data buffer.
'<
' @param    amount The amount of money to extract from the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankTransferGold(ByVal Amount As Long, ByVal name As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.BankTransferGold)
        
        Call .WriteLong(Amount)
        Call .WriteASCIIString(UCase$(name))
    End With
End Sub


''
' Writes the "BankDepositGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to deposit in the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDepositGold(ByVal Amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankDepositGold" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDepositGold)
        
        Call .WriteLong(Amount)
    End With
End Sub

''
' Writes the "Denounce" message to the outgoing data buffer.
'
' @param    message The message to send with the denounce.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDenounce(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Denounce" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Denounce)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteFriends(ByVal Acc As Byte, ByVal Message As String)
'***************************************************
'Author: Marius
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Friends)
        
        Call .WriteByte(Acc)
        Call .WriteASCIIString(Message)
    End With
End Sub


''
' Writes the "GuildFundate" message to the outgoing data buffer.
'
' @param    clanType The alignment of the clan to be founded.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildFundate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildFundate" message to the outgoing data buffer
'***************************************************

    With outgoingData
        Call .WriteByte(ClientPacketID.GuildFundate)
    End With
End Sub

''
' Writes the "PartyKick" message to the outgoing data buffer.
'
' @param    username The user to kick fro mthe party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GrupoKick)
            
        Call .WriteASCIIString(UserName)
    End With
End Sub



''
' Writes the "GuildMemberList" message to the outgoing data buffer.
'
' @param    guild The guild whose member list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberList(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildMemberList" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.GuildMemberList)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GMMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to the other GMs online.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.GMMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "ShowName" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowName()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowName" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.showName)
    End With
End Sub

Public Sub WriteCancelTorneo()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowName" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.CancelTorneo)
    End With
End Sub

Public Sub WriteCrearTorneo(ByVal rondas As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowName" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.CrearTorneo)
        Call .WriteInteger(rondas)
    End With
End Sub

''
' Writes the "OnlineArmada" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineArmada()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineArmada" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
    Call outgoingData.WriteByte(gMessages.OnlineArmada)
End Sub

''
' Writes the "OnlineCaos" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineCaos()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineCaos" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.OnlineCaos)
End Sub

Public Sub WriteOnlineMilicia()
'***************************************************
'Author: Marius
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.OnlineMilicia)
End Sub

''
' Writes the "GoNearby" message to the outgoing data buffer.
'
' @param    username The suer to approach.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoNearby(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GoNearby" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.GoNearby)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "Comment" message to the outgoing data buffer.
'
' @param    message The message to leave in the log as a comment.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteComment(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Comment" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.comment)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "ServerTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerTime()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerTime" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.serverTime)
End Sub

''
' Writes the "Where" message to the outgoing data buffer.
'
' @param    username The user whose position is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhere(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Where" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.Where)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "CreaturesInMap" message to the outgoing data buffer.
'
' @param    map The map in which to check for the existing creatures.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreaturesInMap(ByVal map As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreaturesInMap" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.CreaturesInMap)
        
        Call .WriteInteger(map)
    End With
End Sub

''
' Writes the "WarpMeToTarget" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpMeToTarget()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarpMeToTarget" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.WarpMeToTarget)
End Sub

''
' Writes the "WarpChar" message to the outgoing data buffer.
'
' @param    username The user to be warped. "YO" represent's the user's char.
' @param    map The map to which to warp the character.
' @param    x The x position in the map to which to waro the character.
' @param    y The y position in the map to which to waro the character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpChar(ByVal UserName As String, ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarpChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.WarpChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteInteger(map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub



Public Sub WriteSilence(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Silence" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.Silence)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "SOSShowList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSShowList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SOSShowList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.SOSShowList)
End Sub

''
' Writes the "SOSRemove" message to the outgoing data buffer.
'
' @param    username The user whose SOS call has been already attended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSRemove(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SOSRemove" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.SOSRemove)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GoToChar" message to the outgoing data buffer.
'
' @param    username The user to be approached.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoToChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GoToChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.GoToChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "invisible" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInvisible()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "invisible" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.invisible)
End Sub

''
' Writes the "GMPanel" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMPanel()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMPanel" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.GMPanel)
End Sub

''
' Writes the "RequestUserList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestUserList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestUserList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.RequestUserList)
End Sub

''
' Writes the "Working" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorking()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Working" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.Working)
End Sub

''
' Writes the "Hiding" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHiding()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Hiding" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.Hiding)
End Sub

''
' Writes the "Jail" message to the outgoing data buffer.
'
' @param    username The user to be sent to jail.
' @param    reason The reason for which to send him to jail.
' @param    time The time (in minutes) the user will have to spend there.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteJail(ByVal UserName As String, ByVal reason As String, ByVal time As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Jail" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.Jail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
        
        Call .WriteByte(time)
    End With
End Sub

''
' Writes the "KillNPC" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPC()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillNPC" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.KillNPC)
End Sub

''
' Writes the "WarnUser" message to the outgoing data buffer.
'
' @param    username The user to be warned.
' @param    reason Reason for the warning.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarnUser(ByVal UserName As String, ByVal reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarnUser" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.WarnUser)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
    End With
End Sub

Public Sub WriteSlashSlash(ByVal Commands As String)
'***************************************************
'Author: Marius
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.SlashSlash)
        
        Call .WriteASCIIString(Commands)
    End With
End Sub

''
' Writes the "EditChar" message to the outgoing data buffer.
'
' @param    UserName    The user to be edited.
' @param    editOption  Indicates what to edit in the char.
' @param    arg1        Additional argument 1. Contents depend on editoption.
' @param    arg2        Additional argument 2. Contents depend on editoption.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEditChar(ByVal UserName As String, ByVal editOption As eEditOptions, ByVal arg1 As String, ByVal arg2 As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "EditChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.EditChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteByte(editOption)
        
        Call .WriteASCIIString(arg1)
        Call .WriteASCIIString(arg2)
    End With
End Sub


''
' Writes the "ReviveChar" message to the outgoing data buffer.
'
' @param    username The user to eb revived.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReviveChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReviveChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.ReviveChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "OnlineGM" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineGM()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineGM" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.OnlineGM)
End Sub

''
' Writes the "OnlineMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineMap(ByVal map As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'Writes the "OnlineMap" message to the outgoing data buffer
'26/03/2009: Now you don't need to be in the map to use the comand, so you send the map to server
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.OnlineMap)
        
        Call .WriteInteger(map)
    End With
End Sub

''
' Writes the "Kick" message to the outgoing data buffer.
'
' @param    username The user to be kicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Kick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.Kick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WritePejotas(ByVal UserName As String)
'***************************************************
'Author: Marius
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.Pejotas)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub



''
' Writes the "Execute" message to the outgoing data buffer.
'
' @param    username The user to be executed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteExecute(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Execute" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.Execute)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "BanChar" message to the outgoing data buffer.
'
' @param    username The user to be banned.
' @param    reason The reson for which the user is to be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanChar(ByVal UserName As String, ByVal reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BanChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.BanChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteASCIIString(reason)
    End With
End Sub

''
' Writes the "UnbanChar" message to the outgoing data buffer.
'
' @param    username The user to be unbanned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UnbanChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.UnbanChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "NPCFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCFollow()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCFollow" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.NPCFollow)
End Sub

''
' Writes the "SummonChar" message to the outgoing data buffer.
'
' @param    username The user to be summoned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSummonChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SummonChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.SummonChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "SpawnListRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnListRequest()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnListRequest" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.SpawnListRequest)
End Sub



''
' Writes the "SpawnCreature" message to the outgoing data buffer.
'
' @param    creatureIndex The index of the creature in the spawn list to be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnCreature(ByVal creatureIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnCreature" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.SpawnCreature)
        
        Call .WriteInteger(creatureIndex)
    End With
End Sub

''
' Writes the "ResetNPCInventory" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetNPCInventory()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetNPCInventory" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.ResetNPCInventory)
End Sub

''
' Writes the "CleanWorld" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanWorld()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CleanWorld" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.CleanWorld)
End Sub

''
' Writes the "ServerMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.ServerMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

' Writes the "NickToIP" message to the outgoing data buffer.
'
' @param    username The user whose IP is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNickToIP(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NickToIP" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.NickToIP)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "IPToNick" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIPToNick(ByRef Ip() As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "IPToNick" message to the outgoing data buffer
'***************************************************
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.IPToNick)
        
        For i = LBound(Ip()) To UBound(Ip())
            Call .WriteByte(Ip(i))
        Next i
    End With
End Sub

''
' Writes the "GuildOnlineMembers" message to the outgoing data buffer.
'
' @param    guild The guild whose online player list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnlineMembers(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOnlineMembers" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.GuildOnlineMembers)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "TeleportCreate" message to the outgoing data buffer.
'
' @param    map the map to which the teleport will lead.
' @param    x The position in the x axis to which the teleport will lead.
' @param    y The position in the y axis to which the teleport will lead.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportCreate(ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TeleportCreate" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.TeleportCreate)
        
        Call .WriteInteger(map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "TeleportDestroy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportDestroy()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TeleportDestroy" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.TeleportDestroy)
End Sub


''
' Writes the "SetCharDescription" message to the outgoing data buffer.
'
' @param    desc The description to set to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetCharDescription(ByVal Desc As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetCharDescription" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.SetCharDescription)
        
        Call .WriteASCIIString(Desc)
    End With
End Sub

''
' Writes the "ForceMIDIToMap" message to the outgoing data buffer.
'
' @param    midiID The ID of the midi file to play.
' @param    map The map in which to play the given midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIToMap(ByVal midiID As Byte, ByVal map As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceMIDIToMap" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.ForceMIDIToMap)
        
        Call .WriteByte(midiID)
        
        Call .WriteInteger(map)
    End With
End Sub

''
' Writes the "ForceWAVEToMap" message to the outgoing data buffer.
'
' @param    waveID  The ID of the wave file to play.
' @param    Map     The map into which to play the given wave.
' @param    x       The position in the x axis in which to play the given wave.
' @param    y       The position in the y axis in which to play the given wave.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEToMap(ByVal waveID As Byte, ByVal map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceWAVEToMap" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.ForceWAVEToMap)
        
        Call .WriteByte(waveID)
        
        Call .WriteInteger(map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub


''
' Writes the "TalkAsNPC" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalkAsNPC(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TalkAsNPC" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.TalkAsNPC)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "DestroyAllItemsInArea" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyAllItemsInArea()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DestroyAllItemsInArea" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.DestroyAllItemsInArea)
End Sub

''
' Writes the "ItemsInTheFloor" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemsInTheFloor()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ItemsInTheFloor" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.ItemsInTheFloor)
End Sub

''
' Writes the "MakeDumb" message to the outgoing data buffer.
'
' @param    username The name of the user to be made dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumb(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MakeDumb" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.MakeDumb)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "MakeDumbNoMore" message to the outgoing data buffer.
'
' @param    username The name of the user who will no longer be dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumbNoMore(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MakeDumbNoMore" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.MakeDumbNoMore)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "DumpIPTables" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumpIPTables()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DumpIPTables" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.DumpIPTables)
End Sub


''
' Writes the "SetTrigger" message to the outgoing data buffer.
'
' @param    trigger The type of trigger to be set to the tile.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetTrigger" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.SetTrigger)
        
        Call .WriteByte(Trigger)
    End With
End Sub

''
' Writes the "AskTrigger" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAskTrigger()
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 04/13/07
'Writes the "AskTrigger" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.AskTrigger)
End Sub

''
' Writes the "BannedIPList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPList()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BannedIPList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.BannedIPList)
End Sub

''
' Writes the "BannedIPReload" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPReload()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BannedIPReload" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.BannedIPReload)
End Sub

''
' Writes the "GuildBan" message to the outgoing data buffer.
'
' @param    guild The guild whose members will be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildBan(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildBan" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.GuildBan)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "BanIP" message to the outgoing data buffer.
'
' @param    byIp    If set to true, we are banning by IP, otherwise the ip of a given character.
' @param    IP      The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @param    nick    The nick of the player whose ip will be banned.
' @param    reason  The reason for the ban.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanIP(ByVal byIp As Boolean, ByRef Ip() As Byte, ByVal Nick As String, ByVal reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BanIP" message to the outgoing data buffer
'***************************************************
    If byIp And UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.BanIP)
        
        Call .WriteBoolean(byIp)
        
        If byIp Then
            For i = LBound(Ip()) To UBound(Ip())
                Call .WriteByte(Ip(i))
            Next i
        Else
            Call .WriteASCIIString(Nick)
        End If
        
        Call .WriteASCIIString(reason)
    End With
End Sub

''
' Writes the "UnbanIP" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanIP(ByRef Ip() As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UnbanIP" message to the outgoing data buffer
'***************************************************
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.UnbanIP)
        
        For i = LBound(Ip()) To UBound(Ip())
            Call .WriteByte(Ip(i))
        Next i
    End With
End Sub

''
' Writes the "CreateItem" message to the outgoing data buffer.
'
' @param    itemIndex The index of the item to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateItem(ByVal itemIndex As Long, ByVal count As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.CreateItem)
        
        If itemIndex < 1 Or count < 1 Then
            Exit Sub
        End If
        
        Call .WriteInteger(itemIndex)
        Call .WriteInteger(count)
    End With
End Sub

''
' Writes the "DestroyItems" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyItems()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DestroyItems" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.DestroyItems)
End Sub

''
' Writes the "ChaosLegionKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Chaos Legion.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChaosLegionKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.ChaosLegionKick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RoyalArmyKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Royal Army.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RoyalArmyKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.RoyalArmyKick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteMiliciaKick(ByVal UserName As String)
'***************************************************
'Author: Marius
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.MiliciaKick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ForceMIDIAll" message to the outgoing data buffer.
'
' @param    midiID The id of the midi file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIAll(ByVal midiID As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceMIDIAll" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.ForceMIDIAll)
        
        Call .WriteByte(midiID)
    End With
End Sub

''
' Writes the "ForceWAVEAll" message to the outgoing data buffer.
'
' @param    waveID The id of the wave file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEAll(ByVal waveID As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceWAVEAll" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.ForceWAVEAll)
        
        Call .WriteByte(waveID)
    End With
End Sub


''
' Writes the "TileBlockedToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTileBlockedToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TileBlockedToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.TileBlockedToggle)
End Sub

''
' Writes the "KillNPCNoRespawn" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPCNoRespawn()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillNPCNoRespawn" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.KillNPCNoRespawn)
End Sub

''
' Writes the "KillAllNearbyNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillAllNearbyNPCs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillAllNearbyNPCs" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.KillAllNearbyNPCs)
End Sub

''
' Writes the "LastIP" message to the outgoing data buffer.
'
' @param    username The user whose last IPs are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLastIP(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LastIP" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.LastIP)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "SystemMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to all players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSystemMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SystemMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.SystemMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "CreateNPC" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPC(ByVal NPCIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateNPC" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.CreateNPC)
        
        Call .WriteInteger(NPCIndex)
    End With
End Sub

''
' Writes the "CreateNPCWithRespawn" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPCWithRespawn(ByVal NPCIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateNPCWithRespawn" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.CreateNPCWithRespawn)
        
        Call .WriteInteger(NPCIndex)
    End With
End Sub

''
' Writes the "NavigateToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NavigateToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.NavigateToggle)
End Sub

''
' Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerOpenToUsersToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.ServerOpenToUsersToggle)
End Sub

''
' Writes the "TurnOffServer" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnOffServer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TurnOffServer" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.TurnOffServer)
End Sub

''
' Writes the "TurnCriminal" message to the outgoing data buffer.
'
' @param    username The name of the user to turn into criminal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnCriminal(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TurnCriminal" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.TurnCriminal)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ResetFactions" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any faction.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetFactions(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetFactions" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.ResetFactions)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RemoveCharFromGuild" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharFromGuild(ByVal UserName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharFromGuild" message to the outgoing data buffer
'***************************************************
    With outgoingData
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.RemoveCharFromGuild)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub




''
' Writes the "ToggleCentinelActivated" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteToggleCentinelActivated()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ToggleCentinelActivated" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.ToggleCentinelActivated)
End Sub

''
' Writes the "DoBackup" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoBackup()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DoBackup" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.DoBackUp)
End Sub

''
' Writes the "ShowGuildMessages" message to the outgoing data buffer.
'
' @param    guild The guild to listen to.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildMessages(ByVal guild As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGuildMessages" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.ShowGuildMessages)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "SaveMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveMap()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SaveMap" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.SaveMap)
End Sub

''
' Writes the "ChangeMapInfoPK" message to the outgoing data buffer.
'
' @param    isPK True if the map is PK, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoPK(ByVal isPK As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMapInfoPK" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.ChangeMapInfoPK)
        
        Call .WriteBoolean(isPK)
    End With
End Sub

''
' Writes the "ChangeMapInfoBackup" message to the outgoing data buffer.
'
' @param    backup True if the map is to be backuped, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoBackup(ByVal backup As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMapInfoBackup" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.ChangeMapInfoBackup)
        
        Call .WriteBoolean(backup)
    End With
End Sub

''
' Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer.
'
' @param    restrict NEWBIES (only newbies), NO (everyone), ARMADA (just Armadas), CAOS (just caos) or FACCION (Armadas & caos only)
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoRestricted(ByVal restrict As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.ChangeMapInfoRestricted)
        
        Call .WriteASCIIString(restrict)
    End With
End Sub

''
' Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer.
'
' @param    nomagic TRUE if no magic is to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoMagic(ByVal nomagic As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.ChangeMapInfoNoMagic)
        
        Call .WriteBoolean(nomagic)
    End With
End Sub

''
' Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer.
'
' @param    noinvi TRUE if invisibility is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvi(ByVal noinvi As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.ChangeMapInfoNoInvi)
        
        Call .WriteBoolean(noinvi)
    End With
End Sub
                            
''
' Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer.
'
' @param    noresu TRUE if resurection is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoResu(ByVal noresu As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.ChangeMapInfoNoResu)
        
        Call .WriteBoolean(noresu)
    End With
End Sub
                        
''
' Writes the "ChangeMapInfoLand" message to the outgoing data buffer.
'
' @param    land options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoLand(ByVal land As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoLand" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.ChangeMapInfoLand)
        
        Call .WriteASCIIString(land)
    End With
End Sub
                        
''
' Writes the "ChangeMapInfoZone" message to the outgoing data buffer.
'
' @param    zone options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoZone(ByVal zone As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoZone" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.ChangeMapInfoZone)
        
        Call .WriteASCIIString(zone)
    End With
End Sub

''
' Writes the "SaveChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveChars()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SaveChars" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.SaveChars)
End Sub

''
' Writes the "CleanSOS" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanSOS()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CleanSOS" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.CleanSOS)
End Sub

''
' Writes the "KickAllChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKickAllChars()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KickAllChars" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.KickAllChars)
End Sub

''
' Writes the "ReloadNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadNPCs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadNPCs" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.ReloadNPCs)
End Sub

''
' Writes the "ReloadServerIni" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadServerIni()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadServerIni" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.ReloadServerIni)
End Sub

''
' Writes the "ReloadSpells" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadSpells()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadSpells" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.ReloadSpells)
End Sub

''
' Writes the "ReloadObjects" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadObjects()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadObjects" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.ReloadObjects)
End Sub

''
' Writes the "Restart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestart()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Restart" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.Restart)
End Sub

''

''
' Writes the "Ignored" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIgnored()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Ignored" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.MessagesGM)
        Call outgoingData.WriteByte(gMessages.Ignored)
End Sub

''
' Writes the "CheckSlot" message to the outgoing data buffer.
'
' @param    UserName    The name of the char whose slot will be checked.
' @param    slot        The slot to be checked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCheckSlot(ByVal UserName As String, ByVal slot As Byte)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "CheckSlot" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MessagesGM)
        Call .WriteByte(gMessages.CheckSlot)
        Call .WriteASCIIString(UserName)
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "Ping" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePing()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/01/2007
'Writes the "Ping" message to the outgoing data buffer
'***************************************************
    'Prevent the timer from being cut
    If pingTime <> 0 Then Exit Sub
    
    Call outgoingData.WriteByte(ClientPacketID.Ping)
    
    ' Avoid computing errors due to frame rate
Call FlushBuffer
    DoEvents
    
    pingTime = 1
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Sends all data existing in the buffer
'***************************************************
    Dim sndData As String
    
    With outgoingData
        If .length = 0 Then _
            Exit Sub
        
        
        sndData = .ReadASCIIStringFixed(.length)
        
        Call SendData(sndData)
    End With
End Sub
Sub Login()

    If EstadoLogin = E_MODO.Normal Then
        Call WriteLoginExistingChar
    ElseIf EstadoLogin = E_MODO.CrearNuevoPj Then
        Call WriteLoginNewChar
    ElseIf EstadoLogin = E_MODO.ConectarCuenta Then
        Call WriteLoginAccount
    End If
    DoEvents
    Call FlushBuffer
    
   
End Sub

''
' Sends the data using the socket controls in the MainForm.
'
' @param    sdData  The data to be sent to the server.

Private Sub SendData(ByRef sdData As String)

    
    If Not frmMain.Socket1.IsWritable Then
   
        'Put data back in the bytequeue
        Call outgoingData.WriteASCIIStringFixed(sdData)
        
        Exit Sub
    End If
    
    If Not frmMain.Socket1.Connected Then Exit Sub


    'Dim data() As Byte
    'data = StrConv(sdData, vbFromUnicode)
    'Security.NAC_E_Byte data, Security.Redundance
    'sdData = StrConv(data, vbUnicode)
    
    
    Call frmMain.Socket1.Write(sdData, Len(sdData))
    
    On Error GoTo error:
        frmMain.Winsock1.SendData sdData
        Exit Sub
error:
        MsgBox Err.Description
        
End Sub
Private Sub HandleParticle()
    Dim Particula As Integer
    Dim Y As Byte
    Dim X As Byte
    Dim Life As Integer

    'Remove packet ID
    Call incomingData.ReadByte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    
    Particula = incomingData.ReadInteger()


'If Not InMapBounds(x, y) And Not Particula < 1 And Not Particula > TotalStreams Then
        Call TileEngine.General_Particle_Create(Particula, X, Y)
'End If
End Sub
Private Sub HandleCharParticle()
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Particula As Integer
    Dim Life As Integer
    Dim CharIndex As Integer
    
    Particula = incomingData.ReadInteger()
    CharIndex = incomingData.ReadInteger
    
    If Not Particula < 1 And Not Particula > TotalStreams Then
        Call TileEngine.General_Char_Particle_Create(Particula, CharIndex)
    End If
End Sub
Private Sub HandleDestParticle()
    Dim Y As Byte
    Dim X As Byte

    'Remove packet ID
    Call incomingData.ReadByte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()

   ' If Not InMapBounds(x, y) And MapData(x, y).particle_group_index Then
        Call TileEngine.Particle_Group_Remove(MapData(X, Y).particle_group_index)
   ' End If
End Sub
Private Sub HandleDestCharParticle()
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Particula As Integer
    Dim CharIndex As Integer
    
    Particula = incomingData.ReadInteger()
    CharIndex = incomingData.ReadInteger
    
    If Not Particula < 1 And Not Particula > TotalStreams Then
        Call TileEngine.Char_Particle_Group_Remove(CharIndex, Particula)
    End If
End Sub
Public Sub HandleAddPj()

    If incomingData.length < 11 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim name As String
    Dim Head As Integer, body As Integer
    Dim Casco As Byte, Arma As Byte, Escudo As Byte, Nivel As Byte, Clase As Byte, Mapa As Integer, Index As Byte, color As Byte, tipPet As Byte
    
    With incomingData
        Call .ReadByte
        Index = .ReadByte
        name = .ReadASCIIString
        Head = .ReadInteger()
        body = .ReadInteger()
        Casco = .ReadByte()
        Arma = .ReadByte()
        Escudo = .ReadByte()
        Nivel = .ReadByte()
        Mapa = .ReadInteger()
        Clase = .ReadByte()
        color = .ReadByte()
        
        tipPet = .ReadByte
    End With
    
    cPJ(Index - 1).nombre = name
    cPJ(Index - 1).Head = Head
    cPJ(Index - 1).body = body
    cPJ(Index - 1).Casco = Casco
    cPJ(Index - 1).Weapon = Arma
    cPJ(Index - 1).Shield = Escudo
    cPJ(Index - 1).Nivel = Nivel
    cPJ(Index - 1).Clase = Clase
    cPJ(Index - 1).Mapa = Mapa
    cPJ(Index - 1).color = color
    cPJ(Index - 1).tipPet = tipPet

    Call TileEngine.Char_Account_Render(Index)
End Sub



Private Sub HandleShowAccount()
    'Remove packet ID
    Call incomingData.ReadByte
    
    
    frmPanelAccount.Show
    frmConnect.Visible = False
    frmPanelAccount.lblAccData(0).Caption = UserAccount
    
    Dim i As Long
    
    For i = 0 To 9
        cPJ(i).body = 0
        cPJ(i).Casco = 0
        cPJ(i).Clase = 0
        cPJ(i).color = 0
        cPJ(i).Head = 0
        cPJ(i).Mapa = 0
        cPJ(i).Nivel = 0
        cPJ(i).nombre = ""
        cPJ(i).Shield = 0
        cPJ(i).tipPet = 0
        cPJ(i).Weapon = 0
        
        frmPanelAccount.lblAccData(i + 1) = ""
    Next i
    
    frmPanelAccount.lblCharData(0) = ""
    frmPanelAccount.lblCharData(1) = ""
    frmPanelAccount.lblCharData(2) = ""
    
    cCursores.Parse_Form frmPanelAccount
End Sub
Public Sub WriteLoginExistingChar()

    With outgoingData

        Call .WriteByte(ClientPacketID.LoginExistingChar)
        Call .WriteByte(frmPanelAccount.Seleccionado + 1)
    End With
End Sub
Public Sub WriteLoginAccount()
    With outgoingData
        Call .WriteByte(ClientPacketID.ConnectAccount)
        Call .WriteASCIIString(UserAccount)
        Call .WriteASCIIString(UserPassword)
        'Call .WriteByte(11)
    End With
End Sub
Public Sub WriteLoginNewAccount()
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNewAccount)
        
        Call .WriteASCIIString(UserAccount)
        Call .WriteASCIIString(UserPassword)
        Call .WriteASCIIString(UserEmail)
    End With
End Sub
Public Sub WriteLoginNewChar()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LoginNewChar" message to the outgoing data buffer
'***************************************************
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginNewChar)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(UserAccount)
        
        Call .WriteByte(UserRaza)
        Call .WriteByte(UserSexo)
        Call .WriteByte(UserClase)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(UserSkills(i))
        Next i
        
        Call .WriteByte(UserHogar)
        
        Call .WriteByte(frmCrearPersonaje.lbAtt(0).Caption)
        Call .WriteByte(frmCrearPersonaje.lbAtt(1).Caption)
        Call .WriteByte(frmCrearPersonaje.lbAtt(3).Caption)
        Call .WriteByte(frmCrearPersonaje.lbAtt(4).Caption)
        Call .WriteByte(frmCrearPersonaje.lbAtt(2).Caption)
        
        Call .WriteByte(PetType)
        Call .WriteASCIIString(PetName)
        
        Call .WriteInteger(frmCrearPersonaje.Actual)
        
    End With
End Sub
Private Sub HandleFuerza()
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim xVal As Byte
    
    xVal = incomingData.ReadByte()
    
    frmMain.lblFU.Caption = str(xVal)
    
    If xVal >= 35 Then
        frmMain.lblFU.ForeColor = RGB(192, 64, 0)
    Else
        frmMain.lblFU.ForeColor = RGB(255, 255, 255)
    End If
End Sub
Private Sub HandleAgilidad()
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim xVal As Byte
    
    xVal = incomingData.ReadByte()
    
    frmMain.lblAG.Caption = str(xVal)
    
    If xVal >= 35 Then
        frmMain.lblAG.ForeColor = RGB(192, 64, 0)
    Else
        frmMain.lblAG.ForeColor = RGB(255, 255, 255)
    End If
End Sub

Public Sub WriteSOfrecer(ByVal sI As Byte, ByVal cant As Long)
    With outgoingData
        Call .WriteByte(ClientPacketID.Subasta)
        Call .WriteByte(1)
        Call .WriteLong(cant)
        Call .WriteByte(sI)
    End With
End Sub

Public Sub WriteSSubastar(ByVal slot As Byte, ByVal Cantidad As Integer, ByVal optDura As Byte, ByVal prcOfert As Long, Optional ByVal fnlOfert)
    With outgoingData
        Call .WriteByte(ClientPacketID.Subasta)
        Call .WriteByte(2)
        Call .WriteByte(slot)
        Call .WriteInteger(Cantidad)
        Call .WriteByte(optDura)
        Call .WriteLong(prcOfert)
        Call .WriteLong(fnlOfert)
    End With
End Sub
Public Sub WriteSComprar(ByVal sI As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.Subasta)
        Call .WriteByte(3)
        Call .WriteByte(sI)
    End With
End Sub
Public Sub WriteSRequest()
    With outgoingData
        Call .WriteByte(ClientPacketID.Subasta)
        Call .WriteByte(0)
    End With
End Sub

Private Sub HandleSubastRequest()
    Dim i As Byte, cant As Byte, sI As Byte
    With incomingData
        Call .ReadByte
        cant = .ReadByte
        frmSubastas.LimpiarSubastas
        For i = 1 To cant
            sI = .ReadByte
            lstSubastas(i).active = True
            lstSubastas(i).mnDura = .ReadByte
            lstSubastas(i).hsDura = .ReadByte
            lstSubastas(i).actOfert = .ReadLong
            lstSubastas(i).fnlOfert = .ReadLong
            lstSubastas(i).cant = .ReadInteger
            lstSubastas(i).OBJIndex = .ReadASCIIString
            lstSubastas(i).nckCmprdor = .ReadASCIIString
            lstSubastas(i).nckVndedor = .ReadASCIIString
            lstSubastas(i).grhindex = .ReadLong
        Next i
        frmSubastas.Show , frmMain
        Call frmSubastas.RefreshList
    End With
End Sub
Private Sub HandleHora()
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
    Call incomingData.ReadByte
    tHora = incomingData.ReadByte
    tMinuto = incomingData.ReadByte
    
    frmMain.imgHora.ToolTipText = TileEngine.Get_Time_String
    
    TileEngine.Meteo_Change_Time
End Sub
Private Sub HandleGrupo()
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Call incomingData.ReadByte
    
    GrupoIndex = incomingData.ReadInteger
    IsLeader = (incomingData.ReadByte = 1)
End Sub
Private Sub HandleGrupoForm()
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Call incomingData.ReadByte
    
    Dim NumMembers As Byte, i As Long
    NumMembers = incomingData.ReadByte
    
    frmGrupo.lstGrupo.Clear
    
    For i = 1 To NumMembers
        frmGrupo.lstGrupo.AddItem incomingData.ReadASCIIString
    Next i
    
    frmGrupo.Show , frmMain
    
    frmGrupo.cmdExpulsar.Enabled = IsLeader
    frmGrupo.cmdInvitar.Enabled = IsLeader
End Sub
Public Sub WriteRequestGrupo()
    outgoingData.WriteByte ClientPacketID.RequestGrupo
End Sub
Public Sub WriteDuelo(ByVal opt As Byte)
    outgoingData.WriteByte ClientPacketID.Duelo
    outgoingData.WriteByte opt
End Sub
Sub HandleMessages()
    incomingData.ReadByte
    
    Dim msg As Byte
    Dim CharIndex As Integer
    Dim Daño As Integer
    
    msg = incomingData.ReadByte
    
    
    Select Case msg
        Case 0
            AddToChat "¡¡Estás muerto!!"
            
        Case 1
            AddToChat "¡¡Estás muerto!! Solo podes usar items cuando estas vivo."
        
        Case 2
            AddToChat "No puedes atacar estando muerto."
            
        Case 3
            AddToChat "¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos."
    
        Case 4
            AddToChat "No podés atacar porque estas muerto"
            
        Case 5
            AddToChat "No podés robar porque estas muerto"
            
        Case 6
            AddToChat "No podes lanzar hechizos porque estas muerto."
        
        Case 7
            AddToChat "¡¡Estás muerto!! Solo podés usar meditar cuando estás vivo."
        
        Case 8
            AddToChat "El personaje no existe!!"
            
        Case 9
            AddToChat "El mensaje ha sido borrado correctamente."
        
        Case 10
            AddToChat "Haz extraido el item!!"
            
        Case 11
            AddToChat "El mensaje ha sido entregado correctamente."
            
        Case 12
            AddToChat "Hubo un error en el sistema y el mensaje no ha podido ser entregado."
            
        Case 13
            AddToChat "No tienes esa cantidad."
            
        Case 14
            AddToChat "Haz click derecho sobre el veterinario."
        
        Case 15
            AddToChat "Tienes que poseer 65 skills en domar animales para adoptar una mascota."
    
        Case 16
            AddToChat "No tienes clan."
            
        Case 17
            AddToChat "No eres lider del clan."
        
        Case 18
            AddToChat "El clan se elimino con existo."
            
        Case 19
            AddToChat "Hubo un error al intentar borrar el clan."
    
        Case 20
            AddToChat "El lider del clan ha decidido eliminar el clan por ende has sido expulsado automaticamente."
    
        Case 21
            AddtoRichTextBox frmMain.RecCombat, "Has ganado " & incomingData.ReadLong() & " puntos de experiencia.", 51, 183, 247, True, False
    
        Case 22
            AddToChat "El arma que posees ocupa tus dos manos y no puedes tener el escudo."
        
        Case 23
            AddToChat "El arma abarca tus dos manos, desocupalas sacandote este escudo!!"
            
        Case 24
            CharIndex = incomingData.ReadInteger
            Daño = incomingData.ReadInteger
            
            TileEngine.Char_Dialog_Create CharIndex, IIf(Daño > 140, "¡" & Daño & "!", Daño), -65536, 2
            
        Case 25
            CharIndex = incomingData.ReadInteger
            
            TileEngine.Char_Dialog_Create CharIndex, "'Falla'", -65536, 1
              
        Case 26
            TileEngine.Char_Dialog_Create char_current, "'Falla'", -65536, 1
    
        Case 27
            AddToChat "En " & incomingData.ReadByte & " segundos cerrará el juego..."
    
        Case 28
            AddToChat "No podés lanzar ese hechizo a un muerto."
            
        Case 29
            AddToChat "No puedes estar invisible navegando!!"
            
        Case 30
            AddToChat "¡El hechizo no tiene efecto!"
            
        Case 31
            AddToChat "¡No puedes ponerte invisible mientras te encuentres saliendo!"
            
        Case 32
            AddToChat "¡La invisibilidad no funciona aquí!"
    
        Case 33
            AddToChat "No puedes beneficiar a ese tipo de gente."
            
        Case 34
            AddToChat "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, 2
    
        Case 35
            AddToChat "¡¡ No perteneces a una facción !!"
            
        Case 36
           AddToChat "¡Has ganado " & incomingData.ReadLong() & " monedas de oro!", FontTypeNames.FONTTYPE_BROWNB, 2
      
        
        Case 37
            AddtoRichTextBox frmMain.RecCombat, "Has ganado " & incomingData.ReadLong() & " puntos de experiencia.", 51, 183, 247, True, False
    
        Case 38
            CharIndex = incomingData.ReadInteger
            Daño = incomingData.ReadInteger
            
            AddToChat charlist(CharIndex).nombre & " te ha quitado " & Daño & " puntos de daño.", FontTypeNames.FONTTYPE_FIGHT, 2
        
        Case 39
            CharIndex = incomingData.ReadInteger
            Daño = incomingData.ReadInteger
            
            AddToChat "Le has quitado " & Daño & " puntos de vida a " & charlist(CharIndex).nombre, FontTypeNames.FONTTYPE_FIGHT, 2
            
        Case 40
            CharIndex = incomingData.ReadInteger
            Daño = incomingData.ReadInteger
            
            AddToChat "¡Has mejorado tu skill " & SkillsNames(CharIndex) & " en un punto!. Ahora tienes " & Daño & " pts.", FontTypeNames.FONTTYPE_INFO, 2
    
        Case 41
            AddToChat "El usuario no posee espacio en su casilla."
            
        Case 42
            AddToChat "/salir cancelado."
            
        Case 43
            AddToChat "Gracias por jugar InmortalAO."
            
        Case 44
            frmMensaje.msg.Caption = "Contraseña incorrecta o la cuenta ya esta conectada."
            frmMensaje.Show
            
        Case 45
            frmMensaje.msg.Caption = "La cuenta no existe. Si todavia no tienes una puedes crearla desde nuestro sitio web www.inmortalao.com.ar"
            frmMensaje.Show
            
        Case 46
            frmMensaje.msg.Caption = "Servidor restringido a administradores."
            frmMensaje.Show
         
        Case 47
            frmMensaje.msg.Caption = "Demasiado tiempo inactivo."
            frmMensaje.Show
            
        Case 48
            frmMensaje.msg.Caption = "La version que posees es antigua. Puedes actualizarla automaticamente ejecutando el launcher. Para mas informacion visite www.Inmortalao.com.ar"
            frmMensaje.Show

    End Select
End Sub

Public Sub WriteExtractItem(ByVal Index As Byte)
    With outgoingData
        .WriteByte ClientPacketID.ExtraerItem
        
        .WriteByte Index
    End With
End Sub
Public Sub WriteBorrarMensaje(ByVal Index As Byte)
    With outgoingData
        .WriteByte ClientPacketID.BorrarMensaje
        
        .WriteByte Index
    End With
End Sub
Public Sub WriteEnviarMensaje(ByRef mensaje As String, ByVal slot As Byte, ByRef para As String, ByVal Cantidad As Integer)
    With outgoingData
        .WriteByte ClientPacketID.EnviarMensaje
        
        .WriteByte slot
        
        If Len(mensaje) > 200 Then mensaje = Left$(mensaje, 200)
        .WriteASCIIString mensaje
        
        If Len(para) > 30 Then para = Left$(para, 30)
        .WriteASCIIString para
        
        .WriteInteger Cantidad
    End With
End Sub
Sub HandleShorFamiliarForm()
    incomingData.ReadByte
    
    frmSeleccionFamiliar.Show , frmMain
End Sub
Sub WriteAdoptarMascota()
    outgoingData.WriteByte ClientPacketID.AdoptarMascota
    
    outgoingData.WriteByte PetType
    outgoingData.WriteASCIIString PetName
End Sub
Sub WriteDelClan()
    outgoingData.WriteByte ClientPacketID.DelClan
End Sub
Sub HandleCharMsgStatus()
    With incomingData
        Dim CharIndex As Integer
        Dim tempstr As String
        Dim lngPorcVida As Long
    
        Dim btClase As Byte, btRaza As Byte, btNivel As Byte, btStatus As Byte
        Dim btMatrimonioLen As Byte, btDescLen As Byte, btRed As Byte, btGreen As Byte, btBlue As Byte
        Dim St1 As Byte, St2 As Byte, st As String
        
        .ReadByte
    
        CharIndex = .ReadInteger
        btStatus = .ReadByte
        lngPorcVida = .ReadLong
        St1 = .ReadByte
        St2 = .ReadByte
        btMatrimonioLen = .ReadByte
        btClase = .ReadByte
        btNivel = .ReadByte
        btRaza = .ReadByte
        btDescLen = .ReadByte

        st = Generate_Char_Status(lngPorcVida, BoolToByte((St2 And StatEx.Paralizado)), BoolToByte((St2 And StatEx.Inmovilizado)), _
                BoolToByte((St1 And Stat.Envenenado)), BoolToByte((St1 And Stat.Trabajando)), BoolToByte((St1 And Stat.Silenciado)), BoolToByte((St1 And Stat.Ciego)), _
                BoolToByte((St1 And Stat.Incinerado)), BoolToByte((St1 And Stat.Transformado)), BoolToByte((St1 And Stat.Comerciand)), _
                BoolToByte((St1 And Stat.Inactivo)))
                
        tempstr = charlist(CharIndex).nombre & " (" & ListaClases(btClase) & " " & ListaRazas(btRaza) & " "
        
        If btNivel = 255 Then
            tempstr = tempstr & "??"
        Else
            tempstr = tempstr & btNivel
        End If
        
        tempstr = tempstr & " " & st & ")"
        
        Select Case btStatus
            Case 1 'Imperial
                tempstr = tempstr & " <Imperial>"
                btRed = 32
                btGreen = 81
                btBlue = 251
            Case 2 'Renegado
                tempstr = tempstr & " <Renegado>"
                btRed = 114
                btGreen = 115
                btBlue = 108
            Case 3 'Republicano
                tempstr = tempstr & " <Republicano>"
                btRed = 204
                btGreen = 107
                btBlue = 0
            Case 5 'Caos
                tempstr = tempstr & " <Fuerzas del caos> <" & General_Tittle_Caos(.ReadByte) & ">"
                btRed = 196
                btGreen = 0
                btBlue = 15
            Case 6 'Imperial
                tempstr = tempstr & " <Ejercito Real> <" & General_Tittle_Real(.ReadByte) & ">"
                btRed = 32
                btGreen = 81
                btBlue = 251
            Case 7 'Republicano
                tempstr = tempstr & " <Milicia Republicana> <" & General_Tittle_Milicia(.ReadByte) & ">"
                btRed = 204
                btGreen = 107
                btBlue = 0
            
            Case 10
                tempstr = tempstr & " <CONSEJERO>"
                btRed = 2
                btGreen = 162
                btBlue = 38
            Case 11
                tempstr = tempstr & " <SEMI-DIOS>"
                btRed = 2
                btGreen = 162
                btBlue = 38
            Case 12
                tempstr = tempstr & " <DIOS>"
                btRed = 2
                btGreen = 162
                btBlue = 38
            Case 13
                tempstr = tempstr & " <ADMIN>"
                btRed = 2
                btGreen = 162
                btBlue = 38
        End Select

        
        If charlist(CharIndex).offClanX > 0 Then
            tempstr = tempstr & " " & charlist(CharIndex).clan & ""
        End If
        
        If btMatrimonioLen > 0 Then
            If St2 And StatEx.Hombre Then
                tempstr = tempstr & " <Marido de " & .ReadASCIIStringFixed(btMatrimonioLen) & ">"
            Else
                tempstr = tempstr & " <Mujer de " & .ReadASCIIStringFixed(btMatrimonioLen) & ">"
            End If
        End If
        
        If btDescLen > 0 Then
            tempstr = tempstr & " - " & .ReadASCIIStringFixed(btDescLen)
        End If
        
        AddtoRichTextBox frmMain.RecChat, tempstr, btRed, btGreen, btBlue, 1

    End With
End Sub

Sub WriteChatFaccion(ByRef chat As String)
    With outgoingData
        .WriteByte ClientPacketID.ChatFaccion
        
        .WriteASCIIString chat
    End With
End Sub
Sub HandleMensajeSigno()
    With incomingData
        .ReadByte
        
        Mensajes = (.ReadByte = 1)
        
        If Mensajes Then
            frmMain.nuevocorreo.Visible = True
        Else
            frmMain.nuevocorreo.Visible = False
        End If
    End With
End Sub
Public Sub WriteDragAndDrop(ByVal s1 As Byte, ByVal s2 As Byte)
    outgoingData.WriteByte ClientPacketID.DragAndDrop
    outgoingData.WriteByte s1
    outgoingData.WriteByte s2
End Sub
