Attribute VB_Name = "Protocol"

Option Explicit

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

''
'The last existing client packet id.
Private Const LAST_CLIENT_PACKET_ID As Byte = 103

''
'Auxiliar ByteQueue used as buffer to generate messages not intended to be sent right away.
'Specially usefull to create a message once and send it over to several clients.
Private auxiliarBuffer As New clsByteQueue

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
    online                  '53
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
    Friends                 '107 /FADD /FDEL /FLIST /FMSG
End Enum

Public Enum gMessages
    GMMessage               '/GMSG
    showName                '/SHOWNAME
    OnlineArmada            '/ONLINEREAL
    OnlineCaos              '/ONLINECAOS
    OnlineMilicia           '/ONLINEMILI
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
    Invisible               '/INVISIBLE
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

''
' Handles incoming data.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIncomingData(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'
'***************************************************

On Error Resume Next
    Dim packetID As String
    
    packetID = UserList(UserIndex).incomingData.PeekByte()

    'Does the packet requires a logged user??
    If Not (packetID = ClientPacketID.ConnectAccount _
      Or packetID = ClientPacketID.LoginExistingChar _
      Or packetID = ClientPacketID.LoginNewChar Or packetID = ClientPacketID.CreateNewAccount) Then

        'Is the user actually logged?
        If Not UserList(UserIndex).flags.UserLogged Then
            Call CloseSocket(UserIndex)
            Exit Sub
        
        'He is logged. Reset idle counter if id is valid.
        ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
            UserList(UserIndex).Counters.IdleCount = 0
        End If
    ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
        UserList(UserIndex).Counters.IdleCount = 0
        
        'Is the user logged?
        If UserList(UserIndex).flags.UserLogged Then
       
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    End If

    
    Select Case packetID
        Case ClientPacketID.ConnectAccount
            Call HandleLoginAccount(UserIndex)
         
        Case ClientPacketID.CreateNewAccount
            'Call HandleLoginNewAccount(UserIndex)
            
        Case ClientPacketID.LoginExistingChar       'OLOGIN
            Call HandleLoginExistingChar(UserIndex)
     
        Case ClientPacketID.LoginNewChar            'NLOGIN
            Call HandleLoginNewChar(UserIndex)
        
        Case ClientPacketID.Talk                    ';
            Call HandleTalk(UserIndex)
        
        Case ClientPacketID.Whisper                 '\
            Call HandleWhisper(UserIndex)
        
        Case ClientPacketID.Walk                    'M
            Call HandleWalk(UserIndex)
        
        Case ClientPacketID.RequestPositionUpdate   'RPU
            Call HandleRequestPositionUpdate(UserIndex)
        
        Case ClientPacketID.Attack                  'AT
            Call HandleAttack(UserIndex)
            
        Case ClientPacketID.PickUp                  'AG
            Call HandlePickUp(UserIndex)
        
        Case ClientPacketID.CombatModeToggle        'TAB        - SHOULD BE HANLDED JUST BY THE CLIENT!!
            Call HanldeCombatModeToggle(UserIndex)
        
        Case ClientPacketID.SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
            Call HandleSafeToggle(UserIndex)
        
        Case ClientPacketID.RequestEstadisticas
            Call HandleRequestEstadisticas(UserIndex)
            
        Case ClientPacketID.RequestGuildLeaderInfo
            Call HandleRequestGuildLeaderInfo(UserIndex)

        Case ClientPacketID.CommerceEnd             'FINCOM
            Call HandleCommerceEnd(UserIndex)
        
        Case ClientPacketID.UserCommerceEnd         'FINCOMUSU
            Call HandleUserCommerceEnd(UserIndex)
        
        Case ClientPacketID.BankEnd                 'FINBAN
            Call HandleBankEnd(UserIndex)
        
        Case ClientPacketID.BankTransferGold
            Call HandleBankTransferGold(UserIndex)
        
        Case ClientPacketID.UserCommerceOk          'COMUSUOK
            Call HandleUserCommerceOk(UserIndex)
        
        Case ClientPacketID.UserCommerceReject      'COMUSUNO
            Call HandleUserCommerceReject(UserIndex)
        
        Case ClientPacketID.Drop                    'TI
            Call HandleDrop(UserIndex)
        
        Case ClientPacketID.CastSpell               'LH
            Call HandleCastSpell(UserIndex)
        
        Case ClientPacketID.LeftClick               'LC
            Call HandleLeftClick(UserIndex)
        
        Case ClientPacketID.DoubleClick             'RC
            Call HandleDoubleClick(UserIndex)
        
        Case ClientPacketID.Work                    'UK
            Call HandleWork(UserIndex)
        
        Case ClientPacketID.UseItem                 'USA
            Call HandleUseItem(UserIndex)
        
        Case ClientPacketID.CraftBlacksmith         'CNS
            Call HandleCraftBlacksmith(UserIndex)
        
        Case ClientPacketID.CraftCarpenter          'CNC
            Call HandleCraftCarpenter(UserIndex)
        
        Case ClientPacketID.CraftSastre
        Call HandleCraftSastre(UserIndex)
        
         Case ClientPacketID.Craftalquimia
        Call HandleCraftalquimia(UserIndex)
        
        
        Case ClientPacketID.WorkLeftClick           'WLC
            Call HandleWorkLeftClick(UserIndex)
        
        Case ClientPacketID.CreateNewGuild          'CIG
            Call HandleCreateNewGuild(UserIndex)
        
        Case ClientPacketID.SpellInfo               'INFS
            Call HandleSpellInfo(UserIndex)
        
        Case ClientPacketID.EquipItem               'EQUI
            Call HandleEquipItem(UserIndex)
        
        Case ClientPacketID.ChangeHeading           'CHEA
            Call HandleChangeHeading(UserIndex)
        
        Case ClientPacketID.ModifySkills            'SKSE
            Call HandleModifySkills(UserIndex)
        
        Case ClientPacketID.Train                   'ENTR
            Call HandleTrain(UserIndex)
        
        Case ClientPacketID.CommerceBuy             'COMP
            Call HandleCommerceBuy(UserIndex)
        
        Case ClientPacketID.BankExtractItem         'RETI
            Call HandleBankExtractItem(UserIndex)
        
        Case ClientPacketID.CommerceSell            'VEND
            Call HandleCommerceSell(UserIndex)
        
        Case ClientPacketID.BankDeposit             'DEPO
            Call HandleBankDeposit(UserIndex)
        
        Case ClientPacketID.ForumPost               'DEMSG
            Call HandleForumPost(UserIndex)
        
        Case ClientPacketID.MoveSpell               'DESPHE
            Call HandleMoveSpell(UserIndex)
        
        Case ClientPacketID.ClanCodexUpdate         'DESCOD
            Call HandleClanCodexUpdate(UserIndex)
        
        Case ClientPacketID.UserCommerceOffer       'OFRECER
            Call HandleUserCommerceOffer(UserIndex)
        
        Case ClientPacketID.GuildRequestJoinerInfo  'ENVCOMEN
            Call HandleGuildRequestJoinerInfo(UserIndex)

        Case ClientPacketID.GuildNewWebsite         'NEWWEBSI
            Call HandleGuildNewWebsite(UserIndex)
        
        Case ClientPacketID.GuildAcceptNewMember    'ACEPTARI
            Call HandleGuildAcceptNewMember(UserIndex)
        
        Case ClientPacketID.GuildRejectNewMember    'RECHAZAR
            Call HandleGuildRejectNewMember(UserIndex)
        
        Case ClientPacketID.GuildKickMember         'ECHARCLA
            Call HandleGuildKickMember(UserIndex)
        
        Case ClientPacketID.GuildUpdateNews         'ACTGNEWS
            Call HandleGuildUpdateNews(UserIndex)
        
        Case ClientPacketID.GuildMemberInfo         '1HRINFO<
            Call HandleGuildMemberInfo(UserIndex)

        Case ClientPacketID.GuildRequestMembership  'SOLICITUD
            Call HandleGuildRequestMembership(UserIndex)
        
        Case ClientPacketID.GuildRequestDetails     'CLANDETAILS
            Call HandleGuildRequestDetails(UserIndex)
                  
        Case ClientPacketID.online                  '/ONLINE
            Call HandleOnline(UserIndex)
        'Add Marius
        Case ClientPacketID.Pena                    '/PENA
            Call HandlePena(UserIndex)
            
        Case ClientPacketID.RequestStats            '/EST
            Call HandleRequestStats(UserIndex)
        '\Add
        
        Case ClientPacketID.Quit                    '/SALIR
            Call HandleQuit(UserIndex)
        
        Case ClientPacketID.GuildLeave              '/SALIRCLAN
            Call HandleGuildLeave(UserIndex)
        
        Case ClientPacketID.RequestAccountState     '/BALANCE
            Call HandleRequestAccountState(UserIndex)
        
        Case ClientPacketID.PetStand                '/QUIETO
            Call HandlePetStand(UserIndex)
        
        Case ClientPacketID.PetFollow               '/ACOMPAÑAR
            Call HandlePetFollow(UserIndex)
        
        Case ClientPacketID.TrainList               '/ENTRENAR
            Call HandleTrainList(UserIndex)
        
        Case ClientPacketID.Rest                    '/DESCANSAR
            Call HandleRest(UserIndex)
        
        Case ClientPacketID.Meditate                '/MEDITAR
            Call HandleMeditate(UserIndex)
        
        Case ClientPacketID.Resucitate              '/RESUCITAR
            Call HandleResucitate(UserIndex)
        
        Case ClientPacketID.Heal                    '/CURAR
            Call HandleHeal(UserIndex)
        
        Case ClientPacketID.Help                    '/AYUDA
            Call HandleHelp(UserIndex)
        
        Case ClientPacketID.CommerceStart           '/COMERCIAR
            Call HandleCommerceStart(UserIndex)
        
        Case ClientPacketID.BankStart               '/BOVEDA
            Call HandleBankStart(UserIndex)
        
        Case ClientPacketID.Enlist                  '/ENLISTAR
            Call HandleEnlist(UserIndex)
        
        Case ClientPacketID.Information             '/INFORMACION
            Call HandleInformation(UserIndex)
        
        Case ClientPacketID.Reward                  '/RECOMPENSA
            Call HandleReward(UserIndex)
        
        Case ClientPacketID.UpTime                  '/UPTIME
            Call HandleUpTime(UserIndex)
        
        Case ClientPacketID.GrupoLeave              '/SALIRGrupo
            Call HandleGrupoLeave(UserIndex)
        
        Case ClientPacketID.GuildMessage            '/CMSG
            Call HandleGuildMessage(UserIndex)
        
        Case ClientPacketID.GrupoMessage            '/PMSG
            Call HandleGrupoMessage(UserIndex)
        
        Case ClientPacketID.CentinelReport          '/CENTINELA
            Call HandleCentinelReport(UserIndex)
        
        Case ClientPacketID.GuildOnline             '/ONLINECLAN
            Call HandleGuildOnline(UserIndex)

        Case ClientPacketID.RoleMasterRequest       '/ROL
            Call HandleRoleMasterRequest(UserIndex)
        
        Case ClientPacketID.Subasta
            Call HandleSubasta(UserIndex)
                
        Case ClientPacketID.GMRequest               '/GM
            Call HandleGMRequest(UserIndex)
        
        Case ClientPacketID.bugReport               '/_BUG
            Call HandleBugReport(UserIndex)
        
        Case ClientPacketID.ChangeDescription       '/DESC
            Call HandleChangeDescription(UserIndex)
        
        Case ClientPacketID.Gamble                  '/APOSTAR
            Call HandleGamble(UserIndex)

        Case ClientPacketID.LeaveFaction            '/RETIRAR ( with no arguments )
            Call HandleLeaveFaction(UserIndex)
        
        Case ClientPacketID.BankExtractGold         '/RETIRAR ( with arguments )
            Call HandleBankExtractGold(UserIndex)
        
        Case ClientPacketID.BankDepositGold         '/DEPOSITAR
            Call HandleBankDepositGold(UserIndex)
        
        Case ClientPacketID.Denounce                '/DENUNCIAR
            Call HandleDenounce(UserIndex)
        
        Case ClientPacketID.Friends                 '/FADD /FDEL /FLIST /FMSG
            Call HandleFriends(UserIndex)
        
        Case ClientPacketID.GuildFundate            '/FUNDARCLAN
            Call HandleGuildFundate(UserIndex)
        
        Case ClientPacketID.GrupoKick               '/ECHARGrupo
            Call HandleGrupoKick(UserIndex)
            
        Case ClientPacketID.Ping                    '/PING
            Call HandlePing(UserIndex)
        
        Case ClientPacketID.Casamiento
            Call HandleCasament(UserIndex)
 
        Case ClientPacketID.Acepto
            Call handleacepto(UserIndex)
 
        Case ClientPacketID.Divorcio
            Call handledivorcio(UserIndex)

        
        Case ClientPacketID.MessagesGM
            Call HandleMessagesGM(UserIndex)
            
        Case ClientPacketID.RequestGrupo
            Call HandleRequestGrupo(UserIndex)
            
        Case ClientPacketID.Duelo
            Call HandleDuelo(UserIndex)
            
        Case ClientPacketID.BorrarMensaje
            Call HandleBorrarMensaje(UserIndex)
            
        Case ClientPacketID.EnviarMensaje
            Call HandleEnviarMensaje(UserIndex)
        
        Case ClientPacketID.ExtraerItem
            Call HandleExtractItem(UserIndex)
            
        Case ClientPacketID.AdoptarMascota
            Call HandleAdoptarMascota(UserIndex)
            
        Case ClientPacketID.DelClan
            Call HandleDelClan(UserIndex)
            
        Case ClientPacketID.ChatFaccion
            Call HandleChatFaccion(UserIndex)
            
        Case ClientPacketID.DragAndDrop
            Call HandleDragAndDrop(UserIndex)
        
        Case ClientPacketID.Hogar
            Call HandleHogar(UserIndex)
            
        Case ClientPacketID.Participar          '/PARTICIPAR <EVENTO>
            Call HandleParticipar(UserIndex)

        Case Else
            Call CloseSocket(UserIndex)


    End Select
    
    'Done with this packet, move on to next one or send everything if no more packets found
    If UserList(UserIndex).incomingData.length > 0 And err.Number = 0 Then
        err.Clear
        Call HandleIncomingData(UserIndex)
    
    ElseIf err.Number <> 0 And Not err.Number = UserList(UserIndex).incomingData.NotEnoughDataErrCode Then
        'An error ocurred, log it and kick player.
        Call LogError("Error: " & err.Number & " [" & err.description & "] " & " Source: " & err.source & _
                        vbTab & " HelpFile: " & err.HelpFile & vbTab & " HelpContext: " & err.HelpContext & _
                        vbTab & " LastDllError: " & err.LastDllError & vbTab & _
                        " - UserIndex: " & UserIndex & " - producido al manejar el paquete: " & CStr(packetID) & " Name: " & UserList(UserIndex).name)
        
            Call CloseSocket(UserIndex)

        
    Else
    
        'Flush buffer - send everything that has been written
        Call FlushBuffer(UserIndex)
    End If
End Sub

    ''
' Handles the "LoginExistingChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginExistingChar(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    Dim UI As Integer
    Dim UserName As String
    
    'Remove packet ID
    Call buffer.ReadByte
    UI = buffer.ReadByte
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
    UserName = leePjSqlCuenta(UserList(UserIndex).IndexAccount, UI)

    If BANCheckDB(UserName) Then
        Call WriteErrorMsg(UserIndex, "Se te ha prohibido la entrada a InmortalAO debido a su mal comportamiento. Puedes consultar el reglamento y el sistema de soporte desde www.InmortalAO.com.ar")
    Else
        Call ConnectUser(UserIndex, UserName, UserList(UserIndex).account)
    End If

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "LoginNewChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginNewChar(ByVal UserIndex As Integer)


'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
If UserList(UserIndex).incomingData.length < 46 Then
    err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
    Exit Sub
End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim UserName As String
    Dim account As String
    Dim IndexAccount As Long 'Add Nod kopfnickend
    
    Dim skills(NUMSKILLS - 1) As Byte
    Dim race As eRaza
    Dim gender As eGenero
    Dim homeland As eCiudad
    Dim Class As eClass
    
    Dim Fuerza As Byte, Agilidad As Byte, Carisma As Byte, constitucion As Byte, Inteligencia As Byte
    Dim Cabeza As Integer, petTipe As eMascota, petName As String
    
    If PuedeCrearPersonajes = 0 Then
        Call WriteErrorMsg(UserIndex, "La creacion de personajes en este servidor se ha deshabilitado.")
        Call FlushBuffer(UserIndex)
        Exit Sub
    End If
    
  '  If ServerSoloGMs <> 0 Then
  '      Call WriteErrorMsg(Userindex, "Servidor restringido a administradores. Consulte la página oficial o el foro oficial para mas información.")
  '      Call FlushBuffer(Userindex)
  '      Exit Sub
  '  End If
    
    UserName = buffer.ReadASCIIString()
    account = buffer.ReadASCIIString()
  
    race = buffer.ReadByte()
    gender = buffer.ReadByte()
    Class = buffer.ReadByte()
    
    Call buffer.ReadBlock(skills, NUMSKILLS)
    
    homeland = buffer.ReadByte()
    
    Fuerza = buffer.ReadByte()
    Agilidad = buffer.ReadByte()
    Carisma = buffer.ReadByte()
    constitucion = buffer.ReadByte()
    Inteligencia = buffer.ReadByte()
    
    petTipe = buffer.ReadByte()
    petName = buffer.ReadASCIIString()
    
    Cabeza = buffer.ReadInteger()

    Call ConnectNewUser(UserIndex, UserName, account, race, gender, Class, skills, "", homeland, _
                        Fuerza, Agilidad, Inteligencia, Carisma, constitucion, Cabeza, petTipe, petName)

    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub


''
' Handles the "Talk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalk(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
        
    'Remove packet ID
    Call buffer.ReadByte
        
    Dim chat As String
    Dim TalkMode As Byte
        
    chat = buffer.ReadASCIIString()
    chat = Replace(chat, "~", "")
    TalkMode = buffer.ReadByte
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
    If UserList(UserIndex).Counters.Silenciado <> 0 Then
        If UserList(UserIndex).flags.UltimoMensaje <> 60 Then
            Call WriteConsoleMsg(1, UserIndex, "Los administrador te han silenciado por mensajes reiterados. Espere a ser desilenciado. Gracias.", FontTypeNames.FONTTYPE_BROWNI)
            UserList(UserIndex).flags.UltimoMensaje = 60
            Exit Sub
        End If
    End If
    UserList(UserIndex).Counters.Habla = UserList(UserIndex).Counters.Habla + 1
        
    Select Case TalkMode
        Case 1 'Normal
            Call TalkNormal(UserIndex, chat)
                
        Case 2 ' Gritar
            Call TalkGritar(UserIndex, chat)
                
        Case 3 ' Global
            Call TalkGlobal(UserIndex, chat)
        
    End Select
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub


''
' Handles the "Whisper" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat As String
        Dim targetPriv As PlayerType
        Dim nameTarget As String
        Dim targetUserIndex As Integer
        
        nameTarget = buffer.ReadASCIIString()
        chat = buffer.ReadASCIIString()
        chat = Replace(chat, "~", "")
        targetUserIndex = NameIndex(nameTarget)
        
        If UserList(UserIndex).Counters.Silenciado <> 0 Then
            If UserList(UserIndex).flags.UltimoMensaje <> 60 Then
                Call WriteConsoleMsg(1, UserIndex, "Los administrador te han silenciado por mensajes reiterados. Espere a ser desilenciado. Gracias.", FontTypeNames.FONTTYPE_BROWNI)
                UserList(UserIndex).flags.UltimoMensaje = 60
                Exit Sub
            End If
        End If
    
        .Counters.Habla = .Counters.Habla + 1
        
        If .flags.Muerto Then
            Call WriteMsg(UserIndex, 0)
        Else
            If targetUserIndex = 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
            Else
                'Manda mensaje privados solo a los que no son dioses
                If (EsCONSE(UserIndex) Or (Not EsCONSE(targetUserIndex))) Then
                    If LenB(chat) <> 0 Then
                        If Not EstaPCarea(UserIndex, targetUserIndex) Then
                            Call WriteConsoleMsg(1, UserIndex, UserList(UserIndex).name & ">" & chat, FontTypeNames.FONTTYPE_PRIVADO)
                            Call WriteConsoleMsg(1, targetUserIndex, UserList(UserIndex).name & ">" & chat, FontTypeNames.FONTTYPE_PRIVADO)
                            Call FlushBuffer(targetUserIndex)
                        Else
                            Call WriteChatOverHead(UserIndex, chat, .Char.CharIndex, RGB(182, 226, 29))
                            Call WriteChatOverHead(targetUserIndex, chat, .Char.CharIndex, RGB(182, 226, 29))
                            Call FlushBuffer(targetUserIndex)
                        End If
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "Walk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim dummy As Long
    Dim TempTick As Long
    Dim heading As eHeading
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        heading = .incomingData.ReadByte()
        
        .flags.TimesWalk = .flags.TimesWalk + 1
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        If .flags.Paralizado = 0 Then
            If .flags.Meditando Then
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                .Char.fx = 0
                .Char.loops = 0
                
                Call WriteMeditateToggle(UserIndex)
                Call WriteConsoleMsg(1, UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_BROWNI)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageDestCharParticle(UserList(UserIndex).Char.CharIndex, ParticleToLevel(UserIndex)))
            End If
            
            'If Not IntervaloPermiteMover(UserIndex) Then
            '    Exit Sub
            'End If
            
            'Move user
            Call MoveUserChar(UserIndex, heading)
            
            'Si se mueve chau portal
            If UserList(UserIndex).Counters.CreoTeleport = True And UserList(UserIndex).Counters.TimeTeleport < 5 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageDestParticle(UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY))

                UserList(UserIndex).flags.TiroPortalL = 0
                UserList(UserIndex).Counters.TimeTeleport = 0
                UserList(UserIndex).Counters.CreoTeleport = False
            End If
           
            'Desc Nod kopfnickend
            'Si estamos resucitando y nos movemos, seguimos resucitando
            'If UserList(UserIndex).flags.Resucitando = 1 Then
            '    UserList(UserIndex).flags.Resucitando = 0
            '    UserList(UserIndex).Counters.IntervaloRevive = 0
            '    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageDestCharParticle(UserList(UserIndex).Char.CharIndex, 22))
            'End If
            
            If UserList(UserIndex).flags.Trabajando = True Then
                UserList(UserIndex).flags.Trabajando = False
                
                Call WriteConsoleMsg(1, UserIndex, "Dejas de trabajar.", FontTypeNames.FONTTYPE_BROWNI)
            End If
            
            If .flags.Entrenando = 1 Then
                Call WriteConsoleMsg(1, UserIndex, "Dejas de trabajar.", FontTypeNames.FONTTYPE_BROWNI)
                .flags.Entrenando = 0
            End If
            
            'Stop resting if needed
            If .flags.Descansar Then
                .flags.Descansar = False
                
                Call WriteRestOK(UserIndex)
                Call WriteConsoleMsg(1, UserIndex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else    'paralized
            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1
                
                Call WriteConsoleMsg(1, UserIndex, "No podes moverte porque estas paralizado.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    UserList(UserIndex).incomingData.ReadByte
    
    Call WritePosUpdate(UserIndex)
End Sub

''
' Handles the "Attack" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAttack(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/01/08
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'If dead, can't attack
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 2)
            Exit Sub
        End If
        
        'If not in combat mode, can't attack
        If Not .flags.ModoCombate Then
            Call WriteConsoleMsg(1, UserIndex, "No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'If user meditates, can't attack
        If .flags.Meditando Then
            Exit Sub
        End If
        
        'If equiped weapon is ranged, can't attack this way
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
                Call WriteConsoleMsg(1, UserIndex, "No podés usar así esta arma.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Attack!
        Call UsuarioAtaca(UserIndex)
        
        'I see you...
        If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            If .flags.Invisible = 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                Call WriteConsoleMsg(1, UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
End Sub

''
' Handles the "PickUp" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePickUp(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'If dead, it can't pick up objects
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            Exit Sub
        End If
        
        Call GetObj(UserIndex)
    End With
End Sub

''
' Handles the "CombatModeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HanldeCombatModeToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.ModoCombate Then
            Call WriteConsoleMsg(1, UserIndex, "Has salido del modo combate.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(1, UserIndex, "Has pasado al modo combate.", FontTypeNames.FONTTYPE_INFO)
        End If
        
        .flags.ModoCombate = Not .flags.ModoCombate
    End With
End Sub

''
' Handles the "SafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSafeToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Seguro Then
            Call WriteSafeModeOff(UserIndex)
        Else
            Call WriteSafeModeOn(UserIndex)
        End If
        
        .flags.Seguro = Not .flags.Seguro
    End With
End Sub

''
' Handles the "RequestGuildLeaderInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestGuildLeaderInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    UserList(UserIndex).incomingData.ReadByte
    
    Call modGuilds.SendGuildLeaderInfo(UserIndex)
End Sub

''
' Handles the "RequestSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestEstadisticas(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Debug.Print "Quiere estadisticas"
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteSendSkills(UserIndex)
    Call WriteMiniStats(UserIndex)
    Call WriteAttributes(UserIndex)
End Sub


''
' Handles the "CommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    'User quits commerce mode
    UserList(UserIndex).flags.Comerciando = False
    Call WriteCommerceEnd(UserIndex)
End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Quits commerce mode with user
        If .ComUsu.DestUsu > 0 And UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
            Call WriteConsoleMsg(1, .ComUsu.DestUsu, .name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(.ComUsu.DestUsu)
            
            'Send data in the outgoing buffer of the other user
            Call FlushBuffer(.ComUsu.DestUsu)
        End If
        
        Call FinComerciarUsu(UserIndex)
    End With
End Sub

''
' Handles the "BankEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'User exits banking mode
        .flags.Comerciando = False
        Call WriteBankEnd(UserIndex)
    End With
End Sub

''
' Handles the "UserCommerceOk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOk(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    'Trade accepted
    Call AceptarComercioUsu(UserIndex)
End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceReject(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim otherUser As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        otherUser = .ComUsu.DestUsu
        
        'Offer rejected
        If otherUser > 0 Then
            If UserList(otherUser).flags.UserLogged Then
                Call WriteConsoleMsg(1, otherUser, .name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(otherUser)
                
                'Send data in the outgoing buffer of the other user
                Call FlushBuffer(otherUser)
            End If
        End If
        
        Call WriteConsoleMsg(1, UserIndex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_TALK)
        Call FinComerciarUsu(UserIndex)
    End With
End Sub

''
' Handles the "Drop" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim Slot As Byte
    Dim Amount As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        
        
        
        'low rank admins can't drop item. Neither can the dead nor those sailing.
        If .flags.Comerciando Or .flags.Navegando = 1 Or .flags.Muerto = 1 Then Exit Sub

        'Are we dropping gold or other items??
        If Slot = FLAGORO Then
            If Amount > 10000 Then Exit Sub 'Don't drop too much gold
            
            'Add Marius Deshabilitamos esta funcion de tirar oro por que bugea las puertas
            'Correjimos lo que bugeaba las puertas era el limpiar mundo, pero igual bugea todo, por que tiran mucho oro
            If EsDIOS(UserIndex) Then
                Call TirarOro(Amount, UserIndex)
                Call WriteUpdateGold(UserIndex)
            End If
        Else
            'Only drop valid slots
            If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
                If .Invent.Object(Slot).ObjIndex = 0 Then
                    Exit Sub
                End If
                
                'Add Marius no se pueden tirar objetos caos o reales
                If ObjData(.Invent.Object(Slot).ObjIndex).Caos = 0 And _
                   ObjData(.Invent.Object(Slot).ObjIndex).Real = 0 And _
                   ObjData(.Invent.Object(Slot).ObjIndex).Milicia = 0 And _
                   ObjData(.Invent.Object(Slot).ObjIndex).Shop < 50 Then
                    Call DropObj(UserIndex, Slot, Amount, .Pos.map, .Pos.x, .Pos.Y)
                End If
            End If
        End If
    End With
End Sub

''
' Handles the "CastSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCastSpell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Spell As Byte
        
        Spell = .incomingData.ReadByte()
        
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            Exit Sub
        End If
        
        .flags.Hechizo = Spell
        
        If .flags.Hechizo < 1 Then
            .flags.Hechizo = 0
        ElseIf .flags.Hechizo > MAXUSERHECHIZOS Then
            .flags.Hechizo = 0
        End If
    End With
End Sub

''
' Handles the "LeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim x As Byte
        Dim Y As Byte
        
        x = .ReadByte()
        Y = .ReadByte()
        
        Call LookatTile(UserIndex, UserList(UserIndex).Pos.map, x, Y)
    End With
End Sub

''
' Handles the "DoubleClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDoubleClick(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim x As Byte
        Dim Y As Byte
        
        x = .ReadByte()
        Y = .ReadByte()
        
        
        Call Accion(UserIndex, UserList(UserIndex).Pos.map, x, Y)
    End With
End Sub

''
' Handles the "Work" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWork(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Skill As eSkill
        
        Skill = .incomingData.ReadByte()
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            Exit Sub
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill
            Case Robar, Magia, Domar
                If Skill = Magia Then
                    
                    If UserList(UserIndex).flags.Hechizo = 0 Then Exit Sub
                    
                    
                    'castelli metamorfosis
                    If Hechizos(UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)).AutoLanzar = 1 Then
                        If PuedeLanzar(UserIndex, UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)) Then
                            UserList(UserIndex).flags.TargetUser = UserIndex
                            Call HandleHechizoUsuario(UserIndex, UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo))
                            Exit Sub
                        Else
                            Exit Sub
                        End If
                    End If
                    'castelli metamorfosis
                
                    If Hechizos(UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)).Anillo = 1 Then
                        If TieneObjetos(ANILLO_ESPECTRAL, 1, UserIndex) Or TieneObjetos(ANILLO_PENUMBRAS, 1, UserIndex) Then
                            Call WriteWorkRequestTarget(UserIndex, Skill)
                        Else
                            Call WriteConsoleMsg(1, UserIndex, "Para poder utilizar este hechizo debes poseer el Anillo Espectral o Anillo de las Penumbras.", FontTypeNames.FONTTYPE_WARNING)
                        End If
                    ElseIf Hechizos(UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)).Anillo = 2 Then
                        If TieneObjetos(ANILLO_PENUMBRAS, 1, UserIndex) Then
                            Call WriteWorkRequestTarget(UserIndex, Skill)
                        Else
                            Call WriteConsoleMsg(1, UserIndex, "Para poder utilizar este hechizo debes poseer el Anillo de las Penumbras.", FontTypeNames.FONTTYPE_WARNING)
                        End If
                    Else
                        Call WriteWorkRequestTarget(UserIndex, Skill)
                    End If
                Else
                    Call WriteWorkRequestTarget(UserIndex, Skill)
                End If
                
            Case Ocultarse
                If .flags.Navegando = 1 Then
                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 3 Then
                        Call WriteConsoleMsg(1, UserIndex, "No podés ocultarte si estás navegando.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 3
                    End If
                    '[/CDT]
                    Exit Sub
                End If
                
                If UserList(UserIndex).flags.Montando = 1 Then
                    If Not UserList(UserIndex).flags.UltimoMensaje = 3 Then
                        Call WriteConsoleMsg(1, UserIndex, "||No podes ocultarte si estas sobre una montura.", FontTypeNames.FONTTYPE_INFO)
                        UserList(UserIndex).flags.UltimoMensaje = 3
                    End If
                    Exit Sub
                End If
                
                'Add Nod Kopfnickend Si no te podes tirar invi, tampoco ocultarte xD
                'No usar invi mapas InviSinEfecto
                If MapInfo(UserList(UserIndex).Pos.map).InviSinEfecto > 0 Then
                    If Not .flags.UltimoMensaje = 3 Then
                        Call WriteConsoleMsg(1, UserIndex, "No podés ocultarte en este mapa.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 3
                    End If
                    Exit Sub
                End If
                '\Add
                
                If .flags.Oculto = 1 Then
                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 2 Then
                        Call WriteConsoleMsg(1, UserIndex, "Ya estás oculto.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 2
                    End If
                    '[/CDT]
                    Exit Sub
                End If
                
                Call DoOcultarse(UserIndex)
        End Select
    End With
End Sub


''
' Handles the "UseItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        
        Slot = .incomingData.ReadByte()
        
        If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
            If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
        End If
        
        If .flags.Meditando Then
            Exit Sub    'The error message should have been provided by the client.
        End If
        
        If ObjData(.Invent.Object(Slot).ObjIndex).OBJType = eOBJType.otItemsMagicos Then
            Call EquiparInvItem(UserIndex, Slot)
        ElseIf ObjData(.Invent.Object(Slot).ObjIndex).OBJType = eOBJType.otArmadura Then
            Call EquiparInvItem(UserIndex, Slot)
        ElseIf ObjData(.Invent.Object(Slot).ObjIndex).OBJType = eOBJType.otNudillos Then
            Call EquiparInvItem(UserIndex, Slot)
        Else
            Call UseInvItem(UserIndex, Slot)
        End If
    End With
End Sub

''
' Handles the "CraftBlacksmith" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftBlacksmith(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        Dim Cant As Integer
        
        Item = .ReadInteger()
        Cant = .ReadInteger()
        
        If Item < 1 Or Cant < 1 Then Exit Sub
        
        If ObjData(Item).SkHerreria = 0 Then Exit Sub
        
        Call HerreroConstruirItem(UserIndex, Item, Cant)
    End With
End Sub

''
' Handles the "CraftCarpenter" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftCarpenter(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        Dim Cant As Integer
        
        Item = .ReadInteger()
        Cant = .ReadInteger()
        
        If Item < 1 Or Cant < 1 Or Cant > 1000 Then Exit Sub
        
        If ObjData(Item).SkCarpinteria = 0 Then Exit Sub
        
        Call CarpinteroConstruirItem(UserIndex, Item, Cant)
    End With
End Sub



Private Sub HandleCraftalquimia(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        Dim Cant As Integer
        
        Item = .ReadInteger()
        Cant = .ReadInteger()
        
        If Item < 1 Or Cant < 1 Or Cant > 1000 Then Exit Sub
        
        If ObjData(Item).SkPociones = 0 Then Exit Sub
        
        Call druidaConstruirItem(UserIndex, Item, Cant)
    End With
End Sub





' Handles the "CraftCarpenter" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftSastre(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        Dim Cant As Integer
        
        Item = .ReadInteger()
        Cant = .ReadInteger()
        
        If Item < 1 Or Cant < 1 Or Cant > 1000 Then Exit Sub
        
        If ObjData(Item).SkSastreria = 0 Then Exit Sub
        
        Call SastreConstruirItem(UserIndex, Item, Cant)
    End With
End Sub



''
' Handles the "WorkLeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo Errhandler

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim x As Byte
        Dim Y As Byte
        Dim Skill As eSkill
        Dim DummyInt As Integer
        Dim tU As Integer   'Target user
        Dim tN As Integer   'Target NPC
        
        x = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        Skill = .incomingData.ReadByte()
        
        
        If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando Or Not InMapBounds(.Pos.map, x, Y) Then
            Exit Sub
        End If
        
        If Not InRangoVision(UserIndex, x, Y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill
            Case eSkill.arrojadizas
                
                DummyInt = 0
                If .Invent.WeaponEqpObjIndex = 0 Then 'Este chabon quiere bugear el sistem xD
                    DummyInt = 1
                ElseIf .Invent.WeaponEqpSlot < 1 Or .Invent.WeaponEqpSlot > MAX_INVENTORY_SLOTS Then
                    DummyInt = 1
                Else
                    If ObjData(.Invent.WeaponEqpObjIndex).SubTipo <> 5 And _
                       ObjData(.Invent.WeaponEqpObjIndex).SubTipo <> 6 Then
                        DummyInt = 1
                    Else
                        If .Invent.Object(.Invent.WeaponEqpSlot).Amount < 1 Then
                            DummyInt = 2
                        End If
                    End If
                End If
                
                If DummyInt <> 0 Then
                    Exit Sub
                Else
                    'Quitamos stamina
                    If .Stats.MinSTA >= 10 Then
                        Call QuitarSta(UserIndex, RandomNumber(1, 10))
                    Else
                        If .Genero = eGenero.Hombre Then
                            Call WriteConsoleMsg(1, UserIndex, "Estas muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(1, UserIndex, "Estas muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        Exit Sub
                    End If
                    
                    Call LookatTile(UserIndex, .Pos.map, x, Y)
                    
                    tU = .flags.TargetUser
                    tN = .flags.TargetNPC
                    If tU > 0 Then
                        'Only allow to atack if the other one can retaliate (can see us)
                        If Abs(UserList(tU).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
                            Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                            Exit Sub
                        End If
                        
                        'Prevent from hitting self
                        If tU = UserIndex Then
                            Call WriteConsoleMsg(1, UserIndex, "¡No puedes atacarte a vos mismo!", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        'Attack!
                        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub 'TODO: Por ahora pongo esto para solucionar lo anterior.
                        Call UsuarioAtacaUsuario(UserIndex, tU)
                    ElseIf tN > 0 Then
                        'Only allow to atack if the other one can retaliate (can see us)
                        If Abs(Npclist(tN).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(Npclist(tN).Pos.x - .Pos.x) > RANGO_VISION_X Then
                            Call WriteConsoleMsg(1, UserIndex, "Estas demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                            Exit Sub
                        End If
                        
                        'Is it attackable???
                        If Npclist(tN).Attackable <> 0 Then
                            'Attack!
                            Call UsuarioAtacaNpc(UserIndex, tN)
                        End If
                    End If
   
                     DummyInt = .Invent.WeaponEqpSlot
                     
                     'Take 1 arrow away - we do it AFTER hitting, since if Ammo Slot is 0 it gives a rt9 and kicks players
                     Call QuitarUserInvItem(UserIndex, DummyInt, 1)
                     
                     Call UpdateUserInv(False, UserIndex, DummyInt)
                End If
                
            Case eSkill.Proyectiles
            
                'Check attack interval
                'If Not IntervaloPermiteAtacar(UserIndex, False) Then Exit Sub
                'Check Magic interval
                'If Not IntervaloPermiteLanzarSpell(UserIndex, False) Then Exit Sub
                'Check bow's interval
                'If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
                
                'Make sure the item is valid and there is ammo equipped.
                With .Invent
                    If .WeaponEqpObjIndex = 0 Then
                        DummyInt = 1
                    ElseIf .WeaponEqpSlot < 1 Or .WeaponEqpSlot > MAX_INVENTORY_SLOTS Then
                        DummyInt = 1
                    ElseIf .MunicionEqpSlot < 1 Or .MunicionEqpSlot > MAX_INVENTORY_SLOTS Then
                        DummyInt = 1
                    ElseIf .MunicionEqpObjIndex = 0 Then
                        DummyInt = 1
                    ElseIf ObjData(.WeaponEqpObjIndex).proyectil <> 1 Then
                        DummyInt = 2
                    ElseIf ObjData(.MunicionEqpObjIndex).OBJType <> eOBJType.otFlechas Then
                        DummyInt = 1
                    ElseIf .Object(.MunicionEqpSlot).Amount < 1 Then
                        DummyInt = 1
                    End If
                    
                    If DummyInt <> 0 Then
                        If DummyInt = 1 Then
                            Call WriteConsoleMsg(1, UserIndex, "No tenés municiones.", FontTypeNames.FONTTYPE_INFO)
                            
                            Call Desequipar(UserIndex, .WeaponEqpSlot)
                        End If
                        
                        Call Desequipar(UserIndex, .MunicionEqpSlot)
                        Exit Sub
                    End If
                End With
                
                'Quitamos stamina
                If .Stats.MinSTA >= 10 Then
                    Call QuitarSta(UserIndex, RandomNumber(1, 10))
                Else
                    If .Genero = eGenero.Hombre Then
                        Call WriteConsoleMsg(1, UserIndex, "Estas muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(1, UserIndex, "Estas muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    Exit Sub
                End If
                
                Call LookatTile(UserIndex, .Pos.map, x, Y)
                
                tU = .flags.TargetUser
                tN = .flags.TargetNPC
                
                'Validate target
                If tU > 0 Then
                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(UserList(tU).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
                        Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                        Exit Sub
                    End If
                    
                    'Prevent from hitting self
                    If tU = UserIndex Then
                        Call WriteConsoleMsg(1, UserIndex, "¡No puedes atacarte a vos mismo!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Attack!
                    If Not PuedeAtacar(UserIndex, tU) Then Exit Sub 'TODO: Por ahora pongo esto para solucionar lo anterior.
                    Call UsuarioAtacaUsuario(UserIndex, tU)
                ElseIf tN > 0 Then
                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(Npclist(tN).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(Npclist(tN).Pos.x - .Pos.x) > RANGO_VISION_X Then
                        Call WriteConsoleMsg(1, UserIndex, "Estas demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                        Exit Sub
                    End If
                    
                    'Is it attackable???
                    If Npclist(tN).Attackable <> 0 Then
                        
                        'Attack!
                        Call UsuarioAtacaNpc(UserIndex, tN)
                    End If
                End If
                
                With .Invent
                    DummyInt = .MunicionEqpSlot
                    
                    'Take 1 arrow away - we do it AFTER hitting, since if Ammo Slot is 0 it gives a rt9 and kicks players
                    Call QuitarUserInvItem(UserIndex, DummyInt, 1)
                    
                    If .Object(DummyInt).Amount > 0 Then
                        'QuitarUserInvItem unequipps the ammo, so we equip it again
                        .MunicionEqpSlot = DummyInt
                        .MunicionEqpObjIndex = .Object(DummyInt).ObjIndex
                        .Object(DummyInt).Equipped = 1
                    Else
                        .MunicionEqpSlot = 0
                        .MunicionEqpObjIndex = 0
                    End If
                    Call UpdateUserInv(False, UserIndex, DummyInt)
                End With
                '-----------------------------------
            
            Case eSkill.Magia
                'Check the map allows spells to be casted.
                If MapInfo(.Pos.map).MagiaSinEfecto > 0 Then
                    Call WriteConsoleMsg(1, UserIndex, "Una fuerza oscura te impide canalizar tu energía", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
                
                'Target whatever is in that tile
                Call LookatTile(UserIndex, .Pos.map, x, Y)
                
                'If it's outside range log it and exit
                If Abs(.Pos.x - x) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
                    Call LogCheating("Ataque fuera de rango de " & .name & "(" & .Pos.map & "/" & .Pos.x & "/" & .Pos.Y & ") ip: " & .ip & " a la posicion (" & .Pos.map & "/" & x & "/" & Y & ")")
                    Exit Sub
                End If
                
                
                'Check bow's interval
                'If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
                
                
                'Check Spell-Hit interval
                'If Not IntervaloPermiteGolpeMagia(UserIndex) Then
                    'Check Magic interval
                    'If Not IntervaloPermiteLanzarSpell(UserIndex) Then
                        'Exit Sub
                    'End If
                'End If
                
                    
                'Check intervals and cast
                If .flags.Hechizo > 0 Then
                    Call LanzarHechizo(.flags.Hechizo, UserIndex)
                    .flags.Hechizo = 0
                Else
                    Call WriteConsoleMsg(1, UserIndex, "¡Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Robar
                'Does the map allow us to steal here?
                If MapInfo(.Pos.map).Pk Then
                    
                    'Check interval
                    If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                    
                    'Target whatever is in that tile
                    Call LookatTile(UserIndex, UserList(UserIndex).Pos.map, x, Y)
                    
                    tU = .flags.TargetUser
                    
                    If tU > 0 And tU <> UserIndex Then
                        'Can't steal administrative players
                        If Not EsCONSE(tU) Then
                            If UserList(tU).flags.Muerto = 0 Then
                                 If Abs(.Pos.x - x) + Abs(.Pos.Y - Y) > 2 Then
                                     Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                     Exit Sub
                                 End If
                                 
                                 '17/09/02
                                 'Check the trigger
                                 If MapData(UserList(tU).Pos.map, x, Y).Trigger = eTrigger.ZONASEGURA Then
                                     Call WriteConsoleMsg(1, UserIndex, "No podés robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                     Exit Sub
                                 End If
                                 
                                 If MapData(.Pos.map, .Pos.x, .Pos.Y).Trigger = eTrigger.ZONASEGURA Then
                                     Call WriteConsoleMsg(1, UserIndex, "No podés robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                     Exit Sub
                                 End If
                                 
                                 Call DoRobar(UserIndex, tU)
                            End If
                        End If
                    Else
                        Call WriteConsoleMsg(1, UserIndex, "No a quien robarle!.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(1, UserIndex, "¡No puedes robar en zonas seguras!.", FontTypeNames.FONTTYPE_INFO)
                End If
            Case eSkill.Domar
                'Modificado 25/11/02
                'Optimizado y solucionado el bug de la doma de
                'criaturas hostiles.
                
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                    
                'Target whatever is that tile
                Call LookatTile(UserIndex, .Pos.map, x, Y)
                tN = .flags.TargetNPC
                
                If tN > 0 Then
                    If Npclist(tN).flags.domable > 0 Then
                        If Abs(.Pos.x - x) + Abs(.Pos.Y - Y) > 2 Then
                            Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        If Npclist(tN).flags.AttackedBy <> 0 Then
                            Call WriteConsoleMsg(1, UserIndex, "No podés domar una criatura que está luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        Call DoDomar(UserIndex, tN)
                    Else
                        Call WriteConsoleMsg(1, UserIndex, "No podés domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(1, UserIndex, "No hay ninguna criatura alli!.", FontTypeNames.FONTTYPE_INFO)
                End If
                
            Case 250
'                Call LookatTile(UserIndex, .Pos.map, X, Y)
'
'                tN = .flags.TargetUser ' Persona a formar grupo
'
'                If UserIndex = tN Then WriteConsoleMsg 1, UserIndex, "No puedes mandar grupo a ti mismo!", FontTypeNames.FONTTYPE_BROWNI: Exit Sub
'
'                If tN > 0 Then
'                    If UserList(UserIndex).GrupoIndex = 0 Then
'                        If UserList(tN).GrupoIndex = 0 Then
'                            'Aca crea uno nuevo
'                            If UserList(UserIndex).flags.Solicito = tN Then
'                                Call mdGrupo.CrearGrupo(tN, UserIndex)
'                            Else
'                                WriteConsoleMsg 1, UserIndex, "La peticion ha llegado. Ahora solo debes esperar una respuesta.", FontTypeNames.FONTTYPE_BROWNI
'                                WriteConsoleMsg 1, tN, UserList(UserIndex).name & "  tiene intenciones de formar grupo. Haz click en grupo y luego en él si deseas aceptar.", FontTypeNames.FONTTYPE_BROWNI
'                                UserList(tN).flags.Solicito = UserIndex
'                            End If
'                        Else
'                            If UserList(tN).flags.Solicito = UserIndex Then
'                                Call mdGrupo.ApruebaSolicitud(tN, UserIndex)
'                                UserList(tN).flags.Solicito = 0
'                            Else
'                                WriteConsoleMsg 1, UserIndex, "La peticion ha llegado. Ahora solo debes esperar una respuesta.", FontTypeNames.FONTTYPE_BROWNI
'                                WriteConsoleMsg 1, tN, UserList(UserIndex).name & " desea entrar a tu grupo. Click en grupo>invitar y luego en el para aceptarlo.", FontTypeNames.FONTTYPE_BROWNI
'                                Call mdGrupo.SolicitarIngresoAGrupo(UserIndex)
'                            End If
'                        End If
'                    Else
'                        If EsLider(UserIndex) Then
'                            If UserList(tN).GrupoSolicitud = UserList(UserIndex).GrupoIndex Then
'                                Call BroadCastGrupo(UserIndex, UserList(tN).name & " ha sido aceptado en el grupo.")
'                                Call mdGrupo.AprobarIngresoAGrupo(UserIndex, tN)
'                            Else
'                                If UserList(UserIndex).flags.Solicito <> tN Then
'                                    If UserList(UserIndex).flags.Solicito > 0 Then
'                                        If UserList(UserList(UserIndex).flags.Solicito).flags.UserLogged Then
'                                            WriteConsoleMsg 1, UserList(UserIndex).flags.Solicito, "La peticion de grupo ha sido cancelada.", FontTypeNames.FONTTYPE_BROWNI
'                                        End If
'                                    End If
'                                    WriteConsoleMsg 1, UserIndex, "La peticion ha llegado. Ahora solo debes esperar una respuesta.", FontTypeNames.FONTTYPE_BROWNI
'                                    WriteConsoleMsg 1, tN, UserList(UserIndex).name & " te está invitando en su grupo. Menu>Grupo si deseas entrar en el.", FontTypeNames.FONTTYPE_BROWNI
'                                    UserList(UserIndex).flags.Solicito = tN
'                                End If
'                            End If
'                        End If
'                    End If
'                End If

        End Select
    End With

Exit Sub
Errhandler:
    Call LogError("Error en HandleWorkLeftClick - Error: " & err.Number & " - Desc: " & err.description & " Name:" & UserList(UserIndex).name)
    
End Sub

''
' Handles the "CreateNewGuild" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateNewGuild(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 9 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim desc As String
        Dim GuildName As String
        Dim site As String
        Dim temp As String
        Dim codex() As String
        Dim errorStr As String
        
        desc = buffer.ReadASCIIString()
        GuildName = buffer.ReadASCIIString()
        site = buffer.ReadASCIIString()
        
        temp = Replace$(buffer.ReadASCIIString, "|", " ")
        Debug.Print "@" & temp
        codex = Split(temp, SEPARATOR)
        
        If modGuilds.CrearNuevoClan(UserIndex, desc, GuildName, site, codex, .FundandoGuildAlineacion, errorStr) Then
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg(1, .name & " fundó el clan " & GuildName & " de alineación " & modGuilds.GuildAlignment(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(SND_NEWCLAN, NO_3D_SOUND, NO_3D_SOUND))

            
            'Update tag
             Call RefreshCharStatus(UserIndex)
        Else
            Call WriteConsoleMsg(1, UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "SpellInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpellInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim spellSlot As Byte
        Dim Spell As Integer
        
        spellSlot = .incomingData.ReadByte()
        
        'Validate slot
        If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(1, UserIndex, "¡Primero selecciona el hechizo.!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate spell in the slot
        Spell = .Stats.UserHechizos(spellSlot)
        If Spell > 0 And Spell < NumeroHechizos + 1 Then
            With Hechizos(Spell)
                'Send information
                Call WriteConsoleMsg(1, UserIndex, "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & vbCrLf _
                                               & "Nombre:" & .Nombre & vbCrLf _
                                               & "Descripción:" & .desc & vbCrLf _
                                               & "Skill requerido: " & .MinSkill & " de magia." & vbCrLf _
                                               & "Mana necesario: " & .ManaRequerido & vbCrLf _
                                               & "Stamina necesaria: " & .StaRequerido & vbCrLf _
                                               & "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%", FontTypeNames.FONTTYPE_INFO)
            End With
        End If
    End With
End Sub

''
' Handles the "EquipItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEquipItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim itemSlot As Byte
        
        itemSlot = .incomingData.ReadByte()
        
        'Dead users can't equip items
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        
        'Validate item slot
        If itemSlot > MAX_INVENTORY_SLOTS Or itemSlot < 1 Then Exit Sub
        
        If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub
        
        Call EquiparInvItem(UserIndex, itemSlot)
    End With
End Sub

''
' Handles the "ChangeHeading" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeHeading(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 06/28/2008
'Last Modified By: NicoNZ
' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
' 06/28/2008: NicoNZ - Sólo se puede cambiar si está inmovilizado.
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim heading As eHeading
        Dim posX As Integer
        Dim posY As Integer
                
        heading = .incomingData.ReadByte()
        
        If 1 = 0 Then
            Select Case heading
                Case eHeading.NORTH
                    posY = -1
                Case eHeading.EAST
                    posX = 1
                Case eHeading.SOUTH
                    posY = 1
                Case eHeading.WEST
                    posX = -1
            End Select
            
            If LegalPos(.Pos.map, .Pos.x + posX, .Pos.Y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then
                Exit Sub
            End If
        End If
        
        'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
        If heading > 0 And heading < 5 Then
            .Char.heading = heading
            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
    End With
End Sub

''
' Handles the "ModifySkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleModifySkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 1 + NUMSKILLS Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim i As Long
        Dim count As Integer
        Dim points(1 To NUMSKILLS) As Byte
        
        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For i = 1 To NUMSKILLS
            points(i) = .incomingData.ReadByte()
            
            If points(i) < 0 Then
                Call LogHackAttemp(.name & " IP:" & .ip & " trató de hackear los skills.")
                .Stats.SkillPts = 0
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            
            count = count + points(i)
        Next i
        
        If count > .Stats.SkillPts Then
            Call LogHackAttemp(.name & " IP:" & .ip & " trató de hackear los skills.")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        With .Stats
            For i = 1 To NUMSKILLS
                .SkillPts = .SkillPts - points(i)
                .UserSkills(i) = .UserSkills(i) + points(i)
                
                'Client should prevent this, but just in case...
                If .UserSkills(i) > 100 Then
                    .SkillPts = .SkillPts + .UserSkills(i) - 100
                    .UserSkills(i) = 100
                End If
            Next i
        End With
    End With
End Sub

''
' Handles the "Train" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrain(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim SpawnedNpc As Integer
        Dim petIndex As Byte
        
        petIndex = .incomingData.ReadByte()
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        If Npclist(.flags.TargetNPC).npctype <> eNPCType.Entrenador Then Exit Sub
        
        If Npclist(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
            If petIndex > 0 And petIndex < Npclist(.flags.TargetNPC).nrocriaturas + 1 Then
                'Create the creature
                SpawnedNpc = SpawnNpc(Npclist(.flags.TargetNPC).Criaturas(petIndex).NpcIndex, Npclist(.flags.TargetNPC).Pos, True, False)
                
                If SpawnedNpc > 0 Then
                    Npclist(SpawnedNpc).MaestroNpc = .flags.TargetNPC
                    Npclist(.flags.TargetNPC).Mascotas = Npclist(.flags.TargetNPC).Mascotas + 1
                End If
            End If
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No puedo traer más criaturas, mata las existentes!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
        End If
    End With
End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
            
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        
        'Only if in commerce mode....
        If Not .flags.Comerciando Then
            Call WriteConsoleMsg(1, UserIndex, "No estás comerciando", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'User compra el item
        Call Comercio(eModoComercio.Compra, UserIndex, .flags.TargetNPC, Slot, Amount)
    End With
End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿Es el banquero?
        If Npclist(.flags.TargetNPC).npctype <> eNPCType.Banquero Then
            Exit Sub
        End If
        
        'User retira el item del slot
        Call UserRetiraItem(UserIndex, Slot, Amount)
    End With
End Sub

''
' Handles the "CommerceSell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceSell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        
        'User compra el item del slot
        Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNPC, Slot, Amount)
    End With
End Sub

''
' Handles the "BankDeposit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDeposit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).npctype <> eNPCType.Banquero Then
            Exit Sub
        End If
        
        'User deposita el item del slot rdata
        Call UserDepositaItem(UserIndex, Slot, Amount)
    End With
End Sub

''
' Handles the "ForumPost" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForumPost(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim file As String
        Dim title As String
        Dim msg As String
        Dim postFile As String
        
        Dim Handle As Integer
        Dim i As Long
        Dim count As Integer
        
        title = buffer.ReadASCIIString()
        msg = buffer.ReadASCIIString()
        
        If .flags.TargetObj > 0 Then
            file = App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & ".for"
            
            If FileExist(file, vbNormal) Then
                count = val(GetVar(file, "INFO", "CantMSG"))
                
                'If there are too many messages, delete the forum
                If count > MAX_MENSAJES_FORO Then
                    For i = 1 To count
                        Kill App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & i & ".for"
                    Next i
                    Kill App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & ".for"
                    count = 0
                End If
            Else
                'Starting the forum....
                count = 0
            End If
            
            Handle = FreeFile()
            postFile = Left$(file, Len(file) - 4) & CStr(count + 1) & ".for"
            
            'Create file
            Open postFile For Output As Handle
            Print #Handle, title
            Print #Handle, msg
            Close #Handle
            
            'Update post count
            Call WriteVar(file, "INFO", "CantMSG", count + 1)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "MoveSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim dir As Integer
        
        If .ReadBoolean() Then
            dir = 1
        Else
            dir = -1
        End If
        
        Call DesplazarHechizo(UserIndex, dir, .ReadByte())
    End With
End Sub

''
' Handles the "ClanCodexUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleClanCodexUpdate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Dim temp As String
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim desc As String
        Dim codex() As String
        
        desc = buffer.ReadASCIIString()
        
        temp = Replace$(buffer.ReadASCIIString, "|", " ")
        codex = Split(temp, SEPARATOR)
        
        Call modGuilds.ChangeCodexAndDesc(desc, codex, .GuildIndex)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "UserCommerceOffer" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOffer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        Dim Slot As Byte
        Dim tUser As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadLong()
        
        'Get the other player
        tUser = .ComUsu.DestUsu
        
        'If amount is invalid, or slot is invalid and it's not gold, then ignore it.
        If ((Slot < 1 Or Slot > MAX_INVENTORY_SLOTS) And Slot <> FLAGORO) _
                        Or Amount <= 0 Then Exit Sub
        
        'Is the other player valid??
        If tUser < 1 Or tUser > MaxUsers Then Exit Sub
        
        'Is the commerce attempt valid??
        If UserList(tUser).ComUsu.DestUsu <> UserIndex Then
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        End If
        
        'Is he still logged??
        If Not UserList(tUser).flags.UserLogged Then
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        Else
            'Is he alive??
            If UserList(tUser).flags.Muerto = 1 Then
                Call FinComerciarUsu(UserIndex)
                Exit Sub
            End If
            
            'Has he got enough??
            If Slot = FLAGORO Then
                'gold
                If Amount > .Stats.GLD Then
                    Call WriteConsoleMsg(1, UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            Else
                'inventory
                If Amount > .Invent.Object(Slot).Amount Then
                    Call WriteConsoleMsg(1, UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
                
                'Add Marius No se pueden mandar cosas faccionarias por comercio xD y tampoco cosas de shop >= 50
                Debug.Print ">Ofertar"
                Debug.Print ">shop: " & ObjData(.Invent.Object(Slot).ObjIndex).Shop
                If ObjData(.Invent.Object(Slot).ObjIndex).Caos > 0 Or ObjData(.Invent.Object(Slot).ObjIndex).Real > 0 Or ObjData(.Invent.Object(Slot).ObjIndex).Milicia > 0 Or ObjData(.Invent.Object(Slot).ObjIndex).Shop > 49 Then
                    Debug.Print ">No lo dejamos"
                    Call WriteConsoleMsg(1, UserIndex, "No puedes comerciar este tipo de objetos.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
                '\Add
                
            End If
            
            'Prevent offer changes (otherwise people would ripp off other players)
            If .ComUsu.Objeto > 0 Then
                Call WriteConsoleMsg(1, UserIndex, "No puedes cambiar tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
            
            'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
            If .flags.Navegando = 1 Then
                If .Invent.BarcoSlot = Slot Then
                    Call WriteConsoleMsg(1, UserIndex, "No podés vender tu barco mientras lo estés usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
            
            'Add Marius
            If .flags.Montando = 1 Then
                If .Invent.MonturaSlot = Slot Then
                    Call WriteConsoleMsg(1, UserIndex, "No podés vender tu montura mientras lo estés usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
            '\Add
            
            .ComUsu.Objeto = Slot
            .ComUsu.Cant = Amount
            
            'If the other one had accepted, we turn that back and inform of the new offer (just to be cautious).
            If UserList(tUser).ComUsu.Acepto = True Then
                UserList(tUser).ComUsu.Acepto = False
                Call WriteConsoleMsg(1, tUser, .name & " ha cambiado su oferta.", FontTypeNames.FONTTYPE_TALK)
            End If
            
            Call EnviarObjetoTransaccion(tUser)
        End If
    End With
End Sub


''
' Handles the "GuildRequestJoinerInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestJoinerInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim User As String
        Dim details As String
        
        User = buffer.ReadASCIIString()
        
        details = modGuilds.a_DetallesAspirante(UserIndex, User)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(1, UserIndex, "El personaje no ha mandado solicitud, o no estás habilitado para verla.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteShowUserRequest(UserIndex, details)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub


''
' Handles the "GuildNewWebsite" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildNewWebsite(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.ActualizarWebSite(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "GuildAcceptNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptNewMember(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim errorStr As String
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If Not modGuilds.a_AceptarAspirante(UserIndex, UserName, errorStr) Then
            Call WriteConsoleMsg(1, UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                Call modGuilds.m_ConectarMiembroAClan(tUser, .GuildIndex)
                Call RefreshCharStatus(tUser)
            End If
            
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg(1, UserName & " ha sido aceptado como miembro del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(SND_NEWMEMBER, NO_3D_SOUND, NO_3D_SOUND))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "GuildRejectNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectNewMember(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim errorStr As String
        Dim UserName As String
        Dim reason As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        reason = buffer.ReadASCIIString()
        
        If Not modGuilds.a_RechazarAspirante(UserIndex, UserName, errorStr) Then
            Call WriteConsoleMsg(1, UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call WriteConsoleMsg(1, tUser, errorStr & " : " & reason, FontTypeNames.FONTTYPE_GUILD)
            Else
                'hay que grabar en el char su rechazo
                Call modGuilds.a_RechazarAspiranteChar(UserName, .GuildIndex, reason)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "GuildKickMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildKickMember(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim GuildIndex As Integer
        
        UserName = buffer.ReadASCIIString()
        
        GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)
        
        If GuildIndex > 0 Then
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(1, UserName & " fue expulsado del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessagePlayWave(SND_OUT, NO_3D_SOUND, NO_3D_SOUND))
        Else
            Call WriteConsoleMsg(1, UserIndex, "No puedes expulsar ese personaje del clan.", FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "GuildUpdateNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildUpdateNews(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.ActualizarNoticias(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "GuildMemberInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.SendDetallesPersonaje(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "GuildRequestMembership" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestMembership(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim application As String
        Dim errorStr As String
        
        guild = buffer.ReadASCIIString()
        application = buffer.ReadASCIIString()
        
        If Not modGuilds.a_NuevoAspirante(UserIndex, guild, application, errorStr) Then
           Call WriteConsoleMsg(1, UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
           Call WriteConsoleMsg(1, UserIndex, "Tu solicitud ha sido enviada. Espera prontas noticias del líder de " & guild & ".", FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "GuildRequestDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestDetails(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.SendGuildDetails(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "Online" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnline(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim count As Long
    Dim quienes As String
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then
            Call WriteConsoleMsg(1, UserIndex, "Servidor> Online: " & CStr(NumUsers) & "/" & CStr(MaxUsers), FontTypeNames.FONTTYPE_SERVER)
        Else
            For i = 1 To LastUser
                If LenB(UserList(i).name) <> 0 Then
                    If EsFacc(i) Then
                        quienes = quienes & UserList(i).name & " (" & UserList(i).Pos.map & " " & UserList(i).Pos.x & " " & UserList(i).Pos.Y & "), "
                    End If
                    count = count + 1
                End If
            Next i
            
            NumUsers = count
            
            Call WriteConsoleMsg(1, UserIndex, "Servidor> Online: " & CStr(NumUsers) & "/" & CStr(MaxUsers) & ". Gms online: " & Left$(quienes, Len(quienes) - 2), FontTypeNames.FONTTYPE_SERVER)
            
            Dim time As Long
            Dim UpTimeStr As String
            
            'Get total time in seconds
            time = ((GetTickCount() And &H7FFFFFFF) - tInicioServer) \ 1000
            
            'Get times in dd:hh:mm:ss format
            UpTimeStr = (time Mod 60) & " segundos."
            time = time \ 60
            
            UpTimeStr = (time Mod 60) & " minutos, " & UpTimeStr
            time = time \ 60
            
            UpTimeStr = (time Mod 24) & " horas, " & UpTimeStr
            time = time \ 24
            
            If time = 1 Then
                UpTimeStr = time & " día, " & UpTimeStr
            Else
                UpTimeStr = time & " días, " & UpTimeStr
            End If
            
            Call WriteConsoleMsg(1, UserIndex, "Servidor> Corriendo hace " & UpTimeStr, FontTypeNames.FONTTYPE_SERVER)
        End If

    End With
End Sub
'Add Marius
Private Sub HandlePena(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .Pos.map = Prision.map And .Counters.Pena <> 0 Then
            Call WriteConsoleMsg(1, UserIndex, "Servidor> Te restan " & .Counters.Pena & " minutos", FontTypeNames.FONTTYPE_SERVER)
        Else
            Call WriteConsoleMsg(1, UserIndex, "Servidor> Tu no estas en prisión!", FontTypeNames.FONTTYPE_SERVER)
        End If
    End With
End Sub
Private Sub HandleRequestStats(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call SendUserStatsTxt(UserIndex, UserIndex)
    End With
End Sub
'\Add

''
' Handles the "Quit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ)
'If user is invisible, it automatically becomes
'visible before doing the countdown to exit
'04/15/2008 - No se reseteaban lso contadores de invi ni de ocultar. (NicoNZ)
'***************************************************
    Dim tUser As Integer
    Dim isNotVisible As Boolean
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Paralizado = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "No puedes salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        
        'exit secure commerce
        If .ComUsu.DestUsu > 0 Then
            tUser = .ComUsu.DestUsu
            
            If UserList(tUser).flags.UserLogged Then
                If UserList(tUser).ComUsu.DestUsu = UserIndex Then
                    Call WriteConsoleMsg(1, tUser, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(tUser)
                End If
            End If
            
            Call WriteConsoleMsg(1, UserIndex, "Comercio cancelado. ", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(UserIndex)
        End If
        
        Call Cerrar_Usuario(UserIndex)
    End With
End Sub

''
' Handles the "GuildLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildLeave(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim GuildIndex As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'obtengo el guildindex
        GuildIndex = m_EcharMiembroDeClan(UserIndex, .name)
        
        If GuildIndex > 0 Then
            Call WriteConsoleMsg(1, UserIndex, "Dejas el clan.", FontTypeNames.FONTTYPE_GUILD)
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(1, .name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
        Else
            Call WriteConsoleMsg(1, UserIndex, "Tu no puedes salir de ningún clan.", FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub

''
' Handles the "RequestAccountState" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAccountState(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim earnings As Integer
    Dim percentage As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't check their accounts
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(1, UserIndex, "Primero tenes que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
            Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case Npclist(.flags.TargetNPC).npctype
            Case eNPCType.Banquero
                Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Case eNPCType.Timbero
                If Not .flags.Privilegios And PlayerType.User Then
                    earnings = Apuestas.Ganancias - Apuestas.Perdidas
                    
                    If earnings >= 0 And Apuestas.Ganancias <> 0 Then
                        percentage = Int(earnings * 100 / Apuestas.Ganancias)
                    End If
                    
                    If earnings < 0 And Apuestas.Perdidas <> 0 Then
                        percentage = Int(earnings * 100 / Apuestas.Perdidas)
                    End If
                    
                    Call WriteConsoleMsg(1, UserIndex, "Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & earnings & " (" & percentage & "%) Jugadas: " & Apuestas.Jugadas, FontTypeNames.FONTTYPE_INFO)
                End If
        End Select
    End With
End Sub

''
' Handles the "PetStand" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetStand(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't use pets
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(1, UserIndex, "Primero tenás que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's his pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it!
        Npclist(.flags.TargetNPC).movement = TipoAI.ESTATICO
        
        Call Expresar(.flags.TargetNPC, UserIndex)
    End With
End Sub

''
' Handles the "PetFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetFollow(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(1, UserIndex, "Primero tenás que seleccionar un personaje, hace click izquierdo sobre ál.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it
        Call FollowAmo(.flags.TargetNPC)
        
        Call Expresar(.flags.TargetNPC, UserIndex)
    End With
End Sub

''
' Handles the "TrainList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrainList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(1, UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's the trainer
        If Npclist(.flags.TargetNPC).npctype <> eNPCType.Entrenador Then Exit Sub
        
        Call WriteTrainerCreatureList(UserIndex, .flags.TargetNPC)
    End With
End Sub





''
' Handles the "Rest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRest(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        
        If HayOBJarea(.Pos, FOGATA) Then
            Call WriteRestOK(UserIndex)
            
            If Not .flags.Descansar Then
                Call WriteConsoleMsg(1, UserIndex, "Te acomodás junto a la fogata y comenzás a descansar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .flags.Descansar = Not .flags.Descansar
        Else
            If .flags.Descansar Then
                Call WriteRestOK(UserIndex)
                Call WriteConsoleMsg(1, UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
                
                .flags.Descansar = False
                Exit Sub
            End If
            
            Call WriteConsoleMsg(1, UserIndex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleCasament(ByVal UserIndex As Integer)
'Fix Marius le pusimos el buffer auxiliar
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler

With UserList(UserIndex)
    
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(.incomingData)
      
    Call buffer.ReadByte    'acá borraremos de la memoria el byte identificador.
    
    
    Dim nick As String
    Dim index As Integer
    
    nick = buffer.ReadASCIIString
    index = NameIndex(nick)
    
    'Dead people can't leave a faction.. they can't talk...
    If .flags.Muerto = 1 Then
        Call WriteMsg(UserIndex, 0)
            
    'Validate target NPC
    ElseIf .flags.TargetNPC = 0 Then
        Call WriteConsoleMsg(1, UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
    
    ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
        Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        
    ElseIf Npclist(.flags.TargetNPC).npctype <> eNPCType.Revividor Then
    
    ElseIf .Genero = UserList(index).Genero Then
        Call WriteConsoleMsg(1, UserIndex, "Personajes del mismo genero no pueden casarce.", FontTypeNames.FONTTYPE_BROWNI)
    
    ElseIf .flags.toyCasado = 1 Then
        Call WriteConsoleMsg(1, UserIndex, "Te encuentras casado !", FontTypeNames.FONTTYPE_BROWNI)
    
    ElseIf UserList(index).flags.Muerto = 1 Then
        Call WriteConsoleMsg(1, UserIndex, "Esta muerto!!!", FontTypeNames.FONTTYPE_BROWNI)
    ElseIf UserList(index).flags.yaOfreci = 1 Then
        Call WriteConsoleMsg(1, UserIndex, "Ya le ofrecieron", FontTypeNames.FONTTYPE_BROWNI)
    
    ElseIf UserList(index).flags.toyCasado = 1 Then
        Call WriteConsoleMsg(1, UserIndex, "Se encuentra casado !", FontTypeNames.FONTTYPE_BROWNI)
        
    Else
    
        Call WriteConsoleMsg(1, index, .name & " quiere casarse contigo, si aceptas escribe /ACEPTO " & .name, FontTypeNames.FONTTYPE_BROWNI)
        
        Call WriteConsoleMsg(1, UserIndex, "Le ofreciste casamiento a " & UserList(index).name, FontTypeNames.FONTTYPE_BROWNI)
        
        UserList(index).flags.yaOfreci = 1
        .flags.yaOfreci = 1
    
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call .incomingData.CopyBuffer(buffer)
End With


Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

Private Sub handleacepto(ByVal UserIndex As Integer)
'Fix Marius le pusimos el buffer auxiliar
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler

With UserList(UserIndex)
 
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(.incomingData)
    
    Call buffer.ReadByte    'acá borraremos de la memoria el byte identificador.

     
    Dim nick As String
    Dim index As Integer
     
    nick = buffer.ReadASCIIString
    index = NameIndex(nick)
     
    If index <= 0 Then
         
    'Dead people can't leave a faction.. they can't talk...
    ElseIf .flags.Muerto = 1 Then
        Call WriteMsg(UserIndex, 0)
    
    'Validate target NPC
    ElseIf .flags.TargetNPC = 0 Then
        Call WriteConsoleMsg(1, UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
        
    ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
        Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
    
    ElseIf UserList(index).flags.yaOfreci = 0 Then
        Call WriteConsoleMsg(1, UserIndex, "No te ofrecio matrimonio.", FontTypeNames.FONTTYPE_BROWNI)
    
    Else
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, .name & " y " & UserList(index).name & " se han unido en matrimonio!", FontTypeNames.FONTTYPE_TALK))
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(161, NO_3D_SOUND, NO_3D_SOUND)) 'Casamiento
        
        .flags.miPareja = UserList(index).name
        UserList(index).flags.miPareja = .name
        .flags.toyCasado = 1
        UserList(index).flags.toyCasado = 1
        
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call .incomingData.CopyBuffer(buffer)
End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
 
End Sub
 
Private Sub handledivorcio(ByVal UserIndex As Integer)
'Fix Marius le pusimos el buffer auxiliar
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
 
With UserList(UserIndex)
 
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(.incomingData)
    
    Call buffer.ReadByte    'acá borraremos de la memoria el byte identificador.
     
    Dim nick As String
    Dim index As Integer
     
    nick = buffer.ReadASCIIString
    index = NameIndex(nick)
     
    If index <= 0 Then
        Call WriteConsoleMsg(1, UserIndex, nick & " No esta online en este momento, intenta mas tarde.", FontTypeNames.FONTTYPE_TALK)
        
    'Dead people can't leave a faction.. they can't talk...
    ElseIf .flags.Muerto = 1 Then
        Call WriteMsg(UserIndex, 0)
        
    ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
        Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
    
    ElseIf Npclist(.flags.TargetNPC).npctype <> eNPCType.Revividor Then
        Call WriteConsoleMsg(1, UserIndex, "Primero clickea al Sacerdote", FontTypeNames.FONTTYPE_INFO)
     
    ElseIf .flags.toyCasado = 0 Then
        Call WriteConsoleMsg(1, UserIndex, "No estás casado con nadie", FontTypeNames.FONTTYPE_TALK)
     
    ElseIf UCase$(.flags.miPareja) <> UCase$(nick) Then
        Call WriteConsoleMsg(1, UserIndex, nick & " No es tu pareja. ¿Estas ebrio?", FontTypeNames.FONTTYPE_TALK)
    Else
        
        .flags.miPareja = ""
        UserList(index).flags.miPareja = ""
        .flags.toyCasado = 0
        UserList(index).flags.toyCasado = 0
        UserList(index).flags.yaOfreci = 0
        .flags.yaOfreci = 0
        
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call .incomingData.CopyBuffer(buffer)
End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error

End Sub
''
' Handles the "Meditate" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleMeditate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/08 (NicoNZ)
'Arreglé un bug que mandaba un index de la meditacion diferente
'al que decia el server.
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If UserList(UserIndex).flags.Metamorfosis = 1 Then
            Call WriteConsoleMsg(2, UserIndex, "No puede meditar estando transformado.", FontTypeNames.FONTTYPE_BROWNI)
            Exit Sub
        End If
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 7)
            Exit Sub
        End If
        
        If .flags.Navegando = 1 Then
            Call WriteConsoleMsg(2, UserIndex, "Estas navegando!!!", FontTypeNames.FONTTYPE_BROWNI)
            Exit Sub
        End If
        
        If .flags.Montando = 1 Then
            Call WriteConsoleMsg(2, UserIndex, "Estas montando!!!", FontTypeNames.FONTTYPE_BROWNI)
            Exit Sub
        End If
        
        Call WriteMeditateToggle(UserIndex)
        
        If .flags.Meditando Then _
           Call WriteConsoleMsg(2, UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_BROWNI)
        
        .flags.Meditando = Not .flags.Meditando

        'Barrin 3/10/03 Tiempo de inicio al meditar
        If .flags.Meditando Then
            Call WriteConsoleMsg(21, UserIndex, "Comienzas a meditar.", FontTypeNames.FONTTYPE_BROWNI)

            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateCharParticle(UserList(UserIndex).Char.CharIndex, ParticleToLevel(UserIndex)))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageDestCharParticle(UserList(UserIndex).Char.CharIndex, ParticleToLevel(UserIndex)))
        End If
    End With
End Sub

''
' Handles the "Resucitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResucitate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(2, UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate NPC and make sure player is dead
        If (Npclist(.flags.TargetNPC).npctype <> eNPCType.Revividor _
            And (Npclist(.flags.TargetNPC).npctype <> eNPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) _
            Or .flags.Muerto = 0 Then Exit Sub
        
        'Make sure it's close enough
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(2, UserIndex, "El sacerdote no puede resucitarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call RevivirUsuario(UserIndex)
        Call WriteConsoleMsg(2, UserIndex, "¡¡Hás sido resucitado!!", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Heal" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHeal(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(1, UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If (Npclist(.flags.TargetNPC).npctype <> eNPCType.Revividor _
            And Npclist(.flags.TargetNPC).npctype <> eNPCType.ResucitadorNewbie) _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(1, UserIndex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .Stats.minhp = .Stats.maxhp
        
        Call WriteUpdateHP(UserIndex)
        
        Call WriteConsoleMsg(1, UserIndex, "¡¡Hás sido curado!!", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub


''
' Handles the "Help" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHelp(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call SendHelp(UserIndex)
End Sub

''
' Handles the "CommerceStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            Exit Sub
        End If
        
        'Add Marius
        If UserList(.flags.TargetUser).flags.solDuelo <> 0 Then
            Call WriteConsoleMsg(1, UserIndex, "El usuario esta decidiendo un duelo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        '\Add
        
        'Is it already in commerce mode??
        If .flags.Comerciando Then
            Call WriteConsoleMsg(1, UserIndex, "Ya estás comerciando", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            'Does the NPC want to trade??
            If Npclist(.flags.TargetNPC).comercia = 0 Then
                If LenB(Npclist(.flags.TargetNPC).desc) <> 0 Then
                    Call WriteChatOverHead(UserIndex, "No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                End If
                
                Exit Sub
            End If
            
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Start commerce....
            Call IniciarComercioNPC(UserIndex)
        '[Alejo]
        ElseIf .flags.TargetUser > 0 Then
            'User commerce...
            
            'Is the other one dead??
            If UserList(.flags.TargetUser).flags.Muerto = 1 Then
                Call WriteConsoleMsg(1, UserIndex, "¡¡No puedes comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is it me??
            If .flags.TargetUser = UserIndex Then
                Call WriteConsoleMsg(1, UserIndex, "No puedes comerciar con vos mismo...", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Check distance
            If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is he already trading?? is it with me or someone else??
            If UserList(.flags.TargetUser).flags.Comerciando = True And _
                UserList(.flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                Call WriteConsoleMsg(1, UserIndex, "No puedes comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Initialize some variables...
            .ComUsu.DestUsu = .flags.TargetUser
            .ComUsu.DestNick = UserList(.flags.TargetUser).name
            .ComUsu.Cant = 0
            .ComUsu.Objeto = 0
            .ComUsu.Acepto = False
            
            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(UserIndex, .flags.TargetUser)
        Else
            Call WriteConsoleMsg(1, UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "BankStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankStart(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            Exit Sub
        End If
        
        If .flags.Comerciando Then
            Call WriteConsoleMsg(1, UserIndex, "Ya estás comerciando", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'If it's the banker....
            If Npclist(.flags.TargetNPC).npctype = eNPCType.Banquero Then
                Call IniciarDeposito(UserIndex, False)
            End If
        Else
            Call WriteConsoleMsg(1, UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Enlist" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEnlist(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(1, UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).npctype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(1, UserIndex, "Debes acercarte más.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).flags.faccion = 1 Then
            Call EnlistarArmadaReal(UserIndex)
        ElseIf Npclist(.flags.TargetNPC).flags.faccion = 3 Then
            Call EnlistarCaos(UserIndex)
        ElseIf Npclist(.flags.TargetNPC).flags.faccion = 2 Then
            Call EnlistarMilicia(UserIndex)
        End If
    End With
End Sub

''
' Handles the "Information" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInformation(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(1, UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).npctype <> eNPCType.Noble _
                Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).flags.faccion = 0 Then
             If .faccion.ArmadaReal = 0 Then
                 Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, cada 100 criminales que derrotes te daré una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
             If .faccion.FuerzasCaos = 0 Then
                 Call WriteChatOverHead(UserIndex, "No perteneces a la legión oscura!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te daré una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
    End With
End Sub

''
' Handles the "Reward" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReward(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(1, UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).npctype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).flags.faccion = 1 Then
             If .faccion.ArmadaReal = 0 Then
                 Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call RecompensaArmadaReal(UserIndex)
        ElseIf Npclist(.flags.TargetNPC).flags.faccion = 3 Then
             If .faccion.FuerzasCaos = 0 Then
                 Call WriteChatOverHead(UserIndex, "No perteneces a la legión oscura!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call RecompensaCaos(UserIndex)
        ElseIf Npclist(.flags.TargetNPC).flags.faccion = 2 Then
            If .faccion.Milicia = 0 Then
                 Call WriteChatOverHead(UserIndex, "No perteneces a la tropas milicianas!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call RecompensaMilicia(UserIndex)
        End If
    End With
End Sub

''
' Handles the "UpTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUpTime(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/10/08
'01/10/2008 - Marcos Martinez (ByVal) - Automatic restart removed from the server along with all their assignments and varibles
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Exit Sub
    
    Dim time As Long
    Dim UpTimeStr As String
    
    'Get total time in seconds
    time = ((GetTickCount() And &H7FFFFFFF) - tInicioServer) \ 1000
    
    'Get times in dd:hh:mm:ss format
    UpTimeStr = (time Mod 60) & " segundos."
    time = time \ 60
    
    UpTimeStr = (time Mod 60) & " minutos, " & UpTimeStr
    time = time \ 60
    
    UpTimeStr = (time Mod 24) & " horas, " & UpTimeStr
    time = time \ 24
    
    If time = 1 Then
        UpTimeStr = time & " día, " & UpTimeStr
    Else
        UpTimeStr = time & " días, " & UpTimeStr
    End If
    
    Call WriteConsoleMsg(1, UserIndex, "Server> Corriendo hace " & UpTimeStr, FontTypeNames.FONTTYPE_SERVER)
End Sub

''
' Handles the "GrupoLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGrupoLeave(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call mdGrupo.SalirDeGrupo(UserIndex)
End Sub



''
' Handles the "GuildMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 02/03/09
'02/03/09: ZaMa - Arreglado un indice mal pasado a la funcion de cartel de clanes overhead.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat As String
        
        chat = buffer.ReadASCIIString()
        chat = Replace(chat, "~", "")
        
        If UserList(UserIndex).Counters.Silenciado <> 0 Then
            If UserList(UserIndex).flags.UltimoMensaje <> 60 Then
                Call WriteConsoleMsg(1, UserIndex, "Los administrador te han silenciado por mensajes reiterados. Espere a ser desilenciado. Gracias.", FontTypeNames.FONTTYPE_BROWNI)
                UserList(UserIndex).flags.UltimoMensaje = 60
                Exit Sub
            End If
        End If
        .Counters.Habla = .Counters.Habla + 1
        
        If LenB(chat) <> 0 Then
            
            'Add Marius
            If EsFacc(UserIndex) Then
                Call SendData(SendTarget.ToGM, 0, PrepareMessageGuildChat(.name & "> " & chat))
                'Call SendData(SendTarget.ToGM, UserIndex, PrepareMessageChatOverHead("< " & chat & " >", .Char.CharIndex, vbYellow))
            '\Add
            ElseIf .GuildIndex > 0 Then
                Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageGuildChat(.name & "> " & chat))
                Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageChatOverHead("< " & chat & " >", .Char.CharIndex, vbYellow))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "GrupoMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGrupoMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat As String
        
        chat = buffer.ReadASCIIString()
        chat = Replace(chat, "~", "")
        
        If LenB(chat) <> 0 Then
            Call mdGrupo.BroadCastGrupo(UserIndex, chat)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "CentinelReport" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCentinelReport(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
End With
End Sub
Public Sub WriteSubastRequest(ByVal UserIndex As Integer)
    Dim i As Byte, Cant As Byte
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Subasta)
        For i = 1 To 100
            If lstSubastas(i).active = True Then Cant = Cant + 1
        Next i
        Call .WriteByte(Cant)
        For i = 1 To 100
            If lstSubastas(i).active = True Then
                Call .WriteByte(i)
                Call .WriteByte(lstSubastas(i).mnDura)
                Call .WriteByte(lstSubastas(i).hsDura)
                Call .WriteLong(lstSubastas(i).actOfert)
                Call .WriteLong(lstSubastas(i).fnlOfert)
                Call .WriteInteger(lstSubastas(i).Cant)
                Call .WriteASCIIString(ObjData(lstSubastas(i).ObjIndex).name)
                Call .WriteASCIIString(lstSubastas(i).nckCmprdor)
                Call .WriteASCIIString(lstSubastas(i).nckVndedor)
                Call .WriteLong(ObjData(lstSubastas(i).ObjIndex).GrhIndex)
            End If
        Next i
    End With
End Sub
Private Sub HandleSubasta(ByVal UserIndex As Integer)
    With UserList(UserIndex)

        Call .incomingData.ReadByte
        Dim subs As Byte, Cant As Integer, sI As Byte
        subs = .incomingData.ReadByte()
        
        'Dead people can't commerce
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            Exit Sub
        End If
        
        If subs = 0 Then
            Call WriteSubastRequest(UserIndex)
        ElseIf subs = 1 Then
            Cant = .incomingData.ReadLong
            Call sOfrecer(UserIndex, Cant, .incomingData.ReadByte)
        ElseIf subs = 2 Then
            Dim hsDura As Byte, Slot As Byte, ObjIndex As Integer, fnlOfert As Long, prcOfert As Long
            Slot = .incomingData.ReadByte
            Cant = .incomingData.ReadInteger
            hsDura = .incomingData.ReadByte
            prcOfert = .incomingData.ReadLong
            fnlOfert = .incomingData.ReadLong
            If Not Slot < 0 And Not Slot > MAX_INVENTORY_SLOTS Then
                If .Invent.Object(Slot).ObjIndex > 0 Then
                    If Cant > .Invent.Object(Slot).Amount Then Cant = .Invent.Object(Slot).Amount
                    If .Invent.Object(Slot).Equipped Then Call Desequipar(UserIndex, Slot)
                    Call sSubastar(UserIndex, .Invent.Object(Slot).ObjIndex, _
                                    Cant, Abs(fnlOfert), hsDura, prcOfert)
                    Call QuitarUserInvItem(UserIndex, Slot, CInt(Cant))
                    Call UpdateUserInv(False, UserIndex, Slot)
                End If
            End If
        ElseIf subs = 3 Then
            sI = .incomingData.ReadByte
            If modSubastas.lstSubastas(sI).active = True And modSubastas.lstSubastas(sI).fnlOfert <> 0 Then
                sComprar UserIndex, sI
            End If
        End If
    End With
End Sub

''
' Handles the "GuildOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnline(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim onlineList As String
        
        onlineList = modGuilds.m_ListaDeMiembrosOnline(UserIndex, .GuildIndex)
        
        If .GuildIndex <> 0 Then
            Call WriteConsoleMsg(1, UserIndex, "Compañeros de tu clan conectados: " & onlineList, FontTypeNames.FONTTYPE_TALKITALIC)
        Else
            Call WriteConsoleMsg(1, UserIndex, "No pertences a ningún clan.", FontTypeNames.FONTTYPE_TALKITALIC)
        End If
    End With
End Sub


''
' Handles the "RoleMasterRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoleMasterRequest(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim request As String
        
        request = buffer.ReadASCIIString()
        
        If LenB(request) <> 0 Then
            Call WriteConsoleMsg(1, UserIndex, "Su solicitud ha sido enviada", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(1, .name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_TALKITALIC))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "GMRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMRequest(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If Not Ayuda.Existe(.name) Then
            Call WriteConsoleMsg(1, UserIndex, "El mensaje ha sido entregado, ahora sólo debes esperar que se desocupe algún GM.", FontTypeNames.FONTTYPE_INFO)
            Call Ayuda.Push(.name)
        Else
            Call Ayuda.Quitar(.name)
            Call Ayuda.Push(.name)
            Call WriteConsoleMsg(1, UserIndex, "Ya habías mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "BugReport" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBugReport(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Dim N As Integer
        
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim bugReport As String
        
        bugReport = buffer.ReadASCIIString()
        
        Call SendData(SendTarget.ToADMINS, 0, PrepareMessageConsoleMsg(1, "Bug > " & .name & ": " & bugReport, FontTypeNames.FONTTYPE_SERVER))
        
        N = FreeFile
        Open App.Path & "\LOGS\BUGs.log" For Append Shared As N
        Print #N, "Usuario:" & .name & "  Fecha:" & Date & "    Hora:" & time
        Print #N, "BUG:"
        Print #N, bugReport
        Print #N, "########################################################################"
        Close #N
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim description As String
        
        description = buffer.ReadASCIIString()
        
        If Not AsciiValidos(description) Then
            Call WriteConsoleMsg(1, UserIndex, "La descripción tiene caractéres inválidos.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Len(description) > 100 Then
            Call WriteConsoleMsg(1, UserIndex, "La descripción es muy larga.", FontTypeNames.FONTTYPE_BROWNI)
        Else
            .desc = Trim$(description)
            Call WriteConsoleMsg(1, UserIndex, "La descripción ha cambiado.", FontTypeNames.FONTTYPE_INFO)
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub


''
' Handles the "Gamble" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGamble(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Integer
        
        Amount = .incomingData.ReadInteger()
        
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
        ElseIf .flags.TargetNPC = 0 Then
            'Validate target NPC
            Call WriteConsoleMsg(1, UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Npclist(.flags.TargetNPC).npctype <> eNPCType.Timbero Then
            Call WriteChatOverHead(UserIndex, "No tengo ningún interés en apostar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf Amount < 1 Then
            Call WriteChatOverHead(UserIndex, "El mínimo de apuesta es 1 moneda.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf Amount > 10000 Then
            Call WriteChatOverHead(UserIndex, "El máximo de apuesta es 10000 monedas.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf .Stats.GLD < Amount Then
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            If RandomNumber(1, 100) <= 47 Then
                .Stats.GLD = .Stats.GLD + Amount
                Call WriteChatOverHead(UserIndex, "Felicidades! Has ganado " & CStr(Amount) & " monedas de oro!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
                Apuestas.Perdidas = Apuestas.Perdidas + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            Else
                .Stats.GLD = .Stats.GLD - Amount
                Call WriteChatOverHead(UserIndex, "Lo siento, has perdido " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
                Apuestas.Ganancias = Apuestas.Ganancias + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
            End If
            
            Apuestas.Jugadas = Apuestas.Jugadas + 1
            
            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
            Call WriteUpdateGold(UserIndex)
        End If
    End With
End Sub

Private Sub HandleHogar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
             Call WriteConsoleMsg(1, UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).npctype <> eNPCType.Revividor Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
Select Case .Pos.map

Case 34
    .Hogar = 0
    
Case 194
    .Hogar = 1
    
Case 112
    .Hogar = 2
    
Case 20
    .Hogar = 3
    
End Select

    End With
    
    
    Call WriteConsoleMsg(1, UserIndex, "Has seleccionado tu nuevo Hogar!", FONTTYPE_BROWNI)
    
End Sub
''
' Handles the "BankExtractGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        
        Amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
             Call WriteConsoleMsg(1, UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).npctype <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Amount > 0 And Amount <= .Stats.Banco Then
             .Stats.Banco = .Stats.Banco - Amount
             .Stats.GLD = .Stats.GLD + Amount
             Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
             Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
        
        Call WriteUpdateGold(UserIndex)
    End With
End Sub

Public Sub HandleBankTransferGold(ByVal UserIndex As Integer)
'Fix Marius le pusimos el buffer auxiliar
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler

    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Call buffer.ReadByte    'acá borraremos de la memoria el byte identificador.
        
        Dim Cant As Long
        Dim name As String
        Dim UI As Integer
        Cant = buffer.ReadLong
        name = UCase$(buffer.ReadASCIIString)
        
        'Debug.Print "name: " & Name
        'Debug.Print "cant: " & Cant
        'Debug.Print "bank: " & .Stats.Banco
    
        'Checkeamos que tenga el oro
        If Cant > 0 And Cant <= .Stats.Banco Then
               
            'Debug.Print ">Adentro"
               
            If .flags.TargetNPC > 0 And Not Npclist(.flags.TargetNPC).npctype = eNPCType.Banquero Then
                Call WriteConsoleMsg(1, UserIndex, "Primero tenés que seleccionar al banquero, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Else
        
                UI = NameIndex(name)
                If UI <> UserIndex Then
            
                    If UI > 0 Then 'Esta online
                        'Le damos el oro al otro
                        UserList(UI).Stats.Banco = UserList(UI).Stats.Banco + Cant
                        
                        'Le quitamos a este
                        .Stats.Banco = .Stats.Banco - Cant
                        
                        Call WriteChatOverHead(UserIndex, "Se han transferido " & Cant & " monedas de oro a " & name & ". ¡¡Gracias por utilizar el servicio de finanzas Goliath!! Vuelva pronto", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                    Else 'La puta madre esta OFF, a abrir la DB xD
                        
                        If Add_Bank_Gold(name, Cant) Then
                            'Le quitamos a este
                            .Stats.Banco = .Stats.Banco - Cant
                            
                            Call WriteChatOverHead(UserIndex, "Se han transferido " & Cant & " monedas de oro a " & name & ". ¡¡Gracias por utilizar el servicio de finanzas Goliath!! Vuelva pronto", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                        End If
                    End If
                Else
                    Call WriteChatOverHead(UserIndex, "¿Es un chiste? No puedo transferir el oro a su propia cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                End If
            End If
        Else
            Call WriteChatOverHead(UserIndex, "Usted no tiene esa cantidad de dinero y no hacemos prestamos!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "LeaveFaction" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeaveFaction(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim flag As Byte
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        flag = .incomingData.ReadByte
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            Exit Sub
        End If
        
        If .GuildIndex > 0 Then
            If modGuilds.GuildFounder(.GuildIndex) = .name Then
                Call WriteConsoleMsg(1, UserIndex, "¡¡¡Eres el fundador de un Clan. No puedes cambiar de faccion!!!", FontTypeNames.FONTTYPE_GUILD)
                Exit Sub
            End If
        End If
        
        If flag = 1 Then
            
            If (esCiuda(UserIndex) Or esArmada(UserIndex)) Then
                
                Call ResetFacciones(UserIndex, False)
                
                .faccion.Renegado = 1
                
                If .GuildIndex > 0 Then
                    Call modGuilds.m_EcharMiembroDeClan(-1, .name)
                    Call WriteConsoleMsg(1, UserIndex, "Has sido expulsado del clan por tu nueva facción.", FontTypeNames.FONTTYPE_GUILD)
                End If
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))
            
            ElseIf (esRepu(UserIndex) Or esMili(UserIndex)) Then
                
                Call ResetFacciones(UserIndex, False)
                
                .faccion.Renegado = 1
                
                If .GuildIndex > 0 Then
                    Call modGuilds.m_EcharMiembroDeClan(-1, .name)
                    Call WriteConsoleMsg(1, UserIndex, "Has sido expulsado del clan por tu nueva facción.", FontTypeNames.FONTTYPE_GUILD)
                End If
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))
            End If
            
        Else
            'Validate target NPC
            If .flags.TargetNPC <> 0 Then
                If Npclist(.flags.TargetNPC).npctype = eNPCType.Noble Then
               
                   'Quit the Royal Army?
                   If .faccion.ArmadaReal = 1 Then
                       If Npclist(.flags.TargetNPC).flags.faccion = 1 Then
                           Call ExpulsarFaccionReal(UserIndex)
                           Call WriteChatOverHead(UserIndex, "Serás bienvenido a las fuerzas imperiales si deseas regresar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                       Else
                           Call WriteChatOverHead(UserIndex, "¡¡¡Sal de aquí bufón!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                       End If
                    'Quit the Chaos Legion??
                   ElseIf .faccion.FuerzasCaos = 1 Then
                       If Npclist(.flags.TargetNPC).flags.faccion = 3 Then
                           Call ExpulsarFaccionCaos(UserIndex, False)
                           Call WriteChatOverHead(UserIndex, "Ya volverás arrastrandote.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                       Else
                           Call WriteChatOverHead(UserIndex, "Sal de aquí maldito criminal", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                       End If
                   ElseIf .faccion.Milicia = 1 Then
                       If Npclist(.flags.TargetNPC).flags.faccion = 2 Then
                           Call ExpulsarFaccionMilicia(UserIndex, False)
                           Call WriteChatOverHead(UserIndex, "Que tengas un buen camino!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                       Else
                           Call WriteChatOverHead(UserIndex, "Sal de aquí maldito criminal", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                       End If
                   Else
                       Call WriteChatOverHead(UserIndex, "¡No perteneces a ninguna facción!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                   End If
                End If
            Else
                If .faccion.Ciudadano = 1 Or .faccion.Republicano = 1 Then
                    WriteShowMessageBox UserIndex, "", 1, 1
                End If
            End If
        End If
        
        'Add Marius
        If UserList(UserIndex).Pos.map = Bandera_mapa Then
            Call Bandera_Sale(UserIndex)
        End If
        '\Add
    End With
End Sub

''
' Handles the "BankDepositGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDepositGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        
        Amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(1, UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).npctype <> eNPCType.Banquero Then Exit Sub
        
        If Amount > 0 And Amount <= .Stats.GLD Then
            .Stats.Banco = .Stats.Banco + Amount
            .Stats.GLD = .Stats.GLD - Amount
            Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Call WriteUpdateGold(UserIndex)
        Else
            Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
    End With
End Sub

''
' Handles the "Denounce" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Text As String
        
        Text = buffer.ReadASCIIString()
        
        If .flags.Silenciado = 0 Then
            Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(1, .name & " DENUNCIA: " & Text, FontTypeNames.FONTTYPE_TALKITALIC))
            Call WriteConsoleMsg(1, UserIndex, "Denuncia enviada, espere...", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

Private Sub HandleFriends(ByVal UserIndex As Integer)
'***************************************************
'Author: Marius
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Text As String
        Dim tipo As Byte
        
        tipo = buffer.ReadByte()
        Text = buffer.ReadASCIIString()
        
        'Call WriteConsoleMsg(1, UserIndex, "Tipo: " & tipo & " Texto: " & Text, FontTypeNames.FONTTYPE_INFO)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub
''
' Handles the "GuildFundate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildFundate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 1 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    
        Dim Error As String
        
        If PuedenFundarClan <> 1 Then
            Call WriteConsoleMsg(1, UserIndex, "La creación de clanes esta deshabilitada momentaneamente. (No moleste a los GMs ellos no pueden fundar su clan)", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If
        
        If esCiuda(UserIndex) Or esArmada(UserIndex) Then
            .FundandoGuildAlineacion = ALINEACION_IMPERIAL
        ElseIf esRepu(UserIndex) Or esMili(UserIndex) Then
            .FundandoGuildAlineacion = ALINEACION_REPUBLICANO
        ElseIf esCaos(UserIndex) Then
            .FundandoGuildAlineacion = ALINEACION_CAOTICO
        ElseIf esRene(UserIndex) Then
            .FundandoGuildAlineacion = ALINEACION_RENEGADO
        Else
            Call WriteConsoleMsg(1, UserIndex, "Hay un error en su faccion, comuniquese con algun GameMaster", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If
        
        If modGuilds.PuedeFundarUnClan(UserIndex, .FundandoGuildAlineacion, Error) Then
            Call WriteShowGuildFundationForm(UserIndex)
        Else
            .FundandoGuildAlineacion = 0
            Call WriteConsoleMsg(1, UserIndex, Error, FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub

''
' Handles the "GrupoKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGrupoKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/05/09
'Last Modification by: Marco Vanotti (Marco)
'- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use Grupo commands
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If UserPuedeEjecutarComandos(UserIndex) Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call mdGrupo.ExpulsarDeGrupo(UserIndex, tUser)
            Else
                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")
                End If
                
                Call WriteConsoleMsg(1, UserIndex, LCase(UserName) & " no pertenece a tu Grupo.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub


''

''
' Handles the "GuildMemberList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim memberCount As Integer
        Dim i As Long
        Dim UserName As String
        
        guild = buffer.ReadASCIIString()
        
        If EsDIOS(UserIndex) Then
            If (InStrB(guild, "\") <> 0) Then
                guild = Replace(guild, "\", "")
            End If
            If (InStrB(guild, "/") <> 0) Then
                guild = Replace(guild, "/", "")
            End If
            
            If Not FileExist(App.Path & "\guilds\" & guild & "-members.mem") Then
                Call WriteConsoleMsg(1, UserIndex, "No existe el clan: " & guild, FontTypeNames.FONTTYPE_INFO)
            Else
                memberCount = val(GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "INIT", "NroMembers"))
                
                For i = 1 To memberCount
                    UserName = GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "Members", "Member" & i)
                    
                    Call WriteConsoleMsg(1, UserIndex, UserName & "<" & guild & ">", FontTypeNames.FONTTYPE_INFO)
                Next i
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "GMMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        
        message = buffer.ReadASCIIString()
        
        If EsCONSE(UserIndex) Then
            If LenB(message) <> 0 Then
                Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(1, .name & "> " & message, FontTypeNames.FONTTYPE_BROWNB))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "ShowName" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowName(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsDIOS(UserIndex) Then
            .showName = Not .showName 'Show / Hide the name
            
            Call RefreshCharStatus(UserIndex)
        End If
    End With
End Sub

''
' Handles the "OnlineArmada" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineArmada(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        .incomingData.ReadByte
        
        If EsFacc(UserIndex) Then
    
            Dim i As Long
            Dim list As String
    
            For i = 1 To LastUser
                If UserList(i).ConnID <> -1 Then
                    If UserList(i).faccion.ArmadaReal = 1 Then
                        If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.VIP) Then
                            list = list & UserList(i).name & ", "
                        End If
                    End If
                End If
            Next i
            
            If Len(list) > 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Armadas conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, UserIndex, "No hay Armadas conectados", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
    
End Sub

Private Sub HandleOnlineMilicia(ByVal UserIndex As Integer)
'***************************************************
'Author: Marius
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        .incomingData.ReadByte
        
        If EsFacc(UserIndex) Then
    
            Dim i As Long
            Dim list As String
    
            For i = 1 To LastUser
                If UserList(i).ConnID <> -1 Then
                    If UserList(i).faccion.Milicia = 1 Then
                        If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.VIP) Then
                            list = list & UserList(i).name & ", "
                        End If
                    End If
                End If
            Next i
            
            If Len(list) > 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Milicianos conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, UserIndex, "No hay Milicianos conectados", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
    
End Sub

''
' Handles the "OnlineCaos" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineCaos(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        .incomingData.ReadByte
        
        If EsFacc(UserIndex) Then
        
            Dim i As Long
            Dim list As String
    
            For i = 1 To LastUser
                If UserList(i).ConnID <> -1 Then
                    If UserList(i).faccion.FuerzasCaos = 1 Then
                        If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.VIP) Then
                            list = list & UserList(i).name & ", "
                        End If
                    End If
                End If
            Next i
            
            If Len(list) > 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Caos conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, UserIndex, "No hay Caos conectados", FontTypeNames.FONTTYPE_INFO)
            End If
    
        End If
    End With

End Sub

''
' Handles the "GoNearby" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoNearby(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/10/07
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        
        UserName = buffer.ReadASCIIString()
        
        Dim tIndex As Integer
        Dim x As Long
        Dim Y As Long
        Dim i As Long
        Dim found As Boolean
        
        tIndex = NameIndex(UserName)
        
        'Check the user has enough powers
        If EsSEMI(UserIndex) Then
            If tIndex <= 0 Then 'existe el usuario destino?
                Call WriteConsoleMsg(1, UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                For i = 2 To 5 'esto for sirve ir cambiando la distancia destino
                    For x = UserList(tIndex).Pos.x - i To UserList(tIndex).Pos.x + i
                        For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i
                            If MapData(UserList(tIndex).Pos.map, x, Y).UserIndex = 0 Then
                                If LegalPos(UserList(tIndex).Pos.map, x, Y, True, True) Then
                                    Call WarpUserChar(UserIndex, UserList(tIndex).Pos.map, x, Y, True)
                                    found = True
                                    Exit For
                                End If
                            End If
                        Next Y
                        
                        If found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                    Next x
                    
                    If found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                Next i
                
                'No space found??
                If Not found Then
                    Call WriteConsoleMsg(1, UserIndex, "Todos los lugares están ocupados.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "Comment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleComment(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim comment As String
        comment = buffer.ReadASCIIString()
        
        If EsCONSE(UserIndex) Then
            Call LogGM(.name, "Comentario: " & comment)
            Call WriteConsoleMsg(1, UserIndex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "ServerTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    
        If EsCONSE(UserIndex) Then
            Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Hora: " & time & " " & Date, FontTypeNames.FONTTYPE_INFO))
            Call LogGM(.name, "Hora.")
        End If
    
    End With
End Sub

''
' Handles the "Where" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhere(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If EsCONSE(UserIndex) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, UserIndex, "Ubicación  " & UserName & ": " & UserList(tUser).Pos.map & ", " & UserList(tUser).Pos.x & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "CreaturesInMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreaturesInMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 30/07/06
'Pablo (ToxicWaste): modificaciones generales para simplificar la visualización.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim map As Integer
        Dim i, j As Long
        Dim NPCcount1, NPCcount2 As Integer
        Dim NPCcant1() As Integer
        Dim NPCcant2() As Integer
        Dim List1() As String
        Dim List2() As String
        
        map = .incomingData.ReadInteger()
        
        If EsSEMI(UserIndex) Then
        
            If MapaValido(map) Then
                For i = 1 To LastNPC
                    'VB isn't lazzy, so we put more restrictive condition first to speed up the process
                    If Npclist(i).Pos.map = map Then
                        '¿esta vivo?
                        If Npclist(i).flags.NPCActive And Npclist(i).hostile = 1 And Npclist(i).Stats.alineacion = 2 Then
                            If NPCcount1 = 0 Then
                                ReDim List1(0) As String
                                ReDim NPCcant1(0) As Integer
                                NPCcount1 = 1
                                List1(0) = Npclist(i).name & ": (" & Npclist(i).Pos.x & "," & Npclist(i).Pos.Y & ")"
                                NPCcant1(0) = 1
                            Else
                                For j = 0 To NPCcount1 - 1
                                    If Left$(List1(j), Len(Npclist(i).name)) = Npclist(i).name Then
                                        List1(j) = List1(j) & ", (" & Npclist(i).Pos.x & "," & Npclist(i).Pos.Y & ")"
                                        NPCcant1(j) = NPCcant1(j) + 1
                                        Exit For
                                    End If
                                Next j
                                If j = NPCcount1 Then
                                    ReDim Preserve List1(0 To NPCcount1) As String
                                    ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
                                    NPCcount1 = NPCcount1 + 1
                                    List1(j) = Npclist(i).name & ": (" & Npclist(i).Pos.x & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant1(j) = 1
                                End If
                            End If
                        Else
                            If NPCcount2 = 0 Then
                                ReDim List2(0) As String
                                ReDim NPCcant2(0) As Integer
                                NPCcount2 = 1
                                List2(0) = Npclist(i).name & ": (" & Npclist(i).Pos.x & "," & Npclist(i).Pos.Y & ")"
                                NPCcant2(0) = 1
                            Else
                                For j = 0 To NPCcount2 - 1
                                    If Left$(List2(j), Len(Npclist(i).name)) = Npclist(i).name Then
                                        List2(j) = List2(j) & ", (" & Npclist(i).Pos.x & "," & Npclist(i).Pos.Y & ")"
                                        NPCcant2(j) = NPCcant2(j) + 1
                                        Exit For
                                    End If
                                Next j
                                If j = NPCcount2 Then
                                    ReDim Preserve List2(0 To NPCcount2) As String
                                    ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
                                    NPCcount2 = NPCcount2 + 1
                                    List2(j) = Npclist(i).name & ": (" & Npclist(i).Pos.x & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant2(j) = 1
                                End If
                            End If
                        End If
                    End If
                Next i
                
                Call WriteConsoleMsg(1, UserIndex, "Npcs Hostiles en mapa: ", FontTypeNames.FONTTYPE_WARNING)
                If NPCcount1 = 0 Then
                    Call WriteConsoleMsg(1, UserIndex, "No hay NPCS Hostiles", FontTypeNames.FONTTYPE_INFO)
                Else
                    For j = 0 To NPCcount1 - 1
                        Call WriteConsoleMsg(1, UserIndex, NPCcant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
                    Next j
                End If
                Call WriteConsoleMsg(1, UserIndex, "Otros Npcs en mapa: ", FontTypeNames.FONTTYPE_WARNING)
                If NPCcount2 = 0 Then
                    Call WriteConsoleMsg(1, UserIndex, "No hay más NPCS", FontTypeNames.FONTTYPE_INFO)
                Else
                    For j = 0 To NPCcount2 - 1
                        Call WriteConsoleMsg(1, UserIndex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
                    Next j
                End If
                Call LogGM(.name, "Numero enemigos en mapa " & map)
            End If
        End If
    End With
End Sub

''
' Handles the "WarpMeToTarget" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpMeToTarget(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/09
'26/03/06: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim x As Integer
        Dim Y As Integer
        
        If EsFacc(UserIndex) Then
            x = .flags.TargetX
            Y = .flags.TargetY
            
            Call FindLegalPos(UserIndex, .flags.TargetMap, x, Y)
            Call WarpUserChar(UserIndex, .flags.TargetMap, x, Y, True)
            Call LogGM(.name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .Pos.map)
        End If
        
    End With
End Sub

''
' Handles the "WarpChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
'***************************************************
    If UserList(UserIndex).incomingData.length < 7 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim map As Integer
        Dim x As Integer
        Dim Y As Integer
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        map = buffer.ReadInteger()
        x = buffer.ReadByte()
        Y = buffer.ReadByte()
        
        If EsFacc(UserIndex) Then
            If MapaValido(map) And LenB(UserName) <> 0 Then
                
                If EsSEMI(UserIndex) And UCase$(UserName) <> "YO" Then
                    tUser = NameIndex(UserName)
                Else
                    tUser = UserIndex
                End If
            
                If tUser <= 0 Then
                    Call WriteConsoleMsg(1, UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                ElseIf InMapBounds(map, x, Y) Then
                    Call FindLegalPos(tUser, map, x, Y)
                    Call WarpUserChar(tUser, map, x, Y, False)
                    'Call WriteConsoleMsg(1, UserIndex, UserList(tUser).name & " transportado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, "Transportó a " & UserList(tUser).name & " hacia " & "Mapa" & map & " X:" & x & " Y:" & Y)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "Silence" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSilence(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If EsCONSE(UserIndex) Then
            tUser = NameIndex(UserName)
        
            If tUser <= 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                If UserList(tUser).flags.Silenciado = 0 Then
                    UserList(tUser).flags.Silenciado = 120 'Segundos
                    Call WriteConsoleMsg(1, UserIndex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteShowMessageBox(tUser, "ESTIMADO USUARIO, ud ha sido silenciado por los administradores. Sus denuncias serán ignoradas por el servidor de aquí en más. Utilice /GM para contactar un administrador.")
                    
                    'Flush the other user's buffer
                    Call FlushBuffer(tUser)
                Else
                    UserList(tUser).flags.Silenciado = 0
                    Call WriteConsoleMsg(1, UserIndex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "SOSShowList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSShowList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsSEMI(UserIndex) Then
            Call WriteShowSOSForm(UserIndex)
        End If
    End With
End Sub

''
' Handles the "SOSRemove" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSRemove(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        UserName = buffer.ReadASCIIString()
        
        If EsSEMI(UserIndex) Then
            Call Ayuda.Quitar(UserName)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "GoToChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoToChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim x As Integer
        Dim Y As Integer
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If EsFacc(UserIndex) Then
            If tUser <= 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                x = UserList(tUser).Pos.x
                Y = UserList(tUser).Pos.Y + 1
                Call FindLegalPos(UserIndex, UserList(tUser).Pos.map, x, Y)
                
                Call WarpUserChar(UserIndex, UserList(tUser).Pos.map, x, Y, True)
                
                If .flags.AdminInvisible = 0 Then
                    Call WriteConsoleMsg(1, tUser, .name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                    Call FlushBuffer(tUser)
                End If
                
                Call LogGM(.name, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.map & " X:" & UserList(tUser).Pos.x & " Y:" & UserList(tUser).Pos.Y)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "Invisible" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsFacc(UserIndex) Then
            Call DoAdminInvisible(UserIndex)
            Call LogGM(.name, "/INVISIBLE")
        End If
    End With
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMPanel(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsFacc(UserIndex) Then
            Call WriteShowGMPanelForm(UserIndex)
        End If
    End With
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestUserList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'Last modified by: Lucas Tavolaro Ortiz (Tavo)
'I haven`t found a solution to split, so i make an array of names
'***************************************************
    Dim i As Long
    Dim names() As String
    Dim count As Long
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsFacc(UserIndex) Then
            
            ReDim names(1 To LastUser) As String
            count = 1
            
            For i = 1 To LastUser
                If (LenB(UserList(i).name) <> 0 And i <> UserIndex) Then
                    names(count) = UserList(i).name
                    count = count + 1
                End If
            Next i
            
            If count > 1 Then Call WriteUserNameList(UserIndex, names(), count - 1)
            
        End If
    End With
End Sub

''
' Handles the "Working" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorking(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim users As String
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsCONSE(UserIndex) Then
        
            For i = 1 To LastUser
                If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
                    users = users & ", " & UserList(i).name
                    End If
            Next i
            
            If LenB(users) <> 0 Then
                users = Right$(users, Len(users) - 2)
                Call WriteConsoleMsg(1, UserIndex, "Usuarios trabajando: " & users, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, UserIndex, "No hay usuarios trabajando", FontTypeNames.FONTTYPE_INFO)
            End If
        
        End If
    End With
End Sub

''
' Handles the "Hiding" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHiding(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim users As String
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsCONSE(UserIndex) Then
            
            For i = 1 To LastUser
                If (LenB(UserList(i).name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
                    users = users & UserList(i).name & ", "
                End If
            Next i
            
            If LenB(users) <> 0 Then
                users = Left$(users, Len(users) - 2)
                Call WriteConsoleMsg(1, UserIndex, "Usuarios ocultandose: " & users, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, UserIndex, "No hay usuarios ocultandose", FontTypeNames.FONTTYPE_INFO)
            End If
            
        End If
    End With
End Sub

''
' Handles the "Jail" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleJail(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        Dim jailTime As Byte
        Dim count As Byte
        Dim tUser As Integer
        
        Dim Orden As String
        Dim RS As New ADODB.Recordset
        
        UserName = buffer.ReadASCIIString()
        reason = buffer.ReadASCIIString()
        jailTime = buffer.ReadByte()
        
        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")
        End If
        
        '/carcel nick@motivo@<tiempo>
        If EsCONSE(UserIndex) Then
            If LenB(UserName) = 0 Or LenB(reason) = 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(UserName)
                
                If tUser <= 0 Then
                    Call WriteConsoleMsg(1, UserIndex, "El usuario esta Offline", FontTypeNames.FONTTYPE_INFO)
                Else
                    If EsCONSE(tUser) Then
                        Call WriteConsoleMsg(1, UserIndex, "No podés encarcelar a miembros del staff.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf jailTime > 60 And Not EsADMIN(UserIndex) Then
                        Call WriteConsoleMsg(1, UserIndex, "No podés encarcelar por más de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        If (InStrB(UserName, "\") <> 0) Then
                            UserName = Replace(UserName, "\", "")
                        End If
                        If (InStrB(UserName, "/") <> 0) Then
                            UserName = Replace(UserName, "/", "")
                        End If
                        
                        Call Encarcelar(tUser, jailTime, .name, reason)
                        Call LogGM(UserList(UserIndex).name, "/CARCEL " & UserName & "@" & reason & "@" & jailTime)
                        Call SendData(SendTarget.ToADMINS, 0, PrepareMessageConsoleMsg(1, .name & " encarceló a " & UserName & ", " & jailTime & " minutos por: " & reason, FontTypeNames.FONTTYPE_INFO))
                        Call SaveUserSQL(tUser) 'Grabamos para que ningun rollback lo saque de ahi xD
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "KillNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPC(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/22/08 (NicoNZ)
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsSEMI(UserIndex) Then
            
            Dim tNPC As Integer
            Dim auxNPC As npc
            
            tNPC = .flags.TargetNPC
            
            If tNPC > 0 Then
                Call WriteConsoleMsg(1, UserIndex, "RMatas (con posible respawn) a: " & Npclist(tNPC).name, FontTypeNames.FONTTYPE_INFO)
                
                auxNPC = Npclist(tNPC)
                Call QuitarNPC(tNPC)
                Call ReSpawnNpc(auxNPC)
                
                .flags.TargetNPC = 0
            Else
                Call WriteConsoleMsg(1, UserIndex, "Debes hacer click sobre el NPC antes", FontTypeNames.FONTTYPE_INFO)
            End If
            
        End If
    End With
End Sub

''
' Handles the "WarnUser" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        Dim privs As PlayerType
        Dim count As Byte
        
        UserName = buffer.ReadASCIIString()
        reason = buffer.ReadASCIIString()
        
        If EsCONSE(UserIndex) Then
            If LenB(UserName) = 0 Or LenB(reason) = 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
            Else
                
                If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")
                End If
                If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")
                End If
                
                If ExistePersonaje(UserName) Then
                    Call WriteConsoleMsg(1, UserIndex, "Has advertido a " & UCase$(UserName), FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, " advirtio a " & UserName)
                End If
              
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub


Private Sub HandleSlashSlash(ByVal UserIndex As Integer)
'***************************************************
'Author: Marius
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    Dim i As Integer
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Comandos As String
               
        Comandos = buffer.ReadASCIIString()
        
        If EsCONSE(UserIndex) Then
            
            Dim TmpArgos() As String
    
            Dim Comando As String
            
            Dim UserName As String
            Dim tUser As Integer
        
            Dim ArgumentosAll() As String
            Dim ArgumentosRaw As String
            Dim Argumentos2() As String
            Dim Argumentos3() As String
            Dim Argumentos4() As String
            Dim CantidadArgumentos As Long
            Dim notNullArguments As Boolean
            
            Dim tmpArr() As String
            Dim tmpInt As Integer
            
            ' TmpArgs: Un array de a lo sumo dos elementos,
            ' el primero es el comando (hasta el primer espacio)
            ' y el segundo elemento es el resto. Si no hay argumentos
            ' devuelve un array de un solo elemento
            TmpArgos = Split(Comandos, " ", 2)
            
            Comando = Trim$(UCase$(TmpArgos(0)))
            
            If UBound(TmpArgos) > 0 Then
                ' El string en crudo que este despues del primer espacio
                ArgumentosRaw = TmpArgos(1)
                
                'veo que los argumentos no sean nulos
                notNullArguments = LenB(Trim$(ArgumentosRaw))
                
                ' Un array separado por blancos, con tantos elementos como
                ' se pueda
                ArgumentosAll = Split(TmpArgos(1), " ")
                
                ' Cantidad de argumentos. En ESTE PUNTO el minimo es 1
                CantidadArgumentos = UBound(ArgumentosAll) + 1
                
                ' Los siguientes arrays tienen A LO SUMO, COMO MAXIMO
                ' 2, 3 y 4 elementos respectivamente. Eso significa
                ' que pueden tener menos, por lo que es imperativo
                ' preguntar por CantidadArgumentos.
                
                Argumentos2 = Split(TmpArgos(1), " ", 2)
                Argumentos3 = Split(TmpArgos(1), " ", 3)
                Argumentos4 = Split(TmpArgos(1), " ", 4)
            Else
                CantidadArgumentos = 0
            End If
            
            
            Select Case Comando
                Case "DESC"
                    If CantidadArgumentos > 0 Then
                        Call WriteConsoleMsg(1, UserIndex, " Parametro nombre: " & ArgumentosRaw, FontTypeNames.FONTTYPE_INFO)
                    End If
                    
                Case "INV"
                    If CantidadArgumentos > 0 Then
                        UserName = Replace(ArgumentosRaw, "+", " ")
                        tUser = NameIndex(UserName)
                        
                        If tUser > 0 Then

                            If EsADMIN(tUser) Then
                                Call WriteConsoleMsg(1, UserIndex, "El chusma de " & UserList(UserIndex).name & " esta revisando tu inventario.", FontTypeNames.FONTTYPE_SERVER)
                            End If

                            Comando = ""
                            For i = 1 To MAX_INVENTORY_SLOTS
                                If UserList(tUser).Invent.Object(i).Amount > 0 Then
                                    Comando = Comando & ObjData(UserList(tUser).Invent.Object(i).ObjIndex).name & " - "
                                    Comando = Comando & UserList(tUser).Invent.Object(i).Amount
                                    If UserList(tUser).Invent.Object(i).Equipped = 1 Then Comando = Comando & " (E)"
                                    Comando = Comando & vbNewLine
                                End If
                            Next i
                            Call WriteConsoleMsg(1, UserIndex, Comando, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(1, UserIndex, "Personaje Offline", FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                
                Case "BOV"
                    If CantidadArgumentos > 0 Then
                        UserName = Replace(ArgumentosRaw, "+", " ")
                        tUser = NameIndex(UserName)
                        
                        If tUser > 0 Then

                            If EsADMIN(tUser) Then
                                Call WriteConsoleMsg(1, UserIndex, "El chusma de " & UserList(UserIndex).name & " esta revisando tu boveda.", FontTypeNames.FONTTYPE_SERVER)
                            End If

                            Comando = ""
                            For i = 1 To MAX_BANCOINVENTORY_SLOTS
                                If UserList(tUser).BancoInvent.Object(i).Amount > 0 Then
                                    Comando = Comando & ObjData(UserList(tUser).BancoInvent.Object(i).ObjIndex).name & " - "
                                    Comando = Comando & UserList(tUser).BancoInvent.Object(i).Amount
                                    Comando = Comando & vbNewLine
                                End If
                            Next i
                            Call WriteConsoleMsg(1, UserIndex, Comando, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(1, UserIndex, "Personaje Offline", FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                    
                Case Else
                    Call WriteConsoleMsg(1, UserIndex, "Ese comando no existe.", FontTypeNames.FONTTYPE_INFO)
                
            End Select
            
            Call LogGM(.name, " // " & ArgumentosRaw)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "EditChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEditChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 02/03/2009
'02/03/2009: ZaMa -  Cuando editas nivel, chequea si el pj peude permanecer en clan faccionario
'***************************************************
    If UserList(UserIndex).incomingData.length < 8 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim opcion As Byte
        Dim arg1 As String
        Dim arg2 As String
        Dim loopC As Byte
        Dim commandString As String
        Dim N As Byte
        
        UserName = Replace(buffer.ReadASCIIString(), "+", " ")
        
        If EsADMIN(UserIndex) And UCase$(UserName) <> "YO" Then
            tUser = NameIndex(UserName)
        Else
            tUser = UserIndex
        End If
        
        opcion = buffer.ReadByte()
        arg1 = buffer.ReadASCIIString()
        arg2 = buffer.ReadASCIIString()
        
        If EsSEMI(UserIndex) Then
        
            Select Case opcion
                Case eEditOptions.eo_Gold
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(1, UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        If val(arg1) > 200000000 Then
                            arg1 = 200000000
                        End If
       
                        UserList(tUser).Stats.GLD = val(arg1)
                        Call WriteUpdateGold(tUser)
                    End If
                
                Case eEditOptions.eo_Experience
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(1, UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        If val(arg1) > 200000000 Then
                            arg1 = 200000000
                        End If
                            
                        UserList(tUser).Stats.Exp = UserList(tUser).Stats.Exp + val(arg1)
                        Call CheckUserLevel(tUser)
                        Call WriteUpdateExp(tUser)
                        
                    End If
                
                Case eEditOptions.eo_Body
                    If tUser > 0 Then
                        Call ChangeUserChar(tUser, val(arg1), UserList(tUser).Char.head, UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                    End If
                
                Case eEditOptions.eo_Head
                    If tUser > 0 Then
                        UserList(tUser).Char.head = val(arg1)
                        Call ChangeUserChar(tUser, UserList(tUser).Char.body, UserList(tUser).Char.head, UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                    End If
                
                Case eEditOptions.eo_CriminalsKilled
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(1, UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        If val(arg1) > MAXUSERMATADOS Then
                            UserList(tUser).faccion.RenegadosMatados = MAXUSERMATADOS
                        Else
                            UserList(tUser).faccion.RenegadosMatados = val(arg1)
                        End If
                    End If
                
                Case eEditOptions.eo_CiticensKilled
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(1, UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        If val(arg1) > MAXUSERMATADOS Then
                            UserList(tUser).faccion.CiudadanosMatados = MAXUSERMATADOS
                        Else
                            UserList(tUser).faccion.CiudadanosMatados = val(arg1)
                        End If
                    End If
                
                Case eEditOptions.eo_Level
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(1, UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        If val(arg1) > STAT_MAXELV Then
                            arg1 = CStr(STAT_MAXELV)
                            Call WriteConsoleMsg(1, UserIndex, "No puedes tener un nivel superior a " & STAT_MAXELV & ".", FONTTYPE_INFO)
                        End If
                        
                        UserList(tUser).Stats.ELV = val(arg1)
    
                    End If
                    
                    Call WriteUpdateUserStats(UserIndex)
                
                Case eEditOptions.eo_Class
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(1, UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        For loopC = 1 To NUMCLASES
                            If UCase$(ListaClases(loopC)) = UCase$(arg1) Then Exit For
                        Next loopC
                        
                        If loopC > NUMCLASES Then
                            Call WriteConsoleMsg(1, UserIndex, "Clase desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            UserList(tUser).Clase = loopC
                        End If
                    End If
                
                Case eEditOptions.eo_Skills
                    For loopC = 1 To NUMSKILLS
                        If UCase$(Replace$(SkillsNames(loopC), " ", "+")) = UCase$(arg1) Then Exit For
                    Next loopC
                    
                    If loopC > NUMSKILLS Then
                        Call WriteConsoleMsg(1, UserIndex, "Skill Inexistente!", FontTypeNames.FONTTYPE_INFO)
                    Else
                        If tUser > 0 Then
                            UserList(tUser).Stats.UserSkills(loopC) = val(arg2)
                        End If
                    End If
                
                Case eEditOptions.eo_SkillPointsLeft
                    If tUser > 0 Then
                        UserList(tUser).Stats.SkillPts = IIf(val(arg1) > 32000, 32000, val(arg1))
                    End If
                
                Case eEditOptions.eo_Sex
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(1, UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        arg1 = UCase$(arg1)
                        If (arg1 = "MUJER") Then
                            UserList(tUser).Genero = eGenero.Mujer
                        ElseIf (arg1 = "HOMBRE") Then
                            UserList(tUser).Genero = eGenero.Hombre
                        End If
                    End If
                
                Case eEditOptions.eo_Raza
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(1, UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        arg1 = UCase$(arg1)
                        If (arg1 = "HUMANO") Then
                            UserList(tUser).Raza = eRaza.Humano
                        ElseIf (arg1 = "ELFO") Then
                            UserList(tUser).Raza = eRaza.Elfo
                        ElseIf (arg1 = "DROW") Then
                            UserList(tUser).Raza = eRaza.Drow
                        ElseIf (arg1 = "ENANO") Then
                            UserList(tUser).Raza = eRaza.Enano
                        ElseIf (arg1 = "GNOMO") Then
                            UserList(tUser).Raza = eRaza.Gnomo
                        ElseIf (arg1 = "ORCO") Then
                            UserList(tUser).Raza = eRaza.Orco
                        End If
                    End If
                Case Else
                    Call WriteConsoleMsg(1, UserIndex, "Comando no permitido.", FontTypeNames.FONTTYPE_INFO)
            End Select
    
            'Log it!
            commandString = "/MOD "
            
            Select Case opcion
                Case eEditOptions.eo_Gold
                    commandString = commandString & "ORO "
                
                Case eEditOptions.eo_Experience
                    commandString = commandString & "EXP "
                
                Case eEditOptions.eo_Body
                    commandString = commandString & "BODY "
                
                Case eEditOptions.eo_Head
                    commandString = commandString & "HEAD "
                
                Case eEditOptions.eo_CriminalsKilled
                    commandString = commandString & "CRI "
                
                Case eEditOptions.eo_CiticensKilled
                    commandString = commandString & "CIU "
                
                Case eEditOptions.eo_Level
                    commandString = commandString & "LEVEL "
                
                Case eEditOptions.eo_Class
                    commandString = commandString & "CLASE "
                
                Case eEditOptions.eo_Skills
                    commandString = commandString & "SKILLS "
                
                Case eEditOptions.eo_SkillPointsLeft
                    commandString = commandString & "SKILLSLIBRES "
                    
                Case eEditOptions.eo_Nobleza
                    commandString = commandString & "NOB "
                    
                Case eEditOptions.eo_Asesino
                    commandString = commandString & "ASE "
                    
                Case eEditOptions.eo_Sex
                    commandString = commandString & "SEX "
                    
                Case eEditOptions.eo_Raza
                    commandString = commandString & "RAZA "
                    
                Case Else
                    commandString = commandString & "UNKOWN "
            End Select
            
            commandString = commandString & arg1 & " " & arg2
            
            Call LogCheating("Por las dudas " & .name & " : " & commandString)
        
        End If
            
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
        
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub


' Handles the "ReviveChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReviveChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim loopC As Byte
        
        UserName = buffer.ReadASCIIString()
        
        
        If (EsSEMI(UserIndex)) Then
            If UCase$(UserName) <> "YO" Then
                tUser = NameIndex(UserName)
            Else
                tUser = UserIndex
            End If
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                With UserList(tUser)
                    'If dead, show him alive (naked).
                    If .flags.Muerto = 1 Then
                        .flags.Muerto = 0
                        
                        Call DarCuerpoDesnudo(tUser)
                        
                        Call ChangeUserChar(tUser, .Char.body, .OrigChar.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        
                        Call WriteConsoleMsg(1, tUser, UserList(UserIndex).name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(1, tUser, UserList(UserIndex).name & " te ha curado.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    
                    .Stats.minhp = .Stats.maxhp
                End With
                
                Call WriteUpdateHP(tUser)
                
                Call FlushBuffer(tUser)
                
                Call LogGM(.name, "Resucito a " & UserName)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "OnlineGM" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineGM(ByVal UserIndex As Integer)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 12/28/06
'
'***************************************************
    Dim i As Long
    Dim list As String

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsCONSE(UserIndex) Then

            For i = 1 To LastUser
                If UserList(i).flags.UserLogged Then
                    If EsCONSE(i) Then _
                        list = list & UserList(i).name & ", "
                End If
            Next i
            
            If LenB(list) <> 0 Then
                list = Left$(list, Len(list) - 2)
                Call WriteConsoleMsg(1, UserIndex, list & ".", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, UserIndex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)
            End If
            
        End If
    End With
End Sub

''
' Handles the "OnlineMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 23/03/2009
'23/03/2009: ZaMa - Ahora no requiere estar en el mapa, sino que por defecto se toma en el que esta, pero se puede especificar otro
'Mod Marius, ahora si funciona antes solo mostraba a los admin en el mapa.
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim map As Integer
        map = .incomingData.ReadInteger
        
        If EsCONSE(UserIndex) Then
        
            Dim loopC As Long
            Dim list As String
    
            For loopC = 1 To LastUser
                If LenB(UserList(loopC).name) <> 0 And UserList(loopC).Pos.map = map Then
                    list = list & UserList(loopC).name & ", "
                End If
            Next loopC
            
            If Len(list) > 2 Then list = Left$(list, Len(list) - 2)
            
            Call WriteConsoleMsg(1, UserIndex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)
            
        End If
    End With
End Sub


''
' Handles the "Kick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If EsCONSE(UserIndex) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(1, UserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToADMINS, 0, PrepareMessageConsoleMsg(1, .name & " echo a " & UserName & ".", FontTypeNames.FONTTYPE_INFO))
                Call CloseSocket(tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

Private Sub HandlePejotas(ByVal UserIndex As Integer)
'***************************************************
'Author: Marius
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim Pjs As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If EsSEMI(UserIndex) Then
            
            Pjs = Base.Pejotas(UserName)
            Call WriteConsoleMsg(1, UserIndex, "Lista de personajes de la cuenta de: " & Pjs, FontTypeNames.FONTTYPE_INFO)
            
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "Execute" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleExecute(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If EsDIOS(UserIndex) Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                If EsDIOS(tUser) And Not tUser = UserIndex Then
                    Call WriteConsoleMsg(1, UserIndex, "Que estas haciendo... a un miembro del staff No! Habla con Marius si tenes alguna duda.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If UserList(tUser).flags.Muerto = 0 Then
                        Call UserDie(tUser)
                        Call SendData(SendTarget.ToADMINS, 0, PrepareMessageConsoleMsg(1, .name & " ha ejecutado a " & UserName, FontTypeNames.FONTTYPE_TALKITALIC))
                        Call LogGM(.name, " ejecuto a " & UserName)
                    Else
                        Call WriteConsoleMsg(1, UserIndex, "Ya esta muerto!.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Else
                Call WriteConsoleMsg(1, UserIndex, "No está online", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "BanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        'Dim buffer As New clsByteQueue
        'Saque esto de la 0.13.3
        Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        
        UserName = buffer.ReadASCIIString()
        reason = buffer.ReadASCIIString()
        
        If EsSEMI(UserIndex) Then
            If ChangeBan(UserName, 1) Then
                Call LogGM(.name, "BAN a " & UserName & ". Razón: " & reason)
                Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(1, .name & " beneó a " & UserName & " por " & reason, FontTypeNames.FONTTYPE_INFO))
            Else
                Call WriteConsoleMsg(1, UserIndex, "No se pudo banear al usuario.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "UnbanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim cantPenas As Byte
        
        UserName = buffer.ReadASCIIString()
        
        If EsSEMI(UserIndex) Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            
            If Not ExistePersonaje(UserName) Then
                Call WriteConsoleMsg(1, UserIndex, "Charfile inexistente (no use +)", FontTypeNames.FONTTYPE_INFO)
            Else
                If BANCheckDB(UserName) Then
                    Call ChangeBan(UserName, 0)
                
                    Call LogGM(.name, "/UNBAN a " & UserName)
                    Call WriteConsoleMsg(1, UserIndex, UserName & " unbanned.", FontTypeNames.FONTTYPE_INFO)
                    Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(1, .name & " desbeneó a " & UserName, FontTypeNames.FONTTYPE_INFO))
                Else
                    Call WriteConsoleMsg(1, UserIndex, UserName & " no esta baneado. Imposible unbanear", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "NPCFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNPCFollow(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsSEMI(UserIndex) Then
        
            If .flags.TargetNPC > 0 Then
                Call DoFollow(.flags.TargetNPC, .name)
                Npclist(.flags.TargetNPC).flags.Inmovilizado = 0
                Npclist(.flags.TargetNPC).flags.Paralizado = 0
                Npclist(.flags.TargetNPC).Contadores.Paralisis = 0
            End If
        
        End If
    End With
End Sub

''
' Handles the "SummonChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSummonChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim x As Integer
        Dim Y As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If EsSEMI(UserIndex) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(1, UserIndex, "El jugador no esta online.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, tUser, .name & " te ha trasportado.", FontTypeNames.FONTTYPE_INFO)
                x = .Pos.x
                Y = .Pos.Y - 1
                Call FindLegalPos(tUser, .Pos.map, x, Y)
                Call WarpUserChar(tUser, .Pos.map, x, Y, True)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "SpawnListRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnListRequest(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsSEMI(UserIndex) Then
        
            Call EnviarSpawnList(UserIndex)
        
        End If
    End With
End Sub




''
' Handles the "SpawnCreature" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnCreature(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim npc As Integer
        npc = .incomingData.ReadInteger()
        
        If EsSEMI(UserIndex) Then
            If npc > 0 And npc <= UBound(Declaraciones.SpawnList()) Then _
              Call SpawnNpc(Declaraciones.SpawnList(npc).NpcIndex, .Pos, True, False)
        End If
    End With
End Sub

''
' Handles the "ResetNPCInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResetNPCInventory(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsDIOS(UserIndex) Then
            If .flags.TargetNPC = 0 Then Exit Sub
            
            Call ResetNpcInv(.flags.TargetNPC)
        End If
    End With
End Sub

''
' Handles the "CleanWorld" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCleanWorld(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If EsSEMI(UserIndex) Then
            Call limpiamundo(UserList(UserIndex).Pos.map)
        End If
        
    End With
End Sub

           
''
' Handles the "ServerMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        If EsSEMI(UserIndex) Then
            If LenB(message) <> 0 Then
                Call LogGM(.name, "Mensaje Broadcast:" & message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, message, FontTypeNames.FONTTYPE_TALK))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub





''
' Handles the "NickToIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNickToIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 24/07/07
'Pablo (ToxicWaste): Agrego para uqe el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim priv As PlayerType
        
        UserName = buffer.ReadASCIIString()
        
        If EsDIOS(UserIndex) Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                If UserList(tUser).flags.Privilegios And priv Then
                    Call WriteConsoleMsg(1, UserIndex, "El ip de " & UserName & " es " & UserList(tUser).ip, FontTypeNames.FONTTYPE_INFO)
                    Dim ip As String
                    Dim Lista As String
                    Dim loopC As Long
                    ip = UserList(tUser).ip
                    For loopC = 1 To LastUser
                        If UserList(loopC).ip = ip Then
                            If LenB(UserList(loopC).name) <> 0 And UserList(loopC).flags.UserLogged Then
                                Lista = Lista & UserList(loopC).name & ", "
                            End If
                        End If
                    Next loopC
                    If LenB(Lista) <> 0 Then Lista = Left$(Lista, Len(Lista) - 2)
                    Call WriteConsoleMsg(1, UserIndex, "Los personajes con ip " & ip & " son: " & Lista, FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(1, UserIndex, "No hay ningun personaje con ese nick", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "IPToNick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleIPToNick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim ip As String
        Dim loopC As Long
        Dim Lista As String
        Dim priv As PlayerType
        
        ip = .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte()
        
        If EsDIOS(UserIndex) Then
            
            Call LogGM(.name, "IP2NICK Solicito los Nicks de IP " & ip)
            For loopC = 1 To LastUser
                If UserList(loopC).ip = ip Then
                    If LenB(UserList(loopC).name) <> 0 And UserList(loopC).flags.UserLogged Then
                        Lista = Lista & UserList(loopC).name & ", "
                    End If
                End If
            Next loopC
            
            If LenB(Lista) <> 0 Then Lista = Left$(Lista, Len(Lista) - 2)
            Call WriteConsoleMsg(1, UserIndex, "Los personajes con ip " & ip & " son: " & Lista, FontTypeNames.FONTTYPE_INFO)
            
        End If
    End With
End Sub

''
' Handles the "GuildOnlineMembers" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnlineMembers(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim GuildName As String
        Dim tGuild As Integer
        
        GuildName = buffer.ReadASCIIString()
        
        If (EsSEMI(UserIndex)) Then
        
            If (InStrB(GuildName, "+") <> 0) Then
                GuildName = Replace(GuildName, "+", " ")
            End If
            
            tGuild = GuildIndex(GuildName)
                
            If tGuild > 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Clan " & UCase$(GuildName) & ": " & _
                    modGuilds.m_ListaDeMiembrosOnline(UserIndex, tGuild), FontTypeNames.FONTTYPE_TALKITALIC)
            End If
        
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "TeleportCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportCreate(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim mapa As Integer
        Dim x As Byte
        Dim Y As Byte
        
        mapa = .incomingData.ReadInteger()
        x = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        If EsSEMI(UserIndex) Then
    
            If Not MapaValido(mapa) Or Not InMapBounds(mapa, x, Y) Then _
                Exit Sub
            
            If MapData(.Pos.map, .Pos.x, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then _
                Exit Sub
            
            If MapData(.Pos.map, .Pos.x, .Pos.Y - 1).TileExit.map > 0 Then _
                Exit Sub
            
            If MapData(mapa, x, Y).ObjInfo.ObjIndex > 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Hay un objeto en el piso en ese lugar", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If MapData(mapa, x, Y).TileExit.map > 0 Then
                Call WriteConsoleMsg(1, UserIndex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            Dim ET As obj
            ET.Amount = 1
            ET.ObjIndex = 378
            
            Call MakeObj(ET, .Pos.map, .Pos.x, .Pos.Y - 1)
            
            With MapData(.Pos.map, .Pos.x, .Pos.Y - 1)
                .TileExit.map = mapa
                .TileExit.x = x
                .TileExit.Y = Y
            End With
            
        End If
    End With
End Sub

''
' Handles the "TeleportDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportDestroy(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        Dim mapa As Integer
        Dim x As Byte
        Dim Y As Byte
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        '/dt
        If EsSEMI(UserIndex) Then
        
            mapa = .flags.TargetMap
            x = .flags.TargetX
            Y = .flags.TargetY
            
            If Not InMapBounds(mapa, x, Y) Then Exit Sub
            
            With MapData(mapa, x, Y)
                If .ObjInfo.ObjIndex = 0 Then Exit Sub
                
                If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport And .TileExit.map > 0 Then
                    Call LogGM(UserList(UserIndex).name, "/DT: " & mapa & "," & x & "," & Y)
                    
                    Call EraseObj(.ObjInfo.Amount, mapa, x, Y)
                    
                    If MapData(.TileExit.map, .TileExit.x, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
                        Call EraseObj(1, .TileExit.map, .TileExit.x, .TileExit.Y)
                    End If
                    
                    .TileExit.map = 0
                    .TileExit.x = 0
                    .TileExit.Y = 0
                End If
            End With
            
        End If
    End With
End Sub


''
' Handles the "SetCharDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetCharDescription(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim tUser As Integer
        Dim desc As String
        
        desc = buffer.ReadASCIIString()
        
        If EsSEMI(UserIndex) Then
            tUser = .flags.TargetUser
            If tUser > 0 Then
                UserList(tUser).desc = desc
            Else
                Call WriteConsoleMsg(1, UserIndex, "Haz click sobre un personaje antes!", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "ForceMIDIToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HanldeForceMIDIToMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim midiID As Byte
        Dim mapa As Integer
        
        midiID = .incomingData.ReadByte
        mapa = .incomingData.ReadInteger
        
        'Solo dioses
        If EsDIOS(UserIndex) Then
            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(mapa, 50, 50) Then
                mapa = .Pos.map
            End If
        
            If midiID = 0 Then
                'Ponemos el default del mapa
                Call SendData(SendTarget.ToMap, mapa, PrepareMessagePlayMidi(MapInfo(.Pos.map).Music))
            Else
                'Ponemos el pedido por el GM
                Call SendData(SendTarget.ToMap, mapa, PrepareMessagePlayMidi(midiID))
            End If
        End If
    End With
End Sub

''
' Handles the "ForceWAVEToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEToMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim waveID As Byte
        Dim mapa As Integer
        Dim x As Byte
        Dim Y As Byte
        
        waveID = .incomingData.ReadByte()
        mapa = .incomingData.ReadInteger()
        x = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        'Solo dioses
        If EsDIOS(UserIndex) Then
        'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(mapa, x, Y) Then
                mapa = .Pos.map
                x = .Pos.x
                Y = .Pos.Y
            End If
            
            'Ponemos el pedido por el GM
            Call SendData(SendTarget.ToMap, mapa, PrepareMessagePlayWave(waveID, x, Y))
        End If
    End With
End Sub


''
' Handles the "TalkAsNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalkAsNPC(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If EsSEMI(UserIndex) Then
            'Asegurarse haya un NPC seleccionado
            If .flags.TargetNPC > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(message, Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Else
                Call WriteConsoleMsg(1, UserIndex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "DestroyAllItemsInArea" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyAllItemsInArea(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsCONSE(UserIndex) Then
            
            Dim x As Long
            Dim Y As Long
            
            For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
                For x = .Pos.x - MinXBorder + 1 To .Pos.x + MinXBorder - 1
                    If x > 0 And Y > 0 And x < 101 And Y < 101 Then
                        If MapData(.Pos.map, x, Y).ObjInfo.ObjIndex > 0 Then
                            If ItemNoEsDeMapa(MapData(.Pos.map, x, Y).ObjInfo.ObjIndex) Then
                                Call EraseObj(MAX_INVENTORY_OBJS, .Pos.map, x, Y)
                            End If
                        End If
                    End If
                Next x
            Next Y
            
            Call LogGM(UserList(UserIndex).name, "/MASSDEST")
            
        End If
    End With
End Sub

''
' Handles the "ItemsInTheFloor" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleItemsInTheFloor(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsSEMI(UserIndex) Then
            
            Dim tObj As Integer
            Dim Lista As String
            Dim x As Long
            Dim Y As Long
            
            For x = 5 To 95
                For Y = 5 To 95
                    tObj = MapData(.Pos.map, x, Y).ObjInfo.ObjIndex
                    If tObj > 0 Then
                        If ObjData(tObj).OBJType <> eOBJType.otArboles Then
                            Call WriteConsoleMsg(1, UserIndex, "(" & x & "," & Y & ") " & ObjData(tObj).name, FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                Next Y
            Next x
            
        End If
    End With
End Sub

''
' Handles the "MakeDumb" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumb(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If EsSEMI(UserIndex) Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumb(tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "MakeDumbNoMore" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumbNoMore(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If EsSEMI(UserIndex) Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumbNoMore(tUser)
                Call FlushBuffer(tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "DumpIPTables" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDumpIPTables(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsADMIN(UserIndex) Then
        
            Call SecurityIp.DumpTables
        
        End If
    End With
End Sub


''
' Handles the "SetTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetTrigger(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim tTrigger As Byte
        Dim tLog As String
        
        tTrigger = .incomingData.ReadByte()
        
        If EsADMIN(UserIndex) Then
        
            If tTrigger >= 0 Then
                MapData(.Pos.map, .Pos.x, .Pos.Y).Trigger = tTrigger
                tLog = "Trigger " & tTrigger & " en mapa " & .Pos.map & " " & .Pos.x & "," & .Pos.Y
                
                Call LogGM(.name, tLog)
                Call WriteConsoleMsg(1, UserIndex, tLog, FontTypeNames.FONTTYPE_INFO)
            End If
        
        End If
    End With
End Sub

''
' Handles the "AskTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAskTrigger(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 04/13/07
'
'***************************************************
    Dim tTrigger As Byte
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsADMIN(UserIndex) Then
        
            tTrigger = MapData(.Pos.map, .Pos.x, .Pos.Y).Trigger
            
            Call LogGM(.name, "Miro el trigger en " & .Pos.map & "," & .Pos.x & "," & .Pos.Y & ". Era " & tTrigger)
            
            Call WriteConsoleMsg(1, UserIndex, _
                "Trigger " & tTrigger & " en mapa " & .Pos.map & " " & .Pos.x & ", " & .Pos.Y _
                , FontTypeNames.FONTTYPE_INFO)
                
        End If
    End With
End Sub

''
' Handles the "BannedIPList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPList(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsDIOS(UserIndex) Then
            
            Dim Lista As String
            Dim loopC As Long
            
            Call LogGM(.name, "/BANIPLIST")
            
            For loopC = 1 To BanIps.count
                Lista = Lista & BanIps.Item(loopC) & ", "
            Next loopC
            
            If LenB(Lista) <> 0 Then Lista = Left$(Lista, Len(Lista) - 2)
            
            Call WriteConsoleMsg(1, UserIndex, Lista, FontTypeNames.FONTTYPE_INFO)
            
        End If
    End With
End Sub

''
' Handles the "BannedIPReload" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPReload(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsDIOS(UserIndex) Then
            
            Call BanIpGuardar
            Call BanIpCargar
            
        End If
    End With
End Sub

''
' Handles the "GuildBan" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildBan(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim GuildName As String
        Dim cantMembers As Integer
        Dim loopC As Long
        Dim member As String
        Dim count As Byte
        Dim tIndex As Integer
        Dim tFile As String
        
        GuildName = buffer.ReadASCIIString()
        
        If EsDIOS(UserIndex) Then
            tFile = App.Path & "\guilds\" & GuildName & "-members.mem"
            
            If Not FileExist(tFile) Then
                Call WriteConsoleMsg(1, UserIndex, "No existe el clan: " & GuildName, FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(1, .name & " banned al clan " & UCase$(GuildName), FontTypeNames.FONTTYPE_FIGHT))
                
                'baneamos a los miembros
                Call LogGM(.name, "BANCLAN a " & UCase$(GuildName))
                
                cantMembers = val(GetVar(tFile, "INIT", "NroMembers"))
                
                For loopC = 1 To cantMembers
                    member = GetVar(tFile, "Members", "Member" & loopC)
                    'member es la victima
                    Call ban(member, "Administracion del servidor", "Clan Banned")
                    
                    Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(1, "   " & member & "<" & GuildName & "> ha sido expulsado del servidor.", FontTypeNames.FONTTYPE_FIGHT))
                    
                    tIndex = NameIndex(member)
                    If tIndex > 0 Then
                        'esta online
                        
                        Call ChangeBan(member, 7)
                        
                        UserList(tIndex).flags.ban = 1
                        Call CloseSocket(tIndex)
                    End If
                    
                    Call ChangeBan(member, 7)
                Next loopC
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "BanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/12/08
'Agregado un CopyBuffer porque se producia un bucle
'inifito al intentar banear una ip ya baneada. (NicoNZ)
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim bannedIP As String
        Dim tUser As Integer
        Dim reason As String
        Dim i As Long
        
        ' Is it by ip??
        If buffer.ReadBoolean() Then
            bannedIP = buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte()
        Else
            tUser = NameIndex(buffer.ReadASCIIString())
            
            If EsDIOS(UserIndex) Then
                If tUser <= 0 Then
                    Call WriteConsoleMsg(1, UserIndex, "El personaje no está online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    bannedIP = UserList(tUser).ip
                End If
            End If
        End If
        
        reason = buffer.ReadASCIIString()
        
        If LenB(bannedIP) > 0 Then
            If EsDIOS(UserIndex) Then
                Call LogGM(.name, "/BanIP " & bannedIP & " por " & reason)
                
                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(1, UserIndex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                    Call .incomingData.CopyBuffer(buffer) ' Agregado porque sino no se sacaba del
                                                          ' buffer y se hacia un bucle infinito. (NicoNZ) 05/12/2008
                                        Set buffer = Nothing
                    Exit Sub
                End If
                
                Call BanIpAgrega(bannedIP)
                Call SendData(SendTarget.ToADMINS, 0, PrepareMessageConsoleMsg(1, .name & " baneó la IP " & bannedIP & " por " & reason, FontTypeNames.FONTTYPE_FIGHT))
                
                'Find every player with that ip and ban him!
                For i = 1 To LastUser
                    If UserList(i).ConnIDValida Then
                        If UserList(i).ip = bannedIP Then
                            Call ChangeBan(UserList(i).name, 1)
                        End If
                    End If
                Next i
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "UnbanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim bannedIP As String
        
        bannedIP = .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte()
        
        If EsDIOS(UserIndex) Then
        
            If BanIpQuita(bannedIP) Then
                Call WriteConsoleMsg(1, UserIndex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, UserIndex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        End If
    End With
End Sub

''
' Handles the "CreateItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim tObj As Integer
        tObj = .incomingData.ReadInteger()
        
        If tObj < 1 Or tObj > NumObjDatas Then _
            Exit Sub
        
        'Is the object not null?
        If LenB(ObjData(tObj).name) = 0 Then Exit Sub
        
        Dim Piso As WorldPos
        Dim Objeto As obj

        Objeto.Amount = .incomingData.ReadInteger()
        Objeto.ObjIndex = tObj
        
        If EsADMIN(UserIndex) Then
            
            Call LogGM(.name, "/CI: " & tObj)
            
            'Mod Nod Kopfnickend
            'Los items creados ahora van al inventario, sino hay lugar van al piso
            'Piso = TirarItemAlPiso(.Pos, Objeto)
            If Not MeterItemEnInventario(UserIndex, Objeto) Then
                Piso = TirarItemAlPiso(.Pos, Objeto)
            End If
            '/Mod
    
            'Call MakeObj(Objeto, .Pos.map, Piso.X, Piso.Y)
            
        End If
    End With
End Sub

''
' Handles the "DestroyItems" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyItems(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsCONSE(UserIndex) Then
            
            If MapData(.Pos.map, .Pos.x, .Pos.Y).ObjInfo.ObjIndex = 0 Then Exit Sub
            
            Call LogGM(.name, "/DEST")
            
            If ObjData(MapData(.Pos.map, .Pos.x, .Pos.Y).ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport Then
                Call WriteConsoleMsg(1, UserIndex, "No puede destruir teleports así. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call EraseObj(10000, .Pos.map, .Pos.x, .Pos.Y)
            
        End If
    End With
End Sub

''
' Handles the "ChaosLegionKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim Indexpj As Integer
        Dim RS As New ADODB.Recordset
        Dim str As String
        
        UserName = buffer.ReadASCIIString()
        
        If EsFaccCaos(UserIndex) Or EsSEMI(UserIndex) Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            tUser = NameIndex(UserName)
            
            Call LogGM(.name, "ECHO DEL CAOS A: " & UserName)
    
            If tUser > 0 Then
                UserList(tUser).faccion.FuerzasCaos = 100
                UserList(tUser).faccion.Rango = 0
                UserList(tUser).faccion.Renegado = 1
                Call WriteConsoleMsg(1, UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(1, tUser, .name & " te ha expulsado en forma definitiva de las fuerzas del caos.", FontTypeNames.FONTTYPE_FIGHT)
                
                Call QuitarItemsFaccionarios(tUser)
    
                Call SendData(SendTarget.ToPCArea, tUser, PrepareMessageCharStatus(UserList(tUser).Char.CharIndex, UserTypeColor(tUser)))
                Call FlushBuffer(tUser)
                
                Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(1, .name & " Echo del caos a " & UserName, FontTypeNames.FONTTYPE_SERVER))
            Else
                Indexpj = GetIndexPJ(UserName)
                If Indexpj > 0 Then
                    Set RS = DB_Conn.Execute("SELECT * FROM `charfaccion` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
                    str = "UPDATE `charfaccion` SET"
                    
                    str = str & " IndexPJ=" & RS!Indexpj
                    str = str & ",EjercitoReal=" & RS!EjercitoReal
                    str = str & ",EjercitoCaos=100"
                    str = str & ",EjercitoMili=" & RS!EjercitoMili
                    str = str & ",Republicano=" & RS!Republicano
                    str = str & ",Ciudadano=" & RS!Ciudadano
                    str = str & ",Rango=0"
                    str = str & ",Renegado=1"
                    str = str & ",CiudMatados=" & RS!CiudMatados
                    str = str & ",ReneMatados=" & RS!ReneMatados
                    str = str & ",RepuMatados=" & RS!RepuMatados
                    str = str & ",CaosMatados=" & RS!CaosMatados
                    str = str & ",ArmadaMatados=" & RS!ArmadaMatados
                    str = str & ",MiliMatados=" & RS!MiliMatados
                    str = str & " WHERE IndexPJ=" & RS!Indexpj
                    
                    Call DB_Conn.Execute(str)
                    
                    Set RS = Nothing
                    
                    Call WriteConsoleMsg(1, UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                    Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(1, .name & " Echo del caos a " & UserName, FontTypeNames.FONTTYPE_SERVER))
                Else
                    Call WriteConsoleMsg(1, UserIndex, UserName & " no se encuentra en la base de datos.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "RoyalArmyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim Indexpj As Integer
        Dim RS As New ADODB.Recordset
        Dim str As String
        
        UserName = buffer.ReadASCIIString()
        
        If EsFaccImpe(UserIndex) Or EsSEMI(UserIndex) Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            tUser = NameIndex(UserName)
            
            Call LogGM(.name, "ECHO DE LA REAL A: " & UserName)
            
            If tUser > 0 Then
                UserList(tUser).faccion.ArmadaReal = 100
                UserList(tUser).faccion.Rango = 0
                UserList(tUser).faccion.Renegado = 1
                Call WriteConsoleMsg(1, UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(1, tUser, .name & " te ha expulsado en forma definitiva de las fuerzas reales.", FontTypeNames.FONTTYPE_FIGHT)

                Call QuitarItemsFaccionarios(tUser)
    
                Call SendData(SendTarget.ToPCArea, tUser, PrepareMessageCharStatus(UserList(tUser).Char.CharIndex, UserTypeColor(tUser)))
                Call FlushBuffer(tUser)
                Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(1, .name & " Echo de la armada a " & UserName, FontTypeNames.FONTTYPE_SERVER))
            Else
                Indexpj = GetIndexPJ(UserName)
                If Indexpj > 0 Then
                    Set RS = DB_Conn.Execute("SELECT * FROM `charfaccion` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
                    str = "UPDATE `charfaccion` SET"
                    
                    str = str & " IndexPJ=" & RS!Indexpj
                    str = str & ",EjercitoReal=100"
                    str = str & ",EjercitoCaos=" & RS!EjercitoCaos
                    str = str & ",EjercitoMili=" & RS!EjercitoMili
                    str = str & ",Republicano=" & RS!Republicano
                    str = str & ",Ciudadano=" & RS!Ciudadano
                    str = str & ",Rango=0"
                    str = str & ",Renegado=1"
                    str = str & ",CiudMatados=" & RS!CiudMatados
                    str = str & ",ReneMatados=" & RS!ReneMatados
                    str = str & ",RepuMatados=" & RS!RepuMatados
                    str = str & ",CaosMatados=" & RS!CaosMatados
                    str = str & ",ArmadaMatados=" & RS!ArmadaMatados
                    str = str & ",MiliMatados=" & RS!MiliMatados
                    str = str & " WHERE IndexPJ=" & RS!Indexpj & " LIMIT 1"
                    
                    Call DB_Conn.Execute(str)
                    Call WriteConsoleMsg(1, UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                    Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(1, .name & " Echo de la armada a " & UserName, FontTypeNames.FONTTYPE_SERVER))
                Else
                    Call WriteConsoleMsg(1, UserIndex, UserName & " no se encuentra en la base de datos.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub


Private Sub HandleMiliciaKick(ByVal UserIndex As Integer)
'***************************************************
' Marius (re que fue un Copiar & Pegar y le cambie 2 valores...)
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim Indexpj As Integer
        Dim RS As New ADODB.Recordset
        Dim str As String
        
        UserName = buffer.ReadASCIIString()
        
        If EsFaccRepu(UserIndex) Or EsSEMI(UserIndex) Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            tUser = NameIndex(UserName)
            
            Call LogGM(.name, "ECHO DE LA MILICIA A: " & UserName)
            
            If tUser > 0 Then
                UserList(tUser).faccion.Milicia = 100
                UserList(tUser).faccion.Rango = 0
                UserList(tUser).faccion.Renegado = 1
                Call WriteConsoleMsg(1, UserIndex, UserName & " expulsado de la milicia republicana y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(1, tUser, .name & " te ha expulsado en forma definitiva de la milicia republicana.", FontTypeNames.FONTTYPE_FIGHT)

                Call QuitarItemsFaccionarios(tUser)
    
                Call SendData(SendTarget.ToPCArea, tUser, PrepareMessageCharStatus(UserList(tUser).Char.CharIndex, UserTypeColor(tUser)))
                Call FlushBuffer(tUser)
                Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(1, .name & " Echo de la milicia a " & UserName, FontTypeNames.FONTTYPE_SERVER))
            Else
                Indexpj = GetIndexPJ(UserName)
                If Indexpj > 0 Then
                    Set RS = DB_Conn.Execute("SELECT * FROM `charfaccion` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
                    str = "UPDATE `charfaccion` SET"
                    
                    str = str & " IndexPJ=" & RS!Indexpj
                    str = str & ",EjercitoReal=" & RS!EjercitoReal
                    str = str & ",EjercitoCaos=" & RS!EjercitoCaos
                    str = str & ",EjercitoMili=100"
                    str = str & ",Republicano=" & RS!Republicano
                    str = str & ",Ciudadano=" & RS!Ciudadano
                    str = str & ",Rango=0"
                    str = str & ",Renegado=1"
                    str = str & ",CiudMatados=" & RS!CiudMatados
                    str = str & ",ReneMatados=" & RS!ReneMatados
                    str = str & ",RepuMatados=" & RS!RepuMatados
                    str = str & ",CaosMatados=" & RS!CaosMatados
                    str = str & ",ArmadaMatados=" & RS!ArmadaMatados
                    str = str & ",MiliMatados=" & RS!MiliMatados
                    str = str & " WHERE IndexPJ=" & RS!Indexpj & " LIMIT 1"
                    
                    Call DB_Conn.Execute(str)
                    Call WriteConsoleMsg(1, UserIndex, UserName & " expulsado de la milicia republicana y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                    Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(1, .name & " Echo de la milicia a " & UserName, FontTypeNames.FONTTYPE_SERVER))
                Else
                    Call WriteConsoleMsg(1, UserIndex, UserName & " no se encuentra en la base de datos.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "ForceMIDIAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceMIDIAll(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim midiID As Byte
        midiID = .incomingData.ReadByte()
        
        If EsDIOS(UserIndex) Then
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, .name & " broadcast musica: " & midiID, FontTypeNames.FONTTYPE_SERVER))
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMidi(midiID))
            
        End If
    End With
End Sub

''
' Handles the "ForceWAVEAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEAll(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim waveID As Byte
        waveID = .incomingData.ReadByte()
        
        If EsDIOS(UserIndex) Then
        
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))
        
        End If
    End With
End Sub

''
' Handles the "TileBlockedToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTileBlockedToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsCONSE(UserIndex) Then
            Call LogGM(.name, "/BLOQ")
            
            If MapData(.Pos.map, .Pos.x, .Pos.Y).Blocked = 0 Then
                MapData(.Pos.map, .Pos.x, .Pos.Y).Blocked = 1
            Else
                MapData(.Pos.map, .Pos.x, .Pos.Y).Blocked = 0
            End If
            
            Call Bloquear(True, .Pos.map, .Pos.x, .Pos.Y, MapData(.Pos.map, .Pos.x, .Pos.Y).Blocked)
        End If
    End With
End Sub

''
' Handles the "KillNPCNoRespawn" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPCNoRespawn(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsDIOS(UserIndex) Then
            If .flags.TargetNPC = 0 Then Exit Sub
            
            Call QuitarNPC(.flags.TargetNPC)
            Call LogGM(.name, "/MATA " & Npclist(.flags.TargetNPC).name)
        End If
    End With
End Sub

''
' Handles the "KillAllNearbyNPCs" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillAllNearbyNPCs(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsSEMI(UserIndex) Then
            
            Dim x As Long
            Dim Y As Long
            
            For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
                For x = .Pos.x - MinXBorder + 1 To .Pos.x + MinXBorder - 1
                    If x > 0 And Y > 0 And x < 101 And Y < 101 Then
                        If MapData(.Pos.map, x, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(.Pos.map, x, Y).NpcIndex)
                    End If
                Next x
            Next Y
            Call LogGM(.name, "/MASSKILL")
            
        End If
    End With
End Sub

''
' Handles the "LastIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLastIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim Lista As String
        Dim loopC As Byte
        Dim validCheck As Boolean
        Dim RS      As New ADODB.Recordset
        Dim Indexpj As Integer
    
        UserName = buffer.ReadASCIIString()
        
        If EsDIOS(UserIndex) Then
            'Handle special chars
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            If (InStrB(UserName, "+") <> 0) Then
                UserName = Replace(UserName, "+", " ")
            End If
            
            Indexpj = GetIndexPJ(UserName)
            Set RS = DB_Conn.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
            If (RS.BOF Or RS.EOF) = False Then
                Call WriteConsoleMsg(1, UserIndex, "La ultima IP con la que " & UserName & " se conectó es: " & RS!LastIP, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, UserIndex, UserName & " no se encuentra en la base de datos.", FontTypeNames.FONTTYPE_INFO)
            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "Ignored" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIgnored(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Ignore the user
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If EsDIOS(UserIndex) Then
            .flags.AdminPerseguible = Not .flags.AdminPerseguible
        End If
    End With
End Sub

''
' Handles the "CheckSlot" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleCheckSlot(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 09/09/2008 (NicoNZ)
'Check one Users Slot in Particular from Inventory
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        'Reads the UserName and Slot Packets
        Dim UserName As String
        Dim Slot As Byte
        Dim tIndex As Integer
        
        UserName = buffer.ReadASCIIString() 'Que UserName?
        Slot = buffer.ReadByte() 'Que Slot?
        
        If EsDIOS(UserIndex) Then
            tIndex = NameIndex(UserName)  'Que user index?
            
            Call LogGM(.name, .name & " Checkeo el slot " & Slot & " de " & UserName)
               
            If tIndex > 0 Then
                If Slot > 0 And Slot <= MAX_INVENTORY_SLOTS Then
                    If UserList(tIndex).Invent.Object(Slot).ObjIndex > 0 Then
                        Call WriteConsoleMsg(1, UserIndex, " Objeto " & Slot & ") " & ObjData(UserList(tIndex).Invent.Object(Slot).ObjIndex).name & " Cantidad:" & UserList(tIndex).Invent.Object(Slot).Amount, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(1, UserIndex, "No hay Objeto en slot seleccionado", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(1, UserIndex, "Slot Inválido.", FontTypeNames.FONTTYPE_TALK)
                End If
            Else
                Call WriteConsoleMsg(1, UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handles the "Restart" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleRestart(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Restart the game
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
    
        If EsADMIN(UserIndex) Then
            'time and Time BUG!
            Call LogGM(.name, .name & " reinicio el mundo")
            
            Call ReiniciarServidor(True)
            
        End If
    End With
End Sub

''
' Handles the "ReloadObjects" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadObjects(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the objects
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If EsDIOS(UserIndex) Then
            
            Call LogGM(.name, .name & " ha recargado los objetos.")
            
            Call LoadOBJData
            
        End If
    End With
End Sub

''
' Handles the "ReloadSpells" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadSpells(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the spells
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If EsDIOS(UserIndex) Then
            
            Call LogGM(.name, .name & " ha recargado los hechizos.")
            
            Call CargarHechizos
            
        End If
        
    End With
End Sub

''
' Handle the "ReloadServerIni" message.
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadServerIni(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the Server`s INI
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If EsDIOS(UserIndex) Then
            
            Call LogGM(.name, .name & " ha recargado los INITs y el BALANCE.")
            
            Call LoadSini
            Call LoadBalance
            
        End If
        
    End With
End Sub

''
' Handle the "ReloadNPCs" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadNPCs(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the Server`s NPC
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If EsDIOS(UserIndex) Then
             
            Call LogGM(.name, .name & " ha recargado los NPCs.")
        
            Call CargaNpcsDat
        
            Call WriteConsoleMsg(1, UserIndex, "Npcs.dat recargado.", FontTypeNames.FONTTYPE_INFO)
            
        End If
    End With
End Sub

''
' Handle the "KickAllChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleKickAllChars(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Kick all the chars that are online
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If EsADMIN(UserIndex) Then
            Call LogGM(.name, .name & " ha echado a todos los personajes.")
            
            Call EcharPjsNoPrivilegiados
        End If
    End With
End Sub


''
' Handle the "CleanSOS" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCleanSOS(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Clean the SOS
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If EsDIOS(UserIndex) Then
        
            Call LogGM(.name, .name & " ha borrado los SOS")
        
            Call Ayuda.Reset
        
        End If
    End With
End Sub

''
' Handle the "SaveChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveChars(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Save the characters
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If EsDIOS(UserIndex) Then
            
            Call LogGM(.name, .name & " ha guardado todos los chars")
            
            Call GuardarUsuarios
            
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoBackup" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoBackup(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the backup`s info of the map
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim doTheBackUp As Boolean
        
        doTheBackUp = .incomingData.ReadBoolean()
        
        If EsDIOS(UserIndex) Then
            
            Call LogGM(.name, .name & " ha cambiado la información sobre el BackUp")
            
            'Change the boolean to byte in a fast way
            If doTheBackUp Then
                MapInfo(.Pos.map).backup = 1
            Else
                MapInfo(.Pos.map).backup = 0
            End If
            
            'Change the boolean to string in a fast way
            Call WriteVar(App.Path & MapPath & "mapa" & .Pos.map & ".dat", "Mapa" & .Pos.map, "backup", MapInfo(.Pos.map).backup)
            
            Call WriteConsoleMsg(1, UserIndex, "Mapa " & .Pos.map & " Backup: " & MapInfo(.Pos.map).backup, FontTypeNames.FONTTYPE_INFO)
            
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoPK" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoPK(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the pk`s info of the  map
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim isMapPk As Boolean
        
        isMapPk = .incomingData.ReadBoolean()
        
        If EsDIOS(UserIndex) Then
            
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si es PK el mapa.")
            
            MapInfo(.Pos.map).Pk = isMapPk
            
            'Change the boolean to string in a fast way
            Call WriteVar(App.Path & MapPath & "mapa" & .Pos.map & ".dat", "Mapa" & .Pos.map, "Pk", IIf(isMapPk, "1", "0"))
    
            Call WriteConsoleMsg(1, UserIndex, "Mapa " & .Pos.map & " PK: " & MapInfo(.Pos.map).Pk, FontTypeNames.FONTTYPE_INFO)
            
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoRestricted" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoRestricted(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Restringido -> Options: "NEWBIE", "NO", "ARMADA", "CAOS", "FACCION".
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If EsDIOS(UserIndex) Then
            If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Or tStr = "FACCION" Then
                Call LogGM(.name, .name & " ha cambiado la informacion sobre si es Restringido el mapa.")
                MapInfo(UserList(UserIndex).Pos.map).Restringir = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "Restringir", tStr)
                Call WriteConsoleMsg(1, UserIndex, "Mapa " & .Pos.map & " Restringido: " & MapInfo(.Pos.map).Restringir, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, UserIndex, "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION'", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handle the "ChangeMapInfoNoMagic" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoMagic(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'MagiaSinEfecto -> Options: "1" , "0".
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim nomagic As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        nomagic = .incomingData.ReadBoolean
        
        If EsDIOS(UserIndex) Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar la Magia el mapa.")
            MapInfo(UserList(UserIndex).Pos.map).MagiaSinEfecto = nomagic
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "MagiaSinEfecto", nomagic)
            Call WriteConsoleMsg(1, UserIndex, "Mapa " & .Pos.map & " MagiaSinEfecto: " & MapInfo(.Pos.map).MagiaSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoNoInvi" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvi(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'InviSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim noinvi As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        noinvi = .incomingData.ReadBoolean()
        
        If EsDIOS(UserIndex) Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar Invisibilidad el mapa.")
            MapInfo(UserList(UserIndex).Pos.map).InviSinEfecto = noinvi
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "InviSinEfecto", noinvi)
            Call WriteConsoleMsg(1, UserIndex, "Mapa " & .Pos.map & " InviSinEfecto: " & MapInfo(.Pos.map).InviSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub
            
''
' Handle the "ChangeMapInfoNoResu" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoResu(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'ResuSinEfecto -> Options: "1", "0"
'***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim noresu As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        noresu = .incomingData.ReadBoolean()
        
        If EsDIOS(UserIndex) Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar Resucitar el mapa.")
            MapInfo(UserList(UserIndex).Pos.map).ResuSinEfecto = noresu
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "ResuSinEfecto", noresu)
            Call WriteConsoleMsg(1, UserIndex, "Mapa " & .Pos.map & " ResuSinEfecto: " & MapInfo(.Pos.map).ResuSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoLand" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoLand(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If EsDIOS(UserIndex) Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.name, .name & " ha cambiado la informacion del Terreno del mapa.")
                MapInfo(UserList(UserIndex).Pos.map).Terreno = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "Terreno", tStr)
                Call WriteConsoleMsg(1, UserIndex, "Mapa " & .Pos.map & " Terreno: " & MapInfo(.Pos.map).Terreno, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(1, UserIndex, "Igualmente, el único útil es 'NIEVE' ya que al ingresarlo, la gente muere de frio en el Mapa", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handle the "ChangeMapInfoZone" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoZone(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()

        If EsDIOS(UserIndex) Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.name, .name & " ha cambiado la informacion de la Zona del mapa.")
                MapInfo(UserList(UserIndex).Pos.map).Zona = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "Zona", tStr)
                Call WriteConsoleMsg(1, UserIndex, "Mapa " & .Pos.map & " Zona: " & MapInfo(.Pos.map).Zona, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(1, UserIndex, "Igualmente, el único útil es 'DUNGEON' ya que al ingresarlo, NO se sentirá el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handle the "SaveMap" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Saves the map
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If EsDIOS(UserIndex) Then
            
            Call LogGM(.name, .name & " ha guardado el mapa " & CStr(.Pos.map))
            Call WriteConsoleMsg(1, UserIndex, "Mapa Guardado", FontTypeNames.FONTTYPE_INFO)
            
        End If
    End With
End Sub

''
' Handle the "ShowGuildMessages" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowGuildMessages(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Allows admins to read guild messages
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        
        guild = buffer.ReadASCIIString()
        
        If EsDIOS(UserIndex) Then
            Call modGuilds.GMEscuchaClan(UserIndex, guild)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handle the "DoBackUp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDoBackUp(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Show guilds messages
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If EsADMIN(UserIndex) Then
        
            Call LogGM(.name, .name & " ha hecho un backup")
        
        End If

    End With
End Sub

''
' Handle the "ToggleCentinelActivated" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleToggleCentinelActivated(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/26/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Activate or desactivate the Centinel
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If EsDIOS(UserIndex) Then
            
            
            
        End If

    End With
End Sub
''
' Handle the "HandleCreateNPC" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim NpcIndex As Integer
        
        NpcIndex = .incomingData.ReadInteger()
        
        If EsSEMI(UserIndex) Then
            
            NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False)
            
            If NpcIndex <> 0 Then
                Call LogGM(.name, "Sumoneo a " & Npclist(NpcIndex).name & " en mapa " & .Pos.map)
            End If
            
        End If
    End With
End Sub


''
' Handle the "CreateNPCWithRespawn" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPCWithRespawn(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim NpcIndex As Integer
        
        NpcIndex = .incomingData.ReadInteger()
        
        If EsDIOS(UserIndex) Then
            
            NpcIndex = SpawnNpc(NpcIndex, .Pos, True, True)
            
            If NpcIndex <> 0 Then
                Call LogGM(.name, "Sumoneo con respawn " & Npclist(NpcIndex).name & " en mapa " & .Pos.map)
            End If
            
        End If
    End With
End Sub


''
' Handle the "NavigateToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNavigateToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/12/07
'
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If EsSEMI(UserIndex) Then
            
            If .flags.Navegando = 1 Then
                .flags.Navegando = 0
            Else
                .flags.Navegando = 1
            End If
            
            'Tell the client that we are navigating.
            Call WriteNavigateToggle(UserIndex)
            
        End If
    End With
End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If EsADMIN(UserIndex) Then
            
            If ServerSoloGMs > 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
                ServerSoloGMs = 0
            Else
                Call WriteConsoleMsg(1, UserIndex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
                ServerSoloGMs = 1
            End If
            
        End If
    End With
End Sub

''
' Handle the "TurnOffServer" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnOffServer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'Turns off the server
'***************************************************
    Dim Handle As Integer
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If EsADMIN(UserIndex) Then
            
            Call LogGM(.name, "/APAGAR")
            Call SendData(SendTarget.ToADMINS, 0, PrepareMessageConsoleMsg(1, .name & " VA A APAGAR EL SERVIDOR!!!", FontTypeNames.FONTTYPE_FIGHT))
            
            Unload frmMain
            
        End If
    End With
End Sub

''
' Handle the "TurnCriminal" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnCriminal(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If EsDIOS(UserIndex) Then
            tUser = NameIndex(UserName)
            If tUser > 0 Then _
                Call VolverRenegado(tUser)
        End If
                
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handle the "ResetFactions" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleResetFactions(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If EsDIOS(UserIndex) Then
            Call LogGM(.name, "/RAJAR " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call ResetFacciones(tUser)
                UserList(tUser).faccion.Renegado = 1
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handle the "RemoveCharFromGuild" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRemoveCharFromGuild(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim GuildIndex As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If EsDIOS(UserIndex) Then
            Call LogGM(.name, "/RAJARCLAN " & UserName)
            
            GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)
            
            If GuildIndex = 0 Then
                Call WriteConsoleMsg(1, UserIndex, "No pertenece a ningún clan o es fundador.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, UserIndex, "Expulsado.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(1, UserName & " ha sido expulsado del clan por los administradores del servidor.", FontTypeNames.FONTTYPE_GUILD))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handle the "SystemMessage" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/29/06
'Send a message to all the users
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        If EsADMIN(UserIndex) Then
            Call LogGM(.name, "Mensaje de sistema:" & message)
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(message))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

''
' Handle the "Ping" message
'
' @param userIndex The index of the user sending the message

Public Sub HandlePing(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Show guilds messages
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Call WritePong(UserIndex)
    End With
End Sub

''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Logged" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Logged)
    
    
        UserList(UserIndex).Redundance = RandomNumber(15, 250)
    Call UserList(UserIndex).outgoingData.WriteByte(UserList(UserIndex).Redundance)
    
    
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RemoveDialogs)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal CharIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRemoveCharDialog(CharIndex))
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "NavigateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NavigateToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NavigateToggle)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub
Public Sub WriteEquitateToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NavigateToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.EquitateToggle)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub
''
' Writes the "Disconnect" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Disconnect" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Disconnect)

Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteDisconnect2(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Disconnect" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Disconnect2)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceEnd" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceEnd)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankEnd" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankEnd)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceInit)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal UserIndex As Integer, ByVal goliath As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankInit)
    Call UserList(UserIndex).outgoingData.WriteByte(goliath)
    If goliath = 1 Then
        Call UserList(UserIndex).outgoingData.WriteLong(UserList(UserIndex).Stats.Banco)
        Call UserList(UserIndex).outgoingData.WriteByte(UserList(UserIndex).BancoInvent.nroitems)
    End If
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceInit)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceEnd)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowBlacksmithForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowBlacksmithForm)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowCarpenterForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowCarpenterForm)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


Public Sub WriteShowSastreForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowSastreForm)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub



Public Sub WriteShowalquimiaForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowAlquimiaForm)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


''
' Writes the "NPCSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCSwing(ByVal UserIndex As Integer)

'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCSwing" message to the given user's outgoing data buffer
'***************************************************

On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NPCSwing)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "NPCKillUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCKillUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCKillUser" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NPCKillUser)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlockedWithShieldUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlockedWithShieldUser" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlockedWithShieldUser)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlockedWithShieldOther" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldOther(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlockedWithShieldOther" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlockedWithShieldOther)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserSwing(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserSwing" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserSwing)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


''
' Writes the "SafeModeOn" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeModeOn(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SafeModeOn" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.SafeModeOn)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "SafeModeOff" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeModeOff(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SafeModeOff" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.SafeModeOff)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "NobilityLost" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNobilityLost(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NobilityLost" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NobilityLost)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CantUseWhileMeditating" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCantUseWhileMeditating(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CantUseWhileMeditating" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CantUseWhileMeditating)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateSta)
        Call .WriteInteger(IIf(UserList(UserIndex).Stats.MinSTA < 0, 0, UserList(UserIndex).Stats.MinSTA))
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateMana)
        Call .WriteInteger(IIf(UserList(UserIndex).Stats.MinMAN < 0, 0, UserList(UserIndex).Stats.MinMAN))
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHP)
        Call .WriteInteger(IIf(UserList(UserIndex).Stats.minhp < 0, 0, UserList(UserIndex).Stats.minhp))
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateGold" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateGold)
        Call .WriteLong(UserList(UserIndex).Stats.GLD)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateExp" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateExp)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal UserIndex As Integer, ByVal map As Integer, ByVal Version As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMap" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeMap)
        Call .WriteInteger(map)
        Call .WriteInteger(Version)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PosUpdate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.posUpdate)
        Call .WriteByte(UserList(UserIndex).Pos.x)
        Call .WriteByte(UserList(UserIndex).Pos.Y)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "NPCHitUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCHitUser(ByVal UserIndex As Integer, ByVal Target As PartesCuerpo, ByVal damage As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCHitUser" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.NPCHitUser)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserHitNPC" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    damage The number of HP lost by the target creature.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHitNPC(ByVal UserIndex As Integer, ByVal damage As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserHitNPC" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserHitNPC)
        
        'It is a long to allow the "drake slayer" (matadracos) to kill the great red dragon of one blow.
        Call .WriteLong(damage)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserAttackedSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex       User to which the message is intended.
' @param    attackerIndex   The user index of the user that attacked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserAttackedSwing(ByVal UserIndex As Integer, ByVal attackerIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserAttackedSwing" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserAttackedSwing)
        Call .WriteInteger(UserList(attackerIndex).Char.CharIndex)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserHittedByUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    attackerChar Char index of the user hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHittedByUser(ByVal UserIndex As Integer, ByVal Target As PartesCuerpo, ByVal attackerChar As Integer, ByVal damage As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserHittedByUser" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserHittedByUser)
        Call .WriteInteger(attackerChar)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserHittedUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    attackedChar Char index of the user hitted.
' @param    damage The number of HP lost by the oponent hitted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHittedUser(ByVal UserIndex As Integer, ByVal Target As PartesCuerpo, ByVal attackedChar As Integer, ByVal damage As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserHittedUser" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserHittedUser)
        Call .WriteInteger(attackedChar)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatOverHead(ByVal UserIndex As Integer, ByVal chat As String, ByVal CharIndex As Integer, ByVal color As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChatOverHead" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageChatOverHead(chat, CharIndex, color))
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsoleMsg(ByVal console As Byte, ByVal UserIndex As Integer, ByVal chat As String, ByVal FontIndex As FontTypeNames)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsg(console, chat, FontIndex))
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildChat" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildChat(ByVal UserIndex As Integer, ByVal chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildChat" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageGuildChat(chat))
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal message As String, Optional ByVal preg As Byte = 0, Optional ByVal action As Byte = 0)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(message)
        Call .WriteByte(preg)
        If Not preg = 0 Then Call .WriteByte(action)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserIndexInServer)
        Call .WriteInteger(UserIndex)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserCharIndexInServer)
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, ByVal body As Integer, ByVal head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal x As Byte, ByVal Y As Byte, ByVal Weapon As Integer, ByVal Shield As Integer, _
                                ByVal fx As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal name As String, _
                                ByVal privileges As Byte, ByVal account As Boolean, ByVal bandera As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterCreate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterCreate(body, head, heading, CharIndex, x, Y, Weapon, Shield, fx, FXLoops, _
                                                            helmet, name, privileges, account, bandera))
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal UserIndex As Integer, ByVal CharIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterRemove" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(CharIndex))
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterMove(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal x As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterMove" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterMove(CharIndex, x, Y))
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteForceCharMove(ByVal UserIndex, ByVal Direccion As eHeading)
'***************************************************
'Author: ZaMa
'Last Modification: 26/03/2009
'Writes the "ForceCharMove" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageForceCharMove(Direccion))
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterChange" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterChange(ByVal UserIndex As Integer, ByVal body As Integer, ByVal head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal Weapon As Integer, ByVal Shield As Integer, _
                                ByVal fx As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterChange" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterChange(body, head, heading, CharIndex, Weapon, Shield, fx, FXLoops, helmet))
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub
Public Sub WriteCharStatus(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal Status As Byte)
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharStatus(CharIndex, Status))
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ObjectCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectCreate(ByVal UserIndex As Integer, ByVal x As Byte, ByVal Y As Byte, ByVal ObjIndex As Integer, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ObjectCreate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectCreate(x, Y, ObjIndex, Amount))
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal x As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ObjectDelete" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectDelete(x, Y))
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockPosition(ByVal UserIndex As Integer, ByVal x As Byte, ByVal Y As Byte, ByVal Blocked As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlockPosition" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(x)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub
''
' Writes the "PlayMidi" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMidi(ByVal UserIndex As Integer, ByVal midi As Byte, Optional ByVal loops As Integer = -1)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PlayMidi" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayMidi(midi, loops))
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "PlayWave" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayWave(ByVal UserIndex As Integer, ByVal wave As Byte, ByVal x As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/08/07
'Last Modified by: Rapsodius
'Added X and Y positions for 3D Sounds
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(wave, x, Y))
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildList(ByVal UserIndex As Integer, ByRef guildList() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim Tmp As String
    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildList)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAreaChanged(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AreaChanged" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AreaChanged)
        Call .WriteByte(UserList(UserIndex).Pos.x)
        Call .WriteByte(UserList(UserIndex).Pos.Y)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PauseToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePauseToggle())
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub
''
' Writes the "CreateFX" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateFX(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal fx As Integer, ByVal FXLoops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateFX" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateFX(CharIndex, fx, FXLoops))
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateUserStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateUserStats)
        Call .WriteInteger(UserList(UserIndex).Stats.maxhp)
        Call .WriteInteger(IIf(UserList(UserIndex).Stats.minhp < 0, 0, UserList(UserIndex).Stats.minhp))
        Call .WriteInteger(UserList(UserIndex).Stats.MaxMAN)
        Call .WriteInteger(IIf(UserList(UserIndex).Stats.MinMAN < 0, 0, UserList(UserIndex).Stats.MinMAN))
        Call .WriteInteger(UserList(UserIndex).Stats.MaxSTA)
        Call .WriteInteger(IIf(UserList(UserIndex).Stats.MinSTA < 0, 0, UserList(UserIndex).Stats.MinSTA))
        Call .WriteLong(UserList(UserIndex).Stats.GLD)
        Call .WriteByte(UserList(UserIndex).Stats.ELV)
        Call .WriteLong(UserList(UserIndex).Stats.ELU)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkRequestTarget(ByVal UserIndex As Integer, ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WorkRequestTarget" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.WorkRequestTarget)
        Call .WriteByte(Skill)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeInventorySlot)
        Call .WriteByte(Slot)
        
        Dim ObjIndex As Integer
        Dim obData As ObjData
        
        ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)
        End If
        
        Call .WriteInteger(ObjIndex)
        Call .WriteInteger(UserList(UserIndex).Invent.Object(Slot).Amount)
        Call .WriteBoolean(UserList(UserIndex).Invent.Object(Slot).Equipped)
        Call .WriteInteger(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.maxhit)
        Call .WriteInteger(obData.minhit)
        Call .WriteInteger(obData.def)
        Call .WriteSingle(SalePrice(obData.valor))
        Call .WriteByte(IIf(obData.MinELV < UserList(UserIndex).Stats.ELV And SexoPuedeUsarItem(UserIndex, ObjIndex) = True And FaccionPuedeUsarItem(UserIndex, ObjIndex) = True And ClasePuedeUsarItem(UserIndex, ObjIndex) = True, 1, 0))
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeBankSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeBankSlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeBankSlot)
        Call .WriteByte(Slot)
        
        Dim ObjIndex As Integer
        Dim obData As ObjData
        
        ObjIndex = UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex
        
        Call .WriteInteger(ObjIndex)
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)
        End If
        
        Call .WriteInteger(UserList(UserIndex).BancoInvent.Object(Slot).Amount)
        Call .WriteInteger(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.maxhit)
        Call .WriteInteger(obData.minhit)
        Call .WriteInteger(obData.def)
        Call .WriteLong(obData.valor)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal Slot As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeSpellSlot)
        Call .WriteByte(Slot)
        Call .WriteInteger(UserList(UserIndex).Stats.UserHechizos(Slot))
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Atributes" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Atributes" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.atributes)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion))
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithWeapons(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
'Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim obj As ObjData
    Dim validIndexes() As Integer
    Dim count As Integer
    
    ReDim validIndexes(1 To UBound(ArmasHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithWeapons)
        
        For i = 1 To UBound(ArmasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(UserIndex).Clase), 0) Then
                If Not (ObjData(ArmasHerrero(i)).LingH = 0 And ObjData(ArmasHerrero(i)).LingO = 0 And ObjData(ArmasHerrero(i)).LingP = 0) Then
                    count = count + 1
                    validIndexes(count) = i
                End If
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(count)
        
        ' Write the needed data of each object
        For i = 1 To count
            obj = ObjData(ArmasHerrero(validIndexes(i)))
            Call .WriteASCIIString(obj.name)
            Call .WriteInteger(obj.LingH)
            Call .WriteInteger(obj.LingP)
            Call .WriteInteger(obj.LingO)
            Call .WriteInteger(ArmasHerrero(validIndexes(i)))
        Next i
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlacksmithArmors" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithArmors(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
'Writes the "BlacksmithArmors" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim obj As ObjData
    Dim validIndexes() As Integer
    Dim count As Integer
    
    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithArmors)
        
        For i = 1 To UBound(ArmadurasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmadurasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(UserIndex).Clase), 0) Then
                If Not (ObjData(ArmadurasHerrero(i)).LingH = 0 And ObjData(ArmadurasHerrero(i)).LingO = 0 And ObjData(ArmadurasHerrero(i)).LingP = 0) Then
                    count = count + 1
                    validIndexes(count) = i
                End If
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(count)
        
        ' Write the needed data of each object
        For i = 1 To count
            obj = ObjData(ArmadurasHerrero(validIndexes(i)))
            Call .WriteASCIIString(obj.name)
            Call .WriteInteger(obj.LingH)
            Call .WriteInteger(obj.LingP)
            Call .WriteInteger(obj.LingO)
            Call .WriteInteger(ArmadurasHerrero(validIndexes(i)))
        Next i
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CarpenterObjects" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCarpenterObjects(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CarpenterObjects" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim obj As ObjData
    Dim validIndexes() As Integer
    Dim count As Integer
    
    ReDim validIndexes(1 To UBound(ObjCarpintero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CarpenterObjects)
        
        For i = 1 To UBound(ObjCarpintero())
            ' Can the user create this object? If so add it to the list....
            If ObjCarpintero(i) <> 0 Then
                If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(UserList(UserIndex).Clase) Then
                    count = count + 1
                    validIndexes(count) = i
                End If
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(count)
        
        ' Write the needed data of each object
        For i = 1 To count
            obj = ObjData(ObjCarpintero(validIndexes(i)))
            Call .WriteASCIIString(obj.name)
            Call .WriteInteger(obj.Madera)
            Call .WriteInteger(ObjCarpintero(validIndexes(i)))
        Next i
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub



Public Sub WriteAlquimiaObjects(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CarpenterObjects" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim obj As ObjData
    Dim validIndexes() As Integer
    Dim count As Integer
    
    ReDim validIndexes(1 To UBound(ObjDruida()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AlquimiaObjects)
        
        For i = 1 To UBound(ObjDruida())
            ' Can the user create this object? If so add it to the list....
            If ObjDruida(i) <> 0 Then
                If ObjData(ObjDruida(i)).SkPociones <= UserList(UserIndex).Stats.UserSkills(eSkill.alquimia) \ Modalquimia(UserList(UserIndex).Clase) Then
                    count = count + 1
                    validIndexes(count) = i
                End If
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(count)
        
        ' Write the needed data of each object
        For i = 1 To count
            obj = ObjData(ObjDruida(validIndexes(i)))
            Call .WriteASCIIString(obj.name)
            Call .WriteInteger(obj.raies)
            Call .WriteInteger(ObjDruida(validIndexes(i)))
        Next i
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub



''
' Writes the "RestOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RestOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RestOK)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ErrorMsg" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageErrorMsg(message))
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Blind" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Blind" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Blind)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Dumb" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Dumb" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Dumb)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub



''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeNPCInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef obj As obj, ByVal price As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 06/13/08
'Last Modified by: Nicolas Ezequiel Bouhid (NicoNZ)
'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim ObjInfo As ObjData
    
    If obj.ObjIndex >= LBound(ObjData()) And obj.ObjIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(obj.ObjIndex)
    End If
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeNPCInventorySlot)
        Call .WriteByte(Slot)
        Call .WriteInteger(obj.Amount)
        Call .WriteSingle(price)
        Call .WriteInteger(ObjInfo.GrhIndex)
        Call .WriteInteger(obj.ObjIndex)
        Call .WriteByte(ObjInfo.OBJType)
        Call .WriteInteger(ObjInfo.maxhit)
        Call .WriteInteger(ObjInfo.minhit)
        Call .WriteInteger(ObjInfo.def)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHungerAndThirst)
        Call .WriteByte(UserList(UserIndex).Stats.MinAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MinHAM)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


''
' Writes the "MiniStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMiniStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MiniStats" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MiniStats)
        
        Call .WriteLong(UserList(UserIndex).faccion.CiudadanosMatados)
        Call .WriteLong(UserList(UserIndex).faccion.RenegadosMatados)
        Call .WriteLong(UserList(UserIndex).faccion.RepublicanosMatados)
        
        Call .WriteLong(UserList(UserIndex).faccion.ArmadaMatados)
        Call .WriteLong(UserList(UserIndex).faccion.CaosMatados)
        Call .WriteLong(UserList(UserIndex).faccion.MilicianosMatados)
        
        Call .WriteLong(UserList(UserIndex).Stats.VecesMuertos)
        
        Call .WriteInteger(UserList(UserIndex).Stats.NPCsMuertos)
        
        Call .WriteByte(UserList(UserIndex).Clase)
        Call .WriteByte(UserList(UserIndex).Raza)
        Call .WriteByte(UserList(UserIndex).Genero)
        
        If esRene(UserIndex) Then
            .WriteByte 1
        ElseIf esArmada(UserIndex) Then
            .WriteByte 2
        ElseIf esCaos(UserIndex) Then
            .WriteByte 3
        ElseIf esMili(UserIndex) Then
            .WriteByte 4
        ElseIf esRepu(UserIndex) Then
            .WriteByte 5
        ElseIf esCiuda(UserIndex) Then
            .WriteByte 6
        End If
        
        Call .WriteInteger(UserList(UserIndex).Stats.SkillPts)
        
        Call .WriteByte(UserList(UserIndex).masc.TieneFamiliar)
        If UserList(UserIndex).masc.TieneFamiliar = 1 Then
            .WriteASCIIString UserList(UserIndex).masc.Nombre
            
            .WriteByte UserList(UserIndex).masc.ELV
            .WriteLong UserList(UserIndex).masc.ELU
            .WriteLong UserList(UserIndex).masc.Exp
            
            .WriteInteger UserList(UserIndex).masc.minhp
            .WriteInteger UserList(UserIndex).masc.maxhp
            
            .WriteInteger UserList(UserIndex).masc.minhit
            .WriteInteger UserList(UserIndex).masc.maxhit
            
            .WriteByte UserList(UserIndex).masc.tipo
        End If
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub



''
' Writes the "AddForumMsg" message to the given user's outgoing data buffer.
'
' @param    title The title of the message to display.
' @param    message The message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAddForumMsg(ByVal UserIndex As Integer, ByVal title As String, ByVal message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AddForumMsg" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AddForumMsg)
        Call .WriteASCIIString(title)
        Call .WriteASCIIString(message)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowForumForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowForumForm" message to the given user's outgoing data buffer
'Add Marius: Los muerto no abren correo
'***************************************************
On Error GoTo Errhandler
    If UserList(UserIndex).flags.Muerto = 0 Then
        Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowForumForm)
    End If
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub




Public Sub WriteAddCorreoMsg(ByVal UserIndex As Integer, ByVal cIndex As Byte)
'***************************************************
'Author: Jose Castelli
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AddCorreoMsg)
        Call .WriteByte(cIndex)
        Call .WriteASCIIString(UserList(UserIndex).Correos(cIndex).Mensaje)
        Call .WriteASCIIString(UserList(UserIndex).Correos(cIndex).De)
        Call .WriteInteger(UserList(UserIndex).Correos(cIndex).Item)
        Call .WriteInteger(UserList(UserIndex).Correos(cIndex).Cantidad)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


Public Sub WriteShowCorreoForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Jose Castelli
'***************************************************
On Error GoTo Errhandler
    If UserList(UserIndex).flags.Muerto = 0 Then
        Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.showCorreoForm)
    End If
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub





''
' Writes the "SetInvisible" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal Invisible As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetInvisible" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetInvisible(CharIndex, Invisible))
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditateToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MeditateToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.MeditateToggle)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlindNoMore" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlindNoMore)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DumbNoMore" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.DumbNoMore)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SendSkills" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SendSkills)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(UserList(UserIndex).Stats.UserSkills(i))
        Next i
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainerCreatureList(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TrainerCreatureList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim str As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.TrainerCreatureList)
        
        For i = 1 To Npclist(NpcIndex).nrocriaturas
            str = str & Npclist(NpcIndex).Criaturas(i).NpcName & SEPARATOR
        Next i
        
        If LenB(str) > 0 Then _
            str = Left$(str, Len(str) - 1)
        
        Call .WriteASCIIString(str)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

 
''
' Writes the "CharacterInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    charName The requested char's name.
' @param    race The requested char's race.
' @param    class The requested char's class.
' @param    gender The requested char's gender.
' @param    level The requested char's level.
' @param    gold The requested char's gold.
' @param    reputation The requested char's reputation.
' @param    previousPetitions The requested char's previous petitions to enter guilds.
' @param    currentGuild The requested char's current guild.
' @param    previousGuilds The requested char's previous guilds.
' @param    RoyalArmy True if tha char belongs to the Royal Army.
' @param    CaosLegion True if tha char belongs to the Caos Legion.
' @param    citicensKilled The number of citicens killed by the requested char.
' @param    criminalsKilled The number of criminals killed by the requested char.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterInfo(ByVal UserIndex As Integer, ByVal charName As String, ByVal race As eRaza, ByVal Class As eClass, _
                            ByVal gender As eGenero, ByVal level As Byte, ByVal gold As Long, ByVal bank As Long, _
                            ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As String, _
                            ByVal realMatados As Integer, ByVal MiliMatados As Integer, ByVal CaosMatados As Integer, _
                            ByVal CiudMatados As Integer, ByVal RepuMatados As Integer, ByVal ReneMatados As Integer, _
                            ByVal Real As Byte, ByVal Mili As Byte, ByVal Caos As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterInfo" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CharacterInfo)
        
        Call .WriteASCIIString(charName)
        Call .WriteByte(race)
        Call .WriteByte(Class)
        Call .WriteByte(gender)
        
        Call .WriteByte(level)
        Call .WriteLong(gold)
        Call .WriteLong(bank)

        Call .WriteASCIIString(previousPetitions)
        Call .WriteASCIIString(currentGuild)
        Call .WriteASCIIString(previousGuilds)
        
        If Real = 1 Then
            .WriteByte 1
        ElseIf Mili = 1 Then
            .WriteByte 2
        ElseIf Caos = 1 Then
            .WriteByte 3
        Else
            .WriteByte 0
        End If
        
        .WriteInteger realMatados
        .WriteInteger MiliMatados
        .WriteInteger CaosMatados
        .WriteInteger CiudMatados
        .WriteInteger RepuMatados
        .WriteInteger ReneMatados
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @param    guildNews The guild's news.
' @param    joinRequests The list of chars which requested to join the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeaderInfo(ByVal UserIndex As Integer, ByRef guildList() As String, ByRef MemberList() As String, _
                            ByVal guildNews As String, ByRef joinRequests() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildLeaderInfo)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Prepare guild member's list
        Tmp = vbNullString
        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Store guild news
        Call .WriteASCIIString(guildNews)
        
        ' Prepare the join request's list
        Tmp = vbNullString
        For i = LBound(joinRequests()) To UBound(joinRequests())
            Tmp = Tmp & joinRequests(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "GuildDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildName The requested guild's name.
' @param    founder The requested guild's founder.
' @param    foundationDate The requested guild's foundation date.
' @param    leader The requested guild's current leader.
' @param    URL The requested guild's website.
' @param    memberCount The requested guild's member count.
' @param    electionsOpen True if the clan is electing it's new leader.
' @param    alignment The requested guild's alignment.
' @param    enemiesCount The requested guild's enemy count.
' @param    alliesCount The requested guild's ally count.
' @param    antifactionPoints The requested guild's number of antifaction acts commited.
' @param    codex The requested guild's codex.
' @param    guildDesc The requested guild's description.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDetails(ByVal UserIndex As Integer, ByVal GuildName As String, ByVal founder As String, ByVal foundationDate As String, _
                            ByVal leader As String, ByVal URL As String, ByVal memberCount As Integer, _
                            ByVal alignment As String, ByVal antifactionPoints As String, ByRef codex As String, ByVal guildDesc As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildDetails" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildDetails)
        
        Call .WriteASCIIString(GuildName)
        Call .WriteASCIIString(founder)
        Call .WriteASCIIString(foundationDate)
        Call .WriteASCIIString(leader)
        Call .WriteASCIIString(URL)
        
        Call .WriteInteger(memberCount)
        Call .WriteASCIIString(alignment)
        Call .WriteASCIIString(antifactionPoints)
        
        Call .WriteASCIIString(codex)
        
        Call .WriteASCIIString(guildDesc)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildFundationForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGuildFundationForm)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal UserIndex As Integer, Optional ByVal posUpdate As Boolean = True)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/12/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Writes the "ParalizeOK" message to the given user's outgoing data buffer
'And updates user position
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ParalizeOK)
    If posUpdate Then Call WritePosUpdate(UserIndex)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowUserRequest)
        
        Call .WriteASCIIString(details)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "TradeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTradeOK(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TradeOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.TradeOK)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BankOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankOK(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankOK)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeUserTradeSlot)
        
        Call .WriteInteger(ObjIndex)
        Call .WriteLong(Amount)
        Call .WriteInteger(ObjData(ObjIndex).GrhIndex)
        Call .WriteByte(ObjData(ObjIndex).OBJType)
        Call .WriteInteger(ObjData(ObjIndex).maxhit)
        Call .WriteInteger(ObjData(ObjIndex).minhit)
        Call .WriteInteger(ObjData(ObjIndex).def)
        Call .WriteLong(SalePrice(ObjData(ObjIndex).valor))
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


''
' Writes the "SpawnList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByRef npcNames() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SpawnList)
        
        For i = LBound(npcNames()) To UBound(npcNames())
            Tmp = Tmp & npcNames(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSOSForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowSOSForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSOSForm)
        
        For i = 1 To Ayuda.Longitud
            Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
        Next i
        
        If LenB(Tmp) <> 0 Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGMPanelForm)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal UserIndex As Integer, ByRef userNamesList() As String, ByVal Cant As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06 NIGO:
'Writes the "UserNameList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserNameList)
        
        ' Prepare user's names list
        For i = 1 To Cant
            Tmp = Tmp & userNamesList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Pong" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Pong" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Pong)
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Sends all data existing in the buffer
'***************************************************
    Dim sndData As String
    
    With UserList(UserIndex).outgoingData
        If .length = 0 Then _
            Exit Sub
        
        sndData = .ReadASCIIStringFixed(.length)
        
        Call EnviarDatosASlot(UserIndex, sndData)
        DoEvents
    End With
End Sub

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, ByVal Invisible As Boolean) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "SetInvisible" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetInvisible)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(Invisible)
        
        PrepareMessageSetInvisible = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ChatOverHead" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageChatOverHead(ByVal chat As String, ByVal CharIndex As Integer, ByVal color As Long, Optional ByVal chata As Byte = 0) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ChatOverHead" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChatOverHead)
        Call .WriteASCIIString(chat)
        Call .WriteInteger(CharIndex)
        
        ' Write rgb channels and save one byte from long :D
        Call .WriteByte(color And &HFF)
        Call .WriteByte((color And &HFF00&) \ &H100&)
        Call .WriteByte((color And &HFF0000) \ &H10000)
        
        Call .WriteByte(chata)
        
        PrepareMessageChatOverHead = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageConsoleMsg(ByVal console As Byte, ByVal chat As String, ByVal FontIndex As FontTypeNames) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ConsoleMsg" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ConsoleMsg)
        Call .WriteASCIIString(chat)
        Call .WriteByte(console)
        Call .WriteByte(FontIndex)
        
        PrepareMessageConsoleMsg = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "CreateFX" message and returns it.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, ByVal fx As Integer, ByVal FXLoops As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CreateFX" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateFX)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(fx)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCreateFX = .ReadASCIIStringFixed(.length)
    End With
End Function
Public Function PrepareMessageCreateFXMap(ByVal x As Byte, ByVal Y As Byte, ByVal fx As Integer, ByVal FXLoops As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateFXMap)
        Call .WriteByte(x)
        Call .WriteByte(Y)
        Call .WriteInteger(fx)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCreateFXMap = .ReadASCIIStringFixed(.length)
    End With
End Function
''
' Prepares the "PlayWave" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayWave(ByVal wave As Integer, ByVal x As Byte, ByVal Y As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/08/07
'Last Modified by: Rapsodius
'Added X and Y positions for 3D Sounds
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayWave)
        Call .WriteInteger(wave)
        Call .WriteByte(x)
        Call .WriteByte(Y)
        
        PrepareMessagePlayWave = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageGuildChat(ByVal chat As String) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "GuildChat" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.GuildChat)
        Call .WriteASCIIString(chat)
        
        PrepareMessageGuildChat = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageShowMessageBox(ByVal chat As String) As String
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "ShowMessageBox" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(chat)
        
        PrepareMessageShowMessageBox = .ReadASCIIStringFixed(.length)
    End With
End Function


''
' Prepares the "PlayMidi" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMidi(ByVal midi As Byte, Optional ByVal loops As Integer = -1) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "GuildChat" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayMidi)
        Call .WriteByte(midi)
        Call .WriteInteger(loops)
        
        PrepareMessagePlayMidi = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "PauseToggle" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PauseToggle)
        PrepareMessagePauseToggle = .ReadASCIIStringFixed(.length)
    End With
End Function


''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal x As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ObjectDelete" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectDelete)
        Call .WriteByte(x)
        Call .WriteByte(Y)
        
        PrepareMessageObjectDelete = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageBlockPosition(ByVal x As Byte, ByVal Y As Byte, ByVal Blocked As Boolean) As String
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "BlockPosition" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(x)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)
        
        PrepareMessageBlockPosition = .ReadASCIIStringFixed(.length)
    End With
    
End Function

''
' Prepares the "ObjectCreate" message and returns it.
'
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectCreate(ByVal x As Byte, ByVal Y As Byte, ByVal ObjIndex As Integer, ByVal Amount As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'prepares the "ObjectCreate" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectCreate)
        Call .WriteByte(x)
        Call .WriteByte(Y)
        Call .WriteInteger(ObjIndex)
        Call .WriteByte(ObjData(ObjIndex).OBJType)
        Call .WriteInteger(Amount)
        
        PrepareMessageObjectCreate = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterRemove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterRemove)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageCharacterRemove = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RemoveCharDialog)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageRemoveCharDialog = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterCreate(ByVal body As Integer, ByVal head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal x As Byte, ByVal Y As Byte, ByVal Weapon As Integer, ByVal Shield As Integer, _
                                ByVal fx As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal name As String, _
                                ByVal privileges As Byte, ByVal account As Boolean, ByVal bandera As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterCreate" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterCreate)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(head)
        Call .WriteByte(heading)
        Call .WriteByte(x)
        Call .WriteByte(Y)
        Call .WriteInteger(Weapon)
        Call .WriteInteger(Shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(fx)
        Call .WriteInteger(FXLoops)
        Call .WriteASCIIString(name)
        Call .WriteByte(privileges)
        Call .WriteBoolean(account)
        Call .WriteByte(bandera)
        
        PrepareMessageCharacterCreate = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "CharacterChange" message and returns it.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterChange(ByVal body As Integer, ByVal head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal Weapon As Integer, ByVal Shield As Integer, _
                                ByVal fx As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterChange" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChange)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(head)
        Call .WriteByte(heading)
        Call .WriteInteger(Weapon)
        Call .WriteInteger(Shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(fx)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCharacterChange = .ReadASCIIStringFixed(.length)
    End With
End Function
Public Function PrepareMessageCharStatus(ByVal CharIndex As Integer, ByVal priv As Byte) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharStatus)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(priv)
        
        PrepareMessageCharStatus = .ReadASCIIStringFixed(.length)
    End With
End Function
''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, ByVal x As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterMove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterMove)
        Call .WriteInteger(CharIndex Xor 12459)
        Call .WriteByte(x)
        Call .WriteByte(Y)
        
        PrepareMessageCharacterMove = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading) As String
'***************************************************
'Author: ZaMa
'Last Modification: 26/03/2009
'Prepares the "ForceCharMove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ForceCharMove)
        Call .WriteByte(Direccion)
        
        PrepareMessageForceCharMove = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, tipeuser As Byte, Tag As String) As String
'***************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modification: 04/07/07
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'Prepares the "UpdateTagAndStatus" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.UpdateTagAndStatus)
        
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
        Call .WriteByte(tipeuser)
        Call .WriteASCIIString(Tag)
        
        PrepareMessageUpdateTagAndStatus = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal message As String) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ErrorMsg" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ErrorMsg)
        Call .WriteASCIIString(message)
        
        PrepareMessageErrorMsg = .ReadASCIIStringFixed(.length)
    End With
End Function


Public Function PrepareMessageCreateParticle(ByVal x As Integer, ByVal Y As Byte, ByVal Particle As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ParticleCreate)
        Call .WriteByte(x)
        Call .WriteByte(Y)
        Call .WriteInteger(Particle)

        PrepareMessageCreateParticle = .ReadASCIIStringFixed(.length)
    End With
End Function
Public Function PrepareMessageCreateCharParticle(ByVal CharIndex As Integer, ByVal Particle As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharParticleCreate)
        Call .WriteInteger(Particle)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageCreateCharParticle = .ReadASCIIStringFixed(.length)
    End With
End Function
Public Function PrepareMessageDestParticle(ByVal x As Integer, ByVal Y As Byte) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.DestParticle)
        Call .WriteByte(x)
        Call .WriteByte(Y)

        PrepareMessageDestParticle = .ReadASCIIStringFixed(.length)
    End With
End Function
Public Function PrepareMessageDestCharParticle(ByVal CharIndex As Integer, ByVal Particle As Integer) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.DestCharParticle)
        Call .WriteInteger(Particle)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageDestCharParticle = .ReadASCIIStringFixed(.length)
    End With
End Function
Public Sub HandleLoginNewAccount(ByVal UserIndex As Integer)
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    'Dim UserName        As String
    'Dim UserPassword    As String
    'Dim UserEmail       As String
    'Dim UserAnswer      As String
    'Dim UserQuestion    As Byte
    
    'UserName = buffer.ReadASCIIString()
    'UserPassword = buffer.ReadASCIIString()
    'UserEmail = buffer.ReadASCIIString()
    
    'Call SaveCuenta(UserIndex, UserName, UserPassword, UserEmail)
End Sub
Public Sub HandleLoginAccount(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim UserName        As String
    Dim UserPassword    As String
    Dim Version As Byte
    
    UserName = buffer.ReadASCIIString()
    
    UserPassword = buffer.ReadASCIIString()
    'Version = buffer.ReadByte()
    
    'If Version <> ULTIMAVERSION Then
    '    Call WriteMsg(UserIndex, 48)
    '    Call FlushBuffer(UserIndex)
    '    Call CloseSocket(UserIndex)
    '    Exit Sub
    'End If
    
    'If buffer.length > 0 Then
    '    Dim Pass As String
    '    Pass = buffer.ReadASCIIStringFixed(16)
   '
   '     If Pass = "pompompon sdeaer" Then
   '         UserList(UserIndex).Stats.PuedeStaff = 1
   '     Else
   '         UserList(UserIndex).Stats.PuedeStaff = 0
   '     End If
   ' Else
   '     UserList(UserIndex).Stats.PuedeStaff = 0
   ' End If

    Call ConectarCuenta(UserIndex, UserName, UserPassword)

    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
        
End Sub

Public Sub WriteShowAccount(ByVal UserIndex As Integer)
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowAccount)

Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


Public Sub WriteAgilidad(ByVal UserIndex As Integer)
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Agilidad)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(2))
    End With
End Sub
Public Sub WriteFuerza(ByVal UserIndex As Integer)
    With UserList(UserIndex).outgoingData

        Call .WriteByte(ServerPacketID.Fuerza)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(1))
    End With
End Sub
Public Function WriteCreateCharParticle(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal Particle As Integer, ByVal life As Integer) As String
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CharParticleCreate)
        Call .WriteInteger(Particle)
        Call .WriteInteger(life)
        Call .WriteInteger(CharIndex)
    End With
End Function
Sub WriteHora(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        .outgoingData.WriteASCIIStringFixed PrepareMessageHora
    End With
End Sub
Function PrepareMessageHora() As String
    With auxiliarBuffer
        .WriteByte ServerPacketID.hora
        
        .WriteByte tHora
        .WriteByte tMinuto
        
        PrepareMessageHora = .ReadASCIIStringFixed(.length)
    End With
End Function

Public Sub WriteAddPj(ByVal UserIndex As Integer, ByVal nameUser As String, ByVal index As Byte)
    Dim RS As New ADODB.Recordset
    Dim ipj As Integer
    Dim head As Integer, body As Integer, casco As Byte, Weapon As Byte, map As Integer
    Dim Shield As Byte, Nivel As Byte, Clase As Byte, color As Byte, tipPet As Byte

    Set RS = DB_Conn.Execute("SELECT (IndexPJ) FROM `charflags` WHERE Nombre='" & nameUser & "' LIMIT 1")
        If RS.BOF Or RS.EOF Then
            Exit Sub
        End If

        ipj = RS!Indexpj
    Set RS = Nothing
    
    Set RS = DB_Conn.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
        
        head = RS!head
        head = CInt(head)
       
        body = RS!body
        body = CInt(body)
        
        casco = RS!casco
        casco = CByte(casco)
        
        Weapon = RS!Arma
        Weapon = CByte(Weapon)
        
        Shield = RS!Escudo
        Shield = CByte(Shield)
    Set RS = Nothing
    
    Set RS = DB_Conn.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
              
        Nivel = RS!ELV
    Set RS = Nothing
    
    Set RS = DB_Conn.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub

        map = RS!mapa
        Clase = RS!Clase
    Set RS = Nothing
    
    Set RS = DB_Conn.Execute("SELECT * FROM `charmascotafami` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
        
        tipPet = RS!tipo
    Set RS = Nothing

    If EsDIOS(UserIndex) Then
        color = 5
    Else
        color = UserTypeColorAcc(ipj)
    End If
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AddPJ)
        Call .WriteByte(index)
        Call .WriteASCIIString(nameUser)
        Call .WriteInteger(head)
        Call .WriteInteger(body)
        Call .WriteByte(casco)
        Call .WriteByte(Weapon)
        Call .WriteByte(Shield)
        Call .WriteByte(Nivel)
        Call .WriteInteger(map)
        Call .WriteByte(Clase)
        Call .WriteByte(color)
        Call .WriteByte(tipPet)
    End With
    
    Exit Sub
     
Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub




Public Sub WriteTejiblesObjects(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CarpenterObjects" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim obj As ObjData
    Dim validIndexes() As Integer
    Dim count As Integer
    
    ReDim validIndexes(1 To UBound(ObjSastre()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SastreObjects)
        
        For i = 1 To UBound(ObjSastre())
   
            ' Can the user create this object? If so add it to the list....
            If ObjSastre(i) <> 0 Then

                If ObjData(ObjSastre(i)).SkSastreria <= UserList(UserIndex).Stats.UserSkills(eSkill.Sastreria) \ ModSastreria(UserList(UserIndex).Clase) Then
                    count = count + 1
                    validIndexes(count) = i
                End If
            End If
        Next i
        
  
        ' Write the number of objects in the list
        Call .WriteInteger(count)
        
        ' Write the needed data of each object
              
        For i = 1 To count

            obj = ObjData(ObjSastre(validIndexes(i)))
            Call .WriteASCIIString(obj.name)
            Call .WriteInteger(obj.PielLobo)
            Call .WriteInteger(obj.PielOso)
            Call .WriteInteger(obj.PielLoboInvernal)
            Call .WriteInteger(ObjSastre(validIndexes(i)))
        Next i
        
    End With
Exit Sub

Errhandler:
    If err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub
Public Sub WriteGrupo(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        .outgoingData.WriteByte ServerPacketID.Grupo
        
        .outgoingData.WriteInteger UserList(UserIndex).GrupoIndex
        .outgoingData.WriteByte EsLider(UserIndex)
    End With
End Sub
Public Sub HandleRequestGrupo(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        .incomingData.ReadByte
        
        If UserList(UserIndex).GrupoIndex = 0 Then
            WriteGrupo UserIndex
            Exit Sub
        End If
        
        Call WriteGrupoForm(UserIndex)
    End With
End Sub
Private Sub WriteGrupoForm(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If .GrupoIndex = 0 Then Exit Sub
        
        Call .outgoingData.WriteByte(ServerPacketID.ShowGrupoForm)
        
        Dim NumMembers As Byte, i As Long
        NumMembers = mdGrupo.CantMiembros(UserIndex)
        
        .outgoingData.WriteByte NumMembers
        
        For i = 1 To NumMembers
            .outgoingData.WriteASCIIString mdGrupo.NombreMiembro(UserIndex, i)
        Next i
        
    End With
End Sub
Private Sub HandleDuelo(ByVal UserIndex As Integer)
    'Mannakia
    With UserList(UserIndex)
        .incomingData.ReadByte
        
        Dim opt As Byte
        Dim tU As Integer
        
        opt = .incomingData.ReadByte
        tU = .flags.TargetUser
        
        If .flags.inDuelo <> 0 Then Exit Sub
        
        If .Pos.map = Prision.map Then
            Call WriteConsoleMsg(1, UserIndex, "No retar, estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MapInfo(.Pos.map).Pk = True Then
            Call WriteConsoleMsg(1, UserIndex, "No puedes retar en zona insegura.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If opt = 1 Then 'Esta retando
            If tU > 0 Then
                If UserList(tU).flags.inDuelo = 1 Then
                    WriteConsoleMsg 1, UserIndex, "El personaje ya esta en duelo!!", FontTypeNames.FONTTYPE_BROWNI
                    Exit Sub
                End If
                
                If UserList(tU).flags.ModoCombate = True Then
                    WriteConsoleMsg 1, UserIndex, "Tu objetivo debe estar sin modo combate.", FontTypeNames.FONTTYPE_BROWNI
                    Exit Sub
                End If
                
                UserList(tU).flags.solDuelo = UserIndex
                WriteShowMessageBox tU, .name, 1, 2
            Else
                WriteConsoleMsg 1, UserIndex, "Haz click en el personaje.", FontTypeNames.FONTTYPE_BROWNI
                Exit Sub
            End If
        ElseIf opt = 3 Then
            tU = .flags.solDuelo
            If tU = 0 Then Exit Sub 'F L A S H A S T E
            
            WriteConsoleMsg 1, tU, .name & " ha rechazado el duelo.", FontTypeNames.FONTTYPE_BROWNI
            
            .flags.inDuelo = 0
            .flags.vicDuelo = 0
            .flags.solDuelo = 0
            
            UserList(tU).flags.inDuelo = 0
            UserList(tU).flags.vicDuelo = 0
            UserList(tU).flags.solDuelo = 0
        Else 'Lo acepta :O
            tU = .flags.solDuelo
            If tU = 0 Then Exit Sub 'F L A S H A S T E
            
            .flags.inDuelo = 1
            .flags.vicDuelo = tU
            
            UserList(tU).flags.inDuelo = 1
            UserList(tU).flags.vicDuelo = UserIndex
            
            .flags.solDuelo = 0
            UserList(tU).flags.solDuelo = 1
            
            Call entra_lista_espera(UserIndex, duelo_espera)
            Call entra_lista_espera(tU, duelo_espera)
            
            UserList(UserIndex).evento = 7
            Call entrar_arena(duelo_espera) 'Vemos si hay alguien mas en espera y si hay lo mandamos adentro de la arena
        End If
    End With
End Sub
Private Sub HandleMessagesGM(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        .incomingData.ReadByte
        
        Select Case .incomingData.PeekByte
            'GM messages
            Case gMessages.GMMessage               '/GMSG
                Call HandleGMMessage(UserIndex)     'Conse
            
            Case gMessages.GuildMemberList         '/MIEMBROSCLAN
                Call HandleGuildMemberList(UserIndex) 'Dios
            
            Case gMessages.showName                '/SHOWNAME
                Call HandleShowName(UserIndex)      'Dios
            
            Case gMessages.OnlineArmada            '/ONLINEARMY
                Call HandleOnlineArmada(UserIndex) 'Faccion
            
            Case gMessages.OnlineCaos               '/ONLINECAOS
                Call HandleOnlineCaos(UserIndex)    'Faccion
            
            Case gMessages.OnlineMilicia           '/ONLINEMILI
                Call HandleOnlineMilicia(UserIndex) 'Faccion
            
            Case gMessages.GoNearby                '/IRCERCA
                Call HandleGoNearby(UserIndex)      'Semi
            
            Case gMessages.comment                 '/REM
                Call HandleComment(UserIndex)       'Conse
            
            Case gMessages.serverTime              '/HORA
                Call HandleServerTime(UserIndex)    'Conse
            
            Case gMessages.Where                   '/DONDE
                Call HandleWhere(UserIndex)         'Conse
            
            Case gMessages.CreaturesInMap          '/NENE
                Call HandleCreaturesInMap(UserIndex)    'Semi
            
            Case gMessages.WarpMeToTarget          '/TELEPLOC
                Call HandleWarpMeToTarget(UserIndex)    'Faccion
            
            Case gMessages.WarpChar                '/TELEP
                Call HandleWarpChar(UserIndex)      'facc solo YO y Semi cualquiera
            
            Case gMessages.Silence                 '/SILENCIAR
                Call HandleSilence(UserIndex)       'Conse
            
            Case gMessages.SOSShowList             '/SHOW SOS
                Call HandleSOSShowList(UserIndex)   'Semi
            
            Case gMessages.SOSRemove               'SOSDONE
                Call HandleSOSRemove(UserIndex)     'Semi
            
            Case gMessages.GoToChar                '/IRA
                Call HandleGoToChar(UserIndex)      'Faccion
            
            Case gMessages.Invisible               '/INVISIBLE
                Call HandleInvisible(UserIndex)     'Faccion
            
            Case gMessages.GMPanel                 '/PANELGM
                Call HandleGMPanel(UserIndex)       'Faccion
            
            Case gMessages.RequestUserList         'LISTUSU
                Call HandleRequestUserList(UserIndex)   'faccion
            
            Case gMessages.Working                 '/TRABAJANDO
                Call HandleWorking(UserIndex)       'Conse
            
            Case gMessages.Hiding                  '/OCULTANDO
                Call HandleHiding(UserIndex)        'Conse
            
            Case gMessages.Jail                    '/CARCEL
                Call HandleJail(UserIndex)          'Semi
            
            Case gMessages.KillNPC                 '/RMATA
                Call HandleKillNPC(UserIndex)       'Semi
            
            Case gMessages.WarnUser                '/ADVERTENCIA
                Call HandleWarnUser(UserIndex)      'Conse
                
            Case gMessages.SlashSlash                '//
                Call HandleSlashSlash(UserIndex)      'Conse
            
            Case gMessages.EditChar                '/MOD
                Call HandleEditChar(UserIndex)      'Admin
    
            Case gMessages.ReviveChar              '/REVIVIR
                Call HandleReviveChar(UserIndex)    'Semi
            
            Case gMessages.OnlineGM                '/ONLINEGM
                Call HandleOnlineGM(UserIndex)      'Conse
            
            Case gMessages.OnlineMap               '/ONLINEMAP
                Call HandleOnlineMap(UserIndex)     'Conse
                
            Case gMessages.Kick                    '/ECHAR
                Call HandleKick(UserIndex)          'Conse
            
            Case gMessages.Pejotas                 '/PEJOTAS
                Call HandlePejotas(UserIndex)      'Semi
                
            Case gMessages.Execute                 '/EJECUTAR
                Call HandleExecute(UserIndex)       'Dios
                
            Case gMessages.BanChar                 '/BAN
                Call HandleBanChar(UserIndex)       'Dios
                
            Case gMessages.UnbanChar               '/UNBAN
                Call HandleUnbanChar(UserIndex)     'Dios
                
            Case gMessages.NPCFollow               '/SEGUIR
                Call HandleNPCFollow(UserIndex)     'Semi
                
            Case gMessages.SummonChar              '/SUM
                Call HandleSummonChar(UserIndex)    'Semi
                
            Case gMessages.SpawnListRequest        '/CC
                Call HandleSpawnListRequest(UserIndex)  'Semi
      
                
            Case gMessages.SpawnCreature           'SPA
                Call HandleSpawnCreature(UserIndex) 'Semi
                
            Case gMessages.ResetNPCInventory       '/RESETINV
                Call HandleResetNPCInventory(UserIndex) 'Dios
                
            Case gMessages.CleanWorld              '/LIMPIAR
                Call HandleCleanWorld(UserIndex)    'Semi
                
            Case gMessages.ServerMessage           '/RMSG
                Call HandleServerMessage(UserIndex) 'Semi
                
            Case gMessages.NickToIP                '/NICK2IP
                Call HandleNickToIP(UserIndex)      'Dios
            
            Case gMessages.IPToNick                '/IP2NICK
                Call HandleIPToNick(UserIndex)      'Dios
                
            Case gMessages.GuildOnlineMembers      '/ONCLAN
                Call HandleGuildOnlineMembers(UserIndex) 'Semi
            
            Case gMessages.TeleportCreate          '/CT
                Call HandleTeleportCreate(UserIndex)    'Semi
                
            Case gMessages.TeleportDestroy         '/DT
                Call HandleTeleportDestroy(UserIndex) 'Semi
                
            Case gMessages.SetCharDescription      '/SETDESC
                Call HandleSetCharDescription(UserIndex) 'Semi
            
            Case gMessages.ForceMIDIToMap          '/FORCEMIDIMAP
                Call HanldeForceMIDIToMap(UserIndex) 'Dios
                
            Case gMessages.ForceWAVEToMap          '/FORCEWAVMAP
                Call HandleForceWAVEToMap(UserIndex) 'Dios
                
            Case gMessages.TalkAsNPC               '/TALKAS
                Call HandleTalkAsNPC(UserIndex)     'Semi
            
            Case gMessages.DestroyAllItemsInArea   '/MASSDEST
                Call HandleDestroyAllItemsInArea(UserIndex) 'Conse
                
    
            Case gMessages.ItemsInTheFloor         '/PISO
                Call HandleItemsInTheFloor(UserIndex) 'Semi
                
            Case gMessages.MakeDumb                '/ESTUPIDO
                Call HandleMakeDumb(UserIndex)      'Semi
                
            Case gMessages.MakeDumbNoMore          '/NOESTUPIDO
                Call HandleMakeDumbNoMore(UserIndex) 'Semi
                
            Case gMessages.DumpIPTables            '/DUMPSECURITY"
                Call HandleDumpIPTables(UserIndex) 'Admin
    
            Case gMessages.SetTrigger              '/TRIGGER
                Call HandleSetTrigger(UserIndex)    'Admin
            
            Case gMessages.AskTrigger               '/TRIGGER
                Call HandleAskTrigger(UserIndex)    'Admin
                
            Case gMessages.BannedIPList            '/BANIPLIST
                Call HandleBannedIPList(UserIndex)  'Dios
            
            Case gMessages.BannedIPReload          '/BANIPRELOAD
                Call HandleBannedIPReload(UserIndex) 'Dios
            
            Case gMessages.GuildBan                '/BANCLAN
                Call HandleGuildBan(UserIndex)      'Dios
            
            Case gMessages.BanIP                   '/BANIP
                Call HandleBanIP(UserIndex)         'Dios
            
            Case gMessages.UnbanIP                 '/UNBANIP
                Call HandleUnbanIP(UserIndex)       'Dios
            
            Case gMessages.CreateItem              '/CI
                Call HandleCreateItem(UserIndex)    'Admin
            
            Case gMessages.DestroyItems            '/DEST
                Call HandleDestroyItems(UserIndex)  'Conse
            
            Case gMessages.ChaosLegionKick         '/NOCAOS
                Call HandleChaosLegionKick(UserIndex) 'Faccion o Semi
            
            Case gMessages.RoyalArmyKick           '/NOREAL
                Call HandleRoyalArmyKick(UserIndex) 'Faccion o Semi
            
            Case gMessages.MiliciaKick           '/NOREAL
                Call HandleMiliciaKick(UserIndex) 'Faccion o Semi
            
            Case gMessages.ForceMIDIAll            '/FORCEMIDI
                Call HandleForceMIDIAll(UserIndex)  'Dios
            
            Case gMessages.ForceWAVEAll            '/FORCEWAV
                Call HandleForceWAVEAll(UserIndex) 'Dios
            
            Case gMessages.TileBlockedToggle       '/BLOQ
                Call HandleTileBlockedToggle(UserIndex)
            
            Case gMessages.KillNPCNoRespawn        '/MATA
                Call HandleKillNPCNoRespawn(UserIndex) 'Dios
            
            Case gMessages.KillAllNearbyNPCs       '/MASSKILL
                Call HandleKillAllNearbyNPCs(UserIndex) 'Dios
            
            Case gMessages.LastIP                  '/LASTIP
                Call HandleLastIP(UserIndex)        'Dios
            
            Case gMessages.SystemMessage           '/SMSG
                Call HandleSystemMessage(UserIndex) 'Admin
            
            Case gMessages.CreateNPC               '/ACC
                Call HandleCreateNPC(UserIndex) 'Semi
            
            Case gMessages.CreateNPCWithRespawn    '/RACC
                Call HandleCreateNPCWithRespawn(UserIndex) 'Dios
                
            Case gMessages.NavigateToggle          '/NAVE
                Call HandleNavigateToggle(UserIndex) 'Semi
            
            Case gMessages.ServerOpenToUsersToggle '/HABILITAR
                Call HandleServerOpenToUsersToggle(UserIndex) 'Admin
            
            Case gMessages.TurnOffServer           '/APAGAR
                Call HandleTurnOffServer(UserIndex) 'Admin
            
            Case gMessages.TurnCriminal            '/CONDEN
                Call HandleTurnCriminal(UserIndex) 'Dios
            
            Case gMessages.ResetFactions           '/RAJAR
                Call HandleResetFactions(UserIndex) 'Dios
            
            Case gMessages.RemoveCharFromGuild     '/RAJARCLAN
                Call HandleRemoveCharFromGuild(UserIndex) 'Dios
            
            Case gMessages.ToggleCentinelActivated '/CENTINELAACTIVADO
                Call HandleToggleCentinelActivated(UserIndex) 'Dios
            
            Case gMessages.DoBackUp                '/DOBACKUP
                Call HandleDoBackUp(UserIndex)  'Admin
            
            Case gMessages.ShowGuildMessages       '/SHOWCMSG
                Call HandleShowGuildMessages(UserIndex) 'Dios
            
            Case gMessages.SaveMap                 '/GUARDAMAPA
                Call HandleSaveMap(UserIndex)       'Dios
            
            Case gMessages.ChangeMapInfoPK         '/MODMAPINFO PK
                Call HandleChangeMapInfoPK(UserIndex) 'Dios
            
            Case gMessages.ChangeMapInfoBackup     '/MODMAPINFO BACKUP
                Call HandleChangeMapInfoBackup(UserIndex) 'Dios
        
            Case gMessages.ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
                Call HandleChangeMapInfoRestricted(UserIndex) 'Dios
                
            Case gMessages.ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
                Call HandleChangeMapInfoNoMagic(UserIndex) 'Dios
                
            Case gMessages.ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
                Call HandleChangeMapInfoNoInvi(UserIndex) 'Dios
                
            Case gMessages.ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
                Call HandleChangeMapInfoNoResu(UserIndex) 'Dios
                
            Case gMessages.ChangeMapInfoLand       '/MODMAPINFO TERRENO
                Call HandleChangeMapInfoLand(UserIndex) 'Dios
                
            Case gMessages.ChangeMapInfoZone       '/MODMAPINFO ZONA
                Call HandleChangeMapInfoZone(UserIndex) 'Dios
            
            Case gMessages.SaveChars               '/GRABAR
                Call HandleSaveChars(UserIndex)     'Dios
            
            Case gMessages.CleanSOS                '/BORRAR SOS
                Call HandleCleanSOS(UserIndex)      'Dios
    
            Case gMessages.KickAllChars            '/ECHARTODOSPJS
                Call HandleKickAllChars(UserIndex)  'Admin
            
            Case gMessages.ReloadNPCs              '/RELOADNPCS
                Call HandleReloadNPCs(UserIndex)    'Dios
            
            Case gMessages.ReloadServerIni         '/RELOADSINI
                Call HandleReloadServerIni(UserIndex) 'Dios
            
            Case gMessages.ReloadSpells            '/RELOADHECHIZOS
                Call HandleReloadSpells(UserIndex)  'Dios
            
            Case gMessages.ReloadObjects           '/RELOADOBJ
                Call HandleReloadObjects(UserIndex) 'Dios
            
            Case gMessages.Restart                 '/REINICIAR
                Call HandleRestart(UserIndex)       'Admin
            
            Case gMessages.Ignored                 '/IGNORADO
                Call HandleIgnored(UserIndex)       'Dios
            
            Case gMessages.CheckSlot               '/SLOT
                Call HandleCheckSlot(UserIndex)     'Dios
                
            Case gMessages.CrearTorneo             '/CREARTORNEO
                Call HandleCrearTorneo(UserIndex)   'Semi
                
            Case gMessages.CancelTorneo            '/CANCELTORNEO
                Call HandleCancelTorneo(UserIndex)  'Semi
            
        End Select
    End With
End Sub
Public Sub WriteMsg(ByVal UserIndex As Integer, ByVal msg As Byte, Optional ByVal arg1 As String = "", Optional arg2 As String = "")
    With UserList(UserIndex)
        .outgoingData.WriteByte ServerPacketID.Messages
        
        .outgoingData.WriteByte msg
        
        If msg = 21 Or msg = 36 Then
            .outgoingData.WriteLong CLng(val(arg1))
        ElseIf msg = 24 Or msg = 38 Or msg = 39 Or msg = 40 Or msg = 43 Or msg = 44 Then
            .outgoingData.WriteInteger CInt(val(arg1))
            .outgoingData.WriteInteger CInt(val(arg2))
        ElseIf msg = 25 Then
            .outgoingData.WriteInteger CInt(val(arg1))
        ElseIf msg = 27 Then
            .outgoingData.WriteByte CInt(val(arg1))
        End If
    End With
End Sub

Public Sub HandleExtractItem(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.length < 2 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
    End If
    
On Error GoTo err
    With UserList(UserIndex)
        Dim index As Byte
        Dim elItem As obj
        Dim miPos As WorldPos
        
        .incomingData.ReadByte
        index = .incomingData.ReadByte
        
        If index > 0 And index < (MENSAJES_TOPE_CORREO + 1) Then
            If .Correos(index).Item <> 0 Then
                elItem.ObjIndex = .Correos(index).Item
                elItem.Amount = .Correos(index).Cantidad
                
                
                If Not MeterItemEnInventario(UserIndex, elItem) Then
                    miPos = .Pos
                    TirarItemAlPiso miPos, elItem
                End If
                
                .Correos(index).De = ""
                .Correos(index).Item = 0
                .Correos(index).Mensaje = ""
                .Correos(index).Cantidad = 0
                
                Call Quitarcorreosql(.Correos(index).idmsj)
                WriteAddCorreoMsg UserIndex, index
                
                WriteMsg UserIndex, 10
            End If
        End If
    End With
    
    'Add Marius Cada vez que extrae algo del correo se guarda el pj, Por si copro algo en el shop
    'que no se lo cage el rollback. =)
    Call SaveUserSQL(UserIndex)
    '\Add
    
    Exit Sub
err:
    LogError "Error en HandleExtracItem"
End Sub
Public Sub HandleBorrarMensaje(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.length < 2 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
    End If
    
On Error GoTo err
    With UserList(UserIndex)
        Dim index As Byte
        
        .incomingData.ReadByte
        index = .incomingData.ReadByte
    
        If index > 0 And index < (MENSAJES_TOPE_CORREO + 1) Then
            If .Correos(index).De <> "" Then
                .Correos(index).De = ""
                .Correos(index).Item = 0
                .Correos(index).Mensaje = ""
                .Correos(index).Cantidad = 0
                
                Call Quitarcorreosql(.Correos(index).idmsj)
                WriteAddCorreoMsg UserIndex, index
                
                WriteMsg UserIndex, 9
            End If
        End If
    End With
    
    Exit Sub
err:
    LogError "Error en HandleBorrarMensaje"
End Sub
Public Sub HandleEnviarMensaje(ByVal UserIndex As Integer)
'Fix Marius le pusimos el buffer auxiliar
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    Debug.Print ">Mensaje"
    With UserList(UserIndex)
        Dim index As Byte
        Dim Mensaje As String
        Dim Slot As Byte
        Dim para As String
        Dim tI As Integer
        Dim Cantidad As Integer
        Dim salir As Boolean
        salir = False
        
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Call buffer.ReadByte    'acá borraremos de la memoria el byte identificador.
        
        Slot = buffer.ReadByte
        Debug.Print ">slot: " & Slot
        Mensaje = buffer.ReadASCIIString
        Debug.Print ">msg: " & Mensaje
        para = buffer.ReadASCIIString
        Debug.Print ">para: " & para
        Cantidad = buffer.ReadInteger
        Debug.Print ">cant: " & Cantidad
        
        If Len(Mensaje) = 0 Then Mensaje = " "
        
        If Cantidad < 1 And Slot > 0 And Slot < MAX_INVENTORY_SLOTS Then
            WriteMsg UserIndex, 13
            salir = True
        Else
            If Slot > 0 And Slot <= MAX_INVENTORY_SLOTS Then
                If .Invent.Object(Slot).Amount < Cantidad Then
                    WriteMsg UserIndex, 13
                    salir = True
                End If
            End If
        End If
        Debug.Print ">Pasó la priemra verificacion"
        
        'Add Marius No se pueden mandar cosas faccionarias por correo xD y tampoco cosas de shop >= 50
        If salir = False And Cantidad > 0 And Slot > 0 And Slot <= MAX_INVENTORY_SLOTS And _
           ObjData(.Invent.Object(Slot).ObjIndex).Caos = 1 Or _
           ObjData(.Invent.Object(Slot).ObjIndex).Real = 1 Or _
           ObjData(.Invent.Object(Slot).ObjIndex).Milicia = 1 Or _
           ObjData(.Invent.Object(Slot).ObjIndex).Shop >= 49 Then
        
                Call WriteConsoleMsg(1, UserIndex, "No puedes enviar objetos Faccionarios o Vips por correo.", FontTypeNames.FONTTYPE_INFO)
                salir = True
        
        End If
        Debug.Print ">Pasó la segunda verificacion"
        '\Add
        
        If salir = False Then
            tI = NameIndex(para)
            If tI > 0 Then 'Esta Online
                If EntregarMsgOn(UserIndex, tI, Mensaje, Slot, Cantidad) Then
                    If Slot > 0 And Slot < MAX_INVENTORY_SLOTS + 1 Then
                        QuitarUserInvItem UserIndex, Slot, Cantidad
                        UpdateUserInv False, UserIndex, Slot
                    End If
                    
                    WriteMsg UserIndex, 11
                Else
                    WriteMsg UserIndex, 41
                End If
            Else
                If ExistePersonaje(para) Then 'Esta Offline
                    If EntregarMsgOff(UserIndex, para, Mensaje, Slot, Cantidad) Then
                        If Slot > 0 And Slot < MAX_INVENTORY_SLOTS + 1 Then
                            QuitarUserInvItem UserIndex, Slot, Cantidad
                            UpdateUserInv False, UserIndex, Slot
                        End If
                       
                        WriteMsg UserIndex, 11
                    Else
                        WriteMsg UserIndex, 41
                    End If
                Else 'FLASHATE
                    WriteMsg UserIndex, 8
                End If
            End If
            
            'Add Marius Cada vez que mete algo que se guarde el pj para que no dupliquen con rollback
            Call SaveUserSQL(UserIndex)
            '\Add
        End If
        Debug.Print ">Pasó la tercera verificacion"
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub
Sub WriteShowFamiliarForm(ByVal UserIndex As Integer)
    UserList(UserIndex).outgoingData.WriteByte ServerPacketID.ShowFamiliarForm
End Sub
Sub HandleAdoptarMascota(ByVal UserIndex As Integer)
'Fix Marius le pusimos el buffer auxiliar
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler

    With UserList(UserIndex)
        
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
    
        Call buffer.ReadByte    'acá borraremos de la memoria el byte identificador.
        
        Dim tipe As eMascota
        Dim name As String
        Dim ii As Long
        
        tipe = buffer.ReadByte
        name = buffer.ReadASCIIString
        
        If .flags.TargetNPC > 0 Then
            If Npclist(.flags.TargetNPC).npctype = 11 Then
                If .Stats.UserSkills(eSkill.Domar) < 65 Then
                    WriteMsg UserIndex, 15
                Else
                    For ii = 1 To 35
                        If .Stats.UserHechizos(ii) = 0 Then
                            EntregarMascota UserIndex, tipe, name
                            .Stats.UserHechizos(ii) = 59
                            Call UpdateUserHechizos(False, UserIndex, CByte(ii))
                            Exit For
                        End If
                    Next
                End If
            Else
                WriteMsg UserIndex, 14
            End If
        Else
            WriteMsg UserIndex, 14
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub
Sub HandleDelClan(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        .incomingData.ReadByte
        
        If .GuildIndex = 0 Then
            WriteMsg UserIndex, 16
            Exit Sub
        End If
        
        If Not (UCase$(GuildLeader(.GuildIndex)) = UCase$(.name)) Then
            WriteMsg UserIndex, 17
            Exit Sub
        End If
        
        If PuedeBorrarClan <> 1 Then
            Call WriteConsoleMsg(1, UserIndex, "La eliminación de clanes esta deshabilitada momentaneamente. (No moleste a los GMs ellos no pueden borrar su clan)", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If
        
        'Siempre retorna false asi que dejamos comentado esto xD
        If GuildDelete(.GuildIndex) Then
            WriteMsg UserIndex, 18
        Else
            WriteMsg UserIndex, 19
        End If
        
        .GuildIndex = 0
        'Refrescamos el status
        Call RefreshCharStatus(UserIndex)
        
    End With
End Sub
Sub WriteCharMsgStatus(ByVal UserIndex As Integer, ByVal tI As Integer)
    With UserList(UserIndex)
        Dim St1 As Byte, St2 As Byte
        
        .outgoingData.WriteByte ServerPacketID.CharMsgStatus
        
        .outgoingData.WriteInteger UserList(tI).Char.CharIndex
        
        If UserList(tI).flags.Privilegios = PlayerType.Conse Then
            .outgoingData.WriteByte 10
        ElseIf UserList(tI).flags.Privilegios = PlayerType.Semi Then
            .outgoingData.WriteByte 11
        ElseIf UserList(tI).flags.Privilegios = PlayerType.Dios Then
            .outgoingData.WriteByte 12
        ElseIf UserList(tI).flags.Privilegios = PlayerType.Admin Then
            .outgoingData.WriteByte 13
        
        
        ElseIf UserList(tI).faccion.ArmadaReal = 1 Then
            .outgoingData.WriteByte 6
        ElseIf UserList(tI).faccion.Milicia = 1 Then
            .outgoingData.WriteByte 7
        ElseIf UserList(tI).faccion.FuerzasCaos = 1 Then
            .outgoingData.WriteByte 5
        ElseIf UserList(tI).faccion.Ciudadano = 1 Then
            .outgoingData.WriteByte 1
        ElseIf UserList(tI).faccion.Renegado = 1 Then
            .outgoingData.WriteByte 2
        ElseIf UserList(tI).faccion.Republicano = 1 Then
            .outgoingData.WriteByte 3
        Else
            .outgoingData.WriteByte 2
        End If
    
        .outgoingData.WriteLong CLng(((UserList(tI).Stats.minhp / 100) / (UserList(tI).Stats.maxhp / 100)) * 100)

        St1 = Generate_Char_Stat(tI)
        St2 = Generate_Char_StatEx(tI)
        
        .outgoingData.WriteByte St1
        .outgoingData.WriteByte St2
        
        If UserList(tI).flags.toyCasado = 1 Then
            .outgoingData.WriteByte Len(UserList(tI).flags.miPareja)
        Else
            .outgoingData.WriteByte 0
        End If

        .outgoingData.WriteByte UserList(tI).Clase
        
        'Add Marius Visibilidad del Nivel - Mithrandir
        Dim LevelMith As Integer
        LevelMith = 5
        
        If EsCONSE(tI) Then
            .outgoingData.WriteByte 255
        ElseIf .Stats.ELV >= UserList(tI).Stats.ELV + LevelMith Or .Stats.ELV >= UserList(tI).Stats.ELV Or _
            .Stats.ELV + LevelMith >= UserList(tI).Stats.ELV Then
            .outgoingData.WriteByte UserList(tI).Stats.ELV
        Else
            .outgoingData.WriteByte 255
        End If
        '\Add
                
        '.outgoingData.WriteByte UserList(tI).Stats.ELV
        
        .outgoingData.WriteByte UserList(tI).Raza
        
        If Len(UserList(tI).desc) > 0 Then
            .outgoingData.WriteByte Len(UserList(tI).desc)
        Else
            .outgoingData.WriteByte 0
        End If
        
        If UserList(tI).faccion.ArmadaReal = 1 Or UserList(tI).faccion.Milicia = 1 Or UserList(tI).faccion.FuerzasCaos = 1 Then
            .outgoingData.WriteByte UserList(tI).faccion.Rango
        End If

        If UserList(tI).flags.toyCasado = 1 Then
            .outgoingData.WriteASCIIStringFixed UserList(tI).flags.miPareja
        End If
        
        If Len(UserList(tI).desc) > 0 Then
            .outgoingData.WriteASCIIStringFixed UserList(tI).desc
        End If
    
    End With
End Sub
Sub HandleChatFaccion(ByVal UserIndex As Integer)
'Fix Marius le pusimos el buffer auxiliar
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Call buffer.ReadByte    'acá borraremos de la memoria el byte identificador.
        
        Dim chat As String
        chat = buffer.ReadASCIIString
        
        .Counters.Habla = .Counters.Habla + 1
        If UserList(UserIndex).Counters.Silenciado <> 0 Then
            If UserList(UserIndex).flags.UltimoMensaje <> 60 Then
                Call WriteConsoleMsg(1, UserIndex, "Los administrador te han silenciado por mensajes reiterados. Espere a ser desilenciado. Gracias.", FontTypeNames.FONTTYPE_BROWNI)
                UserList(UserIndex).flags.UltimoMensaje = 60
            End If
        ElseIf .flags.Muerto Then
            Call WriteMsg(UserIndex, 0)
        End If
        
        chat = Replace(chat, "~", "") 'Add Marius No colores
        
        If .faccion.Milicia = 1 Then
            If EsFaccRepu(UserIndex) And chat = "CUERNO" Then
                Call SendData(SendTarget.Tomili, 0, PrepareMessagePlayWave(139, NO_3D_SOUND, NO_3D_SOUND)) 'Cuerno faccion
            Else
                Call SendData(SendTarget.Tomili, 0, PrepareMessageConsoleMsg(3, .name & ">" & chat, FontTypeNames.FONTTYPE_FACCION_REPU))
            End If
        ElseIf .faccion.ArmadaReal = 1 Then
            If EsFaccImpe(UserIndex) And chat = "CUERNO" Then
                Call SendData(SendTarget.ToReal, 0, PrepareMessagePlayWave(139, NO_3D_SOUND, NO_3D_SOUND)) 'Cuerno faccion
            Else
                Call SendData(SendTarget.ToReal, 0, PrepareMessageConsoleMsg(3, .name & ">" & chat, FontTypeNames.FONTTYPE_FACCION_IMPE))
            End If
        ElseIf .faccion.FuerzasCaos = 1 Then
            If EsFaccCaos(UserIndex) And chat = "CUERNO" Then
                Call SendData(SendTarget.ToCaos, 0, PrepareMessagePlayWave(139, NO_3D_SOUND, NO_3D_SOUND)) 'Cuerno faccion
            Else
                Call SendData(SendTarget.ToCaos, 0, PrepareMessageConsoleMsg(3, .name & ">" & chat, FontTypeNames.FONTTYPE_FACCION_CAOS))
            End If
        Else
            Call WriteMsg(UserIndex, 35)
            'FlushBuffer UserIndex
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub
Sub WriteMensajeSigno(ByVal UserIndex As Integer)
    UserList(UserIndex).outgoingData.WriteByte ServerPacketID.MensajeSigno
    UserList(UserIndex).outgoingData.WriteByte UserList(UserIndex).cVer
End Sub
Public Sub HandleDragAndDrop(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        Dim s1 As Byte, s2 As Byte
        s1 = .incomingData.ReadByte
        s2 = .incomingData.ReadByte
        
        If .flags.Comerciando Then
            'Call SendData(SendTarget.ToADMINS, 0, PrepareMessageConsoleMsg(1, .Name & " Intentó cagar  por comercio! xD", FontTypeNames.FONTTYPE_FIGHT))
            Exit Sub
        End If
        
        If s1 < 1 Or s1 > MAX_INVENTORY_SLOTS Then _
            Exit Sub
        If s2 < 1 Or s2 > MAX_INVENTORY_SLOTS Then _
            Exit Sub
            
        SwapObjects UserIndex, s1, s2
    End With
End Sub


Private Sub HandleCancelTorneo(ByVal UserIndex As Integer)
'***************************************************
'Author: Jose Ignacio Castelli (Fedodok)
'***************************************************


        Call UserList(UserIndex).incomingData.ReadByte

        
        If EsSEMI(UserIndex) Then
            Call torneo_terminar
        End If
        

End Sub

Private Sub HandleCrearTorneo(ByVal UserIndex As Integer)
'***************************************************
'Author: Jose Ignacio Castelli (Fedudok)
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Torneos As Integer
        Torneos = buffer.ReadInteger
        
        If (EsSEMI(UserIndex)) Then

            If (Torneos > 3 And Torneos < 51) Then
                Call torneo_iniciar(Torneos)
            Else
                Call WriteConsoleMsg(1, UserIndex, "Las rondas deven ser mayores que 3 y menores que 51.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub

Private Sub HandleParticipar(ByVal UserIndex As Integer)
'*********************************************************************
'Author: Jose Ignacio Castelli (Fedudok) Y su super compañero Marius
'*********************************************************************

    If UserList(UserIndex).incomingData.length < 3 Then
        err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim evento As String
        
        evento = buffer.ReadASCIIString()
        evento = UCase$(evento)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
        
        If MapInfo(.Pos.map).Pk = True Then
            Call WriteConsoleMsg(1, UserIndex, "No puedes ingresar al evento en zona insegura.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        ElseIf UserList(UserIndex).flags.Invisible = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "No puedes ir a eventos estando invisible!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        ElseIf UserList(UserIndex).flags.Oculto = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "No puedes ir a eventos estando oculto!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        ElseIf UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "No puedes ir a eventos estando muerto!!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        ElseIf UserList(UserIndex).Stats.ELV < 25 Then
            Call WriteConsoleMsg(1, UserIndex, "Debes ser lvl 25 o mas para entrar al evento!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        'Add Marius Seguimos con los presos
        ElseIf UserList(UserIndex).Pos.map = Prision.map Then
            Call WriteConsoleMsg(1, UserIndex, "No podes estra al evento estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        '\Add
        
        If UserList(UserIndex).evento <> 0 Then
            Call salir_listas_espera(UserIndex)
        End If
        
        'If Evento = "T1VS1" And Torneo_Activo Then
            'Call Torneos_Entra(UserIndex)
        If evento = "T1VS1" And Torneo_estado = 1 Then
            UserList(UserIndex).evento = 1
            Call torneo_entrar(UserIndex)
        
        ElseIf evento = "T2VS2" Then
            'UserList(UserIndex).evento = 2
            Call WriteConsoleMsg(1, UserIndex, "Muy pronto!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        
        ElseIf evento = "CARRERA" And Carrera_estado Then
            UserList(UserIndex).evento = 3
            Call Carrera_Entra(UserIndex)
        
        ElseIf evento = "BANDERA" And Bandera_estado Then
            UserList(UserIndex).evento = 4
            Call Bandera_Entra(UserIndex)
        
        ElseIf evento = "A1VS1" And arenas_estado Then
            UserList(UserIndex).evento = 5
            Call entra_lista_espera(UserIndex, arena_espera)
            Call entrar_arena(arena_espera) 'Vemos si hay alguien mas en espera y si hay lo mandamos adentro de la arena
        
        ElseIf evento = "A2VS2" Then
            'UserList(UserIndex).evento = 6
            Call WriteConsoleMsg(1, UserIndex, "Muy pronto!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        
        ElseIf Not (Torneo_estado = 1 And Carrera_estado And arenas_estado And Bandera_estado) Then
            Call WriteConsoleMsg(1, UserIndex, "No hay ningun evento en este momento. Entra en http://inmortalao.com.ar/eventos.php Para enterarte de los eventos automáticos.", FontTypeNames.FONTTYPE_GUILD)
        
        Else
            Call WriteConsoleMsg(1, UserIndex, "El comando /PARTICIPAR ya no funciona. Para entrar a los eventos hazlo desde el boton Eventos en el Menu de tu pantalla.", FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
        
    End With
    
Errhandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        err.Raise Error
End Sub
