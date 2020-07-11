VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Inmortal AO"
   ClientHeight    =   2250
   ClientLeft      =   1950
   ClientTop       =   1530
   ClientWidth     =   6390
   ControlBox      =   0   'False
   DrawMode        =   9  'Not Mask Pen
   FillColor       =   &H00C0C0C0&
   FillStyle       =   3  'Vertical Line
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2250
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Carrera"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Captura la Bandera"
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   960
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Eventos"
      Height          =   1335
      Left            =   4080
      TabIndex        =   8
      Top             =   360
      Width           =   2175
      Begin VB.CommandButton Command3 
         Caption         =   "Arena"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "¡Cerrar Servidor!"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   5895
   End
   Begin VB.Timer minuto 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5880
      Top             =   0
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   3960
      Top             =   0
   End
   Begin VB.Timer packetResend 
      Interval        =   10
      Left            =   4440
      Top             =   0
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   0
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   3000
      Top             =   0
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4920
      Top             =   0
   End
   Begin VB.Timer npcataca 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   5400
      Top             =   0
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3480
      Top             =   0
   End
   Begin VB.Frame asdasd 
      Caption         =   "Reload"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3735
      Begin VB.CommandButton Command2 
         Caption         =   "Balance"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton reloadIni 
         Caption         =   ".ini's"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton ReloadObjs 
         Caption         =   "Objs"
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton ReloadHechis 
         Caption         =   "Hechis"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton reloadNpc 
         Caption         =   "Npc's"
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label lbl_online 
      Caption         =   "Online: 0"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMaximizar 
         Caption         =   "&Maximizar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public ESCUCHADAS As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONUP = &H205

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, id As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = id
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Sub CheckIdleUser()
    Dim iUserIndex As Long
    
    On Error GoTo hayerror
    
    'CASTELLI /// Cambiamos de To maxuser a lastuser
    'Vos decis, conexion activa? , y usuario logueado?
    'entonces para que carajo haces un mega for hasta 100 todo el tiempo
    'si lo unico que revisas son las conexiones activas :S:S:S
    'Inentendible
    

    
    For iUserIndex = 1 To LastUser

       
       'Conexion activa? y es un usuario loggeado?
       If UserList(iUserIndex).ConnID <> -1 And UserList(iUserIndex).flags.UserLogged = True Then
            
            'Actualiza el contador de inactividad
            UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
            
            'Add by Nod se intenta no desconectar a los GMS
            If EsGM(iUserIndex) Then
                UserList(iUserIndex).Counters.IdleCount = 1
            End If
            '\Add
            
            If UserList(iUserIndex).Counters.IdleCount >= 20 Then
                Call WriteShowMessageBox(iUserIndex, "Demasiado tiempo inactivo. Has sido desconectado..")
                'mato los comercios seguros
                If UserList(iUserIndex).ComUsu.DestUsu > 0 Then
                    If UserList(UserList(iUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                        If UserList(UserList(iUserIndex).ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
                            Call WriteConsoleMsg(1, UserList(iUserIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
                            Call FinComerciarUsu(UserList(iUserIndex).ComUsu.DestUsu)
                            Call FlushBuffer(UserList(iUserIndex).ComUsu.DestUsu) 'flush the buffer to send the message right away
                        End If
                    End If
                    Call FinComerciarUsu(iUserIndex)
                End If
                Call Cerrar_Usuario(iUserIndex)
            End If
        End If
    Next iUserIndex
    
    
    Exit Sub
    
hayerror:
    LogError ("Error en CHECKIDLEUSER: " & err.Number & " Desc: " & err.description)
    
    
End Sub

Private Sub Auditoria_Timer()
On Error GoTo errhand


    Call PasarSegundo 'sistema de desconexion de 10 segs

    Exit Sub

errhand:

'Call LogError("Error en Timer Auditoria. Err: " & err.description & " - " & err.Number)
End Sub

Private Sub AutoSave_Timer()

On Error GoTo Errhandler

'Static MinutosLatsClean As Integer
'Dim i As Integer
    ' Cada 1 minuto limpia un mapa
    ' Son 850 mapas tardaria mas de 14 horas en limpiar todos :S
 'Call limpiaritemsmundo
    
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    'Call ModAreas.AreasOptimizacion
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

  '  modSubastas.sLoop
    
   ' If MinutosLatsClean >= 15 Then
   '     Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
   '     'Add Nod kopfnickend prevenimos que haga un overflood
    '   MinutosLatsClean = 0
    'Else
    '    MinutosLatsClean = MinutosLatsClean + 1
    'End If
    
    
    Call PurgarPenas
    Call CheckIdleUser
    

Exit Sub
Errhandler:
    Call LogError("Error en TimerAutoSave " & err.Number & ": " & err.description)
    Resume Next
End Sub

Public Sub InitMain(ByVal F As Byte)

Dim i As Integer
Dim S As String
Dim nid As NOTIFYICONDATA

    S = "InmortalAO"
    nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
    i = Shell_NotifyIconA(NIM_ADD, nid)
        
    'If WindowState <> vbMinimized Then WindowState = vbMinimized
    'Visible = False

End Sub

Private Sub Command1_Click()
    Call mnusalir_Click
End Sub

Private Sub Command2_Click()
    LoadBalance
End Sub


Private Sub Command3_Click()
    If arenas_estado Then
        Call arenas_Cerrar
    Else
        Call arenas_Abrir
    End If
End Sub

Private Sub Command4_Click()
    If Bandera_estado Then
        Call Bandera_Termina
    Else
        Call Bandera_Inicia
    End If
End Sub

Private Sub Command6_Click()
    Call LoadGuildsDB
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
   
   If Not Visible Then
        Select Case x \ Screen.TwipsPerPixelX
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
End Sub

Private Sub QuitarIconoSystray()
On Error Resume Next

    'Borramos el icono del systray
    Dim i As Integer
    Dim nid As NOTIFYICONDATA
    
    nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")
    
    i = Shell_NotifyIconA(NIM_DELETE, nid)
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call QuitarIconoSystray
    
    
    Dim loopC As Integer
    
    For loopC = 1 To LastUser
        If UserList(loopC).ConnID <> -1 Then Call CloseSocket(loopC)
    Next
    
    End

End Sub


Private Sub GameTimer_Timer()
    Static iUserIndex As Long
    Static bEnviarStats As Boolean
    Static bEnviarAyS As Boolean
    
On Error GoTo hayerror
    
    '<<<<<< Procesa eventos de los usuarios >>>>>>


' BY CASTELLI /// REEMPLAZO DE MAXUSERS POR LASTUSER
' No entiendo porq ponen maxusers si solo se debe
' revisar entre los usuarios logueados :S:S:S:S
    
    For iUserIndex = 1 To MaxUsers ' masusers
    

    
        With UserList(iUserIndex)
           'Conexion activa? ' Comenta eso pelotudo que no sabes ni de lo que hablas
           ' poniendo MAXUSERS JAJAJA by castelli... dios mio...
                '¿User valido?
                
                If .ConnID <> -1 Then
                    If .ConnIDValida = True And .flags.UserLogged = True Then
                    
                        bEnviarStats = False
                        bEnviarAyS = False
                    

                        If .flags.Muerto = 0 Then
                            If .flags.Paralizado = 1 Then
                                Call EfectoParalisisUser(iUserIndex)
                            End If
                        
                        If .flags.Ceguera = 1 Or .flags.Estupidez Then
                            Call EfectoCegueEstu(iUserIndex)
                        End If
                        
                        If .flags.Metamorfosis = 1 Then
                            Call EfectoMetamorfosis(iUserIndex)
                        End If
                        
                        If Not EsCONSE(iUserIndex) Then
                            If .flags.Montando = 0 And .flags.Desnudo <> 0 Then
                                Call EfectoFrio(iUserIndex)
                            End If
                        
                            If .flags.Envenenado <> 0 Then
                                Call EfectoVeneno(iUserIndex)
                            End If
                                                
                            If .flags.Incinerado <> 0 Then
                                Call EfectoIncineracion(iUserIndex)
                            End If
                            
                            If .Stats.eCreateTipe <> 0 Then
                                Call EfectoHechizoMagico(iUserIndex)
                            End If
                        End If
                            
                        If .flags.Meditando Then
                            Call DoMeditar(iUserIndex)
                        End If
                        
                        If .flags.AdminInvisible <> 1 Then
                            If .flags.Invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
                            If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)
                        End If
                        If .flags.Trabajando Then DoTrabajar iUserIndex
                        
                        Call DuracionPociones(iUserIndex)
                        
                        Call HambreYSed(iUserIndex, bEnviarAyS)
                        
                        If .flags.Hambre = 0 And .flags.Sed = 0 Then
                            If Not .flags.Descansar Then
                                Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                                If bEnviarStats Then
                                    Call WriteUpdateHP(iUserIndex)
                                    bEnviarStats = False
                                End If
                                
                                Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                                
                                If bEnviarStats Then
                                    Call WriteUpdateSta(iUserIndex)
                                    bEnviarStats = False
                                End If
                                
                            Else
                                Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                                If bEnviarStats Then
                                    Call WriteUpdateHP(iUserIndex)
                                    bEnviarStats = False
                                End If
                                
                                Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                                If bEnviarStats Then
                                    Call WriteUpdateSta(iUserIndex)
                                    bEnviarStats = False
                                End If
                                
                                If .Stats.MaxHP = .Stats.MinHP And .Stats.MaxSTA = .Stats.MinSTA Then
                                    Call WriteRestOK(iUserIndex)
                                    Call WriteConsoleMsg(1, iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                    .flags.Descansar = False
                                End If
                                
                            End If
                        End If
                        
                        If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)
                        
                        If .NroMascotas > 0 Then Call TiempoInvocacion(iUserIndex)
                    Else
                        If .flags.Resucitando <> 0 Then Call DoResucitar(iUserIndex)
                    End If
                    
                    If .Counters.Silenciado > 0 Then
                        .Counters.Silenciado = .Counters.Silenciado - 1
                    End If
                    End If
                ElseIf .ConnID <> -1 And .ConnIDValida = True And .flags.UserLogged = False Then

                    .Counters.IdleCount = .Counters.IdleCount + 1
                    If .Counters.IdleCount > IntervaloParaConexion Then
                        .Counters.IdleCount = 0
                        Call WriteMsg(iUserIndex, 47)
                        Call WriteDisconnect2(iUserIndex) 'Se aplica_
                        'Para cerrar bien bien la conexion_
                        'No como se lo estaba haciendo antes_
                        'Como el culo :P, este error sucedia ya que_
                        'Al aplicar el sistema de cuenta, no se tenia en_
                        'cuenta que debia cerrarce desp tambien la conexion_
                        'mandando mensaje al cliente /// BY CASTELLI
                        Call FlushBuffer(iUserIndex)
                        Call CloseSocket(iUserIndex)
                    End If
                    
                End If
                
                Call FlushBuffer(iUserIndex)
           
           
        End With
  
    Next iUserIndex
    Exit Sub

hayerror:
'    LogError ("Error en GameTimer: " & Err.description & " UserIndex = " & iUserIndex)
    
End Sub


Private Sub mnuMaximizar_click()
    frmMain.Visible = True
End Sub
Public Sub mnusalir_Click()
    If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
        If MsgBox("Desea guardar los usuarios y hacer un backup?", vbYesNo) = vbYes Then
            GuardarUsuarios
        End If

        Call extra_set("online", "0")

        Dim F
        For Each F In Forms
            Unload F
        Next
    End If
End Sub


Private Sub npcataca_Timer()

On Error Resume Next
    Dim npc As Long
    
    For npc = 1 To LastNPC
        Npclist(npc).CanAttack = 1
    Next npc

End Sub

Private Sub packetResend_Timer()
On Error GoTo Errhandler:
    Dim i As Long
    
    For i = 1 To MaxUsers
        If UserList(i).ConnIDValida Then
            If UserList(i).outgoingData.length > 0 Then
                Call EnviarDatosASlot(i, UserList(i).outgoingData.ReadASCIIStringFixed(UserList(i).outgoingData.length))
            End If
        End If
    Next i
    DoEvents
Exit Sub

Errhandler:
'    LogError ("Error en packetResend - Error: " & Err.Number & " - Desc: " & Err.description)
    Resume Next
End Sub





Private Sub ReloadHechis_Click()
    Call CargarHechizos
End Sub

Private Sub reloadIni_Click()
    Call LoadSini
End Sub

Private Sub reloadNpc_Click()
    Call CargaNpcsDat
End Sub

Private Sub ReloadObjs_Click()
    Call LoadOBJData
End Sub

Private Sub TIMER_AI_Timer()

On Error GoTo ErrorHandler

Dim NpcIndex As Long
Dim x As Integer
Dim Y As Integer
Dim UseAI As Integer
Dim mapa As Integer
Dim e_p As Integer


If Not haciendoBK And Not EnPausa Then
    'Update NPCs
    For NpcIndex = 1 To LastNPC

        If Npclist(NpcIndex).flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
            If Npclist(NpcIndex).flags.Paralizado = 1 Then
                Call EfectoParalisisNpc(NpcIndex)
            Else
                'Usamos AI si hay algun user en el mapa
                If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
                   Call EfectoParalisisNpc(NpcIndex)
                End If
                
                mapa = Npclist(NpcIndex).Pos.map
                
                If mapa > 0 Then
                    If MapInfo(mapa).NumUsers > 0 Then
                        If Npclist(NpcIndex).Movement <> TipoAI.ESTATICO Then
                            Call NPCAI(NpcIndex)
                        End If
                    End If
                End If
            End If
        End If
    Next NpcIndex
End If

Exit Sub

ErrorHandler:
    Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).Name & " mapa:" & Npclist(NpcIndex).Pos.map)
'    Call MuereNpc(NpcIndex, 0)
End Sub


Private Sub tPiqueteC_Timer()

    Dim NuevaA As Boolean
    Dim NuevoL As Boolean
    Dim GI As Integer
    
    Dim i As Long
        
    On Error GoTo Errhandler
        For i = 1 To LastUser
            With UserList(i)
                If .flags.UserLogged Then
                    If MapData(.Pos.map, .Pos.x, .Pos.Y).Trigger = eTrigger.ANTIPIQUETE And Not EsCONSE(i) Then
                        .Counters.PiqueteC = .Counters.PiqueteC + 1
                        Call WriteConsoleMsg(1, i, "¡¡¡Estás obstruyendo la vía pública, muévete o serás encarcelado!!!", FontTypeNames.FONTTYPE_INFO)
                        
                        If .Counters.PiqueteC > 10 Then
                            .Counters.PiqueteC = 0
                            Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
                        End If
                    Else
                        .Counters.PiqueteC = 0
                    End If
                    
                    'ustedes se preguntaran que hace esto aca?
                    'bueno la respuesta es simple: el codigo de AO es una mierda y encontrar
                    'todos los puntos en los cuales la alineacion puede cambiar es un dolor de
                    'huevos, asi que lo controlo aca, cada 6 segundos, lo cual es razonable
            
                    GI = .GuildIndex
                    If GI > 0 Then
                        NuevaA = False
                        NuevoL = False
                        If Not modGuilds.m_ValidarPermanencia(i, True, NuevaA, NuevoL) Then Call WriteConsoleMsg(1, i, "Has sido expulsado del clan. ¡El clan ha sumado un punto de antifacción!", FontTypeNames.FONTTYPE_GUILD)
                  
                        If NuevaA Then Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(1, "¡El clan ha pasado a tener alineación neutral!", FontTypeNames.FONTTYPE_GUILD))
                        If NuevoL Then Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(1, "¡El clan tiene un nuevo líder!", FontTypeNames.FONTTYPE_GUILD))
                    End If
                    
                    Call FlushBuffer(i)
                End If
            End With
        Next i
    Exit Sub

Errhandler:
    Call LogError("Error en tPiqueteC_Timer " & err.Number & ": " & err.description)
End Sub

Private Sub minuto_Timer()

On Error Resume Next
    'Add Marius Sistema de mensajes publicitarios e informativos
    mins = mins + 1
    
    If mins = 15 Then
        mins = 0 'Reseteamos la variable
        
        'Mandamos la publicidad
        Call SendPublicidad

        'Guardamos usuarios
        Call GuardarUsuarios
    End If
    '\Add
    
    'Add Marius
    'Mod Torneos 1 vs 1
    'Add Eventos "x2"
    Dim hora As String
    Dim npc As Long
    
    hora = Format(time, "hh:mm")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''' Torneo 1 vs 1 '''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If hora = "18:25" Or hora = "13:25" Then ' Avisos 5 min antes de empezar Torneo
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> En 5 minutos comienza Torneo automatico 1vs1.", FontTypeNames.FONTTYPE_SERVER))
    
    ElseIf hora = "18:30" Or hora = "13:30" Then ' Comienza el torneo
        'Call torneos_auto(RondasAutomatico)
        Call torneo_iniciar(RondasAutomatico)
    ElseIf hora = "18:31" Or hora = "13:31" Then 'Si no hay participantes suficientes, se cancela
        If Torneo_estado = 1 Then Call torneo_terminar
        
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''' Carrera ''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf hora = "20:25" Or hora = "10:25" Then ' Avisos 5 min antes de empezar Carrera
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> En 5 minutos comienza la Carrera. La inscripción cuesta 100k.", FontTypeNames.FONTTYPE_SERVER))
    
    ElseIf hora = "20:30" Or hora = "10:30" Then ' Comienza
        MapInfo(MapaCarrera).InviSinEfecto = True
        MapInfo(MapaCarrera).Pk = False
        Carrera_estado = True
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(252, NO_3D_SOUND, NO_3D_SOUND)) 'Cuerno
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> La carrera comienza en 1 minuto. La inscripción cuesta 100k, para participar ingresa al evento mediante el boton Eventos de tu Menu.", FontTypeNames.FONTTYPE_SERVER))
        MsgEvento = "¡La inscripcion a la Carrera essta abierta. Ingresa desde el botón Eventos de tu Menu!"
        
    ElseIf hora = "20:31" Or hora = "10:31" Then 'Si hay participantes se hace, sino se cancela
        MsgEvento = ""
        If Carrera_puestos > 5 Then
            MapInfo(MapaCarrera).Pk = True
            
            'Ponemos la particula del portal
            Call SendData(SendTarget.ToMap, MapaCarrera, PrepareMessageCreateParticle(47, 56, 16))
            
            'Desbloqueamos las salidas
            MapData(MapaCarrera, 44, 68).Blocked = 0
            Call Bloquear(True, MapaCarrera, 44, 68, 0)
            MapData(MapaCarrera, 50, 68).Blocked = 0
            Call Bloquear(True, MapaCarrera, 50, 68, 0)
            MapData(MapaCarrera, 44, 41).Blocked = 0
            Call Bloquear(True, MapaCarrera, 44, 41, 0)
            MapData(MapaCarrera, 50, 41).Blocked = 0
            Call Bloquear(True, MapaCarrera, 50, 41, 0)

            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> Ya comenzó la Carrera!", FontTypeNames.FONTTYPE_SERVER))
            Carrera_puestos = 255
        Else
            'Transportamos a todos a intermundia
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> La Carrera se canceló por falta de participantes.", FontTypeNames.FONTTYPE_SERVER))
            Carrera_puestos = 0
        End If
        
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''' Arenas 1vs1 ''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf hora = "13:55" Or hora = "20:55" Then ' Avisos 5 min antes de empezar evento
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> En 5 minutos abren las Arenas.", FontTypeNames.FONTTYPE_SERVER))
    
    ElseIf hora = "15:25" Or hora = "22:55" Then ' Avisos 5 min antes de terminar el evento
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> En 5 minutos cierran las Arenas!", FontTypeNames.FONTTYPE_SERVER))
    
    ElseIf hora = "14:00" Or hora = "21:00" Then 'Comienza
        Call arenas_Abrir
        MsgEvento = "¡Las Arenas 1vs1 Estan abiertas. Ingresa desde el botón Eventos de tu Menu!"
        
    ElseIf hora = "15:30" Or hora = "23:00" Then 'Termina
        Call arenas_Cerrar
        MsgEvento = ""
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''' Sms x2 '''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf hora = "18:55" Then ' Avisos 5 min antes de empezar evento
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> En 5 minutos comienza Sms x2. Enterate de mas en http://inmortalao.com.ar/", FontTypeNames.FONTTYPE_SERVER))
    
    ElseIf hora = "19:55" Then ' Avisos 5 min antes de terminar el evento
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> En 5 minutos termina Sms x2. Todavia no mandaste tu SMS, apurate!", FontTypeNames.FONTTYPE_SERVER))
    
    ElseIf hora = "19:00" Then 'Comienza
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(252, NO_3D_SOUND, NO_3D_SOUND)) 'Cuerno
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> Comienza el evento Sms x2. Enterate de mas en http://inmortalao.com.ar/", FontTypeNames.FONTTYPE_SERVER))
        MsgEvento = "¡El evento SMS x2 esta en curso! Enterate de mas en http://inmortalao.com.ar/"
        
    ElseIf hora = "20:00" Then 'Termina
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> Terminó el evento Sms x2.", FontTypeNames.FONTTYPE_SERVER))
        MsgEvento = ""
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''' ExpOro x2 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf hora = "10:55" Or hora = "16:55" Then ' Avisos 5 min antes de empezar ExpOroX2
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> En 5 minutos comienza Expe y Oro x2.", FontTypeNames.FONTTYPE_SERVER))
    
    ElseIf hora = "11:55" Or hora = "17:55" Then ' Avisos 5 min antes de terminar ExpOroX2
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> En 5 minutos termina Expe y Oro x2.", FontTypeNames.FONTTYPE_SERVER))
    
    ElseIf hora = "11:00" Or hora = "17:00" Then ' Comienza ExpOroX2
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(252, NO_3D_SOUND, NO_3D_SOUND)) 'Cuerno
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> Comienza el evento Expe y Oro x2. Duracion 1 hora.", FontTypeNames.FONTTYPE_SERVER))
        MsgEvento = "¡El evento Experiencia y Oro x2 esta en curso!"

        For npc = 1 To MAXNPCS
            Npclist(npc).GiveEXP = Npclist(npc).GiveEXP * 1.8
            Npclist(npc).GiveGLD = Npclist(npc).GiveGLD * 1.8
        Next npc
        
        
        'Pone en la db que el evento comenzó
        Call extra_set("EventoX2", "1")
        
    ElseIf hora = "12:00" Or hora = "18:00" Then ' Termina ExpOroX2
        'Hacemos la verificacion, por si el server estaba apagado y justo se prende sin iniciar el evento, la expe quedaria mas facil de lo habitual si la baja
        If "1" = extra_get("EventoX2") Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> Terminó el evento Expe y Oro x2.", FontTypeNames.FONTTYPE_SERVER))
            MsgEvento = ""
            
            For npc = 1 To MAXNPCS
                Npclist(npc).GiveEXP = Npclist(npc).GiveEXP / 1.8
                Npclist(npc).GiveGLD = Npclist(npc).GiveGLD / 1.8
            Next npc
            
            'Pone en la db que el evento terminó
            Call extra_set("EventoX2", "0")
        End If
    End If
    '\add
    
End Sub
