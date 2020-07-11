Attribute VB_Name = "wskapiAO"
'
'Option Explicit
'
'''
'' Modulo para manejar Winsock
''
'
''Si la variable esta en TRUE , al iniciar el WsApi se crea
''una ventana LABEL para recibir los mensajes. Al detenerlo,
''se destruye.
''Si es FALSE, los mensajes se envian al form frmMain (o el
''que sea).
'#Const WSAPI_CREAR_LABEL = True
'
'Private Const SD_BOTH As Long = &H2
'
'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'
'Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
'Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
'Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
'
'Private Const WS_CHILD = &H40000000
'Public Const GWL_WNDPROC = (-4)
'
'Private Const SIZE_RCVBUF As Long = 8192
'Private Const SIZE_SNDBUF As Long = 8192
'
'Dim TickAcceptConection As Long
'
'''
''Esto es para agilizar la busqueda del slot a partir de un socket dado,
''sino, la funcion BuscaSlotSock se nos come todo el uso del CPU.
''
'' @param Sock sock
'' @param slot slot
''
'Public Type tSockCache
'    Sock As Long
'    Slot As Long
'End Type
'
'Public WSAPISock2Usr As New Collection
'
'' ====================================================================================
'' ====================================================================================
'
'Public OldWProc As Long
'Public ActualWProc As Long
'Public hWndMsg As Long
'
'' ====================================================================================
'' ====================================================================================
'
'Public SockListen As Long
'
'
'' ====================================================================================
'' ====================================================================================
'
'
'Public Sub IniciaWsApi(ByVal hwndParent As Long)
'
'Call LogApiSock("IniciaWsApi")
'Debug.Print "IniciaWsApi"
'
'#If WSAPI_CREAR_LABEL Then
'hWndMsg = CreateWindowEx(0, "STATIC", "AOMSG", WS_CHILD, 0, 0, 0, 0, hwndParent, 0, App.hInstance, ByVal 0&)
'#Else
'hWndMsg = hwndParent
'#End If 'WSAPI_CREAR_LABEL
'
'OldWProc = SetWindowLong(hWndMsg, GWL_WNDPROC, AddressOf WndProc)
'ActualWProc = GetWindowLong(hWndMsg, GWL_WNDPROC)
'
'Dim desc As String
'Call StartWinsock(desc)
'
'End Sub
'
'Public Sub LimpiaWsApi()
'
'Call LogApiSock("LimpiaWsApi")
'
'If WSAStartedUp Then
'    Call EndWinsock
'End If
'
'If OldWProc <> 0 Then
'    SetWindowLong hWndMsg, GWL_WNDPROC, OldWProc
'    OldWProc = 0
'End If
'
'#If WSAPI_CREAR_LABEL Then
'If hWndMsg <> 0 Then
'    DestroyWindow hWndMsg
'End If
'#End If
'
'End Sub
'
'Public Function BuscaSlotSock(ByVal S As Long) As Long
'On Error GoTo hayerror
'
'    BuscaSlotSock = WSAPISock2Usr.Item(CStr(S))
'    Exit Function
'
'hayerror:
'    BuscaSlotSock = -1
'
'End Function
'
'Public Sub AgregaSlotSock(ByVal Sock As Long, ByVal Slot As Long)
'Debug.Print "AgregaSockSlot"
'
'If WSAPISock2Usr.Count > MaxUsers Then
'    Call CloseSocket(Slot)
'    Exit Sub
'End If
'
'WSAPISock2Usr.Add CStr(Slot), CStr(Sock)
'
'End Sub
'
'Public Sub BorraSlotSock(ByVal Sock As Long)
'    Dim Cant As Long
'    Cant = WSAPISock2Usr.Count
'
'On Error Resume Next
'    WSAPISock2Usr.Remove CStr(Sock)
'
'    Debug.Print "BorraSockSlot " & Cant & " -> " & WSAPISock2Usr.Count
'End Sub
'
'
'
'Public Function WndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'On Error Resume Next
'
'    Dim Ret As Long
'    Dim Tmp() As Byte
'    Dim S As Long
'    Dim E As Long
'    Dim N As Integer
'    Dim UltError As Long
'
'    Select Case msg
'        Case 1025
'            S = wParam
'            E = WSAGetSelectEvent(lParam)
'
'            Select Case E
'                Case FD_ACCEPT
'                    If S = SockListen Then
'                        Call EventoSockAccept(S)
'                    End If
'
'            End Select
'
'        Case Else
'            WndProc = CallWindowProc(OldWProc, hWnd, msg, wParam, lParam)
'    End Select
'End Function
'
''Retorna 0 cuando se envió o se metio en la cola,
''retorna <> 0 cuando no se pudo enviar o no se pudo meter en la cola
'Public Function WsApiEnviar(ByVal Slot As Integer, ByRef str As String) As Long
'    Dim Ret As String
'    Dim Retorno As Long
'    Dim data() As Byte
'
'    ReDim Preserve data(Len(str) - 1) As Byte
'
'    data = StrConv(str, vbFromUnicode)
'
'    Retorno = 0
'
'    If UserList(Slot).ConnID <> -1 And UserList(Slot).ConnIDValida Then
'        Ret = send(ByVal UserList(Slot).ConnID, data(0), ByVal UBound(data()) + 1, ByVal 0)
'        If Ret < 0 Then
'            Ret = err.LastDllError
'            If Ret = WSAEWOULDBLOCK Then
'                Call UserList(Slot).outgoingData.WriteASCIIStringFixed(str)
'            End If
'        End If
'    ElseIf UserList(Slot).ConnID <> -1 And Not UserList(Slot).ConnIDValida Then
'        If Not UserList(Slot).Counters.Saliendo Then
'            Retorno = -1
'        End If
'    End If
'
'    WsApiEnviar = Retorno
'End Function
'
'Public Sub LogApiSock(ByVal str As String)
'
'On Error GoTo Errhandler
'
'Dim nfile As Integer
'nfile = FreeFile ' obtenemos un canal
'Open App.Path & "\logs\wsapi.log" For Append Shared As #nfile
'Print #nfile, Date & " " & time & " " & str
'Close #nfile
'
'Exit Sub
'
'Errhandler:
'
'End Sub
'
'Public Sub EventoSockAccept(ByVal SockID As Long)
''==========================================================
''USO DE LA API DE WINSOCK
''========================
'
'    Dim NewIndex As Integer
'    Dim Ret As Long
'    Dim Tam As Long, sa As sockaddr
'    Dim NuevoSock As Long
'    Dim i As Long
'    Dim tStr As String
'
'    Tam = sockaddr_size
'
'    Ret = accept(SockID, sa, Tam)
'
'    If GetTickCount - TickAcceptConection > 200 Then
'        TickAcceptConection = GetTickCount
'    Else
'        Call WSApiCloseSocket(NuevoSock)
'        Exit Sub
'    End If
'
'    If Ret = INVALID_SOCKET Then
'        i = err.LastDllError
'        Call LogCriticEvent("Error en Accept() API " & i & ": " & GetWSAErrorString(i))
'        Exit Sub
'    End If
'
'    If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
'        Call WSApiCloseSocket(NuevoSock)
'        Exit Sub
'    End If
'
'    NuevoSock = Ret
'
'    'Seteamos el tamaño del buffer de entrada
'    If setsockopt(NuevoSock, SOL_SOCKET, SO_RCVBUFFER, SIZE_RCVBUF, 4) <> 0 Then
'        i = err.LastDllError
'        Call LogCriticEvent("Error al setear el tamaño del buffer de entrada " & i & ": " & GetWSAErrorString(i))
'    End If
'    'Seteamos el tamaño del buffer de salida
'    If setsockopt(NuevoSock, SOL_SOCKET, SO_SNDBUFFER, SIZE_SNDBUF, 4) <> 0 Then
'        i = err.LastDllError
'        Call LogCriticEvent("Error al setear el tamaño del buffer de salida " & i & ": " & GetWSAErrorString(i))
'    End If
'
'    'Mariano: Baje la busqueda de slot abajo de CondicionSocket y limite x ip
'    NewIndex = NextOpenUser ' Nuevo indice
'
'    If NewIndex <= MaxUsers Then
'
'        'Make sure both outgoing and incoming data buffers are clean
'        Call UserList(NewIndex).incomingData.ReadASCIIStringFixed(UserList(NewIndex).incomingData.length)
'        Call UserList(NewIndex).outgoingData.ReadASCIIStringFixed(UserList(NewIndex).outgoingData.length)
'
'        UserList(NewIndex).ip = GetAscIP(sa.sin_addr)
'        'Busca si esta banneada la ip
'        For i = 1 To BanIps.Count
'            If BanIps.Item(i) = UserList(NewIndex).ip Then
'                'Call apiclosesocket(NuevoSock)
'                Call WriteErrorMsg(NewIndex, "Su IP se encuentra bloqueada en este servidor.")
'                Call FlushBuffer(NewIndex)
'
'                Call WSApiCloseSocket(NuevoSock)
'                Exit Sub
'            End If
'        Next i
'
'        If NewIndex > LastUser Then LastUser = NewIndex
'
'        UserList(NewIndex).ConnID = NuevoSock
'        UserList(NewIndex).ConnIDValida = True
'
'        Call AgregaSlotSock(NuevoSock, NewIndex)
'        SendEncriptCode NewIndex
'    Else
'        Dim str As String
'        Dim data() As Byte
'
'        str = Protocol.PrepareMessageErrorMsg("El server se encuentra lleno en este momento. Disculpe las molestias ocasionadas.")
'
'        ReDim Preserve data(Len(str) - 1) As Byte
'
'        data = StrConv(str, vbFromUnicode)
'
'        Call send(ByVal NuevoSock, data(0), ByVal UBound(data()) + 1, ByVal 0)
'        Call WSApiCloseSocket(NuevoSock)
'    End If
'End Sub
'
'Public Sub EventoSockRead(ByVal Slot As Integer, ByRef Datos() As Byte)
'    With UserList(Slot)
'
'        Call .incomingData.WriteBlock(Datos)
'
'        If .ConnID <> -1 Then
'            Call HandleIncomingData(Slot)
'        Else
'            Exit Sub
'        End If
'    End With
'End Sub
'
'Public Sub EventoSockClose(ByVal Slot As Integer)
'    'Es el mismo user al que está revisando el centinela??
'    'Si estamos acá es porque se cerró la conexión, no es un /salir, y no queremos banearlo....
'    If Centinela.RevisandoUserIndex = Slot Then _
'        Call modCentinela.CentinelaUserLogout
'
'    If UserList(Slot).flags.UserLogged Then
'        Call CloseSocketSL(Slot)
'        Call Cerrar_Usuario(Slot)
'    Else
'        Call CloseSocket(Slot)
'    End If
'End Sub
'
'
'Public Sub WSApiReiniciarSockets()
'Dim i As Long
'    'Cierra el socket de escucha
'    If SockListen >= 0 Then Call apiclosesocket(SockListen)
'
'    'Cierra todas las conexiones
'    For i = 1 To MaxUsers
'        If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
'            Call CloseSocket(i)
'        End If
'    Next i
'
'    For i = 1 To MaxUsers
'        Set UserList(i).incomingData = Nothing
'        Set UserList(i).outgoingData = Nothing
'    Next i
'
'    ' No 'ta el PRESERVE :p
'    ReDim UserList(1 To MaxUsers)
'    For i = 1 To MaxUsers
'        UserList(i).ConnID = -1
'        UserList(i).ConnIDValida = False
'        UserList(i).KeyCrypt = 0
'        UserList(i).HaveEncript = 0
'
'        Set UserList(i).incomingData = New clsByteQueue
'        Set UserList(i).outgoingData = New clsByteQueue
'    Next i
'
'    LastUser = 1
'    NumUsers = 0
'
'    Call LimpiaWsApi
'    Call Sleep(100)
'    Call IniciaWsApi(frmMain.hWnd)
'    SockListen = ListenForConnect(Puerto, hWndMsg, "")
'
'End Sub
'
'Public Sub WSApiCloseSocket(ByVal socket As Long)
'    Call WSAAsyncSelect(socket, hWndMsg, ByVal 1025, ByVal (FD_CLOSE))
'    Call ShutDown(socket, SD_BOTH)
'End Sub
'
'Public Function CondicionSocket(ByRef lpCallerId As WSABUF, ByRef lpCallerData As WSABUF, ByRef lpSQOS As FLOWSPEC, ByVal Reserved As Long, ByRef lpCalleeId As WSABUF, ByRef lpCalleeData As WSABUF, ByRef Group As Long, ByVal dwCallbackData As Long) As Long
'    Dim sa As sockaddr
'
'    'Check if we were requested to force reject
'
'    If dwCallbackData = 1 Then
'        CondicionSocket = CF_REJECT
'        Exit Function
'    End If
'
'     'Get the address
'
'    CopyMemory sa, ByVal lpCallerId.lpBuffer, lpCallerId.dwBufferLen
'
'
'    If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
'        CondicionSocket = CF_REJECT
'        Exit Function
'    End If
'
'    CondicionSocket = CF_ACCEPT 'En realdiad es al pedo, porque CondicionSocket se inicializa a 0, pero así es más claro....
'End Function
'
'
