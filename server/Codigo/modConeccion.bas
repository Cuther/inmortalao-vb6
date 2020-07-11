Attribute VB_Name = "modConeccion"
Option Explicit
#Const WSAPI_CREAR_LABEL = True

Private Const SD_BOTH As Long = &H2

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const WS_CHILD = &H40000000
Public Const GWL_WNDPROC = (-4)

Private Const SIZE_RCVBUF As Long = 8192
Private Const SIZE_SNDBUF As Long = 8192

Public Type tSockCache
    sock As Long
    Slot As Long
End Type

Public WSAPISock2Usr As New Collection

Public OldWProc As Long
Public ActualWProc As Long
Public hWndMsg As Long

Public SockListen As Long


Public Sub IniciaWsApi(ByVal hwndParent As Long)
    Call LogApiSock("IniciaWsApi")
    Debug.Print "IniciaWsApi"

    #If WSAPI_CREAR_LABEL Then
        hWndMsg = CreateWindowEx(0, "STATIC", "AOMSG", WS_CHILD, 0, 0, 0, 0, hwndParent, 0, App.hInstance, ByVal 0&)
    #Else
        hWndMsg = hwndParent
    #End If

    OldWProc = SetWindowLong(hWndMsg, GWL_WNDPROC, AddressOf WndProc)
    ActualWProc = GetWindowLong(hWndMsg, GWL_WNDPROC)

    Dim desc As String
    Call StartWinsock(desc)

End Sub



Public Function BuscaSlotSock(ByVal S As Long) As Long
On Error GoTo Error

    BuscaSlotSock = WSAPISock2Usr.Item(CStr(S))
    Exit Function

Error:
    BuscaSlotSock = -1
    err.Clear
    
End Function

Public Sub AgregaSlotSock(ByVal sock As Long, ByVal Slot As Long)

    If WSAPISock2Usr.count > MaxUsers Then
        Call CloseSocket(Slot)
        Exit Sub
    End If
    
    
    WSAPISock2Usr.Add CStr(Slot), CStr(sock)
    Debug.Print "AgregaSockSlot"

End Sub

Public Sub BorraSlotSock(ByVal sock As Long)
On Error Resume Next

Dim Cant As Long

    Cant = WSAPISock2Usr.count
    
    WSAPISock2Usr.Remove CStr(sock)
    
    Debug.Print "BorraSockSlot " & Cant & " -> " & WSAPISock2Usr.count

End Sub



Public Function WndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
    Dim Ret As Long
    Dim Tmp() As Byte
    
    Dim S As Long, E As Long
    Dim N As Integer
    
    Dim UltError As Long
    
    WndProc = 0
    
 
    
    Select Case msg
        Case 1025
            S = wParam
            E = WSAGetSelectEvent(lParam)
            'Call LogApiSock("Msg: " & msg & " W: " & wParam & " L: " & lParam)
 
            Select Case E
                
                Case FD_ACCEPT
                    If S = SockListen Then
                        Call EventoSockAccept(S)
                    End If
        
                Case FD_READ
                    N = BuscaSlotSock(S)
                    If N < 0 And S <> SockListen Then
                  
                        Call WSApiCloseSocket(S)
                        Exit Function
                    End If
                    
                    ReDim Preserve Tmp(SIZE_RCVBUF - 1) As Byte
                    
                    Ret = recv(S, Tmp(0), SIZE_RCVBUF, 0)
                    If Ret < 0 Then
                        UltError = err.LastDllError
                        If UltError = WSAEMSGSIZE Then
                            Debug.Print "WSAEMSGSIZE"
                            Ret = SIZE_RCVBUF
                        Else
                            
                            Call LogApiSock("Error en Recv: N=" & N & " S=" & S & " Str=" & GetWSAErrorString(UltError))
                            Call CloseSocketSL(N)
                            Call Cerrar_Usuario(N)
                            Exit Function
                        End If
                    ElseIf Ret = 0 Then
                        Call CloseSocketSL(N)
                        Call Cerrar_Usuario(N)
                    End If
                    ReDim Preserve Tmp(Ret - 1) As Byte
                   
                    Call EventoSockRead(N, Tmp)
                    
                Case FD_CLOSE
                    N = BuscaSlotSock(S)
                                  
                    If S <> SockListen Then
                        Call apiclosesocket(S)
    
                    End If
                    
                        If N > 0 Then
                        
                            Call BorraSlotSock(S)
                            UserList(N).ConnID = -1
                            UserList(N).ConnIDValida = False
                            Call EventoSockClose(N)
                            
                            '////Modificacion por Castelli
                            'Al final de la DESCONEXION se aclara que la_
                            'ID debe ser invalida y -1 asi la conexion se
                            'puede volver a utilizar y no que se use un idex
                            'cada vez mas alto como sucedia y lograr asi
                            'saturar la conexion como sucedia antes :S
                            'pero por suerte lo arregle :P :D
                            'UserList(N).ConnID = -1
                            'UserList(N).ConnIDValida = False
                            'Sacado porq se puso en el Sub resetuserslot
                            'debajo del zeromemory
                            '////Modificacion por Castelli
                            
                        End If
                End Select
            Case Else

            WndProc = CallWindowProc(OldWProc, hWnd, msg, wParam, lParam)
    End Select

End Function

Public Function WsApiEnviar(ByVal Slot As Integer, ByRef str As String) As Long
    

    
    Dim Ret As String
    Dim UltError As Long
    Dim Retorno As Long
    Dim data() As Byte

    ReDim Preserve data(Len(str) - 1) As Byte

    data = StrConv(str, vbFromUnicode)

    Retorno = 0

    If UserList(Slot).ConnID <> -1 And UserList(Slot).ConnIDValida Then
        Ret = send(ByVal UserList(Slot).ConnID, data(0), ByVal UBound(data()) + 1, ByVal 0)
        If Ret < 0 Then
            UltError = err.LastDllError
            If UltError = WSAEWOULDBLOCK Then
                Call UserList(Slot).outgoingData.WriteASCIIStringFixed(str)
            End If
            Retorno = UltError
        End If
    ElseIf UserList(Slot).ConnID <> -1 And Not UserList(Slot).ConnIDValida Then
        If Not UserList(Slot).Counters.Saliendo Then
            Retorno = -1
        End If
    End If

    

    WsApiEnviar = Retorno
End Function

Public Sub LogApiSock(ByVal str As String)
On Error GoTo Errhandler
    Dim nfile As Integer
    
    nfile = FreeFile()
    Open App.Path & "\logs\wsapi.log" For Append Shared As #nfile
        Print #nfile, Date & " " & time & " " & str
    Close #nfile

    Exit Sub

Errhandler:

End Sub

Public Sub EventoSockAccept(ByVal SockID As Long)
    Dim NewIndex As Integer
    Dim Ret As Long
    Dim Tam As Long, sa As sockaddr
    Dim NuevoSock As Long
    Dim i As Long
    Dim tStr As String
    
    Tam = sockaddr_size

    Ret = accept(SockID, sa, Tam)

    If Ret = INVALID_SOCKET Then
        i = err.LastDllError
        Call LogCriticEvent("Error en Accept() API " & i & ": " & GetWSAErrorString(i))
        Exit Sub
    End If
    
    If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
        Call WSApiCloseSocket(NuevoSock)
        Exit Sub
    End If

    NuevoSock = Ret


    
    'Seteamos el tamaño del buffer de entrada
    If setsockopt(NuevoSock, SOL_SOCKET, SO_RCVBUFFER, SIZE_RCVBUF, 4) <> 0 Then
        i = err.LastDllError
        Call LogCriticEvent("Error al setear el tamaño del buffer de entrada " & i & ": " & GetWSAErrorString(i))
    End If
    'Seteamos el tamaño del buffer de salida
    If setsockopt(NuevoSock, SOL_SOCKET, SO_SNDBUFFER, SIZE_SNDBUF, 4) <> 0 Then
        i = err.LastDllError
        Call LogCriticEvent("Error al setear el tamaño del buffer de salida " & i & ": " & GetWSAErrorString(i))
    End If







    'Mariano: Baje la busqueda de slot abajo de CondicionSocket y limite x ip
    NewIndex = NextOpenUser() ' Nuevo indice

Debug.Print "last: " & LastUser
Debug.Print "new: " & NewIndex


    If NewIndex <= MaxUsers Then
        
        
        
        'Make sure both outgoing and incoming data buffers are clean
        Call UserList(NewIndex).incomingData.ReadASCIIStringFixed(UserList(NewIndex).incomingData.length)
        Call UserList(NewIndex).outgoingData.ReadASCIIStringFixed(UserList(NewIndex).outgoingData.length)
        
        UserList(NewIndex).ip = GetAscIP(sa.sin_addr)
        'Busca si esta banneada la ip
        For i = 1 To BanIps.count
            If BanIps.Item(i) = UserList(NewIndex).ip Then
                Call WriteErrorMsg(NewIndex, "Su IP se encuentra bloqueada en este servidor.")
                Call FlushBuffer(NewIndex)
                Call WSApiCloseSocket(NuevoSock)
                Exit Sub
            End If
        Next i
        
        If NewIndex > LastUser Then LastUser = NewIndex
        
        UserList(NewIndex).ConnID = NuevoSock

        UserList(NewIndex).ConnIDValida = True
        
        Call AgregaSlotSock(NuevoSock, NewIndex)

    Else
        Dim str As String
        Dim data() As Byte
        
        str = Protocol.PrepareMessageErrorMsg("El server se encuentra lleno en este momento. Disculpe las molestias ocasionadas.")
        
        ReDim Preserve data(Len(str) - 1) As Byte
        
        data = StrConv(str, vbFromUnicode)

        Call send(ByVal NuevoSock, data(0), ByVal UBound(data()) + 1, ByVal 0)
        Call WSApiCloseSocket(NuevoSock)
    End If

End Sub

Public Sub EventoSockRead(ByVal Slot As Integer, ByRef Datos() As Byte)



    With UserList(Slot)
    
    
           'SEGURIDAD ////////////
    
    
       '    If .flags.UserLogged Then
       '     Security.NAC_D_Byte Datos, UserList(Slot).Redundance
       ' Else
       '     Security.NAC_D_Byte Datos, 13 'DEFAULT
       ' End If
    
        
        'SEGURIDAD ////////////
        
        
    
        Call .incomingData.WriteBlock(Datos)

        If .ConnID <> -1 Then
            Call HandleIncomingData(Slot)
        Else
            Exit Sub
        End If
    End With
End Sub

Public Sub EventoSockClose(ByVal Slot As Integer)
        
    If UserList(Slot).flags.UserLogged = True Then
        Call CloseSocketSL(Slot)
        Cerrar_Usuario Slot
    Else
        CloseSocket Slot
        
    End If
End Sub



Public Sub WSApiCloseSocket(ByVal Socket As Long)
    Call WSAAsyncSelect(Socket, hWndMsg, ByVal 1025, ByVal (FD_CLOSE))
    Call ShutDown(Socket, SD_BOTH)
End Sub

Public Function CondicionSocket(ByRef lpCallerId As WSABUF, ByRef lpCallerData As WSABUF, ByRef lpSQOS As FLOWSPEC, ByVal Reserved As Long, ByRef lpCalleeId As WSABUF, ByRef lpCalleeData As WSABUF, ByRef Group As Long, ByVal dwCallbackData As Long) As Long
    Dim sa As sockaddr
    
    'Check if we were requested to force reject

    If dwCallbackData = 1 Then
        CondicionSocket = CF_REJECT
        Exit Function
    End If
    
     'Get the address

    CopyMemory sa, ByVal lpCallerId.lpBuffer, lpCallerId.dwBufferLen

   
    If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
        CondicionSocket = CF_REJECT
        Exit Function
    End If

    CondicionSocket = CF_ACCEPT 'En realdiad es al pedo, porque CondicionSocket se inicializa a 0, pero así es más claro....
End Function


