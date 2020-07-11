Attribute VB_Name = "Arenas"
'Add Marius Duelos con arenas
Option Explicit

'Add Marius
Public arenas_estado As Boolean
Public Type Arena
    Estado As Boolean
    Pj1 As Integer
    Pj2 As Integer
    
    tipo As Byte
    
    Pos1 As WorldPos
    Pos2 As WorldPos
    
    Tiempo As Byte
End Type
Public Const MAX_ARENAS As Byte = 20
Public Arenas(1 To MAX_ARENAS) As Arena

Public Const MAX_ARENA_ESPERA As Integer = 20
Public arena_espera(1 To MAX_ARENA_ESPERA) As Integer

Public Const MAX_DUELO_ESPERA As Integer = 20
Public duelo_espera(1 To MAX_DUELO_ESPERA) As Integer
'\Add

Public Sub arenas_iniciar()
On Error Resume Next
Dim i As Integer
Dim a As Integer
Dim map As Integer
    
    For i = 1 To MAX_ARENA_ESPERA
        arena_espera(i) = 0
    Next
    
    For i = 1 To MAX_ARENAS
        Arenas(i).Estado = False
        Arenas(i).Pj1 = 0
        Arenas(i).Pj2 = 0
        Arenas(i).Tiempo = 0
        Arenas(i).tipo = 0
    Next
    
    For i = 1 To 5
        If i = 1 Then
            map = 238
            MapInfo(map).Pk = False
        Else ' Cuando se saque el mapa 238 sacar tambien el -1
            map = 851 + (i - 1)
        End If
        
        MapInfo(map).InviSinEfecto = True
        'MapInfo(map).Pk = False
        
        a = a + 1
        Arenas(a).Pos1.map = map
        Arenas(a).Pos1.x = 14
        Arenas(a).Pos1.Y = 77
        Arenas(a).Pos2.map = map
        Arenas(a).Pos2.x = 28
        Arenas(a).Pos2.Y = 87
        
        a = a + 1
        Arenas(a).Pos1.map = map
        Arenas(a).Pos1.x = 72
        Arenas(a).Pos1.Y = 77
        Arenas(a).Pos2.map = map
        Arenas(a).Pos2.x = 86
        Arenas(a).Pos2.Y = 87
        
        a = a + 1
        Arenas(a).Pos1.map = map
        Arenas(a).Pos1.x = 72
        Arenas(a).Pos1.Y = 14
        Arenas(a).Pos2.map = map
        Arenas(a).Pos2.x = 86
        Arenas(a).Pos2.Y = 24
        
        a = a + 1
        Arenas(a).Pos1.map = map
        Arenas(a).Pos1.x = 14
        Arenas(a).Pos1.Y = 14
        Arenas(a).Pos2.map = map
        Arenas(a).Pos2.x = 28
        Arenas(a).Pos2.Y = 24
    Next
    
End Sub

Public Sub arenas_Abrir()
On Error Resume Next
   
    arenas_estado = True
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(252, NO_3D_SOUND, NO_3D_SOUND)) 'Cuerno
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> La Arena 1vs1 esta abierta para participar ingresa al evento haganlo mediante el boton Eventos de su Menu.", FontTypeNames.FONTTYPE_SERVER))

End Sub

Public Sub arenas_Cerrar()
On Error Resume Next
Dim i As Integer
    
    arenas_estado = False
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> Arena 1vs1 cerrada.", FontTypeNames.FONTTYPE_SERVER))

    Call cerrar_arena(5)
    
    For i = 1 To MAX_ARENA_ESPERA
        
        If arena_espera(i) <> 0 Then
            Call WriteConsoleMsg(1, arena_espera(i), "La arena ya cerró, te quedaste sin pelear.", FontTypeNames.FONTTYPE_TALK)
            arena_espera(i) = 0
        End If
        
    Next

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''Funciones genericas'''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Sum(ByVal UserIndex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal fx As Boolean)
On Error Resume Next
Dim NuevaPos As WorldPos
Dim FuturePos As WorldPos
    
    FuturePos.map = map
    FuturePos.x = x
    FuturePos.Y = Y
    
    If UserIndex <> 0 Then
        Call ClosestLegalPos(FuturePos, NuevaPos)
        If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then
            Call WarpUserChar(UserIndex, NuevaPos.map, NuevaPos.x, NuevaPos.Y, fx)
        Else
            Call WarpUserChar(UserIndex, FuturePos.map, FuturePos.x, FuturePos.Y, fx)
        End If
    End If
End Sub

Public Function entrar_arena(Lista() As Integer) As Integer
On Error Resume Next
Dim i As Integer
    
    entrar_arena = 0
    If Lista(1) <> 0 And Lista(2) <> 0 Then
        
        Call salir_arena(Lista(1))
        Call salir_arena(Lista(2))
    
        For i = 1 To MAX_ARENAS
            If Not Arenas(i).Estado Then
                
                Arenas(i).Estado = True
                
                Arenas(i).Pj1 = Lista(1)
                Arenas(i).Pj2 = Lista(2)
                
                Arenas(i).tipo = UserList(Lista(1)).evento
                
                Call WarpUserChar(Arenas(i).Pj1, Arenas(i).Pos1.map, Arenas(i).Pos1.x, Arenas(i).Pos1.Y, True)
                Call WarpUserChar(Arenas(i).Pj2, Arenas(i).Pos2.map, Arenas(i).Pos2.x, Arenas(i).Pos2.Y, True)
                
                Lista(1) = 0
                Lista(2) = 0
                
                Call WriteConsoleMsg(1, Arenas(i).Pj1, "Comienza la pelea!.", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(1, Arenas(i).Pj2, "Comienza la pelea!.", FontTypeNames.FONTTYPE_TALK)
                
                Call BubbleSort(Lista)
                
                entrar_arena = i
                Exit Function
            End If
        Next
    End If
    
End Function

Public Function salir_arena(ByVal UserIndex As Integer) As Integer
On Error Resume Next
Dim i As Integer
Dim tipo As Byte

    salir_arena = 0
    For i = 1 To MAX_ARENAS
        If Arenas(i).Pj1 = UserIndex Or Arenas(i).Pj2 = UserIndex Then
            
            '¿Quien carajo ganó?
            If Arenas(i).Pj1 = UserIndex Then
                Call WriteConsoleMsg(1, Arenas(i).Pj1, "Perdiste la pelea!.", FontTypeNames.FONTTYPE_TALK)
                Call RevivirUsuario(Arenas(i).Pj1)
            Else
                Call WriteConsoleMsg(1, Arenas(i).Pj1, "Ganaste la pelea!.", FontTypeNames.FONTTYPE_TALK)
                salir_arena = Arenas(i).Pj1
            End If
            
            'Seguramente con 1 if se podia hacer, pero ni ganas de pensar :P
            If Arenas(i).Pj2 = UserIndex Then
                Call WriteConsoleMsg(1, Arenas(i).Pj2, "Perdiste la pelea!.", FontTypeNames.FONTTYPE_TALK)
                Call RevivirUsuario(Arenas(i).Pj2)
            Else
                Call WriteConsoleMsg(1, Arenas(i).Pj2, "Ganaste la pelea!.", FontTypeNames.FONTTYPE_TALK)
                salir_arena = Arenas(i).Pj2
            End If
            
            tipo = Arenas(i).tipo
            
            Call Sum(Arenas(i).Pj1, 49, 49, 50, True)
            Call Sum(Arenas(i).Pj2, 49, 48, 50, True)
            
            'Resetemos la data
            Arenas(i).Estado = False
            Arenas(i).Tiempo = 0
            Arenas(i).Pj1 = 0
            Arenas(i).Pj2 = 0
            Arenas(i).tipo = 0
            
            If tipo = 1 Then
                Call Sum(salir_arena, 238, 51, 47, True)
                Call entra_lista_espera(salir_arena, torneo_espera)
                Call entrar_arena(torneo_espera)
                
                If Not torneo_motor() Then
                    Call Sum(salir_arena, 49, 50, 50, True)
                    'No mas premios
                    'Call agregarcreditos(salir_arena)
                    
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> " & UserList(salir_arena).Name & " Es el nuevo ganador del Torneo 1vs1.", FontTypeNames.FONTTYPE_SERVER))
                    'Call WriteConsoleMsg(1, salir_arena, "Se han acreditado 2 creditos en tu cuenta.", FontTypeNames.FONTTYPE_TALK)
                End If
                
            ElseIf tipo = 5 Then
                Call entrar_arena(arena_espera)
            ElseIf tipo = 7 Then
                Call entrar_arena(duelo_espera)
            End If

            
            Exit For
        End If
    Next
    

End Function

Public Sub entra_lista_espera(ByVal UserIndex As Integer, Lista() As Integer)
On Error Resume Next
Dim i As Integer
Dim Slot As Integer
    
    Slot = 0
    
    For i = 1 To UBound(Lista)
        If Lista(i) = UserIndex Then
            Call WriteConsoleMsg(1, UserIndex, "Ya estas en la lista de espera. Ten paciencia.", FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        End If
    Next
    
    For i = 1 To UBound(Lista)
        If Lista(i) = 0 Then
            Slot = i
            Exit For
        End If
    Next
    
    If Slot <> 0 Then
        Lista(Slot) = UserIndex
        Call WriteConsoleMsg(1, UserIndex, "Entraste a la lista de espera, por favor aguarde.", FontTypeNames.FONTTYPE_TALK)
    Else
        Call WriteConsoleMsg(1, UserIndex, "La lista de espera esta llena, intentalo dentro de unos minutos.", FontTypeNames.FONTTYPE_TALK)
    End If

End Sub

Public Sub salir_lista_espera(ByVal UserIndex As Integer, Lista() As Integer)
On Error Resume Next
Dim i As Integer
    
    For i = 1 To UBound(Lista)
        If Lista(i) = UserIndex Then
            Lista(i) = 0
            Call WriteConsoleMsg(1, UserIndex, "Saliste de la lista de espera.", FontTypeNames.FONTTYPE_TALK)
            Exit For
        End If
    Next
    
    Call BubbleSort(Lista)
End Sub

Public Sub salir_listas_espera(ByVal UserIndex As Integer)
On Error Resume Next
    Call salir_lista_espera(UserIndex, arena_espera)
    Call salir_lista_espera(UserIndex, duelo_espera)
    Call salir_lista_espera(UserIndex, torneo_espera)
End Sub

Public Sub cerrar_arena(tipo As Byte)
On Error Resume Next
Dim i As Integer
    
    For i = 1 To MAX_ARENAS
        
        If Arenas(i).tipo = tipo Then 'Solo saca de la arena a los que son de este evento
            Arenas(i).Estado = False
            
            If Arenas(i).Pj1 <> 0 Then
                Call Sum(Arenas(i).Pj1, 49, 50, 50, True)
            End If
            
            If Arenas(i).Pj2 <> 0 Then
                Call Sum(Arenas(i).Pj2, 49, 50, 50, True)
            End If
            
            Arenas(i).Pj1 = 0
            Arenas(i).Pj2 = 0
            Arenas(i).Tiempo = 0
            Arenas(i).tipo = 0
        End If
    Next

End Sub

Public Function BubbleSort(Lista() As Integer) As Integer
On Error Resume Next
Dim p As Integer
Dim c As Integer
Dim h As Integer
    
    BubbleSort = 0
    For p = 1 To (UBound(Lista) - 1)
        For c = 1 To (UBound(Lista) - 1)
            If Lista(c) = 0 Or Lista(c + 1) = 0 Then
                If Lista(c) < Lista(c + 1) Then
                    h = Lista(c)
                    Lista(c) = Lista(c + 1)
                    Lista(c + 1) = h
                End If
            Else
                BubbleSort = BubbleSort + 1
            End If
        Next c
    Next p
End Function
'\Add

Public Function torneo_motor() As Boolean
Dim i As Integer
On Error Resume Next
    
    torneo_motor = False
    For i = 1 To MAX_ARENAS
        If Arenas(i).tipo = 1 Then
            torneo_motor = True
        End If
    Next
    
    If Not torneo_motor And torneo_espera(2) = 0 Then
        torneo_motor = False
    End If
    
End Function
