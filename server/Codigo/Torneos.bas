Attribute VB_Name = "Torneos"
'Add Marius Duelos con arenas
Option Explicit

Public Torneo_estado As Byte
Public Torneo_maximo As Byte
Public Torneo_cont As Integer

Public Const MAX_TORNEO_ESPERA As Integer = 50
Public torneo_espera(1 To MAX_TORNEO_ESPERA) As Integer

Public Sub torneo_iniciar(ByVal maximo As Integer)
On Error Resume Next
    Torneo_estado = 1
    Torneo_maximo = maximo
    Torneo_cont = 0
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(252, NO_3D_SOUND, NO_3D_SOUND)) 'Cuerno
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> Inician las inscripciones para el Torneo 1vs1, sale 100k y hay " & maximo & " cupos, para participar entra desde el boton Eventos de tu Menu.", FontTypeNames.FONTTYPE_SERVER))
End Sub

Public Sub torneo_terminar()
On Error Resume Next
Dim i As Integer
    Torneo_estado = 0
    Torneo_maximo = 0
    Torneo_cont = 0
    
    Call cerrar_arena(1)
    
    For i = 1 To MAX_TORNEO_ESPERA
        
        If torneo_espera(i) <> 0 Then
            Call WriteConsoleMsg(1, torneo_espera(i), "El torneo se canceló, te quedaste sin pelear.", FontTypeNames.FONTTYPE_TALK)
            UserList(torneo_espera(i)).Stats.GLD = UserList(torneo_espera(i)).Stats.GLD + 100000
            Call WriteUpdateGold(torneo_espera(i))
            Call Sum(torneo_espera(i), 49, 48, 50, True)
            torneo_espera(i) = 0
        End If
        
    Next
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> Se canceló el Torneo 1vs1.", FontTypeNames.FONTTYPE_SERVER))
End Sub

Public Sub torneo_entrar(ByVal UserIndex As Integer)
On Error Resume Next
    If Torneo_estado = 2 Then
        Call WriteConsoleMsg(1, UserIndex, "El torneo ya cerró las inscripciones.", FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    ElseIf UserList(UserIndex).Stats.GLD < 100000 Then ' Si no tiene plata el pobre se queda afuera
        Call WriteConsoleMsg(1, UserIndex, "No tiene es dinero suficiente para la inscripcion.", FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    ElseIf Torneo_cont > Torneo_maximo Then
        Torneo_estado = 2
        Call WriteConsoleMsg(1, UserIndex, "El torneo se quedó sin cupos.", FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    End If
    
    Torneo_cont = Torneo_cont + 1
    
    Call entra_lista_espera(UserIndex, torneo_espera)
    Call Sum(UserIndex, 238, 51, 47, True)
    
    'Le sacamos la plata
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 100000
    Call WriteUpdateGold(UserIndex)
    
    Call WriteConsoleMsg(1, UserIndex, "Aguarde mientras es llamado a la Arena.", FontTypeNames.FONTTYPE_TALK)
    
    If Torneo_cont = Torneo_maximo Then
        Torneo_estado = 2
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> El torneo se quedó sin cupos.", FontTypeNames.FONTTYPE_SERVER))
        'Que loco no?
        If entrar_arena(torneo_espera) <> 0 And entrar_arena(torneo_espera) <> 0 And entrar_arena(torneo_espera) <> 0 Then  'Vemos si hay alguien mas en espera y si hay lo mandamos adentro de la arena
            Call entrar_arena(torneo_espera)
        End If
    End If
End Sub

Public Sub torneo_sale(ByVal UserIndex As Integer)
Dim ganador As Integer
On Error Resume Next
    Torneo_cont = Torneo_cont - 1
    
    Call salir_lista_espera(UserIndex, torneo_espera)
    Call Sum(UserIndex, 49, 50, 50, True)

    ganador = salir_arena(UserIndex)
    
    'Los llevamos a intermundia
    Call Sum(UserIndex, 49, 48, 50, True)
    Call Sum(ganador, 238, 51, 47, True)
    Call entra_lista_espera(ganador, torneo_espera)
    
    'Vemos si hay alguien mas en espera y si hay lo mandamos adentro de la arena
    Call entrar_arena(torneo_espera)
End Sub
