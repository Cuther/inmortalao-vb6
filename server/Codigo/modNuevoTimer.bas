Attribute VB_Name = "modNuevoTimer"
Option Explicit

'
' Las siguientes funciones devuelven TRUE o FALSE si el intervalo
' permite hacerlo. Si devuelve TRUE, setean automaticamente el
' timer para que no se pueda hacer la accion hasta el nuevo ciclo.
'




Public Function IntervaloPermiteGolpeUsar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'***************************************************
'Author: ZaMa
'Checks if the time that passed from the last hit is enough for the user to use a potion.
'Last Modification: 06/04/2009
'***************************************************

Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerGolpeUsar >= IntervaloGolpeUsar Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerGolpeUsar = TActual
        End If
        IntervaloPermiteGolpeUsar = True
    Else
        IntervaloPermiteGolpeUsar = False
    End If
    
End Function




' TRABAJO
Public Function IntervaloPermiteTrabajar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerPuedeTrabajar >= IntervaloUserPuedeTrabajar Then  ' Then
        If Actualizar Then UserList(UserIndex).Counters.TimerPuedeTrabajar = TActual
        IntervaloPermiteTrabajar = True
    Else
        IntervaloPermiteTrabajar = False
    End If
End Function

' USAR OBJETOS
Public Function IntervaloPermiteUsar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerUsar >= IntervaloUserPuedeUsar Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerUsar = TActual
            UserList(UserIndex).Counters.failedUsageAttempts = 0
        End If
        IntervaloPermiteUsar = True
    Else
        IntervaloPermiteUsar = False
        
        UserList(UserIndex).Counters.failedUsageAttempts = UserList(UserIndex).Counters.failedUsageAttempts + 1
        
        'Tolerancia arbitraria - 20 es MUY alta, la está chiteando zarpado
        If UserList(UserIndex).Counters.failedUsageAttempts = 20 Then
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(1, UserList(UserIndex).Name & " kicked by the server por posible modificación de intervalos.", FontTypeNames.FONTTYPE_FIGHT))
            
            Call CloseSocket(UserIndex)
        End If
    End If

End Function

Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= IntervaloUserPuedeAtacar Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
            UserList(UserIndex).Counters.failedUsageAttempts = 0
        End If
        IntervaloPermiteAtacar = True
    Else
        IntervaloPermiteAtacar = False
        
        UserList(UserIndex).Counters.failedUsageAttempts = UserList(UserIndex).Counters.failedUsageAttempts + 1
        
        'Tolerancia arbitraria - 20 es MUY alta, la está chiteando zarpado
        If UserList(UserIndex).Counters.failedUsageAttempts = 20 Then
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(1, UserList(UserIndex).Name & " kicked by the server por posible modificación de intervalos.", FontTypeNames.FONTTYPE_FIGHT))
            
            Call CloseSocket(UserIndex)
        End If
    End If

End Function






























' CASTING DE HECHIZOS
Public Function IntervaloPermiteLanzarSpell(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - UserList(UserIndex).Counters.TimerLanzarSpell >= IntervaloUserPuedeCastear Then
    If Actualizar Then
        UserList(UserIndex).Counters.TimerLanzarSpell = TActual
    End If
    IntervaloPermiteLanzarSpell = True
Else
    IntervaloPermiteLanzarSpell = False
End If

End Function



Public Function IntervaloPermiteGolpeMagia(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim TActual As Long
    
    If UserList(UserIndex).Counters.TimerGolpeMagia > UserList(UserIndex).Counters.TimerPuedeAtacar Then
        Exit Function
    End If
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= IntervaloGolpeMagia Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerGolpeMagia = TActual
            UserList(UserIndex).Counters.TimerLanzarSpell = TActual
        End If
        IntervaloPermiteGolpeMagia = True
    Else
        IntervaloPermiteGolpeMagia = False
    End If
End Function

Public Function IntervaloPermiteUsarArcos(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerPuedeUsarArco >= IntervaloFlechasCazadores Then
        If Actualizar Then UserList(UserIndex).Counters.TimerPuedeUsarArco = TActual
        IntervaloPermiteUsarArcos = True
    Else
        IntervaloPermiteUsarArcos = False
    End If

End Function

Public Function IntervaloPermiteMover(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerMove >= IntervaloUserMove Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerMove = TActual
            'UserList(UserIndex).Counters.failedUsageAttempts = 0
        End If
        IntervaloPermiteMover = True
    Else
        IntervaloPermiteMover = False
        
        'UserList(UserIndex).Counters.failedUsageAttempts = UserList(UserIndex).Counters.failedUsageAttempts + 1
        
        'Tolerancia arbitraria - 20 es MUY alta, la está chiteando zarpado
        'If UserList(UserIndex).Counters.failedUsageAttempts = 20 Then
        '    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(1, UserList(UserIndex).Name & " kicked by the server por posible modificación de intervalos.", FontTypeNames.FONTTYPE_FIGHT))
        '
        '    Call CloseSocket(UserIndex)
        'End If
    End If

End Function
