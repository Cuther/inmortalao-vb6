Attribute VB_Name = "mdlCOmercioConUsuario"

'[Alejo]
Option Explicit

Private Const MAX_ORO_LOGUEABLE As Long = 50000
Private Const MAX_OBJ_LOGUEABLE As Long = 1000

Public Type tCOmercioUsuario
    DestUsu As Integer
    DestNick As String
    Objeto As Integer
    Cant As Long
    Acepto As Boolean
End Type

'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Sub IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer)
On Error GoTo Errhandler

'Si ambos pusieron /comerciar entonces
If UserList(Origen).ComUsu.DestUsu = Destino And UserList(Destino).ComUsu.DestUsu = Origen Then
    
    If UserList(Origen).flags.Comerciando Or UserList(Destino).flags.Comerciando Then
        Call WriteConsoleMsg(1, Origen, "No puedes comerciar en este momento", FontTypeNames.FONTTYPE_TALK)
        Call WriteConsoleMsg(1, Destino, "No puedes comerciar en este momento", FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    End If
        
    'Actualiza el inventario del usuario
    Call UpdateUserInv(True, Origen, 0)
    'Decirle al origen que abra la ventanita.
    Call WriteUserCommerceInit(Origen)
    UserList(Origen).flags.Comerciando = True

    'Actualiza el inventario del usuario
    Call UpdateUserInv(True, Destino, 0)
    'Decirle al origen que abra la ventanita.
    Call WriteUserCommerceInit(Destino)
    UserList(Destino).flags.Comerciando = True

    'Call EnviarObjetoTransaccion(Origen)
Else
    'Es el primero que comercia ?
    Call WriteConsoleMsg(1, Destino, UserList(Origen).Name & " desea comerciar. Si deseas aceptar, Escribe /COMERCIAR.", FontTypeNames.FONTTYPE_TALK)
    UserList(Destino).flags.TargetUser = Origen
    
End If

Call FlushBuffer(Destino)

Exit Sub
Errhandler:
    Call LogError("Error en IniciarComercioConUsuario: " & err.description)
End Sub

'envia a AQuien el objeto del otro
Public Sub EnviarObjetoTransaccion(ByVal AQuien As Integer)
Dim ObjInd As Integer
Dim ObjCant As Long

'[Alejo]: En esta funcion se centralizaba el problema
'         de no poder comerciar con mas de 32k de oro.
'         Ahora si funciona!!!

    ObjCant = UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Cant
    If UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto = FLAGORO Then
        ObjInd = iORO
    Else
        ObjInd = UserList(UserList(AQuien).ComUsu.DestUsu).Invent.Object(UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto).ObjIndex
    End If
    
    If ObjCant <= 0 Or ObjInd <= 0 Then Exit Sub
    
    If ObjInd > 0 And ObjCant > 0 Then
        Call WriteChangeUserTradeSlot(AQuien, ObjInd, ObjCant)
        Call FlushBuffer(AQuien)
    End If

End Sub

Public Sub FinComerciarUsu(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If .ComUsu.DestUsu > 0 Then
            Call WriteUserCommerceEnd(UserIndex)
        End If
        
        .ComUsu.Acepto = False
        .ComUsu.Cant = 0
        .ComUsu.DestUsu = 0
        .ComUsu.Objeto = 0
        .ComUsu.DestNick = vbNullString
        .flags.Comerciando = False
    End With
End Sub

Public Sub AceptarComercioUsu(ByVal UserIndex As Integer)
Dim Obj1 As obj, Obj2 As obj
Dim OtroUserIndex As Integer
Dim TerminarAhora As Boolean

TerminarAhora = False

If UserList(UserIndex).ComUsu.DestUsu <= 0 Or UserList(UserIndex).ComUsu.DestUsu > MaxUsers Then
    TerminarAhora = True
End If

OtroUserIndex = UserList(UserIndex).ComUsu.DestUsu

If Not TerminarAhora Then
    If UserList(OtroUserIndex).flags.UserLogged = False Or UserList(UserIndex).flags.UserLogged = False Then
        TerminarAhora = True
    End If
End If

If Not TerminarAhora Then
    If UserList(OtroUserIndex).ComUsu.DestUsu <> UserIndex Then
        TerminarAhora = True
    End If
End If

If Not TerminarAhora Then
    If UserList(OtroUserIndex).Name <> UserList(UserIndex).ComUsu.DestNick Then
        TerminarAhora = True
    End If
End If

If Not TerminarAhora Then
    If UserList(UserIndex).Name <> UserList(OtroUserIndex).ComUsu.DestNick Then
        TerminarAhora = True
    End If
End If

If TerminarAhora = True Then
    Call FinComerciarUsu(UserIndex)
    
    If OtroUserIndex <= 0 Or OtroUserIndex > MaxUsers Then
        Call FinComerciarUsu(OtroUserIndex)
        Call Protocol.FlushBuffer(OtroUserIndex)
    End If
    
    Exit Sub
End If
    

UserList(UserIndex).ComUsu.Acepto = True
TerminarAhora = False

If UserList(OtroUserIndex).ComUsu.Acepto = False Then
    Call WriteConsoleMsg(1, UserIndex, "El otro usuario aun no ha aceptado tu oferta.", FontTypeNames.FONTTYPE_TALK)
    Exit Sub
End If


If UserList(UserIndex).ComUsu.Objeto = FLAGORO Then
    Obj1.ObjIndex = iORO
    If UserList(UserIndex).ComUsu.Cant > UserList(UserIndex).Stats.GLD Then
        Call WriteConsoleMsg(1, UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
        TerminarAhora = True
    End If
Else
    Obj1.Amount = UserList(UserIndex).ComUsu.Cant
    Obj1.ObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.Objeto).ObjIndex
    If Obj1.Amount > UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.Objeto).Amount Then
        Call WriteConsoleMsg(1, UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
        TerminarAhora = True
    End If
End If

If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
    Obj2.ObjIndex = iORO
    If UserList(OtroUserIndex).ComUsu.Cant > UserList(OtroUserIndex).Stats.GLD Then
        Call WriteConsoleMsg(1, OtroUserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
        TerminarAhora = True
    End If
Else
    Obj2.Amount = UserList(OtroUserIndex).ComUsu.Cant
    Obj2.ObjIndex = UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).ObjIndex
    If Obj2.Amount > UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).Amount Then
        Call WriteConsoleMsg(1, OtroUserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
        TerminarAhora = True
    End If
End If

'Por si las moscas...
If TerminarAhora = True Then
    Call FinComerciarUsu(UserIndex)
    
    Call FinComerciarUsu(OtroUserIndex)
    Call FlushBuffer(OtroUserIndex)
    Exit Sub
End If

Call FlushBuffer(OtroUserIndex)

'[CORREGIDO]
'Desde acá corregí el bug que cuando se ofrecian mas de
'10k de oro no le llegaban al destinatario.

'pone el oro directamente en la billetera
If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
    'quito la cantidad de oro ofrecida
    UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD - UserList(OtroUserIndex).ComUsu.Cant
    If UserList(OtroUserIndex).ComUsu.Cant > MAX_ORO_LOGUEABLE Then Call LogCheating(UserList(OtroUserIndex).Name & " solto oro en comercio seguro con " & UserList(UserIndex).Name & ". Cantidad: " & UserList(OtroUserIndex).ComUsu.Cant)
    Call WriteUpdateUserStats(OtroUserIndex)
    'y se la doy al otro
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(OtroUserIndex).ComUsu.Cant
    If UserList(OtroUserIndex).ComUsu.Cant > MAX_ORO_LOGUEABLE Then Call LogCheating(UserList(UserIndex).Name & " recibio oro en comercio seguro con " & UserList(OtroUserIndex).Name & ". Cantidad: " & UserList(OtroUserIndex).ComUsu.Cant)
    'Esta linea del log es al pedo. > Vuelvo a ponerla a pedido del CGMS
    Call WriteUpdateUserStats(UserIndex)
Else
    'Quita el objeto y se lo da al otro
    If MeterItemEnInventario(UserIndex, Obj2) = False Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, Obj2)
    End If
    Call QuitarObjetos(Obj2.ObjIndex, Obj2.Amount, OtroUserIndex)
End If

'pone el oro directamente en la billetera
If UserList(UserIndex).ComUsu.Objeto = FLAGORO Then
    'quito la cantidad de oro ofrecida
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - UserList(UserIndex).ComUsu.Cant
    If UserList(UserIndex).ComUsu.Cant > MAX_ORO_LOGUEABLE Then Call LogCheating(UserList(UserIndex).Name & " soltó oro en comercio seguro con " & UserList(OtroUserIndex).Name & ". Cantidad: " & UserList(UserIndex).ComUsu.Cant)
    Call WriteUpdateUserStats(UserIndex)
    'y se la doy al otro
    UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD + UserList(UserIndex).ComUsu.Cant
    'If UserList(UserIndex).ComUsu.cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(UserList(OtroUserIndex).name & " recibio oro en comercio seguro con " & UserList(UserIndex).name & ". Cantidad: " & UserList(UserIndex).ComUsu.cant)
    'Esta linea del log es al pedo.
    Call WriteUpdateUserStats(OtroUserIndex)
Else
    'Quita el objeto y se lo da al otro
    If MeterItemEnInventario(OtroUserIndex, Obj1) = False Then
        Call TirarItemAlPiso(UserList(OtroUserIndex).Pos, Obj1)
    End If
    Call QuitarObjetos(Obj1.ObjIndex, Obj1.Amount, UserIndex)
    
End If

'[/CORREGIDO] :p

Call UpdateUserInv(True, UserIndex, 0)
Call UpdateUserInv(True, OtroUserIndex, 0)

Call FinComerciarUsu(UserIndex)
Call FinComerciarUsu(OtroUserIndex)
 
End Sub

'[/Alejo]
