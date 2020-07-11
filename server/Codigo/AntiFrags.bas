Attribute VB_Name = "AntiFrags"
Option Explicit
'Si se aumenta el nimero de max buffer killeds hay que agregar campos a la Db no sea boludo como yo que tarde 15 minutos
'en darme cuenta por que carajo tiraba error la carga del pj xD (Si son las 4 de la mañana, andate a dormir, no programes que desprogramas) xD
Public Const MAX_BUFFER_KILLEDS As Byte = 5
Public Const INTERVALOxMUERTE As Long = 75219
 
Function YaMatoUsuario(ByVal MuertoIndex As Integer, ByVal UserIndex As Integer) As Byte
On Error Resume Next
Dim loopC       As Byte
Dim ContMuertes As Byte
     
    YaMatoUsuario = 0
     
    With UserList(UserIndex)
        'Recorro los muertos.
        For loopC = 1 To MAX_BUFFER_KILLEDS
            'If UserList(MuertoIndex).Indexpj = .Matados(loopC) Then
            'Des Marius Experimental
            If UserList(MuertoIndex).IndexAccount = .Matados(loopC) Then
               ContMuertes = ContMuertes + 1
            End If
        Next loopC
    End With
    
    YaMatoUsuario = ContMuertes
    
    If ContMuertes > 2 Then
        Call SendData(SendTarget.ToADMINS, 0, PrepareMessageConsoleMsg(1, UserList(UserIndex).Name & " Mató " & ContMuertes & " al mismo usuario: " & UserList(MuertoIndex).Name, FontTypeNames.FONTTYPE_FIGHT))
    End If
 
End Function
 
Sub AgregarListaMuertos(ByVal MuertoIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next
Dim loopC   As Byte
 
    'Agarro, Corro todos los que tenia matados
    'A un slot abajo, y despues en el primero
    'Guardo este ultimo "NameMuerto"
     
    For loopC = 1 To MAX_BUFFER_KILLEDS
        With UserList(UserIndex)
            If loopC >= MAX_BUFFER_KILLEDS Then Exit For
            
            If GetTickCount > .Matados_timer(loopC + 1) Then
               .Matados(loopC + 1) = 0
               .Matados_timer(loopC + 1) = 0
            End If
            
            .Matados(loopC) = .Matados(loopC + 1)
            .Matados_timer(loopC) = .Matados_timer(loopC + 1)
        End With
    Next loopC
     
    'Con el bucle de arriba, tengo En el
    'Primer y segundo slot el mismo nombre,
    'Ahora seteo el primero con "NameMuerto"
    
    'UserList(UserIndex).Matados(MAX_BUFFER_KILLEDS) = UserList(MuertoIndex).Indexpj
    'Mod Marius Experimental, en vez de guardar el indexpj guardamos el id de la cuenta
    UserList(UserIndex).Matados(MAX_BUFFER_KILLEDS) = UserList(MuertoIndex).IndexAccount
    UserList(UserIndex).Matados_timer(MAX_BUFFER_KILLEDS) = GetTickCount + INTERVALOxMUERTE
End Sub
 
Sub LimpiarListaDeMuertos(ByVal UserIndex As Integer)
On Error Resume Next
Dim loopC   As Byte
 
    For loopC = 1 To MAX_BUFFER_KILLEDS
        UserList(UserIndex).Matados(loopC) = 0
        UserList(UserIndex).Matados_timer(loopC) = 0
    Next loopC
 
End Sub
