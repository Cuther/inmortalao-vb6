Attribute VB_Name = "InvUsuario"

Option Explicit

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean

'17/09/02
'Agregue que la función se asegure que el objeto no es un barco

On Error GoTo hayerror

Dim i As Integer
Dim ObjIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
    ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    If ObjIndex > 0 Then
           If (Not ItemSeCae(ObjIndex)) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    
    End If
Next i

Exit Function

hayerror:
    LogError ("Error en TieneObjetosRobables: " & err.description)



End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo manejador
Dim flag As Boolean

    If ObjIndex = 0 Then Exit Function
    
    'Admins can use ANYTHING!
    If UserList(UserIndex).flags.Privilegios And (PlayerType.User Or PlayerType.VIP) Then
        If ObjData(ObjIndex).ClaseTipo = 0 Then
            If ObjData(ObjIndex).ClaseProhibida(1) > 0 Then
                Dim i As Integer
                For i = 1 To NUMCLASES
                    If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).Clase Then
                        
                        ClasePuedeUsarItem = False
                        Exit Function
                    End If
                Next i
            End If
        Else
            If (UserList(UserIndex).Clase = Gladiador Or _
            UserList(UserIndex).Clase = Guerrero Or _
            UserList(UserIndex).Clase = Paladin Or _
            UserList(UserIndex).Clase = Cazador Or _
            UserList(UserIndex).Clase = Mercenario) Then
                If ObjData(ObjIndex).ClaseTipo = 1 Then
                    ClasePuedeUsarItem = True
                Else
                    ClasePuedeUsarItem = False
                End If
            End If
        End If
    End If
    
    ClasePuedeUsarItem = True

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
             
             If ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, UserIndex, j)
        
        End If
Next j

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)


Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
    UserList(UserIndex).Invent.Object(j).ObjIndex = 0
    UserList(UserIndex).Invent.Object(j).Amount = 0
    UserList(UserIndex).Invent.Object(j).Equipped = 0
Next j

UserList(UserIndex).Invent.NroItems = 0

UserList(UserIndex).Invent.NudiEqpSlot = 0
UserList(UserIndex).Invent.NudiEqpIndex = 0

UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
UserList(UserIndex).Invent.ArmourEqpSlot = 0

UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
UserList(UserIndex).Invent.WeaponEqpSlot = 0

UserList(UserIndex).Invent.CascoEqpObjIndex = 0
UserList(UserIndex).Invent.CascoEqpSlot = 0

UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
UserList(UserIndex).Invent.EscudoEqpSlot = 0

UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
UserList(UserIndex).Invent.AnilloEqpSlot = 0

UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
UserList(UserIndex).Invent.MunicionEqpSlot = 0

UserList(UserIndex).Invent.BarcoObjIndex = 0
UserList(UserIndex).Invent.BarcoSlot = 0

UserList(UserIndex).Invent.MonturaObjIndex = 0
UserList(UserIndex).Invent.MonturaSlot = 0

UserList(UserIndex).Invent.MagicIndex = 0
UserList(UserIndex).Invent.MagicSlot = 0
End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
'***************************************************
On Error GoTo Errhandler

'If Cantidad > 100000 Then Exit Sub

'SI EL Pjta TIENE ORO LO TIRAMOS
If (Cantidad > 0) And (Cantidad <= UserList(UserIndex).Stats.GLD) Then
        Dim i As Byte
        Dim MiObj As obj
        'info debug
        Dim loops As Integer
        
        'Des Nod Kopfncikend (Todo este quilombo para nada, por que despues Cercanos no se usa para nada) asi que lo saque, ahorramos procesos
        'Seguridad Alkon (guardo el oro tirado si supera los 50k)
        'If Cantidad > 50000 Then
        '    Dim j As Integer
        '    Dim k As Integer
        '    Dim M As Integer
        '    Dim Cercanos As String
        '    M = UserList(UserIndex).pos.map
        '    For j = UserList(UserIndex).pos.x - 10 To UserList(UserIndex).pos.x + 10
        '        For k = UserList(UserIndex).pos.Y - 10 To UserList(UserIndex).pos.Y + 10
        '            If InMapBounds(M, j, k) Then
        '                If MapData(M, j, k).UserIndex > 0 Then
        '                    Cercanos = Cercanos & UserList(MapData(M, j, k).UserIndex).Name & ","
        '                End If
        '            End If
        '        Next k
        '    Next j
        'End If
        '/Seguridad
        
        Dim Extra As Long
        Dim TeniaOro As Long
        TeniaOro = UserList(UserIndex).Stats.GLD
        
        If Cantidad > 500000 Then 'Para evitar explotar demasiado
            Extra = Cantidad - 500000
            Cantidad = 500000
        End If
        
        Do While (Cantidad > 0)
            
            If Cantidad > MAX_INVENTORY_OBJS And UserList(UserIndex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.Amount = MAX_INVENTORY_OBJS
                Cantidad = Cantidad - MiObj.Amount
            Else
                MiObj.Amount = Cantidad
                Cantidad = Cantidad - MiObj.Amount
            End If

            MiObj.ObjIndex = iORO
            
            Dim AuxPos As WorldPos
            
            
            If UserList(UserIndex).Clase = eClass.Mercenario And UserList(UserIndex).Invent.BarcoObjIndex = 476 Then
                AuxPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj, False)
                If AuxPos.x <> 0 And AuxPos.Y <> 0 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MiObj.Amount
                End If
            Else
                AuxPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj, True)
                If AuxPos.x <> 0 And AuxPos.Y <> 0 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MiObj.Amount
                End If
            End If
            
            'info debug
            loops = loops + 1
            If loops > 100 Then
                LogError ("Error en tiraroro")
                Exit Sub
            End If
            
        Loop
        
        If TeniaOro = UserList(UserIndex).Stats.GLD Then Extra = 0
        If Extra > 0 Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Extra
        End If
    
End If

Exit Sub

Errhandler:

End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
    If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    With UserList(UserIndex).Invent.Object(Slot)
        If .Amount <= Cantidad Then
            If .Equipped = 1 Then
                Call Desequipar(UserIndex, Slot)
            End If
        End If
        
        'Quita un objeto
        .Amount = .Amount - Cantidad
        '¿Quedan mas?
        If .Amount <= 0 Then
            UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
            .ObjIndex = 0
            .Amount = 0
        End If
    End With
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

Dim NullObj As UserObj
Dim loopC As Long

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Invent.Object(Slot).ObjIndex > 0 Then
        Call ChangeUserInv(UserIndex, Slot, UserList(UserIndex).Invent.Object(Slot))
    Else
        Call ChangeUserInv(UserIndex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For loopC = 1 To MAX_INVENTORY_SLOTS
        'Actualiza el inventario
        If UserList(UserIndex).Invent.Object(loopC).ObjIndex > 0 Then
            Call ChangeUserInv(UserIndex, loopC, UserList(UserIndex).Invent.Object(loopC))
        Else
            Call ChangeUserInv(UserIndex, loopC, NullObj)
        End If
    Next loopC
End If

End Sub

Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer)
    Dim obj As obj
    
    If num > 0 Then
        If num > UserList(UserIndex).Invent.Object(Slot).Amount Then num = UserList(UserIndex).Invent.Object(Slot).Amount
            
            'Check objeto en el suelo
            If MapData(UserList(UserIndex).Pos.map, x, Y).ObjInfo.ObjIndex = 0 Or MapData(UserList(UserIndex).Pos.map, x, Y).ObjInfo.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex Then
                obj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                
                If num + MapData(UserList(UserIndex).Pos.map, x, Y).ObjInfo.Amount > MAX_INVENTORY_OBJS Then
                    num = MAX_INVENTORY_OBJS - MapData(UserList(UserIndex).Pos.map, x, Y).ObjInfo.Amount
                End If
                
                obj.Amount = num
                
                Call MakeObj(obj, map, x, Y)
                Call QuitarUserInvItem(UserIndex, Slot, num)
                Call UpdateUserInv(False, UserIndex, Slot)
                
            Else
                Call WriteConsoleMsg(1, UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
            End If
    End If
End Sub

Sub EraseObj(ByVal num As Integer, ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer)

MapData(map, x, Y).ObjInfo.Amount = MapData(map, x, Y).ObjInfo.Amount - num

If MapData(map, x, Y).ObjInfo.Amount <= 0 Then
    MapData(map, x, Y).ObjInfo.ObjIndex = 0
    MapData(map, x, Y).ObjInfo.Amount = 0
    
    Call modSendData.SendToAreaByPos(map, x, Y, PrepareMessageObjectDelete(x, Y))
End If

End Sub

Sub MakeObj(ByRef obj As obj, ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer)

If obj.ObjIndex > 0 And obj.ObjIndex <= UBound(ObjData) Then

    If MapData(map, x, Y).ObjInfo.ObjIndex = obj.ObjIndex Then
        MapData(map, x, Y).ObjInfo.Amount = MapData(map, x, Y).ObjInfo.Amount + obj.Amount
    Else
        MapData(map, x, Y).ObjInfo = obj
        
        Call modSendData.SendToAreaByPos(map, x, Y, PrepareMessageObjectCreate(x, Y, obj.ObjIndex, obj.Amount))
    End If
End If

End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As obj) As Boolean
On Error GoTo Errhandler

'Call LogTarea("MeterItemEnInventario")
 
Dim x As Integer
Dim Y As Integer
Dim Slot As Byte

'¿el user ya tiene un objeto del mismo tipo?
Slot = 1
Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And _
         UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
   Slot = Slot + 1
   If Slot > MAX_INVENTORY_SLOTS Then
         Exit Do
   End If
Loop
    
'Sino busca un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
           Call WriteConsoleMsg(1, UserIndex, "No podes cargar mas objetos.", FontTypeNames.FONTTYPE_FIGHT)
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If
    
'Mete el objeto
If UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
   UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + MiObj.Amount
Else
   UserList(UserIndex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True
       
Call UpdateUserInv(False, UserIndex, Slot)


Exit Function
Errhandler:

End Function


Sub GetObj(ByVal UserIndex As Integer)

Dim obj As ObjData
Dim MiObj As obj
Dim ObjPos As String

'¿Hay algun obj?
If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex > 0 Then
    '¿Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then
        Dim x As Integer
        Dim Y As Integer
        Dim Slot As Byte
        
        x = UserList(UserIndex).Pos.x
        Y = UserList(UserIndex).Pos.Y
        obj = ObjData(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex)
        MiObj.Amount = MapData(UserList(UserIndex).Pos.map, x, Y).ObjInfo.Amount
        MiObj.ObjIndex = MapData(UserList(UserIndex).Pos.map, x, Y).ObjInfo.ObjIndex
        If MiObj.ObjIndex = 12 Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + MiObj.Amount
            Call WriteUpdateGold(UserIndex)
            Call EraseObj(MapData(UserList(UserIndex).Pos.map, x, Y).ObjInfo.Amount, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y)
        Else
            If MeterItemEnInventario(UserIndex, MiObj) Then
                Call EraseObj(MapData(UserList(UserIndex).Pos.map, x, Y).ObjInfo.Amount, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y)
            End If
        End If
        
    End If
Else
    Call WriteConsoleMsg(1, UserIndex, "No hay nada aquí.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
'Desequipa el item slot del inventario
Dim obj As ObjData


If (Slot < LBound(UserList(UserIndex).Invent.Object)) Or (Slot > UBound(UserList(UserIndex).Invent.Object)) Then
    Exit Sub
ElseIf UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0 Then
    Exit Sub
End If

obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)

Select Case obj.OBJType
    Case eOBJType.otMonturas
        Call DoEquita(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex, Slot)
        
    Case eOBJType.otHerramientas
        UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
        UserList(UserIndex).Invent.AnilloEqpSlot = 0
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        
        If UserList(UserIndex).flags.Trabajando = True Then
            UserList(UserIndex).flags.Trabajando = False
            
            Call WriteConsoleMsg(1, UserIndex, "Dejas de trabajar.", FontTypeNames.FONTTYPE_BROWNI)
        End If
        
    Case eOBJType.otWeapon
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
        UserList(UserIndex).Invent.WeaponEqpSlot = 0
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                
    Case eOBJType.otItemsMagicos
        UserList(UserIndex).Invent.MagicIndex = 0
        UserList(UserIndex).Invent.MagicSlot = 0
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        
        If obj.EfectoMagico = eMagicType.ModificaAtributo Then
            If obj.QueAtributo <> 0 Then
                UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
            End If
        ElseIf obj.EfectoMagico = eMagicType.ModificaSkill Then
            If obj.QueSkill <> 0 Then
                UserList(UserIndex).Stats.UserSkills(obj.QueSkill) = UserList(UserIndex).Stats.UserSkills(obj.QueSkill) - obj.CuantoAumento
            End If
        End If
        
    Case eOBJType.otNudillos
        UserList(UserIndex).Invent.NudiEqpIndex = 0
        UserList(UserIndex).Invent.NudiEqpSlot = 0
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                
    Case eOBJType.otFlechas
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
        UserList(UserIndex).Invent.MunicionEqpSlot = 0
        
    Case eOBJType.otArmadura ' Puede ser un escudo, casco , o vestimenta
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        
        Select Case ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).SubTipo
            Case 0
                UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
                UserList(UserIndex).Invent.ArmourEqpSlot = 0
                Call DarCuerpoDesnudo(UserIndex)
                If Not UserList(UserIndex).flags.Montando = 1 Then Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                
            Case 1
                UserList(UserIndex).Invent.Object(Slot).Equipped = 0
                UserList(UserIndex).Invent.CascoEqpObjIndex = 0
                UserList(UserIndex).Invent.CascoEqpSlot = 0
                
                UserList(UserIndex).Char.CascoAnim = NingunCasco
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

                
            Case 2
                UserList(UserIndex).Invent.Object(Slot).Equipped = 0
                UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
                UserList(UserIndex).Invent.EscudoEqpSlot = 0
                
                UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                
        End Select
End Select

Call WriteUpdateUserStats(UserIndex)
Call UpdateUserInv(False, UserIndex, Slot)

End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo Errhandler

If Not ObjIndex <> 0 Then Exit Function



If ObjData(ObjIndex).Mujer = 1 Then
    SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Hombre
ElseIf ObjData(ObjIndex).Hombre = 1 Then
    SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Mujer
Else
    SexoPuedeUsarItem = True
End If

Exit Function
Errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean

    If Not ObjIndex <> 0 Then Exit Function
    
    If ObjData(ObjIndex).Real > 0 Then
        FaccionPuedeUsarItem = esArmada(UserIndex)
    ElseIf ObjData(ObjIndex).Caos > 0 Then
        FaccionPuedeUsarItem = esCaos(UserIndex)
    ElseIf ObjData(ObjIndex).Milicia > 0 Then
        FaccionPuedeUsarItem = esMili(UserIndex)
    Else
        FaccionPuedeUsarItem = True
    End If

End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
On Error GoTo Errhandler

'Equipa un item del inventario
Dim obj As ObjData
Dim ObjIndex As Integer

ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
obj = ObjData(ObjIndex)

If Not obj.MinELV = 0 Then
    If obj.MinELV > UserList(UserIndex).Stats.ELV Then
        Call WriteConsoleMsg(1, UserIndex, "Debes ser nivel " & obj.MinELV & " para poder utilizar este objeto.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
End If


If obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
     Call WriteConsoleMsg(1, UserIndex, "Solo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
     Exit Sub
End If
        
Select Case obj.OBJType
    Case eOBJType.otMonturas
        If UserList(UserIndex).flags.Metamorfosis = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "Transformado no puedes utilizar este objeto. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        
        Call DoEquita(UserIndex, ObjIndex, Slot)
    Case eOBJType.otHerramientas
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        
        If UserList(UserIndex).Invent.Object(Slot).Equipped Then
            Call Desequipar(UserIndex, Slot)
            Exit Sub
        End If
        
        If UserList(UserIndex).Invent.AnilloEqpSlot <> 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.AnilloEqpSlot)
        End If
        
        UserList(UserIndex).Invent.AnilloEqpSlot = Slot
        UserList(UserIndex).Invent.AnilloEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        UserList(UserIndex).Invent.Object(Slot).Equipped = 1
        
    Case eOBJType.otItemsMagicos
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        
        If UserList(UserIndex).Invent.Object(Slot).Equipped Then
            'Quitamos del inv el item
            Call Desequipar(UserIndex, Slot)
            Exit Sub
        End If
        
        If UserList(UserIndex).Invent.MagicIndex <> 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.MagicSlot)
        End If
        
        UserList(UserIndex).Invent.MagicIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        UserList(UserIndex).Invent.MagicSlot = Slot
        UserList(UserIndex).Invent.Object(Slot).Equipped = 1
        
        If obj.EfectoMagico = eMagicType.ModificaAtributo Then
            If obj.QueAtributo <> 0 Then
                UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) + obj.CuantoAumento
            End If
        ElseIf obj.EfectoMagico = eMagicType.ModificaSkill Then
            If obj.QueSkill <> 0 Then
                UserList(UserIndex).Stats.UserSkills(obj.QueSkill) = UserList(UserIndex).Stats.UserSkills(obj.QueSkill) + obj.CuantoAumento
            End If
        End If
        
    Case eOBJType.otFlechas
       If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        
       If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.MunicionEqpSlot = Slot
                
       Else
            Call WriteConsoleMsg(1, UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
    Case eOBJType.otArmadura
    
        If UserList(UserIndex).flags.Metamorfosis = 1 Then Exit Sub
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        

        If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).SubTipo = 0 Then
            If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
            'Nos aseguramos que puede usarla
            If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
               SexoPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
               CheckRazaUsaRopa(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
               FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
               
               'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot)
                    Call DarCuerpoDesnudo(UserIndex)
                    If Not UserList(UserIndex).flags.Montando = 1 Then Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
        
                'Quita el anterior
                If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
                End If
        
                'Lo equipa
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.ArmourEqpSlot = Slot
                
                'Sonido
                If ObjIndex = 1093 Then 'Armadura de Placas Dorada (RM +10)
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(248, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                End If

                UserList(UserIndex).Char.body = obj.Ropaje
                If Not UserList(UserIndex).flags.Montando = 1 Then Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

                UserList(UserIndex).flags.Desnudo = 0
            Else
                Call WriteConsoleMsg(1, UserIndex, "Tu clase, genero o raza no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
            End If
        ElseIf ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).SubTipo = 1 Then
            If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
            If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot)
                    
                    UserList(UserIndex).Char.CascoAnim = NingunCasco
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
        
                'Quita el anterior
                If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
                End If
        
                'Lo equipa
                
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.CascoEqpSlot = Slot

                UserList(UserIndex).Char.CascoAnim = obj.CascoAnim
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

            Else
                Call WriteConsoleMsg(1, UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
            End If
        ElseIf ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).SubTipo = 2 Then
            If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
                If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                
                If UserList(UserIndex).Invent.WeaponEqpObjIndex <> 0 Then
                    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).DosManos = 1 Then
                        If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).DosManos = 0 Then
                            WriteMsg UserIndex, 22
                            Exit Sub
                        End If
                    End If
                End If
                
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot)
                    
                    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    
                    Exit Sub
                End If
                
                'Quita el anterior
                If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
                End If
                
                'Lo equipa
                
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.EscudoEqpSlot = Slot
                
                'Sonido
                If ObjIndex = 1180 Then 'Escudo de Reflexión (RM +30)
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(28, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                End If
            
                UserList(UserIndex).Char.ShieldAnim = obj.ShieldAnim
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
             Else
                 Call WriteConsoleMsg(1, UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
             End If
        End If
    Case eOBJType.otWeapon, eOBJType.otHerramientas
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        
       If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
            If UserList(UserIndex).Invent.EscudoEqpObjIndex <> 0 Then
                If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).DosManos = 1 Then
                    If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).DosManos = 0 Then
                        WriteMsg UserIndex, 23
                        Exit Sub
                    End If
                End If
            End If
            
            'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                'Quitamos del inv el item
                Call Desequipar(UserIndex, Slot)
                Exit Sub
            End If
            
            'Quitamos el elemento anterior
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
            End If
            
            If UserList(UserIndex).Invent.NudiEqpIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.NudiEqpSlot)
            End If
            
            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            UserList(UserIndex).Invent.WeaponEqpSlot = Slot
            
            'Sonido
            If ObjIndex = 1181 Then 'Báculo de Hechicero (DM +10)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(234, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
            ElseIf ObjIndex = 668 Then 'Harbinger Kin
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(241, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
            ElseIf ObjIndex = 1095 Then 'Espada Argentum
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(245, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
            ElseIf ObjIndex = 1252 Then 'Báculo de Larzüll
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(242, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
            ElseIf ObjIndex = 402 Then 'Espada Mata Dragones
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(254, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
            ElseIf ObjIndex = 1147 Then 'Báculo de Hechicero (DM +20)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(243, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
            End If
            
            
            UserList(UserIndex).Char.WeaponAnim = obj.WeaponAnim
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
       Else
            Call WriteConsoleMsg(1, UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
       
    Case eOBJType.otNudillos
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        
        If ClasePuedeUsarItem(UserIndex, ObjIndex) Then
            'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                'Quitamos del inv el item
                Call Desequipar(UserIndex, Slot)
                'Animacion por defecto
                UserList(UserIndex).Char.WeaponAnim = NingunArma
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                Exit Sub
            End If
            
            'Quitamos el arma si tiene alguna equipada
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
            End If
            
            If UserList(UserIndex).Invent.NudiEqpIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.NudiEqpSlot)
            End If
            
            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.NudiEqpIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            UserList(UserIndex).Invent.NudiEqpSlot = Slot
            
            'Sonido
            If ObjIndex = 1333 Then 'Nudillos de Oro
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(255, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
            End If

            UserList(UserIndex).Char.WeaponAnim = obj.WeaponAnim
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
       Else
            Call WriteConsoleMsg(1, UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
        
End Select

'Actualiza
Call UpdateUserInv(False, UserIndex, Slot)

Exit Sub
Errhandler:
Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & err.Number & " - Error Description : " & err.description & " Name: " & UserList(UserIndex).Name)
End Sub

 Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo Errhandler

If ObjData(ItemIndex).RazaTipo > 0 Then
    'Verifica si la raza puede usar la ropa
    If UserList(UserIndex).Raza = eRaza.Humano Or _
       UserList(UserIndex).Raza = eRaza.Elfo Or _
       UserList(UserIndex).Raza = eRaza.Drow Then
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaTipo = 1)
    ElseIf UserList(UserIndex).Raza = eRaza.Orco Then
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaTipo = 3)
    Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaTipo = 2)
    End If
Else
    'Verifica si la raza puede usar la ropa
    If UserList(UserIndex).Raza = eRaza.Humano Or _
       UserList(UserIndex).Raza = eRaza.Elfo Or _
       UserList(UserIndex).Raza = eRaza.Drow Or _
       UserList(UserIndex).Raza = eRaza.Orco Then
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
    Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
    End If
End If

Exit Function
Errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'*************************************************
'Author: Unknown
'Last modified: 24/01/2007
'Handels the usage of items from inventory box.
'24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legión.
'24/01/2007 Pablo (ToxicWaste) - Utilización nueva de Barco en lvl 20 por clase Pirata y Pescador.
'*************************************************
On Error GoTo hayerror


Dim obj As ObjData
Dim ObjIndex As Integer
Dim TargObj As ObjData
Dim MiObj As obj

With UserList(UserIndex)

If .Invent.Object(Slot).Amount = 0 Then Exit Sub

obj = ObjData(.Invent.Object(Slot).ObjIndex)

If Not obj.MinELV = 0 Then
    If obj.MinELV > .Stats.ELV Then
        Call WriteConsoleMsg(1, UserIndex, "Debes ser nivel " & obj.MinELV & " para poder utilizar este objeto.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
End If


If obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
    Call WriteConsoleMsg(1, UserIndex, "Solo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If


If obj.OBJType = eOBJType.otWeapon Then
    
    If obj.proyectil = 1 Then
        If Not .flags.ModoCombate Then
            Call WriteConsoleMsg(1, UserIndex, "No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
        If Not IntervaloPermiteUsar(UserIndex, False) Then Exit Sub
    Else
        'dagas
        If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
    End If
Else
    If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
End If

ObjIndex = .Invent.Object(Slot).ObjIndex
.flags.TargetObjInvIndex = ObjIndex
.flags.TargetObjInvSlot = Slot

Select Case obj.OBJType
    Case eOBJType.otUseOnce
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If

        'Usa el item
        .Stats.MinHAM = .Stats.MinHAM + obj.MinHAM
        If .Stats.MinHAM > .Stats.MaxHAM Then _
            .Stats.MinHAM = .Stats.MaxHAM
        .flags.Hambre = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        'Sonido
        
        Call ReproducirSonido(SendTarget.ToPCArea, UserIndex, 7)

        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        Call UpdateUserInv(False, UserIndex, Slot)

    Case eOBJType.otGuita
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        
        .Stats.GLD = .Stats.GLD + .Invent.Object(Slot).Amount
        .Invent.Object(Slot).Amount = 0
        .Invent.Object(Slot).ObjIndex = 0
        .Invent.NroItems = .Invent.NroItems - 1
        
        Call UpdateUserInv(False, UserIndex, Slot)
        Call WriteUpdateGold(UserIndex)
        
    Case eOBJType.otWeapon
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        
        If Not .Stats.MinSTA > 0 Then
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(1, UserIndex, "Estas muy cansado", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, UserIndex, "Estas muy cansada", FontTypeNames.FONTTYPE_INFO)
            End If
            Exit Sub
        End If
        
        
        If ObjData(ObjIndex).proyectil = 1 Then
            If .Invent.Object(Slot).Equipped = 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            'liquid: muevo esto aca adentro, para que solo pida modo combate si estamos por usar el arco
            If Not .flags.ModoCombate Then
                Call WriteConsoleMsg(1, UserIndex, "No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            Call WriteWorkRequestTarget(UserIndex, Proyectiles)
        ElseIf ObjData(ObjIndex).SubTipo = 5 Then
            If .Invent.Object(Slot).Equipped = 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'liquid: muevo esto aca adentro, para que solo pida modo combate si estamos por usar el arco
            If Not .flags.ModoCombate Then
                Call WriteConsoleMsg(1, UserIndex, "No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call WriteWorkRequestTarget(UserIndex, arrojadizas)
        Else
            If .flags.TargetObj = Leña Then
                If .Invent.Object(Slot).ObjIndex = DAGA Then
                    If .Invent.Object(Slot).Equipped = 0 Then
                        Call WriteConsoleMsg(1, UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    Call TratarDeHacerFogata(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY, UserIndex)
                End If
            End If
        End If

    
    Case eOBJType.otPociones
        
        If .flags.Muerto = 1 And obj.TipoPocion <> 11 Then 'Pocion de resurreccion
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        
        
        If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then
            Call WriteConsoleMsg(1, UserIndex, "¡¡Debes esperar unos momentos para tomar otra poción!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .flags.TomoPocion = True
        .flags.TipoPocion = obj.TipoPocion
                
        Select Case .flags.TipoPocion
        
            Case 1 'Modif la agilidad
                .flags.DuracionEfecto = obj.DuracionEfecto
        
                'Usa el item
                .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(obj.MinModificador, obj.MaxModificador)
                If .Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then _
                    .Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
  
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.x, .Pos.Y))
                
                Call WriteAgilidad(UserIndex)
                
            Case 2 'Modif la fuerza
                .flags.DuracionEfecto = obj.DuracionEfecto
        
                'Usa el item
                .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(obj.MinModificador, obj.MaxModificador)
                If .Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then _
                    .Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
   
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.x, .Pos.Y))
                
                Call WriteFuerza(UserIndex)
                
            Case 3 'Pocion roja, restaura HP
                'Usa el item
                .Stats.MinHP = .Stats.MinHP + RandomNumber(obj.MinModificador, obj.MaxModificador)
                If .Stats.MinHP > .Stats.MaxHP Then _
                    .Stats.MinHP = .Stats.MaxHP
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.x, .Pos.Y))
            
            Case 4 'Pocion azul, restaura MANA
                'Usa el item
                'nuevo calculo para recargar mana
                .Stats.MinMAN = .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, 4) + .Stats.ELV \ 2 + 40 / .Stats.ELV
                If .Stats.MinMAN > .Stats.MaxMAN Then _
                    .Stats.MinMAN = .Stats.MaxMAN
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.x, .Pos.Y))
                
            Case 5 ' Pocion violeta
                If .flags.Envenenado <> 0 Then
                    .flags.Envenenado = 0
                    Call WriteConsoleMsg(1, UserIndex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
                End If
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.x, .Pos.Y))
            
            Case 6 'Remueve Paralisis
                
                If UserList(UserIndex).Pos.map = Bandera_mapa Then
                    Call WriteConsoleMsg(1, UserIndex, "No puedes usar aquí este objeto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If (UserList(UserIndex).flags.Paralizado = 1 Or UserList(UserIndex).flags.Inmovilizado = 1) Then
                
                    UserList(UserIndex).flags.Inmovilizado = 0
                    UserList(UserIndex).flags.Paralizado = 0
                    
                    'no need to crypt this
                    Call WriteParalizeOK(UserIndex)
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_REMO, .Pos.x, .Pos.Y))
                End If
                
            Case 7 'Metamorfosis Deshablitado
            
            Case 8 'Remueve Ceguera
                If UserList(UserIndex).flags.Ceguera = 1 Then
                    UserList(UserIndex).flags.Ceguera = 0
                    UserList(UserIndex).Counters.Ceguera = 0

                    Call WriteBlindNoMore(UserIndex)
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call FlushBuffer(UserIndex)
                End If
            Case 9 'Remueve Estupidez
                If UserList(UserIndex).flags.Estupidez = 1 Then
                    UserList(UserIndex).flags.Estupidez = 0
                    
                    'no need to crypt this
                    Call WriteDumbNoMore(UserIndex)
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call FlushBuffer(UserIndex)
                End If
                
            Case 10 'Invisibilidad
                 
                If UserList(UserIndex).Counters.Saliendo Then
                    Call WriteConsoleMsg(1, UserIndex, "¡No puedes ponerte invisible mientras te encuentres saliendo!", FontTypeNames.FONTTYPE_WARNING)
                    Exit Sub
                End If
                 
                 'No usar invi mapas InviSinEfecto
                If MapInfo(UserList(UserIndex).Pos.map).InviSinEfecto > 0 Then
                    Call WriteConsoleMsg(1, UserIndex, "¡La invisibilidad no funciona aquí!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(UserIndex).flags.Navegando = 1 Then
                    Call WriteConsoleMsg(1, UserIndex, "No puedes estar invisible navegando!!", FontTypeNames.FONTTYPE_BROWNI)
                    Exit Sub
                End If
    
                UserList(UserIndex).flags.Invisible = 1
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, True))
                Call QuitarUserInvItem(UserIndex, Slot, 1)
            
            'Add Marius
            Case 11 'Resurreccion
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call RevivirUsuario(UserIndex)
                    
                    'no need to crypt this
                    Call WriteDumbNoMore(UserIndex)
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call FlushBuffer(UserIndex)
                End If
            
            Case 12 'Cambio de sexo (falta terminar, cambia body pero no head y divorciarte si estas casado)
                ' DESEQUIPA TODOS LOS OBJETOS
                If .Invent.ArmourEqpObjIndex > 0 Then 'desequipar armadura
                    Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
                End If
                
                If .Invent.WeaponEqpObjIndex > 0 Then 'desequipar arma
                    Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                End If
                
                If .Invent.CascoEqpObjIndex > 0 Then 'desequipar casco
                    Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
                End If
                
                If .Invent.AnilloEqpSlot > 0 Then 'desequipar herramienta
                    Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)
                End If
                
                If .Invent.MunicionEqpObjIndex > 0 Then 'desequipar municiones
                    Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
                End If
                
                If .Invent.MagicIndex > 0 Then 'desequipamos items macigos
                    Call Desequipar(UserIndex, .Invent.MagicSlot)
                End If
                
                If .Invent.EscudoEqpObjIndex > 0 Then 'desequipar escudo
                    Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
                End If
                
                If .Genero = eGenero.Hombre Then
                    .Genero = eGenero.Mujer
                    Call WriteConsoleMsg(1, UserIndex, "Ahora eres mujer!", FontTypeNames.FONTTYPE_BROWNI)
                Else
                    .Genero = eGenero.Hombre
                    Call WriteConsoleMsg(1, UserIndex, "Ahora eres hombre!", FontTypeNames.FONTTYPE_BROWNI)
                End If
                
                Call DarCuerpoDesnudo(UserIndex)
                
                'Refrescamos el user a lo cabeza, la unica forma que conozco
                Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, False)
                
                'no need to crypt this
                Call WriteDumbNoMore(UserIndex)
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call FlushBuffer(UserIndex)
            
            Case 13 'Cambio de cabeza
            
            Case 14 '14 Nharet
                Debug.Print "Pocion nharet"

            Case 15 '15 Manual de Liderazgo
                Debug.Print "Pergamino liderazgo"
                If .Stats.UserSkills(eSkill.Liderazgo) <> 100 Then
                    .Stats.UserSkills(eSkill.Liderazgo) = 100
                    Call WriteMsg(UserIndex, 40, CStr(eSkill.Liderazgo), CStr(.Stats.UserSkills(eSkill.Liderazgo)))
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call WriteConsoleMsg(1, UserIndex, "Ahora tiene 100 puntos en Liderazgo!", FontTypeNames.FONTTYPE_BROWNI)
                    Call FlushBuffer(UserIndex)
                End If

            Case 16 '16 Manual de navegacíon
                Debug.Print "Pergamino navegacion"
                If .Stats.UserSkills(eSkill.Navegacion) <> 100 Then
                    .Stats.UserSkills(eSkill.Navegacion) = 100
                    Call WriteMsg(UserIndex, 40, CStr(eSkill.Navegacion), CStr(.Stats.UserSkills(eSkill.Navegacion)))
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call WriteConsoleMsg(1, UserIndex, "Ahora tiene 100 puntos en Navegacion!", FontTypeNames.FONTTYPE_BROWNI)
                    Call FlushBuffer(UserIndex)
                End If
            Case 17 '17 Manual de equitacíon
                Debug.Print "Pergamino equitacion"
                If .Stats.UserSkills(eSkill.Equitacion) <> 100 Then
                    .Stats.UserSkills(eSkill.Equitacion) = 100
                    Call WriteMsg(UserIndex, 40, CStr(eSkill.Equitacion), CStr(.Stats.UserSkills(eSkill.Equitacion)))
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call WriteConsoleMsg(1, UserIndex, "Ahora tiene 100 puntos en Equitacion!", FontTypeNames.FONTTYPE_BROWNI)
                    Call FlushBuffer(UserIndex)
                End If
            '\Add
       End Select
       
       Call WriteUpdateUserStats(UserIndex)
       Call UpdateUserInv(False, UserIndex, Slot)

     Case eOBJType.otBebidas
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        .Stats.MinAGU = .Stats.MinAGU + obj.MinSed
        If .Stats.MinAGU > .Stats.MaxAGU Then _
            .Stats.MinAGU = .Stats.MaxAGU
        .flags.Sed = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.x, .Pos.Y))
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otLlaves
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        
        If .flags.TargetObj = 0 Then Exit Sub
        TargObj = ObjData(.flags.TargetObj)
        '¿El objeto clickeado es una puerta?
        If TargObj.OBJType = eOBJType.otPuertas Then
            '¿Esta cerrada?
            If TargObj.Cerrada = 1 Then
                  '¿Cerrada con llave?
                  If TargObj.Llave > 0 Then
                     If TargObj.clave = obj.clave Then
         
                        MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex _
                        = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                        .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                        Call WriteConsoleMsg(1, UserIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     Else
                        Call WriteConsoleMsg(1, UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     End If
                  Else
                     If TargObj.clave = obj.clave Then
                        MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex _
                        = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                        Call WriteConsoleMsg(1, UserIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                        .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                        Exit Sub
                     Else
                        Call WriteConsoleMsg(1, UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     End If
                  End If
            Else
                  Call WriteConsoleMsg(1, UserIndex, "No esta cerrada.", FontTypeNames.FONTTYPE_INFO)
                  Exit Sub
            End If
        End If
    
    Case eOBJType.otBotellaVacia
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        If Not HayAgua(.Pos.map, .flags.TargetX, .flags.TargetY) Then
            Call WriteConsoleMsg(1, UserIndex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        MiObj.Amount = 1
        MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexAbierta
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otBotellaLlena
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        .Stats.MinAGU = .Stats.MinAGU + obj.MinSed
        If .Stats.MinAGU > .Stats.MaxAGU Then _
            .Stats.MinAGU = .Stats.MaxAGU
        .flags.Sed = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        MiObj.Amount = 1
        MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexCerrada
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otPergaminos
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        
        If .Stats.MaxMAN > 0 Then
            If .flags.Hambre = 0 And .flags.Sed = 0 Then
                
                'Add Marius pusimos el if para verificar
                If InStr(Hechizos(ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).HechizoIndex).ExclusivoClase, UCase$(ListaClases(UserList(UserIndex).Clase))) <> 0 Or _
                Len(Hechizos(ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).HechizoIndex).ExclusivoClase) = 0 Then
                    Call AgregarHechizo(UserIndex, Slot)
                    Call UpdateUserInv(False, UserIndex, Slot)
                Else
                    Call WriteConsoleMsg(1, UserIndex, "Tú clase no puede aprender este hechizo.", FontTypeNames.FONTTYPE_INFO)
                End If
                
            Else
                Call WriteConsoleMsg(1, UserIndex, "Estas demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(1, UserIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_INFO)
        End If

    Case eOBJType.otMinerales
        If .flags.Muerto = 1 Then
             Call WriteMsg(UserIndex, 1)
             Exit Sub
        End If
        Call WriteWorkRequestTarget(UserIndex, FundirMetal)
        .flags.Lingoteando = Slot
        
    Case eOBJType.otInstrumentos
        If .flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        
        If obj.Real Then '¿Es el Cuerno Real?
            If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                If MapInfo(.Pos.map).Pk = False Then
                    Call WriteConsoleMsg(1, UserIndex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call SendData(SendTarget.ToMap, .Pos.map, PrepareMessagePlayWave(obj.Snd1, .Pos.x, .Pos.Y))
                Exit Sub
            Else
                Call WriteConsoleMsg(1, UserIndex, "Solo Miembros de la Armada Real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        ElseIf obj.Caos Then '¿Es el Cuerno Legión?
            If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                If MapInfo(.Pos.map).Pk = False Then
                    Call WriteConsoleMsg(1, UserIndex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call SendData(SendTarget.ToMap, .Pos.map, PrepareMessagePlayWave(obj.Snd1, .Pos.x, .Pos.Y))
                Exit Sub
            Else
                Call WriteConsoleMsg(1, UserIndex, "Solo Miembros de la Legión Oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        'Si llega aca es porque es o Laud o Tambor o Flauta
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.x, .Pos.Y))
       
    Case eOBJType.otMonturas
        If UserList(UserIndex).flags.Metamorfosis = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "Transformado no puedes utilizar este objeto. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
        
        If LegalPos(.Pos.map, .Pos.x, .Pos.Y) Then
            Call DoEquita(UserIndex, ObjIndex, Slot)
        End If
        
    Case eOBJType.otBarcos
        If UserList(UserIndex).flags.Metamorfosis = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "Transformado no puedes utilizar este objeto. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        If ((LegalPos(.Pos.map, .Pos.x - 1, .Pos.Y, True, False) _
                Or LegalPos(.Pos.map, .Pos.x, .Pos.Y - 1, True, False) _
                Or LegalPos(.Pos.map, .Pos.x + 1, .Pos.Y, True, False) _
                Or LegalPos(.Pos.map, .Pos.x, .Pos.Y + 1, True, False)) _
                And .flags.Navegando = 0) _
                Or .flags.Navegando = 1 Then
            Call DoNavega(UserIndex, obj, Slot)
        Else
            Call WriteConsoleMsg(1, UserIndex, "¡Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)
        End If
        
    Case eOBJType.otPasajes
        If UserList(UserIndex).flags.Metamorfosis = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "Transformado no puedes utilizar este objeto. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 1)
            Exit Sub
        End If
       
        If UserList(UserIndex).flags.TargetNPC <> 0 Then
            If Left$(Npclist(UserList(UserIndex).flags.TargetNPC).Name, 6) <> "Pirata" Then
                Call WriteConsoleMsg(1, UserIndex, "Primero debes hacer click sobre el pirata.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            Call WriteConsoleMsg(1, UserIndex, "Primero debes hacer click sobre el pirata.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
       
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
            Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Not MapaValido(obj.HastaMap) Then
            Call WriteConsoleMsg(1, UserIndex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
       
        If UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) < obj.CantidadSkill Then
            Call WriteConsoleMsg(1, UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
       
        Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
        Call WriteConsoleMsg(1, UserIndex, "Has viajado por varios días, te sientes exhausto!", FontTypeNames.FONTTYPE_BROWNB)
        UserList(UserIndex).Stats.MinAGU = 0
        UserList(UserIndex).Stats.MinHAM = 0
        UserList(UserIndex).flags.Sed = 1
        UserList(UserIndex).flags.Hambre = 1
        
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Agilidad)
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Fuerza)
        
        Call WriteAgilidad(UserIndex)
        Call WriteFuerza(UserIndex)
        
        Call WriteUpdateHungerAndThirst(UserIndex)
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        Call UpdateUserInv(False, UserIndex, Slot)
        
    Case eOBJType.otruna
        'COMIENZO RUNA////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
        Dim FuturePos As WorldPos
        Dim NuevaPos As WorldPos
            
            'Add Nod Kopfnickend Los fantasmas se quedan en prision.
            If .Pos.map = Prision.map Then
                Call WriteConsoleMsg(1, UserIndex, "No puedes usar la runa en este mapa!", FontTypeNames.FONTTYPE_BROWNB)
                Exit Sub
            End If
            '\Add
            
            
            'Add Marius Solucionamos un error con las arenas xD
            If UserList(UserIndex).flags.inDuelo = 1 Then
                Call PerderDuelo(UserIndex)
            End If
            
            If UserList(UserIndex).evento <> 0 Then
                Call salir_arena(UserIndex)
            End If
            
             'Add Marius Captura la Bandera
            If UserList(UserIndex).Pos.map = Bandera_mapa Then
                Call Bandera_muere(UserIndex)
            End If
            '\Add
            '\Add
            
            
            
            Select Case UserList(UserIndex).Hogar
                Case 0
                    FuturePos.map = 34
                    FuturePos.x = 22
                    FuturePos.Y = 60
                Case 1
                    FuturePos.map = 194
                    FuturePos.x = 61
                    FuturePos.Y = 71
                Case 2
                    FuturePos.map = 112
                    FuturePos.x = 25
                    FuturePos.Y = 59
                Case 3
                    FuturePos.map = 20
                    FuturePos.x = 77
                    FuturePos.Y = 26
            End Select

            'Add Marius Carrera
            If mapasEspeciales(UserIndex) Then
                FuturePos.map = 49
                FuturePos.x = 50
                FuturePos.Y = 50
            ElseIf MapInfo(.Pos.map).Pk And UserList(UserIndex).flags.Muerto = 0 Then
                Exit Sub
            End If
            '\Add
            
            'Add Marius Captura la Bandera
            If UserList(UserIndex).Pos.map = Bandera_mapa Then
                If Bandera_estado Then
                    Call Bandera_Sale(UserIndex)
                    Exit Sub
                End If
                FuturePos.map = 49
                FuturePos.x = 50
                FuturePos.Y = 50
            End If
            '\Add
            
            'Add Nod kopfnickend
            'Comprabamos si lo vamos a transportar al mismo lugar donde esta.
            'Para que no runeen al pedo si estan a donde lo vamos a llevar.
            If UserList(UserIndex).Pos.map = FuturePos.map And UserList(UserIndex).Pos.x = FuturePos.x And UserList(UserIndex).Pos.Y = FuturePos.Y Then Exit Sub
            '\Add

            'Mod Marius Lo hacemos mas lindo, que busque la pos mas cercana al punto.
            Call ClosestLegalPos(FuturePos, NuevaPos)
        
            If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then
                Call WarpUserChar(UserIndex, NuevaPos.map, NuevaPos.x, NuevaPos.Y, True)
            Else
                Call WarpUserChar(UserIndex, FuturePos.map, FuturePos.x, FuturePos.Y, True)
            End If
            
            If UserList(UserIndex).flags.Muerto = 1 Then _
                Call RevivirUsuario(UserIndex)
            '\Mod
        'FIN RUNA/////////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
End Select

If UserList(UserIndex).Invent.Object(Slot).Equipped <> 0 Then
    If ObjIndex = SERRUCHO_CARPINTERO Then
        Call EnivarObjConstruibles(UserIndex)
        Call WriteShowCarpenterForm(UserIndex)
    ElseIf ObjIndex = COSTURERO Then
        Call EnivarObjTejibles(UserIndex)
        Call WriteShowSastreForm(UserIndex)
    ElseIf ObjIndex = OLLA Then
        Call EnivarObjalquimia(UserIndex)
        Call WriteShowalquimiaForm(UserIndex)
    End If
End If
        
End With

Exit Sub
hayerror:
    LogError ("Error en Userinvitem: " & err.Number & " Desc: " & err.description & " UI: " & UserIndex & " Name: " & UserList(UserIndex).Name & " Obj:" & ObjIndex)
End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)

Call WriteBlacksmithWeapons(UserIndex)

End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)

Call WriteCarpenterObjects(UserIndex)

End Sub

Sub EnivarObjalquimia(ByVal UserIndex As Integer)

Call WriteAlquimiaObjects(UserIndex)

End Sub

Sub EnivarObjTejibles(ByVal UserIndex As Integer)

Call WriteTejiblesObjects(UserIndex)

End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)

    Call WriteBlacksmithArmors(UserIndex)

End Sub

Sub TirarTodo(ByVal UserIndex As Integer)
On Error GoTo hayerror

    If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).Trigger = 6 Then Exit Sub
    If MapInfo(UserList(UserIndex).Pos.map).Seguro = 1 Then Exit Sub
            
    Call TirarTodosLosItems(UserIndex)
    
    'Dim Cantidad As Long
    'Cantidad = UserList(Userindex).Stats.GLD
    'If UserList(userindex).Stats.GLD < 100001 Then _
        Call TirarOro(UserList(userindex).Stats.GLD, userindex)
    
Exit Sub

hayerror:
    LogError ("Error en TirarTodo: " & err.description)
End Sub

Public Function ItemSeCae(ByVal index As Integer) As Boolean
    
    ItemSeCae = ObjData(index).OBJType <> eOBJType.otLlaves And _
                ObjData(index).OBJType <> eOBJType.otMonturas And _
                ObjData(index).OBJType <> eOBJType.otBarcos And _
                ObjData(index).OBJType <> eOBJType.otMinerales And _
                ObjData(index).OBJType <> eOBJType.otBebidas And _
                ObjData(index).OBJType <> eOBJType.otMapas And _
                ObjData(index).OBJType <> eOBJType.otruna And _
                ObjData(index).OBJType <> 1 And _
                ObjData(index).Caos < 1 And _
                ObjData(index).Real < 1 And _
                ObjData(index).Shop < 49 And _
                ObjData(index).EfectoMagico <> eMagicType.Sacrificio And _
                ObjData(index).Milicia <> 1 And _
                index <> 1275 And _
                index <> 592
End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
On Error GoTo hayerror

    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As obj
    Dim ItemIndex As Integer
    
    'Ponemos aca el carro de mineria
    Dim Carro As Byte
    Dim Minerales As Integer
    Dim Porc As Byte
    Dim Hierro As Integer, Plata As Integer, oro As Integer

    Carro = Have_Obj_Slot(CARROMINERO, UserIndex)
    If Carro > 0 Then
        Hierro = Have_Obj_To_Slot(iMinerales.HierroCrudo, Carro, UserIndex)
        Plata = Have_Obj_To_Slot(iMinerales.PlataCruda, Carro, UserIndex)
        oro = Have_Obj_To_Slot(iMinerales.OroCrudo, Carro, UserIndex)
        
        If Hierro > 0 Then
            Porc = Porc + 1
        End If
        
        If Plata > 0 Then
            Porc = Porc + 1
        End If
        
        If oro > 0 Then
            Porc = Porc + 1
        End If
        
        If Hierro > 0 Then Hierro = Porcentaje(Hierro, (100 - ObjData(UserList(UserIndex).Invent.Object(Carro).ObjIndex).CuantoAumento) / Porc)
        If Plata > 0 Then Plata = Porcentaje(Plata, (100 - ObjData(UserList(UserIndex).Invent.Object(Carro).ObjIndex).CuantoAumento) / Porc)
        If oro > 0 Then oro = Porcentaje(oro, (100 - ObjData(UserList(UserIndex).Invent.Object(Carro).ObjIndex).CuantoAumento) / Porc)
        
        If Porc > 0 Then
            For i = 1 To Carro
                If UserList(UserIndex).Invent.Object(i).ObjIndex = iMinerales.HierroCrudo Then
                    If Hierro > 0 Then
                        TirarObjeto UserIndex, i, Hierro
                        Hierro = Hierro - IIf(UserList(UserIndex).Invent.Object(i).Amount > Hierro, Hierro, UserList(UserIndex).Invent.Object(i).Amount)
                    End If
                ElseIf UserList(UserIndex).Invent.Object(i).ObjIndex = iMinerales.PlataCruda Then
                    If Plata > 0 Then
                        TirarObjeto UserIndex, i, Plata
                        Plata = Plata - IIf(UserList(UserIndex).Invent.Object(i).Amount > Plata, Plata, UserList(UserIndex).Invent.Object(i).Amount)
                    End If
                ElseIf UserList(UserIndex).Invent.Object(i).ObjIndex = iMinerales.OroCrudo Then
                    If oro > 0 Then
                        TirarObjeto UserIndex, i, oro
                        oro = oro - IIf(UserList(UserIndex).Invent.Object(i).Amount > oro, oro, UserList(UserIndex).Invent.Object(i).Amount)
                    End If
                End If
            Next i
        End If
    End If
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
        If ItemIndex > 0 Then
             If ItemSeCae(ItemIndex) Then
                NuevaPos.x = 0
                NuevaPos.Y = 0
                
                'Creo el Obj
                MiObj.Amount = UserList(UserIndex).Invent.Object(i).Amount
                MiObj.ObjIndex = ItemIndex
                'Pablo (ToxicWaste) 24/01/2007
                'Si es pirata y usa un Galeón entonces no explota los items. (en el agua)
                'If UserList(UserIndex).Clase = eClass.Mercenario And UserList(UserIndex).Invent.BarcoObjIndex = 476 Then
                If HayAgua(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y) Then
                    TileLibre UserList(UserIndex).Pos, NuevaPos, MiObj, False, True
                Else
                    TileLibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True
                End If
                
                If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.x, NuevaPos.Y)
                End If
             End If
        End If
    Next i


    Exit Sub
hayerror:
    LogError ("Error en TirarTodosLosItems: " & err.description & " Name: " & UserList(UserIndex).Name)
End Sub
Sub TirarObjeto(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cant As Integer)
    Dim MiObj As obj
    Dim NuevaPos As WorldPos
    
  
    
    If Cant > UserList(UserIndex).Invent.Object(Slot).Amount Then _
        Cant = UserList(UserIndex).Invent.Object(Slot).Amount
    'Creo el Obj
    MiObj.Amount = Cant
    MiObj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    
    If UserList(UserIndex).Clase = eClass.Mercenario And UserList(UserIndex).Invent.BarcoObjIndex = 476 Then
        TileLibre UserList(UserIndex).Pos, NuevaPos, MiObj, False, True
    Else
        TileLibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True
    End If
                
    If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then
        Call DropObj(UserIndex, Slot, Cant, NuevaPos.map, NuevaPos.x, NuevaPos.Y)
    End If
End Sub
Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean

    ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)
Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As obj
Dim ItemIndex As Integer

If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).Trigger = 6 Then Exit Sub

For i = 1 To MAX_INVENTORY_SLOTS
    ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    If ItemIndex > 0 Then
        If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
            NuevaPos.x = 0
            NuevaPos.Y = 0
            
            'Creo MiObj
            MiObj.Amount = UserList(UserIndex).Invent.Object(i).ObjIndex
            MiObj.ObjIndex = ItemIndex
            'Pablo (ToxicWaste) 24/01/2007
            'Tira los Items no newbies en todos lados.
            TileLibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True
            If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then
                If MapData(NuevaPos.map, NuevaPos.x, NuevaPos.Y).ObjInfo.ObjIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.x, NuevaPos.Y)
            End If
        End If
    End If
Next i

End Sub
