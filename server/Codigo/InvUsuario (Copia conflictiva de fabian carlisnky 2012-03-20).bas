Attribute VB_Name = "InvUsuario"

Option Explicit

Public Function TieneObjetosRobables(ByVal userindex As Integer) As Boolean

'17/09/02
'Agregue que la función se asegure que el objeto no es un barco

On Error GoTo hayerror

Dim i As Integer
Dim ObjIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
    ObjIndex = UserList(userindex).Invent.Object(i).ObjIndex
    If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And _
                ObjData(ObjIndex).OBJType <> eOBJType.otMonturas And _
                ObjData(ObjIndex).OBJType <> eOBJType.otBarcos) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    
    End If
Next i

Exit Function

hayerror:
    LogError ("Error en TieneObjetosRobables: " & err.description)



End Function

Function ClasePuedeUsarItem(ByVal userindex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo manejador

'Call LogTarea("ClasePuedeUsarItem")



Dim flag As Boolean
If ObjIndex = 0 Then Exit Function




'Admins can use ANYTHING!
If UserList(userindex).flags.Privilegios And (PlayerType.User Or PlayerType.VIP) Then
    If ObjData(ObjIndex).ClaseTipo = 0 Then
        If ObjData(ObjIndex).ClaseProhibida(1) > 0 Then
            Dim i As Integer
            For i = 1 To NUMCLASES
                If ObjData(ObjIndex).ClaseProhibida(i) = UserList(userindex).Clase Then
                    
                    ClasePuedeUsarItem = False
                    Exit Function
                End If
            Next i
        End If
    Else
        If (UserList(userindex).Clase = Gladiador Or _
        UserList(userindex).Clase = Guerrero Or _
        UserList(userindex).Clase = Paladin Or _
        UserList(userindex).Clase = Cazador Or _
        UserList(userindex).Clase = Mercenario) Then
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

Sub QuitarNewbieObj(ByVal userindex As Integer)
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(userindex).Invent.Object(j).ObjIndex > 0 Then
             
             If ObjData(UserList(userindex).Invent.Object(j).ObjIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(userindex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, userindex, j)
        
        End If
Next j

End Sub

Sub LimpiarInventario(ByVal userindex As Integer)


Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
    UserList(userindex).Invent.Object(j).ObjIndex = 0
    UserList(userindex).Invent.Object(j).Amount = 0
    UserList(userindex).Invent.Object(j).Equipped = 0
Next j

UserList(userindex).Invent.NroItems = 0

UserList(userindex).Invent.NudiEqpSlot = 0
UserList(userindex).Invent.NudiEqpIndex = 0

UserList(userindex).Invent.ArmourEqpObjIndex = 0
UserList(userindex).Invent.ArmourEqpSlot = 0

UserList(userindex).Invent.WeaponEqpObjIndex = 0
UserList(userindex).Invent.WeaponEqpSlot = 0

UserList(userindex).Invent.CascoEqpObjIndex = 0
UserList(userindex).Invent.CascoEqpSlot = 0

UserList(userindex).Invent.EscudoEqpObjIndex = 0
UserList(userindex).Invent.EscudoEqpSlot = 0

UserList(userindex).Invent.AnilloEqpObjIndex = 0
UserList(userindex).Invent.AnilloEqpSlot = 0

UserList(userindex).Invent.MunicionEqpObjIndex = 0
UserList(userindex).Invent.MunicionEqpSlot = 0

UserList(userindex).Invent.BarcoObjIndex = 0
UserList(userindex).Invent.BarcoSlot = 0

UserList(userindex).Invent.MonturaObjIndex = 0
UserList(userindex).Invent.MonturaSlot = 0

UserList(userindex).Invent.MagicIndex = 0
UserList(userindex).Invent.MagicSlot = 0
End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal userindex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
'***************************************************
On Error GoTo Errhandler

'If Cantidad > 100000 Then Exit Sub

'SI EL Pjta TIENE ORO LO TIRAMOS
If (Cantidad > 0) And (Cantidad <= UserList(userindex).Stats.GLD) Then
        Dim i As Byte
        Dim MiObj As Obj
        'info debug
        Dim loops As Integer
        
        'Seguridad Alkon (guardo el oro tirado si supera los 50k)
        If Cantidad > 50000 Then
            Dim j As Integer
            Dim k As Integer
            Dim M As Integer
            Dim Cercanos As String
            M = UserList(userindex).pos.map
            For j = UserList(userindex).pos.x - 10 To UserList(userindex).pos.x + 10
                For k = UserList(userindex).pos.Y - 10 To UserList(userindex).pos.Y + 10
                    If InMapBounds(M, j, k) Then
                        If MapData(M, j, k).userindex > 0 Then
                            Cercanos = Cercanos & UserList(MapData(M, j, k).userindex).Name & ","
                        End If
                    End If
                Next k
            Next j
        End If
        
        '/Seguridad
        Dim Extra As Long
        Dim TeniaOro As Long
        TeniaOro = UserList(userindex).Stats.GLD
        If Cantidad > 500000 Then 'Para evitar explotar demasiado
            Extra = Cantidad - 500000
            Cantidad = 500000
        End If
        
        Do While (Cantidad > 0)
            
            If Cantidad > MAX_INVENTORY_OBJS And UserList(userindex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.Amount = MAX_INVENTORY_OBJS
                Cantidad = Cantidad - MiObj.Amount
            Else
                MiObj.Amount = Cantidad
                Cantidad = Cantidad - MiObj.Amount
            End If

            MiObj.ObjIndex = iORO
            
            Dim AuxPos As WorldPos
            
            If UserList(userindex).Clase = eClass.Mercenario And UserList(userindex).Invent.BarcoObjIndex = 476 Then
                AuxPos = TirarItemAlPiso(UserList(userindex).pos, MiObj, False)
                If AuxPos.x <> 0 And AuxPos.Y <> 0 Then
                    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - MiObj.Amount
                End If
            Else
                AuxPos = TirarItemAlPiso(UserList(userindex).pos, MiObj, True)
                If AuxPos.x <> 0 And AuxPos.Y <> 0 Then
                    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - MiObj.Amount
                End If
            End If
            
            'info debug
            loops = loops + 1
            If loops > 100 Then
                LogError ("Error en tiraroro")
                Exit Sub
            End If
            
        Loop
        If TeniaOro = UserList(userindex).Stats.GLD Then Extra = 0
        If Extra > 0 Then
            UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - Extra
        End If
    
End If

Exit Sub

Errhandler:

End Sub

Sub QuitarUserInvItem(ByVal userindex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
    If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    With UserList(userindex).Invent.Object(Slot)
        If .Amount <= Cantidad Then
            If .Equipped = 1 Then
                Call Desequipar(userindex, Slot)
            End If
        End If
        
        'Quita un objeto
        .Amount = .Amount - Cantidad
        '¿Quedan mas?
        If .Amount <= 0 Then
            UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems - 1
            .ObjIndex = 0
            .Amount = 0
        End If
    End With
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal userindex As Integer, ByVal Slot As Byte)

Dim NullObj As UserObj
Dim LoopC As Long

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(userindex).Invent.Object(Slot).ObjIndex > 0 Then
        Call ChangeUserInv(userindex, Slot, UserList(userindex).Invent.Object(Slot))
    Else
        Call ChangeUserInv(userindex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        'Actualiza el inventario
        If UserList(userindex).Invent.Object(LoopC).ObjIndex > 0 Then
            Call ChangeUserInv(userindex, LoopC, UserList(userindex).Invent.Object(LoopC))
        Else
            Call ChangeUserInv(userindex, LoopC, NullObj)
        End If
    Next LoopC
End If

End Sub

Sub DropObj(ByVal userindex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer)
    Dim Obj As Obj
    
    If num > 0 Then
        If num > UserList(userindex).Invent.Object(Slot).Amount Then num = UserList(userindex).Invent.Object(Slot).Amount
      
        'Check objeto en el suelo
        If MapData(UserList(userindex).pos.map, x, Y).ObjInfo.ObjIndex = 0 Or MapData(UserList(userindex).pos.map, x, Y).ObjInfo.ObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex Then
            Obj.ObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
            
            If num + MapData(UserList(userindex).pos.map, x, Y).ObjInfo.Amount > MAX_INVENTORY_OBJS Then
                num = MAX_INVENTORY_OBJS - MapData(UserList(userindex).pos.map, x, Y).ObjInfo.Amount
            End If
            
            Obj.Amount = num
            
            Call MakeObj(Obj, map, x, Y)
            Call QuitarUserInvItem(userindex, Slot, num)
            Call UpdateUserInv(False, userindex, Slot)
            
        Else
            Call WriteConsoleMsg(1, userindex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
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

Sub MakeObj(ByRef Obj As Obj, ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer)

If Obj.ObjIndex > 0 And Obj.ObjIndex <= UBound(ObjData) Then

    If MapData(map, x, Y).ObjInfo.ObjIndex = Obj.ObjIndex Then
        MapData(map, x, Y).ObjInfo.Amount = MapData(map, x, Y).ObjInfo.Amount + Obj.Amount
    Else
        MapData(map, x, Y).ObjInfo = Obj
        
        Call modSendData.SendToAreaByPos(map, x, Y, PrepareMessageObjectCreate(x, Y, Obj.ObjIndex, Obj.Amount))
    End If
End If

End Sub

Function MeterItemEnInventario(ByVal userindex As Integer, ByRef MiObj As Obj) As Boolean
On Error GoTo Errhandler

'Call LogTarea("MeterItemEnInventario")
 
Dim x As Integer
Dim Y As Integer
Dim Slot As Byte

'¿el user ya tiene un objeto del mismo tipo?
Slot = 1
Do Until UserList(userindex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And _
         UserList(userindex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
   Slot = Slot + 1
   If Slot > MAX_INVENTORY_SLOTS Then
         Exit Do
   End If
Loop
    
'Sino busca un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(userindex).Invent.Object(Slot).ObjIndex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
           Call WriteConsoleMsg(1, userindex, "No podes cargar mas objetos.", FontTypeNames.FONTTYPE_FIGHT)
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems + 1
End If
    
'Mete el objeto
If UserList(userindex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(userindex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
   UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount + MiObj.Amount
Else
   UserList(userindex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True
       
Call UpdateUserInv(False, userindex, Slot)


Exit Function
Errhandler:

End Function


Sub GetObj(ByVal userindex As Integer)

Dim Obj As ObjData
Dim MiObj As Obj
Dim ObjPos As String

'¿Hay algun obj?
If MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).ObjInfo.ObjIndex > 0 Then
    '¿Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then
        Dim x As Integer
        Dim Y As Integer
        Dim Slot As Byte
        
        x = UserList(userindex).pos.x
        Y = UserList(userindex).pos.Y
        Obj = ObjData(MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).ObjInfo.ObjIndex)
        MiObj.Amount = MapData(UserList(userindex).pos.map, x, Y).ObjInfo.Amount
        MiObj.ObjIndex = MapData(UserList(userindex).pos.map, x, Y).ObjInfo.ObjIndex
        If MiObj.ObjIndex = 12 Then
            UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + MiObj.Amount
            Call WriteUpdateGold(userindex)
            Call EraseObj(MapData(UserList(userindex).pos.map, x, Y).ObjInfo.Amount, UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y)
        Else
            If MeterItemEnInventario(userindex, MiObj) Then
                Call EraseObj(MapData(UserList(userindex).pos.map, x, Y).ObjInfo.Amount, UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y)
            End If
        End If
        
    End If
Else
    Call WriteConsoleMsg(1, userindex, "No hay nada aquí.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Sub Desequipar(ByVal userindex As Integer, ByVal Slot As Byte)
'Desequipa el item slot del inventario
Dim Obj As ObjData


If (Slot < LBound(UserList(userindex).Invent.Object)) Or (Slot > UBound(UserList(userindex).Invent.Object)) Then
    Exit Sub
ElseIf UserList(userindex).Invent.Object(Slot).ObjIndex = 0 Then
    Exit Sub
End If

Obj = ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex)

Select Case Obj.OBJType
    Case eOBJType.otMonturas
        Call DoEquita(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex, Slot)
        
    Case eOBJType.otHerramientas
        UserList(userindex).Invent.AnilloEqpObjIndex = 0
        UserList(userindex).Invent.AnilloEqpSlot = 0
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        
        If UserList(userindex).flags.Trabajando = True Then
            UserList(userindex).flags.Trabajando = False
            
            Call WriteConsoleMsg(1, userindex, "Dejas de trabajar.", FontTypeNames.FONTTYPE_BROWNI)
        End If
        
    Case eOBJType.otWeapon
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Invent.WeaponEqpObjIndex = 0
        UserList(userindex).Invent.WeaponEqpSlot = 0
        UserList(userindex).Char.WeaponAnim = NingunArma
        Call ChangeUserChar(userindex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
                
    Case eOBJType.otItemsMagicos
        UserList(userindex).Invent.MagicIndex = 0
        UserList(userindex).Invent.MagicSlot = 0
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        
        If Obj.EfectoMagico = eMagicType.ModificaAtributo Then
            If Obj.QueAtributo <> 0 Then
                UserList(userindex).Stats.UserAtributos(Obj.QueAtributo) = UserList(userindex).Stats.UserAtributos(Obj.QueAtributo) - Obj.CuantoAumento
            End If
        ElseIf Obj.EfectoMagico = eMagicType.ModificaSkill Then
            If Obj.QueSkill <> 0 Then
                UserList(userindex).Stats.UserSkills(Obj.QueSkill) = UserList(userindex).Stats.UserSkills(Obj.QueSkill) - Obj.CuantoAumento
            End If
        End If
        
    Case eOBJType.otNudillos
        UserList(userindex).Invent.NudiEqpIndex = 0
        UserList(userindex).Invent.NudiEqpSlot = 0
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Char.WeaponAnim = NingunArma
        Call ChangeUserChar(userindex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
                
    Case eOBJType.otFlechas
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Invent.MunicionEqpObjIndex = 0
        UserList(userindex).Invent.MunicionEqpSlot = 0
        
    Case eOBJType.otArmadura ' Puede ser un escudo, casco , o vestimenta
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        
        Select Case ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).SubTipo
            Case 0
                UserList(userindex).Invent.ArmourEqpObjIndex = 0
                UserList(userindex).Invent.ArmourEqpSlot = 0
                Call DarCuerpoDesnudo(userindex)
                If Not UserList(userindex).flags.Montando = 1 Then Call ChangeUserChar(userindex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
                
            Case 1
                UserList(userindex).Invent.Object(Slot).Equipped = 0
                UserList(userindex).Invent.CascoEqpObjIndex = 0
                UserList(userindex).Invent.CascoEqpSlot = 0
                
                UserList(userindex).Char.CascoAnim = NingunCasco
                Call ChangeUserChar(userindex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)

                
            Case 2
                UserList(userindex).Invent.Object(Slot).Equipped = 0
                UserList(userindex).Invent.EscudoEqpObjIndex = 0
                UserList(userindex).Invent.EscudoEqpSlot = 0
                
                UserList(userindex).Char.ShieldAnim = NingunEscudo
                Call ChangeUserChar(userindex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
                
        End Select
End Select

Call WriteUpdateUserStats(userindex)
Call UpdateUserInv(False, userindex, Slot)

End Sub

Function SexoPuedeUsarItem(ByVal userindex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo Errhandler

If Not ObjIndex <> 0 Then Exit Function



If ObjData(ObjIndex).Mujer = 1 Then
    SexoPuedeUsarItem = UserList(userindex).Genero <> eGenero.Hombre
ElseIf ObjData(ObjIndex).Hombre = 1 Then
    SexoPuedeUsarItem = UserList(userindex).Genero <> eGenero.Mujer
Else
    SexoPuedeUsarItem = True
End If

Exit Function
Errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal userindex As Integer, ByVal ObjIndex As Integer) As Boolean

If Not ObjIndex <> 0 Then Exit Function

If ObjData(ObjIndex).Real = 1 Then
    FaccionPuedeUsarItem = esArmada(userindex)
ElseIf ObjData(ObjIndex).Caos = 1 Then
    FaccionPuedeUsarItem = esCaos(userindex)
ElseIf ObjData(ObjIndex).Milicia Then
    FaccionPuedeUsarItem = esMili(userindex)
Else
    FaccionPuedeUsarItem = True
End If

End Function

Sub EquiparInvItem(ByVal userindex As Integer, ByVal Slot As Byte)
On Error GoTo Errhandler

'Equipa un item del inventario
Dim Obj As ObjData
Dim ObjIndex As Integer

ObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
Obj = ObjData(ObjIndex)

If Not Obj.MinELV = 0 Then
    If Obj.MinELV > UserList(userindex).Stats.ELV Then
        Call WriteConsoleMsg(1, userindex, "Debes ser nivel " & Obj.MinELV & " para poder utilizar este objeto.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
End If


If Obj.Newbie = 1 And Not EsNewbie(userindex) Then
     Call WriteConsoleMsg(1, userindex, "Solo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
     Exit Sub
End If
        
Select Case Obj.OBJType
    Case eOBJType.otMonturas
        If UserList(userindex).flags.Metamorfosis = 1 Then
            Call WriteConsoleMsg(1, userindex, "Transformado no puedes utilizar este objeto. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(userindex).flags.Muerto = 1 Then
            Call WriteMsg(userindex, 1)
            Exit Sub
        End If
        
        Call DoEquita(userindex, ObjIndex, Slot)
    Case eOBJType.otHerramientas
        If UserList(userindex).Invent.Object(Slot).Equipped Then
            Call Desequipar(userindex, Slot)
            Exit Sub
        End If
        
        If UserList(userindex).Invent.AnilloEqpSlot <> 0 Then
            Call Desequipar(userindex, UserList(userindex).Invent.AnilloEqpSlot)
        End If
        
        UserList(userindex).Invent.AnilloEqpSlot = Slot
        UserList(userindex).Invent.AnilloEqpObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
        UserList(userindex).Invent.Object(Slot).Equipped = 1
        
    Case eOBJType.otItemsMagicos
        If UserList(userindex).Invent.Object(Slot).Equipped Then
            'Quitamos del inv el item
            Call Desequipar(userindex, Slot)
            Exit Sub
        End If
        
        If UserList(userindex).Invent.MagicIndex <> 0 Then
            Call Desequipar(userindex, UserList(userindex).Invent.MagicSlot)
        End If
        
        UserList(userindex).Invent.MagicIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
        UserList(userindex).Invent.MagicSlot = Slot
        UserList(userindex).Invent.Object(Slot).Equipped = 1
        
        If Obj.EfectoMagico = eMagicType.ModificaAtributo Then
            If Obj.QueAtributo <> 0 Then
                UserList(userindex).Stats.UserAtributos(Obj.QueAtributo) = UserList(userindex).Stats.UserAtributos(Obj.QueAtributo) + Obj.CuantoAumento
            End If
        ElseIf Obj.EfectoMagico = eMagicType.ModificaSkill Then
            If Obj.QueSkill <> 0 Then
                UserList(userindex).Stats.UserSkills(Obj.QueSkill) = UserList(userindex).Stats.UserSkills(Obj.QueSkill) + Obj.CuantoAumento
            End If
        End If
        
    Case eOBJType.otFlechas
       If ClasePuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) And _
          FaccionPuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) Then
                
                'Si esta equipado lo quita
                If UserList(userindex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(userindex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(userindex).Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(userindex, UserList(userindex).Invent.MunicionEqpSlot)
                End If
        
                UserList(userindex).Invent.Object(Slot).Equipped = 1
                UserList(userindex).Invent.MunicionEqpObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
                UserList(userindex).Invent.MunicionEqpSlot = Slot
                
       Else
            Call WriteConsoleMsg(1, userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
    Case eOBJType.otArmadura
    
        If UserList(userindex).flags.Metamorfosis = 1 Then Exit Sub


        If ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).SubTipo = 0 Then
            If UserList(userindex).flags.Navegando = 1 Then Exit Sub
            'Nos aseguramos que puede usarla
            If ClasePuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) And _
               SexoPuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) And _
               CheckRazaUsaRopa(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) And _
               FaccionPuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) Then
               
               'Si esta equipado lo quita
                If UserList(userindex).Invent.Object(Slot).Equipped Then
                    Call Desequipar(userindex, Slot)
                    Call DarCuerpoDesnudo(userindex)
                    If Not UserList(userindex).flags.Montando = 1 Then Call ChangeUserChar(userindex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
                    Exit Sub
                End If
        
                'Quita el anterior
                If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
                    Call Desequipar(userindex, UserList(userindex).Invent.ArmourEqpSlot)
                End If
        
                'Lo equipa
                UserList(userindex).Invent.Object(Slot).Equipped = 1
                UserList(userindex).Invent.ArmourEqpObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
                UserList(userindex).Invent.ArmourEqpSlot = Slot
                
                UserList(userindex).Char.body = Obj.Ropaje
                If Not UserList(userindex).flags.Montando = 1 Then Call ChangeUserChar(userindex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)

                UserList(userindex).flags.Desnudo = 0
            Else
                Call WriteConsoleMsg(1, userindex, "Tu clase,genero o raza no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
            End If
        ElseIf ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).SubTipo = 1 Then
            If UserList(userindex).flags.Navegando = 1 Then Exit Sub
            If ClasePuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(userindex).Invent.Object(Slot).Equipped Then
                    Call Desequipar(userindex, Slot)
                    
                    UserList(userindex).Char.CascoAnim = NingunCasco
                    Call ChangeUserChar(userindex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
                    Exit Sub
                End If
        
                'Quita el anterior
                If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
                    Call Desequipar(userindex, UserList(userindex).Invent.CascoEqpSlot)
                End If
        
                'Lo equipa
                
                UserList(userindex).Invent.Object(Slot).Equipped = 1
                UserList(userindex).Invent.CascoEqpObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
                UserList(userindex).Invent.CascoEqpSlot = Slot

                UserList(userindex).Char.CascoAnim = Obj.CascoAnim
                Call ChangeUserChar(userindex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)

            Else
                Call WriteConsoleMsg(1, userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
            End If
        ElseIf ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).SubTipo = 2 Then
            If UserList(userindex).flags.Navegando = 1 Then Exit Sub
                If ClasePuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) And FaccionPuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) Then
                
                If UserList(userindex).Invent.WeaponEqpObjIndex <> 0 Then
                    If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).DosManos = 1 Then
                        If ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).DosManos = 0 Then
                            WriteMsg userindex, 22
                            Exit Sub
                        End If
                    End If
                End If
                
                'Si esta equipado lo quita
                If UserList(userindex).Invent.Object(Slot).Equipped Then
                    Call Desequipar(userindex, Slot)
                    
                    UserList(userindex).Char.ShieldAnim = NingunEscudo
                    Call ChangeUserChar(userindex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
                    
                    Exit Sub
                End If
                
                'Quita el anterior
                If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then
                    Call Desequipar(userindex, UserList(userindex).Invent.EscudoEqpSlot)
                End If
                
                'Lo equipa
                
                UserList(userindex).Invent.Object(Slot).Equipped = 1
                UserList(userindex).Invent.EscudoEqpObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
                UserList(userindex).Invent.EscudoEqpSlot = Slot
                
                UserList(userindex).Char.ShieldAnim = Obj.ShieldAnim
                Call ChangeUserChar(userindex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
             Else
                 Call WriteConsoleMsg(1, userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
             End If
        End If
    Case eOBJType.otWeapon, eOBJType.otHerramientas
       If ClasePuedeUsarItem(userindex, ObjIndex) And _
          FaccionPuedeUsarItem(userindex, ObjIndex) Then
            If UserList(userindex).Invent.EscudoEqpObjIndex <> 0 Then
                If ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).DosManos = 1 Then
                    If ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).DosManos = 0 Then
                        WriteMsg userindex, 23
                        Exit Sub
                    End If
                End If
            End If
            
            'Si esta equipado lo quita
            If UserList(userindex).Invent.Object(Slot).Equipped Then
                'Quitamos del inv el item
                Call Desequipar(userindex, Slot)
                Exit Sub
            End If
            
            'Quitamos el elemento anterior
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                Call Desequipar(userindex, UserList(userindex).Invent.WeaponEqpSlot)
            End If
            
            If UserList(userindex).Invent.NudiEqpIndex > 0 Then
                Call Desequipar(userindex, UserList(userindex).Invent.NudiEqpSlot)
            End If
            
            UserList(userindex).Invent.Object(Slot).Equipped = 1
            UserList(userindex).Invent.WeaponEqpObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
            UserList(userindex).Invent.WeaponEqpSlot = Slot
            
            'Sonido
            Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_SACARARMA, UserList(userindex).pos.x, UserList(userindex).pos.Y))
            
            UserList(userindex).Char.WeaponAnim = Obj.WeaponAnim
            Call ChangeUserChar(userindex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
       Else
            Call WriteConsoleMsg(1, userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
       
    Case eOBJType.otNudillos
        If ClasePuedeUsarItem(userindex, ObjIndex) Then
            'Si esta equipado lo quita
            If UserList(userindex).Invent.Object(Slot).Equipped Then
                'Quitamos del inv el item
                Call Desequipar(userindex, Slot)
                'Animacion por defecto
                UserList(userindex).Char.WeaponAnim = NingunArma
                Call ChangeUserChar(userindex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
                Exit Sub
            End If
            
            'Quitamos el arma si tiene alguna equipada
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                Call Desequipar(userindex, UserList(userindex).Invent.WeaponEqpSlot)
            End If
            
            If UserList(userindex).Invent.NudiEqpIndex > 0 Then
                Call Desequipar(userindex, UserList(userindex).Invent.NudiEqpSlot)
            End If
            
            UserList(userindex).Invent.Object(Slot).Equipped = 1
            UserList(userindex).Invent.NudiEqpIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
            UserList(userindex).Invent.NudiEqpSlot = Slot

            UserList(userindex).Char.WeaponAnim = Obj.WeaponAnim
            Call ChangeUserChar(userindex, UserList(userindex).Char.body, UserList(userindex).Char.Head, UserList(userindex).Char.heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
       Else
            Call WriteConsoleMsg(1, userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
        
End Select

'Actualiza
Call UpdateUserInv(False, userindex, Slot)

Exit Sub
Errhandler:
Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & err.Number & " - Error Description : " & err.description)
End Sub

 Function CheckRazaUsaRopa(ByVal userindex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo Errhandler


 

If ObjData(ItemIndex).RazaTipo > 0 Then
    'Verifica si la raza puede usar la ropa
    If UserList(userindex).Raza = eRaza.Humano Or _
       UserList(userindex).Raza = eRaza.Elfo Or _
       UserList(userindex).Raza = eRaza.Drow Then
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaTipo = 1)
    ElseIf UserList(userindex).Raza = eRaza.Orco Then
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaTipo = 3)
    Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaTipo = 2)
    End If
Else
    'Verifica si la raza puede usar la ropa
    If UserList(userindex).Raza = eRaza.Humano Or _
       UserList(userindex).Raza = eRaza.Elfo Or _
       UserList(userindex).Raza = eRaza.Drow Or _
       UserList(userindex).Raza = eRaza.Orco Then
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
    Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
    End If
End If

Exit Function
Errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal userindex As Integer, ByVal Slot As Byte)
'*************************************************
'Author: Unknown
'Last modified: 24/01/2007
'Handels the usage of items from inventory box.
'24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legión.
'24/01/2007 Pablo (ToxicWaste) - Utilización nueva de Barco en lvl 20 por clase Pirata y Pescador.
'*************************************************
On Error GoTo hayerror


Dim Obj As ObjData
Dim ObjIndex As Integer
Dim TargObj As ObjData
Dim MiObj As Obj

With UserList(userindex)

If .Invent.Object(Slot).Amount = 0 Then Exit Sub

Obj = ObjData(.Invent.Object(Slot).ObjIndex)

If Not Obj.MinELV = 0 Then
    If Obj.MinELV > .Stats.ELV Then
        Call WriteConsoleMsg(1, userindex, "Debes ser nivel " & Obj.MinELV & " para poder utilizar este objeto.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
End If


If Obj.Newbie = 1 And Not EsNewbie(userindex) Then
    Call WriteConsoleMsg(1, userindex, "Solo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Obj.OBJType = eOBJType.otWeapon Then
    If Obj.proyectil = 1 Then
        If Not .flags.ModoCombate Then
            Call WriteConsoleMsg(1, userindex, "No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
        If Not IntervaloPermiteUsar(userindex, False) Then Exit Sub
    Else
        'dagas
        If Not IntervaloPermiteUsar(userindex) Then Exit Sub
    End If
Else
    If Not IntervaloPermiteUsar(userindex) Then Exit Sub
End If

ObjIndex = .Invent.Object(Slot).ObjIndex
.flags.TargetObjInvIndex = ObjIndex
.flags.TargetObjInvSlot = Slot

Select Case Obj.OBJType
    Case eOBJType.otUseOnce
        If .flags.Muerto = 1 Then
            Call WriteMsg(userindex, 1)
            Exit Sub
        End If

        'Usa el item
        .Stats.MinHAM = .Stats.MinHAM + Obj.MinHAM
        If .Stats.MinHAM > .Stats.MaxHAM Then _
            .Stats.MinHAM = .Stats.MaxHAM
        .flags.Hambre = 0
        Call WriteUpdateHungerAndThirst(userindex)
        'Sonido
        
        Call ReproducirSonido(SendTarget.ToPCArea, userindex, 7)

        'Quitamos del inv el item
        Call QuitarUserInvItem(userindex, Slot, 1)
        
        Call UpdateUserInv(False, userindex, Slot)

    Case eOBJType.otGuita
        If .flags.Muerto = 1 Then
            Call WriteMsg(userindex, 1)
            Exit Sub
        End If
        
        .Stats.GLD = .Stats.GLD + .Invent.Object(Slot).Amount
        .Invent.Object(Slot).Amount = 0
        .Invent.Object(Slot).ObjIndex = 0
        .Invent.NroItems = .Invent.NroItems - 1
        
        Call UpdateUserInv(False, userindex, Slot)
        Call WriteUpdateGold(userindex)
        
    Case eOBJType.otWeapon
        If .flags.Muerto = 1 Then
            Call WriteMsg(userindex, 1)
            Exit Sub
        End If
        
        If Not .Stats.MinSTA > 0 Then
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(1, userindex, "Estas muy cansado", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, userindex, "Estas muy cansada", FontTypeNames.FONTTYPE_INFO)
            End If
            Exit Sub
        End If
        
        
        If ObjData(ObjIndex).proyectil = 1 Then
            If .Invent.Object(Slot).Equipped = 0 Then
                Call WriteConsoleMsg(1, userindex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            'liquid: muevo esto aca adentro, para que solo pida modo combate si estamos por usar el arco
            If Not .flags.ModoCombate Then
                Call WriteConsoleMsg(1, userindex, "No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            Call WriteWorkRequestTarget(userindex, Proyectiles)
        ElseIf ObjData(ObjIndex).SubTipo = 5 Then
            If .Invent.Object(Slot).Equipped = 0 Then
                Call WriteConsoleMsg(1, userindex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'liquid: muevo esto aca adentro, para que solo pida modo combate si estamos por usar el arco
            If Not .flags.ModoCombate Then
                Call WriteConsoleMsg(1, userindex, "No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call WriteWorkRequestTarget(userindex, arrojadizas)
        Else
            If .flags.TargetObj = Leña Then
                If .Invent.Object(Slot).ObjIndex = DAGA Then
                    If .Invent.Object(Slot).Equipped = 0 Then
                        Call WriteConsoleMsg(1, userindex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    Call TratarDeHacerFogata(.flags.TargetObjMap, _
                         .flags.TargetObjX, .flags.TargetObjY, userindex)
                End If
            End If
        End If

    
    Case eOBJType.otPociones
        If .flags.Muerto = 1 Then
            Call WriteMsg(userindex, 1)
            Exit Sub
        End If
        
        If Not IntervaloPermiteGolpeUsar(userindex, False) Then
            Call WriteConsoleMsg(1, userindex, "¡¡Debes esperar unos momentos para tomar otra poción!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .flags.TomoPocion = True
        .flags.TipoPocion = Obj.TipoPocion
                
        Select Case .flags.TipoPocion
        
            Case 1 'Modif la agilidad
                .flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If .Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then _
                    .Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
  
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.Y))
                
                Call WriteAgilidad(userindex)
                
            Case 2 'Modif la fuerza
                .flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If .Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then _
                    .Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
   
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.Y))
                
                Call WriteFuerza(userindex)
                
            Case 3 'Pocion roja, restaura HP
                'Usa el item
                .Stats.MinHP = .Stats.MinHP + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If .Stats.MinHP > .Stats.MaxHP Then _
                    .Stats.MinHP = .Stats.MaxHP
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.Y))
            
            Case 4 'Pocion azul, restaura MANA
                'Usa el item
                'nuevo calculo para recargar mana
                .Stats.MinMAN = .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, 4) + .Stats.ELV \ 2 + 40 / .Stats.ELV
                If .Stats.MinMAN > .Stats.MaxMAN Then _
                    .Stats.MinMAN = .Stats.MaxMAN
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.Y))
                
            Case 5 ' Pocion violeta
                If .flags.Envenenado <> 0 Then
                    .flags.Envenenado = 0
                    Call WriteConsoleMsg(1, userindex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
                End If
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.Y))
            Case 6 'Remueve Paralisis
                If UserList(userindex).flags.Paralizado = 1 Or UserList(userindex).flags.Inmovilizado = 1 Then
                
                    UserList(userindex).flags.Inmovilizado = 0
                    UserList(userindex).flags.Paralizado = 0
                    
                    'no need to crypt this
                    Call WriteParalizeOK(userindex)
                    Call QuitarUserInvItem(userindex, Slot, 1)
                    Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_REMO, .pos.x, .pos.Y))
                End If
                
            Case 7 'Metamorfosis Deshablitado
            
            Case 8 'Remueve Ceguera
                If UserList(userindex).flags.Ceguera = 1 Then
                    UserList(userindex).flags.Ceguera = 0
                    UserList(userindex).Counters.Ceguera = 0

                    Call WriteBlindNoMore(userindex)
                    Call QuitarUserInvItem(userindex, Slot, 1)
                    Call FlushBuffer(userindex)
                End If
            Case 9 'Remueve Estupidez
                If UserList(userindex).flags.Estupidez = 1 Then
                    UserList(userindex).flags.Estupidez = 0
                    
                    'no need to crypt this
                    Call WriteDumbNoMore(userindex)
                    Call QuitarUserInvItem(userindex, Slot, 1)
                    Call FlushBuffer(userindex)
                End If
                
            Case 10 'Invisibilidad
                If UserList(userindex).flags.Muerto = 1 Then
                    Call WriteMsg(userindex, 1)
                    Exit Sub
                End If
                 
                If UserList(userindex).Counters.Saliendo Then
                    Call WriteConsoleMsg(1, userindex, "¡No puedes ponerte invisible mientras te encuentres saliendo!", FontTypeNames.FONTTYPE_WARNING)
                    Exit Sub
                End If
                 
                 'No usar invi mapas InviSinEfecto
                If MapInfo(UserList(userindex).pos.map).InviSinEfecto > 0 Then
                    Call WriteConsoleMsg(1, userindex, "¡La invisibilidad no funciona aquí!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(userindex).flags.Navegando = 1 Then
                    Call WriteConsoleMsg(1, userindex, "No puedes estar invisible navegando!!", FontTypeNames.FONTTYPE_BROWNI)
                    Exit Sub
                End If
    
                UserList(userindex).flags.Invisible = 1
                Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageSetInvisible(UserList(userindex).Char.CharIndex, True))
                Call QuitarUserInvItem(userindex, Slot, 1)
                
       End Select
       
       Call WriteUpdateUserStats(userindex)
       Call UpdateUserInv(False, userindex, Slot)

     Case eOBJType.otBebidas
        If .flags.Muerto = 1 Then
            Call WriteMsg(userindex, 1)
            Exit Sub
        End If
        .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed
        If .Stats.MinAGU > .Stats.MaxAGU Then _
            .Stats.MinAGU = .Stats.MaxAGU
        .flags.Sed = 0
        Call WriteUpdateHungerAndThirst(userindex)
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(userindex, Slot, 1)
        
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.Y))
        
        Call UpdateUserInv(False, userindex, Slot)
    
    Case eOBJType.otLlaves
        If .flags.Muerto = 1 Then
            Call WriteMsg(userindex, 1)
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
                     If TargObj.clave = Obj.clave Then
         
                        MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex _
                        = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                        .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                        Call WriteConsoleMsg(1, userindex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     Else
                        Call WriteConsoleMsg(1, userindex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     End If
                  Else
                     If TargObj.clave = Obj.clave Then
                        MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex _
                        = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                        Call WriteConsoleMsg(1, userindex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                        .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                        Exit Sub
                     Else
                        Call WriteConsoleMsg(1, userindex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     End If
                  End If
            Else
                  Call WriteConsoleMsg(1, userindex, "No esta cerrada.", FontTypeNames.FONTTYPE_INFO)
                  Exit Sub
            End If
        End If
    
    Case eOBJType.otBotellaVacia
        If .flags.Muerto = 1 Then
            Call WriteMsg(userindex, 1)
            Exit Sub
        End If
        If Not HayAgua(.pos.map, .flags.TargetX, .flags.TargetY) Then
            Call WriteConsoleMsg(1, userindex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        MiObj.Amount = 1
        MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexAbierta
        Call QuitarUserInvItem(userindex, Slot, 1)
        If Not MeterItemEnInventario(userindex, MiObj) Then
            Call TirarItemAlPiso(.pos, MiObj)
        End If
        
        Call UpdateUserInv(False, userindex, Slot)
    
    Case eOBJType.otBotellaLlena
        If .flags.Muerto = 1 Then
            Call WriteMsg(userindex, 1)
            Exit Sub
        End If
        .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed
        If .Stats.MinAGU > .Stats.MaxAGU Then _
            .Stats.MinAGU = .Stats.MaxAGU
        .flags.Sed = 0
        Call WriteUpdateHungerAndThirst(userindex)
        MiObj.Amount = 1
        MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexCerrada
        Call QuitarUserInvItem(userindex, Slot, 1)
        If Not MeterItemEnInventario(userindex, MiObj) Then
            Call TirarItemAlPiso(.pos, MiObj)
        End If
        
        Call UpdateUserInv(False, userindex, Slot)
    
    Case eOBJType.otPergaminos
        If .flags.Muerto = 1 Then
            Call WriteMsg(userindex, 1)
            Exit Sub
        End If
        
        If .Stats.MaxMAN > 0 Then
            If .flags.Hambre = 0 And _
                .flags.Sed = 0 Then
                Call AgregarHechizo(userindex, Slot)
                Call UpdateUserInv(False, userindex, Slot)
            Else
                Call WriteConsoleMsg(1, userindex, "Estas demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(1, userindex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_INFO)
        End If
    Case eOBJType.otMinerales
        If .flags.Muerto = 1 Then
             Call WriteMsg(userindex, 1)
             Exit Sub
        End If
        Call WriteWorkRequestTarget(userindex, FundirMetal)
        .flags.Lingoteando = Slot
        
    Case eOBJType.otInstrumentos
        If .flags.Muerto = 1 Then
            Call WriteMsg(userindex, 1)
            Exit Sub
        End If
        
        If Obj.Real Then '¿Es el Cuerno Real?
            If FaccionPuedeUsarItem(userindex, ObjIndex) Then
                If MapInfo(.pos.map).Pk = False Then
                    Call WriteConsoleMsg(1, userindex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call SendData(SendTarget.ToMap, .pos.map, PrepareMessagePlayWave(Obj.Snd1, .pos.x, .pos.Y))
                Exit Sub
            Else
                Call WriteConsoleMsg(1, userindex, "Solo Miembros de la Armada Real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        ElseIf Obj.Caos Then '¿Es el Cuerno Legión?
            If FaccionPuedeUsarItem(userindex, ObjIndex) Then
                If MapInfo(.pos.map).Pk = False Then
                    Call WriteConsoleMsg(1, userindex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call SendData(SendTarget.ToMap, .pos.map, PrepareMessagePlayWave(Obj.Snd1, .pos.x, .pos.Y))
                Exit Sub
            Else
                Call WriteConsoleMsg(1, userindex, "Solo Miembros de la Legión Oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        'Si llega aca es porque es o Laud o Tambor o Flauta
        Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(Obj.Snd1, .pos.x, .pos.Y))
       
    Case eOBJType.otMonturas
        If UserList(userindex).flags.Metamorfosis = 1 Then
            Call WriteConsoleMsg(1, userindex, "Transformado no puedes utilizar este objeto. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(userindex).flags.Muerto = 1 Then
            Call WriteMsg(userindex, 1)
            Exit Sub
        End If
        
        If LegalPos(.pos.map, .pos.x, .pos.Y) Then
            Call DoEquita(userindex, ObjIndex, Slot)
        End If
        
    Case eOBJType.otBarcos
        If UserList(userindex).flags.Metamorfosis = 1 Then
            Call WriteConsoleMsg(1, userindex, "Transformado no puedes utilizar este objeto. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        If ((LegalPos(.pos.map, .pos.x - 1, .pos.Y, True, False) _
                Or LegalPos(.pos.map, .pos.x, .pos.Y - 1, True, False) _
                Or LegalPos(.pos.map, .pos.x + 1, .pos.Y, True, False) _
                Or LegalPos(.pos.map, .pos.x, .pos.Y + 1, True, False)) _
                And .flags.Navegando = 0) _
                Or .flags.Navegando = 1 Then
            Call DoNavega(userindex, Obj, Slot)
        Else
            Call WriteConsoleMsg(1, userindex, "¡Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)
        End If
        
    Case eOBJType.otPasajes
        If UserList(userindex).flags.Metamorfosis = 1 Then
            Call WriteConsoleMsg(1, userindex, "Transformado no puedes utilizar este objeto. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        If UserList(userindex).flags.Muerto = 1 Then
            Call WriteMsg(userindex, 1)
            Exit Sub
        End If
       
        If UserList(userindex).flags.TargetNPC <> 0 Then
            If Left$(Npclist(UserList(userindex).flags.TargetNPC).Name, 6) <> "Pirata" Then
                Call WriteConsoleMsg(1, userindex, "Primero debes hacer click sobre el pirata.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            Call WriteConsoleMsg(1, userindex, "Primero debes hacer click sobre el pirata.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
       
        If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 3 Then
            Call WriteConsoleMsg(1, userindex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Not MapaValido(Obj.HastaMap) Then
            Call WriteConsoleMsg(1, userindex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
       
        If UserList(userindex).Stats.UserSkills(eSkill.Navegacion) < Obj.CantidadSkill Then
            Call WriteConsoleMsg(1, userindex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
       
        Call WarpUserChar(userindex, Obj.HastaMap, Obj.HastaX, Obj.HastaY, True)
        Call WriteConsoleMsg(1, userindex, "Has viajado por varios días, te sientes exhausto!", FontTypeNames.FONTTYPE_BROWNB)
        UserList(userindex).Stats.MinAGU = 0
        UserList(userindex).Stats.MinHAM = 0
        UserList(userindex).flags.Sed = 1
        UserList(userindex).flags.Hambre = 1
        
        UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userindex).Stats.UserAtributosBackUP(eAtributos.Agilidad)
        UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributosBackUP(eAtributos.Fuerza)
        
        Call WriteAgilidad(userindex)
        Call WriteFuerza(userindex)
        
        Call WriteUpdateHungerAndThirst(userindex)
        Call QuitarUserInvItem(userindex, Slot, 1)
        Call UpdateUserInv(False, userindex, Slot)
        
    Case eOBJType.otruna
        'COMIENZO RUNA////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
Dim map As Integer
Dim x As Integer
Dim Y As Integer
        If MapInfo(.pos.map).Pk And UserList(userindex).flags.Muerto = 0 Then Exit Sub
        
        Select Case UserList(userindex).Hogar
        
        Case 0
        map = 34
        x = 22
        Y = 60
        Case 1
        map = 194
        x = 61
        Y = 71
        Case 2
        map = 112
        x = 25
        Y = 59
        Case 3
        map = 20
        x = 77
        Y = 26
        
        End Select
        

'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martín Sotuyo Dodero (Maraxus)
If MapData(map, x, Y).userindex <> 0 Or MapData(map, x, Y).NpcIndex <> 0 Then
    Dim FoundPlace As Boolean
    Dim esAgua As Boolean
    Dim tX As Long
    Dim tY As Long
    
    FoundPlace = False
    esAgua = HayAgua(map, x, Y)
    
    For tY = Y - 1 To Y + 1
        For tX = x - 1 To x + 1
            If esAgua Then
                'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
                If LegalPos(map, tX, tY, True, False) Then
                    FoundPlace = True
                    Exit For
                End If
            Else
                'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
                If LegalPos(map, tX, tY, False, True) Then
                    FoundPlace = True
                    Exit For
                End If
            End If
        Next tX
        
        If FoundPlace Then _
            Exit For
    Next tY
    
    If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
        x = tX
        Y = tY
    Else
        'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
        If MapData(map, x, Y).userindex <> 0 Then
            'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
            If UserList(MapData(map, x, Y).userindex).ComUsu.DestUsu > 0 Then
                'Le avisamos al que estaba comerciando que se tuvo que ir.
                If UserList(UserList(MapData(map, x, Y).userindex).ComUsu.DestUsu).flags.UserLogged Then
                    Call FinComerciarUsu(UserList(MapData(map, x, Y).userindex).ComUsu.DestUsu)
                    Call WriteConsoleMsg(1, UserList(MapData(map, x, Y).userindex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                    Call FlushBuffer(UserList(MapData(map, x, Y).userindex).ComUsu.DestUsu)
                End If
                'Lo sacamos.
                If UserList(MapData(map, x, Y).userindex).flags.UserLogged Then
                    Call FinComerciarUsu(MapData(map, x, Y).userindex)
                    Call WriteErrorMsg(MapData(map, x, Y).userindex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                    Call FlushBuffer(MapData(map, x, Y).userindex)
                End If
            End If

        End If
    End If
End If
        
        Call WarpUserChar(userindex, map, x, Y, True)
        
        'FIN RUNA/////////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
        '/////////////////////////////
End Select

If UserList(userindex).Invent.Object(Slot).Equipped <> 0 Then
    If ObjIndex = SERRUCHO_CARPINTERO Then
        Call EnivarObjConstruibles(userindex)
        Call WriteShowCarpenterForm(userindex)
    ElseIf ObjIndex = COSTURERO Then
        Call EnivarObjTejibles(userindex)
        Call WriteShowSastreForm(userindex)
    ElseIf ObjIndex = OLLA Then
        Call EnivarObjalquimia(userindex)
        Call WriteShowalquimiaForm(userindex)
    End If
End If
        
End With

Exit Sub

hayerror:

LogError ("Error en Userinvitem: " & err.Number & " Desc: " & err.description)


End Sub

Sub EnivarArmasConstruibles(ByVal userindex As Integer)

Call WriteBlacksmithWeapons(userindex)

End Sub
 
Sub EnivarObjConstruibles(ByVal userindex As Integer)

Call WriteCarpenterObjects(userindex)

End Sub

Sub EnivarObjalquimia(ByVal userindex As Integer)

Call WriteAlquimiaObjects(userindex)

End Sub

Sub EnivarObjTejibles(ByVal userindex As Integer)

Call WriteTejiblesObjects(userindex)

End Sub

Sub EnivarArmadurasConstruibles(ByVal userindex As Integer)

Call WriteBlacksmithArmors(userindex)

End Sub

Sub TirarTodo(ByVal userindex As Integer)
On Error GoTo hayerror

If MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).Trigger = 6 Then Exit Sub
If MapInfo(UserList(userindex).pos.map).Seguro = 1 Then Exit Sub




Call TirarTodosLosItems(userindex)

Dim Cantidad As Long
Cantidad = UserList(userindex).Stats.GLD

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
            ObjData(index).Caos <> 1 And _
            ObjData(index).Real <> 1 And _
            ObjData(index).Milicia <> 1
            


End Function

Sub TirarTodosLosItems(ByVal userindex As Integer)
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    
    'Ponemos aca el carro de mineria
    Dim Carro As Byte
    Dim Minerales As Integer
    Dim Porc As Byte
    Dim Hierro As Integer, Plata As Integer, oro As Integer

    Carro = Have_Obj_Slot(CARROMINERO, userindex)
    If Carro > 0 Then
        Hierro = Have_Obj_To_Slot(iMinerales.HierroCrudo, Carro, userindex)
        Plata = Have_Obj_To_Slot(iMinerales.PlataCruda, Carro, userindex)
        oro = Have_Obj_To_Slot(iMinerales.OroCrudo, Carro, userindex)
        
        If Hierro > 0 Then
            Porc = Porc + 1
        End If
        
        If Plata > 0 Then
            Porc = Porc + 1
        End If
        
        If oro > 0 Then
            Porc = Porc + 1
        End If
        
        If Hierro > 0 Then Hierro = Porcentaje(Hierro, (100 - ObjData(UserList(userindex).Invent.Object(Carro).ObjIndex).CuantoAumento) / Porc)
        If Plata > 0 Then Plata = Porcentaje(Plata, (100 - ObjData(UserList(userindex).Invent.Object(Carro).ObjIndex).CuantoAumento) / Porc)
        If oro > 0 Then oro = Porcentaje(oro, (100 - ObjData(UserList(userindex).Invent.Object(Carro).ObjIndex).CuantoAumento) / Porc)
        
        If Porc > 0 Then
            For i = 1 To Carro
                If UserList(userindex).Invent.Object(i).ObjIndex = iMinerales.HierroCrudo Then
                    If Hierro > 0 Then
                        TirarObjeto userindex, i, Hierro
                        Hierro = Hierro - IIf(UserList(userindex).Invent.Object(i).Amount > Hierro, Hierro, UserList(userindex).Invent.Object(i).Amount)
                    End If
                ElseIf UserList(userindex).Invent.Object(i).ObjIndex = iMinerales.PlataCruda Then
                    If Plata > 0 Then
                        TirarObjeto userindex, i, Plata
                        Plata = Plata - IIf(UserList(userindex).Invent.Object(i).Amount > Plata, Plata, UserList(userindex).Invent.Object(i).Amount)
                    End If
                ElseIf UserList(userindex).Invent.Object(i).ObjIndex = iMinerales.OroCrudo Then
                    If oro > 0 Then
                        TirarObjeto userindex, i, oro
                        oro = oro - IIf(UserList(userindex).Invent.Object(i).Amount > oro, oro, UserList(userindex).Invent.Object(i).Amount)
                    End If
                End If
            Next i
        End If
    End If
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(userindex).Invent.Object(i).ObjIndex
        If ItemIndex > 0 Then
             If ItemSeCae(ItemIndex) Then
                NuevaPos.x = 0
                NuevaPos.Y = 0
                
                'Creo el Obj
                MiObj.Amount = UserList(userindex).Invent.Object(i).Amount
                MiObj.ObjIndex = ItemIndex
                'Pablo (ToxicWaste) 24/01/2007
                'Si es pirata y usa un Galeón entonces no explota los items. (en el agua)
                If UserList(userindex).Clase = eClass.Mercenario And UserList(userindex).Invent.BarcoObjIndex = 476 Then
                    TileLibre UserList(userindex).pos, NuevaPos, MiObj, False, True
                Else
                    TileLibre UserList(userindex).pos, NuevaPos, MiObj, True, True
                End If
                
                If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(userindex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.x, NuevaPos.Y)
                End If
             End If
        End If
    Next i
End Sub
Sub TirarObjeto(ByVal userindex As Integer, ByVal Slot As Byte, ByVal Cant As Integer)
    Dim MiObj As Obj
    Dim NuevaPos As WorldPos
    
  
    
    If Cant > UserList(userindex).Invent.Object(Slot).Amount Then _
        Cant = UserList(userindex).Invent.Object(Slot).Amount
    'Creo el Obj
    MiObj.Amount = Cant
    MiObj.ObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
    
    If UserList(userindex).Clase = eClass.Mercenario And UserList(userindex).Invent.BarcoObjIndex = 476 Then
        TileLibre UserList(userindex).pos, NuevaPos, MiObj, False, True
    Else
        TileLibre UserList(userindex).pos, NuevaPos, MiObj, True, True
    End If
                
    If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then
        Call DropObj(userindex, Slot, Cant, NuevaPos.map, NuevaPos.x, NuevaPos.Y)
    End If
End Sub
Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean

ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal userindex As Integer)
Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As Obj
Dim ItemIndex As Integer

If MapData(UserList(userindex).pos.map, UserList(userindex).pos.x, UserList(userindex).pos.Y).Trigger = 6 Then Exit Sub

For i = 1 To MAX_INVENTORY_SLOTS
    ItemIndex = UserList(userindex).Invent.Object(i).ObjIndex
    If ItemIndex > 0 Then
        If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
            NuevaPos.x = 0
            NuevaPos.Y = 0
            
            'Creo MiObj
            MiObj.Amount = UserList(userindex).Invent.Object(i).ObjIndex
            MiObj.ObjIndex = ItemIndex
            'Pablo (ToxicWaste) 24/01/2007
            'Tira los Items no newbies en todos lados.
            TileLibre UserList(userindex).pos, NuevaPos, MiObj, True, True
            If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then
                If MapData(NuevaPos.map, NuevaPos.x, NuevaPos.Y).ObjInfo.ObjIndex = 0 Then Call DropObj(userindex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.x, NuevaPos.Y)
            End If
        End If
    End If
Next i

End Sub
