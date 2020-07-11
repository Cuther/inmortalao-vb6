Attribute VB_Name = "modSubastas"
Option Explicit

Type tSubasta
    active As Boolean 'No se usa este slot en el array ?
    
    ObjIndex As Integer 'Objeto qe se subasta
    Cant     As Integer 'Cantidad de objetos
    
    nckVndedor As String * 32 'Nick del vendedor
    nckCmprdor As String * 32 'Nick del que oferto en caso de ser oferta
    
    actOfert As Long 'Actual oferta en Oro
    fnlOfert As Long 'Oferta final en caso de comprar de una
    
    hsDura As Byte 'Duracion
    mnDura As Byte 'Duracion
End Type

'Tamaño de subastas : 8000 bytes en memoria : 7,81kbs
Public lstSubastas(1 To 100) As tSubasta
Public LastSubasta As Byte
Public Sub sFinalizar(ByVal sI As Byte)
On Error Resume Next
    
    If sI > 100 Or sI <= 0 Then Exit Sub 'Si sobrepasa el max
    If lstSubastas(sI).active = False Then Exit Sub 'Si es valida
    
    Dim tIndex As Integer, tGld As Long, tCantObjs As Integer
    Dim objs As obj
        
    objs.Amount = lstSubastas(sI).Cant
    objs.ObjIndex = lstSubastas(sI).ObjIndex
    
    If lstSubastas(sI).nckCmprdor = "" Then 'Nadie :(
        tIndex = NameIndex(lstSubastas(sI).nckVndedor)
        
        If tIndex > 0 Then
            sCancelar tIndex, sI
        Else
            Add_Item_Subast UCase$(lstSubastas(sI).nckVndedor), objs.ObjIndex, objs.Amount
        End If
    Else ':D
        'Primero, le damos el oro al user, o se lo depositamos
        tIndex = NameIndex(lstSubastas(sI).nckVndedor)
        If tIndex > 0 Then 'Esta online?
            UserList(tIndex).Stats.GLD = UserList(tIndex).Stats.GLD + lstSubastas(sI).actOfert
            Call WriteUpdateGold(tIndex)
        Else
            Add_GLD_Subast lstSubastas(sI).nckVndedor, lstSubastas(sI).actOfert
        End If

        'Segundo le depositamos o damos el item al verga este
        tIndex = NameIndex(lstSubastas(sI).nckCmprdor)
        
        If tIndex > 0 Then
            'Tiene el banco lleno ?
            If Not UserList(tIndex).BancoInvent.NroItems = MAX_BANCOINVENTORY_SLOTS Then
                Call sUserDejaObj(tIndex, lstSubastas(sI).ObjIndex, lstSubastas(sI).Cant)
            Else
                If UserList(tIndex).Invent.NroItems = MAX_INVENTORY_SLOTS Then
                    If MapData(UserList(tIndex).Pos.map, UserList(tIndex).Pos.x, UserList(tIndex).Pos.Y).ObjInfo.ObjIndex = 0 Then
                        Call WriteConsoleMsg(1, tIndex, "El item de la subasta ha sido arrojado debajo tuyo.", FontTypeNames.FONTTYPE_INFO)
                        Call MakeObj(objs, UserList(tIndex).Pos.map, UserList(tIndex).Pos.x, UserList(tIndex).Pos.Y)
                    End If
                Else
                    'Bueno tiene lugar en el inventario
                    MeterItemEnInventario tIndex, objs
                End If
            End If
        Else
            Add_Item_Subast UCase$(lstSubastas(sI).nckCmprdor), objs.ObjIndex, objs.Amount
        End If
    End If
    
    sBorrar sI
        
End Sub
Public Sub sSubastar(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer, _
                    ByVal fnlOfert As Long, hsDura As Byte, ByVal prcOfert As Integer)
    'Sistema de subasta programado por mannakia
    
    Dim i As Byte 'Para el bucle
    Dim sI As Byte 'Index de la subasta elegida
    
    For i = 1 To LastSubasta
        If lstSubastas(i).active = False Then sI = i
    Next i
    
    If sI = 0 Then
        If LastSubasta = 100 Then
            Call WriteConsoleMsg(1, UserIndex, "El limite de subasta esta al máximo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        LastSubasta = LastSubasta + 1
        sI = LastSubasta
    End If
    
    With lstSubastas(sI)
        .active = True 'Subasta valida
        
        .actOfert = prcOfert 'Cantidad inicial para ofrecer
        .fnlOfert = fnlOfert 'Cantidad para comprar de una
        
        .Cant = Cantidad 'Cantidad de objeto
        .ObjIndex = ObjIndex 'Que objeto
        
        .nckVndedor = UserList(UserIndex).Name ' Qien pija lo vende?
        If hsDura > 2 And hsDura < 49 Then
            .hsDura = hsDura - 1
        Else
            .hsDura = 2 'Que pt q es
        End If
        
        If .hsDura * 10 > CLng(.fnlOfert * 0.05) Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - .hsDura * 10
        Else
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - .fnlOfert * 0.05
        End If
        
        .mnDura = 59
    End With
End Sub
Public Sub sOfrecer(ByVal UserIndex As Integer, ByVal Cantidad As Integer, ByVal sI As Byte)
On Error Resume Next
    If sI > 100 Or sI <= 0 Then Exit Sub 'Si sobrepasa el max
    If lstSubastas(sI).active = False Then Exit Sub 'Si es valida
    
    If UserList(UserIndex).Stats.GLD < Cantidad Then ' Si tenemos lo que ofrecimos
        Call WriteConsoleMsg(1, UserIndex, "No dispones el oro ofrecido.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If lstSubastas(sI).fnlOfert <= Cantidad And Not lstSubastas(sI).fnlOfert = 0 Then 'Si esta hueviando
        Call WriteConsoleMsg(1, UserIndex, "El oro ofrecido es mayor al de la cantidad final. Para eso compre el objeto.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If lstSubastas(sI).actOfert > Cantidad Then 'Si es vivo y ofrece menos de lo actual
        Call WriteConsoleMsg(1, UserIndex, "Tu oferta debe sobrepasar la oferta actual.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    'Pasamos pruebas y todo OK
    lstSubastas(sI).actOfert = Cantidad
    lstSubastas(sI).nckCmprdor = UserList(UserIndex).Name
    
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Cantidad
    Exit Sub
End Sub
Public Sub sCancelar(ByVal UserIndex As Integer, ByVal sI As Byte)
On Error Resume Next
    If sI > 100 Or sI <= 0 Then Exit Sub 'Si sobrepasa el max
    If lstSubastas(sI).active = False Then Exit Sub 'Si es valida
    
    Dim objs As obj
    objs.Amount = lstSubastas(sI).Cant
    objs.ObjIndex = lstSubastas(sI).ObjIndex
    
    'Primero le devolvemos al subastador
    
    'Tiene el inventario lleno ?
    If UserList(UserIndex).Invent.NroItems = MAX_INVENTORY_SLOTS Then
        'Tambien lleno ?
        If UserList(UserIndex).BancoInvent.NroItems = MAX_BANCOINVENTORY_SLOTS Then
            If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex = 0 Then
                Call WriteConsoleMsg(1, UserIndex, "El item de la subasta ha sido arrojado debajo tuyo.", FontTypeNames.FONTTYPE_INFO)
                Call MakeObj(objs, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y)
            End If
        Else 'Le depositamos el item
            'Agregamos el obj que deposita al banco
            Call sUserDejaObj(UserIndex, lstSubastas(sI).ObjIndex, lstSubastas(sI).Cant)
        End If
    Else
        'Bueno tiene lugar en el inventario
        MeterItemEnInventario UserIndex, objs
    End If
    
    'Segundo le devolvemos al que ofrecio en caso de que halla
    If lstSubastas(sI).nckCmprdor <> "" Then
        Dim tIndex As Integer, tGld As Long
        tIndex = NameIndex(lstSubastas(sI).nckCmprdor)
        If tIndex > 0 Then 'Esta online?
            UserList(tIndex).Stats.GLD = UserList(tIndex).Stats.GLD + lstSubastas(sI).actOfert
            Call WriteUpdateGold(tIndex)
        Else
            Add_GLD_Subast lstSubastas(sI).nckCmprdor, lstSubastas(sI).actOfert
        End If
    End If

    sBorrar sI

End Sub
Public Sub sBorrar(ByVal sI As Byte)
On Error Resume Next
    If sI > 100 Or sI <= 0 Then Exit Sub 'Si sobrepasa el max
    If lstSubastas(sI).active = False Then Exit Sub 'Si es valida
    
    With lstSubastas(sI)
        .active = False
        .actOfert = 0
        .Cant = 0
        .fnlOfert = 0
        .hsDura = 0
        .mnDura = 0
        .nckCmprdor = ""
        .nckVndedor = ""
        .ObjIndex = 0
    End With
End Sub
Public Sub sComprar(ByVal UserIndex As Integer, ByVal sI As Byte)
On Error Resume Next
    If sI > 100 Or sI <= 0 Then Exit Sub 'Si sobrepasa el max
    If lstSubastas(sI).active = False Then Exit Sub 'Si es valida
    
    If UserList(UserIndex).Stats.GLD < lstSubastas(sI).fnlOfert Then  ' Si tenemos lo que ofrecimos
        Call WriteConsoleMsg(1, UserIndex, "No dispones el oro ofrecido.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    'Iniciamos El user que compro esta online o nO ?
    Dim objs As obj
    objs.Amount = lstSubastas(sI).Cant
    objs.ObjIndex = lstSubastas(sI).ObjIndex
    
    'Tiene el inventario lleno ?
    If UserList(UserIndex).Invent.NroItems = MAX_INVENTORY_SLOTS Then
        'Tambien lleno ?
        If UserList(UserIndex).BancoInvent.NroItems = MAX_BANCOINVENTORY_SLOTS Then
            If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex = 0 Then
                Call WriteConsoleMsg(1, UserIndex, "El item de la subasta ha sido arrojado debajo tuyo.", FontTypeNames.FONTTYPE_INFO)
                Call MakeObj(objs, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y)
            End If
        Else 'Le depositamos el item
            'Agregamos el obj que deposita al banco
            Call sUserDejaObj(UserIndex, lstSubastas(sI).ObjIndex, lstSubastas(sI).Cant)
        End If
    Else
        'Bueno tiene lugar en el inventario
        MeterItemEnInventario UserIndex, objs
    End If
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - lstSubastas(sI).fnlOfert
    Call WriteUpdateGold(UserIndex)
    
    'Segundo, le damos el oro al user, o se lo depositamos
    Dim tIndex As Integer, tGld As Long
    tIndex = NameIndex(lstSubastas(sI).nckVndedor)
    If tIndex > 0 Then 'Esta online?
        UserList(tIndex).Stats.GLD = UserList(tIndex).Stats.GLD + lstSubastas(sI).fnlOfert
        Call WriteUpdateGold(tIndex)
    Else
        Add_GLD_Subast lstSubastas(sI).nckVndedor, lstSubastas(sI).fnlOfert
    End If
    
    sBorrar sI

End Sub
Sub sUserDejaObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)
    Dim Slot As Integer
    Dim obji As Integer
    
    If Cantidad < 1 Then Exit Sub
    
    obji = ObjIndex

    Slot = 1
    Do Until UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex = obji And _
        UserList(UserIndex).BancoInvent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
        Slot = Slot + 1
        
        If Slot > MAX_BANCOINVENTORY_SLOTS Then
            Exit Do
        End If
    Loop
    
    'Sino se fija por un slot vacio antes del slot devuelto
    If Slot > MAX_BANCOINVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1
            
            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Call WriteConsoleMsg(1, UserIndex, "No tienes mas espacio en el banco!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Loop
        
        UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems + 1
    End If
    
    If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido
        'Mete el obj en el slot
        If UserList(UserIndex).BancoInvent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
            
            'Menor que MAX_INV_OBJS
            UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex = obji
            UserList(UserIndex).BancoInvent.Object(Slot).Amount = UserList(UserIndex).BancoInvent.Object(Slot).Amount + Cantidad
            
            Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)
        Else
            Call WriteConsoleMsg(1, UserIndex, "El banco no puede cargar tantos objetos.", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
End Sub

Public Sub sLoop()
Dim i As Long
For i = 1 To 100
    If lstSubastas(i).active = True Then
        If Not lstSubastas(i).mnDura = 0 Then lstSubastas(i).mnDura = lstSubastas(i).mnDura - 1
        If lstSubastas(i).mnDura = 0 And lstSubastas(i).hsDura > 0 Then
            lstSubastas(i).hsDura = lstSubastas(i).hsDura - 1
            lstSubastas(i).mnDura = 60
        ElseIf lstSubastas(i).mnDura = 0 Then
            sFinalizar i
        End If
    End If
Next i
End Sub
Sub SaveSubast()
    Dim F As Integer
    
    If FileExist(DatPath & "subastas.ind") Then Kill DatPath & "subastas.ind"
    
    F = FreeFile
    Open DatPath & "subastas.ind" For Binary Access Write As #F
        Put #F, , lstSubastas
    Close #F
    
End Sub
Sub LoadSubast()
    Dim F As Integer
    
    If FileExist(DatPath & "subastas.ind") Then
        F = FreeFile
        Open DatPath & "subastas.ind" For Binary Access Read As #F
            Get #F, , lstSubastas
        Close #F
    End If
End Sub
