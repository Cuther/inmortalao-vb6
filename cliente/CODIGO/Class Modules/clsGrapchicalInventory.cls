VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim InventoryOffset As Long             'Number of lines we scrolled down from topmost
Dim InvSelectedItem As Long             'Currently selected item
Dim TempItem As Integer
Dim TempSlot As Integer

Dim MouseX As Integer
Dim MouseY As Integer
Dim DragItem As Byte
Dim Drag As Byte

Dim WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1
Public Sub SelectGold()
    InvSelectedItem = FLAGORO
End Sub

Public Property Get SelectedItem() As Long
    SelectedItem = InvSelectedItem
End Property

Public Property Get MaxHit(ByVal slot As Byte) As Integer
    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    MaxHit = UserInventory(slot).MaxHit
End Property

Public Property Get MinHit(ByVal slot As Byte) As Integer
    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    MinHit = UserInventory(slot).MinHit
End Property

Public Property Get Def(ByVal slot As Byte) As Integer
    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    Def = UserInventory(slot).Def
End Property

Public Property Get grhindex(ByVal slot As Byte) As Integer
    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    grhindex = UserInventory(slot).grhindex
End Property

Public Property Get Valor(ByVal slot As Byte) As Single
    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    Valor = UserInventory(slot).Valor
End Property

Public Property Get Amount(ByVal slot As Byte) As Long
    If slot = FLAGORO Then
        Amount = UserGLD
    ElseIf slot >= LBound(UserInventory) And slot <= UBound(UserInventory) Then
        Amount = UserInventory(slot).Amount
    End If
End Property

Public Property Get OBJIndex(ByVal slot As Byte) As Integer
    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    OBJIndex = UserInventory(slot).OBJIndex
End Property

Public Property Get OBJType(ByVal slot As Byte) As Integer
    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    OBJType = UserInventory(slot).OBJType
End Property

Public Property Get ItemName(ByVal slot As Byte) As String
    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    ItemName = UserInventory(slot).name
End Property

Public Property Get Equipped(ByVal slot As Byte) As Boolean
    If slot > MAX_INVENTORY_SLOTS Or slot < 1 Then Exit Property
    Equipped = UserInventory(slot).Equipped
End Property

Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Integer, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal eDef As Integer, ByVal eValor As Single, ByVal eName As String, ByVal Puede As Byte)
    If slot < 1 Or slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    With UserInventory(slot)
        .Amount = eAmount
        .Def = eDef
        .Equipped = eEquipped
        .grhindex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .name = eName
        .OBJIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
        .PuedeUsar = Puede
    End With
End Sub

Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
    
    temp_x = X \ 32
    temp_y = Y \ 32
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
    
    'Make sure it's within limits
    If TempItem <= MAX_INVENTORY_SLOTS Then
        'Make sure slot isn't empty
        If UserInventory(TempItem).grhindex Then
            ClickItem = TempItem
        Else
            ClickItem = 0
        End If
    End If
    
    RenderInv = True
End Function

Public Sub DrawInventory()
    Dim i As Byte
    Dim X As Integer
    Dim Y As Integer
    Dim temp_array(3) As Long
    
    If DragItem <> 0 And Drag Then
        If TempSlot <> 0 Then
            X = ((TempSlot - 1) Mod (InventoryWindow.width / 32)) * 32
            Y = ((TempSlot - 1) \ (InventoryWindow.width / 32)) * 32
            Call TileEngine.Draw_Grh_Index(UserInventory(DragItem).grhindex, X, Y, &H64FFFFFF)
        End If
    End If
    
    For i = 1 To UBound(UserInventory)
        If UserInventory(i).grhindex Then
            X = ((i - 1) Mod (InventoryWindow.width / 32)) * 32
            Y = ((i - 1) \ (InventoryWindow.width / 32)) * 32
            
            TileEngine.Engine_Long_To_RGB_List temp_array, -1
            
            If Not UserInventory(i).PuedeUsar = 1 Then
                Call TileEngine.Draw_Grh_Index(UserInventory(i).grhindex, X - 1, Y - 1, -1763311616)
            Else
                Call TileEngine.Draw_Grh_Index(UserInventory(i).grhindex, X - 1, Y - 1, -1)
            End If

            Call TileEngine.Text_Render(str(UserInventory(i).Amount), X - 5, Y + 20, temp_array, 1)
            
            TileEngine.Engine_Long_To_RGB_List temp_array, &HFFFF0000
            If UserInventory(i).Equipped Then
                Call TileEngine.Text_Render("+", X + 20, Y, temp_array, 1)
            End If
        End If
    Next i
    
    If InvSelectedItem > 0 Then
        TileEngine.Draw_Grh_Index 2, (((InvSelectedItem - 1) Mod (InventoryWindow.width / 32)) * 32) - 1, (((InvSelectedItem - 1) \ (InventoryWindow.width / 32)) * 32) - 1, -1
    End If
    
    If DragItem <> 0 And Drag = 1 Then
        Call TileEngine.Draw_Grh_Index(UserInventory(DragItem).grhindex, MouseX - 16, MouseY - 16)
        RenderInv = True
    Else
        RenderInv = False
    End If
    
End Sub

Public Sub Initialize(ByRef InvPic As PictureBox)
    Set InventoryWindow = InvPic

    InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item

End Sub

Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If X < 0 Or Y < 0 Or X > InventoryWindow.width Or Y > InventoryWindow.height Then Exit Sub
        Dim temp_x As Integer
        Dim temp_y As Integer
    
        temp_x = X \ 32
        temp_y = Y \ 32
    
        'Get the currently clickced item
        DragItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
        If Not (DragItem <= 25 And DragItem >= 1) Then
            DragItem = 0
            Exit Sub
        Else
            If UserInventory(DragItem).grhindex = 0 Then
                DragItem = 0
                Exit Sub
            End If
        End If
        
        Drag = 1
        
        RenderInv = True
    End If
End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse move event of the inventory picture box
'***************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    
    temp_x = X \ 32
    temp_y = Y \ 32
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
    TempSlot = TempItem
    
    If TempItem <= 25 And TempItem >= 1 Then
        frmMain.lblInvInfo.Visible = True
        If UserInventory(TempItem).name = "" Then
            frmMain.lblInvInfo.Caption = "Nada"
        Else
            frmMain.lblInvInfo.Caption = UserInventory(TempItem).name '& " || " & "Cantidad: " & UserInventory(TempItem).Amount & vbNewLine & "Def:" & UserInventory(TempItem).Def & " MinHit:" & UserInventory(TempItem).MinHit & " MaxHit:" & UserInventory(TempItem).MaxHit
        End If
    End If
    
    If Button = vbRightButton Then ' And DragItem <> 0 Then
        Drag = 1
    Else
        Drag = 0
    End If
    
    MouseX = CInt(X)
    MouseY = CInt(Y)
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse up event of the inventory picture box
'Check outs which item was clicked
'***************************************************
    'Store previously selected item
    Dim prevSelItem As Long
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.width Or Y > InventoryWindow.height Then _
        Exit Sub
    
    prevSelItem = InvSelectedItem

    'Get the currently clickced item
    InvSelectedItem = ClickItem(CInt(X), CInt(Y))

    If Drag Then
        If Not DragItem = TempItem And Not (TempItem < 1 Or DragItem < 1) Then
            WriteDragAndDrop TempItem, DragItem
            
            Dim i As Integer
    
            For i = LBound(MacroKeys) To UBound(MacroKeys)
                If MacroKeys(i).invslot = DragItem Then
                    MacroKeys(i).invslot = TempItem
                End If
            Next i
            
        End If
        
        DragItem = 0
        Drag = 0
    End If

    RenderInv = True
End Sub

Public Function GetPos(ByVal X As Integer, ByVal Y As Integer) As Byte
    Dim temp_x As Integer
    Dim temp_y As Integer
    
    temp_x = X \ 32
    temp_y = Y \ 32
    
    GetPos = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
End Function

