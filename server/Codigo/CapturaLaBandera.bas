Attribute VB_Name = "CapturaLaBandera"
'Add Marius Captura la bandera
Option Explicit
Public Bandera_estado As Boolean
'Public Bandera_time As Byte
Public Bandera_ronda_impe As Byte
Public Bandera_ronda_repu As Byte

Public Bandera_imperial As Boolean
Public Bandera_republicana As Boolean

Public Bandera_repu_cant As Byte
Public Bandera_impe_cant As Byte


Public Const Bandera_mapa As Integer = 848

'Sums
Public Const Bandera_impe_sum_x As Integer = 14
Public Const Bandera_impe_sum_y As Integer = 64
    
Public Const Bandera_repu_sum_x As Integer = 87
Public Const Bandera_repu_sum_y As Integer = 29

'Bases
Public Const Bandera_impe_base_x1 As Integer = 11
Public Const Bandera_impe_base_y1 As Integer = 13
Public Const Bandera_impe_base_x2 As Integer = 14
Public Const Bandera_impe_base_y2 As Integer = 16

Public Const Bandera_repu_base_x1 As Integer = 87
Public Const Bandera_repu_base_y1 As Integer = 77
Public Const Bandera_repu_base_x2 As Integer = 90
Public Const Bandera_repu_base_y2 As Integer = 80

'Objs Banderas
Public Bandera_impe_obj_num As Integer
Public Bandera_impe_obj_x As Integer
Public Bandera_impe_obj_y As Integer

Public Bandera_repu_obj_num As Integer
Public Bandera_repu_obj_x As Integer
Public Bandera_repu_obj_y As Integer

Public Bandera_vacio_obj_num As Integer

Public Const MAX_BANDO_BANDERA As Byte = 15
Public Bandera_impe(1 To MAX_BANDO_BANDERA) As Integer
Public Bandera_repu(1 To MAX_BANDO_BANDERA) As Integer

Sub Bandera_Sale(ByVal UserIndex As Integer)
On Error Resume Next
    Dim i As Integer
    Dim NuevaPos As WorldPos
    Dim FuturePos As WorldPos

    If (esCiuda(UserIndex) Or esArmada(UserIndex)) Then

        Bandera_impe_cant = Bandera_impe_cant - 1

    ElseIf (esRepu(UserIndex) Or esMili(UserIndex)) Then

        Bandera_repu_cant = Bandera_repu_cant - 1

    End If
    
    
    For i = 1 To MAX_BANDO_BANDERA
        If Bandera_impe(i) = UserIndex Then
            Bandera_impe(i) = 0
            Call BubbleSort(Bandera_impe)
            Exit For
        End If
        
        If Bandera_repu(i) = UserIndex Then
            Bandera_repu(i) = 0
            Call BubbleSort(Bandera_repu)
            Exit For
        End If
    Next
    
    Call Bandera_muere(UserIndex)
    
    FuturePos.map = 49
    FuturePos.x = 50
    FuturePos.Y = 50
    
    UserList(UserIndex).bandera = 0
    

    Call ClosestLegalPos(FuturePos, NuevaPos)

    If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then _
        Call WarpUserChar(UserIndex, NuevaPos.map, NuevaPos.x, NuevaPos.Y, True)

    Call WriteConsoleMsg(1, UserIndex, "Estas fuera del evento Captura la Bandera.", FontTypeNames.FONTTYPE_TALK)

End Sub

Sub Bandera_Inicia()
On Error Resume Next
Dim i As Integer

    Bandera_estado = True
    
    MapInfo(Bandera_mapa).Pk = True
    MapInfo(Bandera_mapa).InviSinEfecto = True
    MapInfo(Bandera_mapa).ResuSinEfecto = 1

    Bandera_impe_cant = 0
    Bandera_repu_cant = 0
    
    Bandera_imperial = True
    Bandera_republicana = True

    'Bandera_time = 0
    Bandera_ronda_impe = 0
    Bandera_ronda_repu = 0

    Bandera_impe_obj_num = 1491
    Bandera_impe_obj_x = 13
    Bandera_impe_obj_y = 15
    
    Bandera_vacio_obj_num = 1494
    
    Bandera_repu_obj_num = 1493
    Bandera_repu_obj_x = 88
    Bandera_repu_obj_y = 78
    
    For i = 1 To MAX_BANDO_BANDERA
        Bandera_impe(i) = 0
        Bandera_repu(i) = 0
    Next
    
    Call Bandera_estandartes(1, True)
    Call Bandera_estandartes(2, True)
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(252, NO_3D_SOUND, NO_3D_SOUND)) 'Cuerno
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> El evento Captura la Bandera empezó, para participar entra desde el boton Eventos de tu Menu.", FontTypeNames.FONTTYPE_SERVER))

End Sub
Sub Bandera_Termina()
On Error Resume Next
Dim i As Integer
Dim NuevaPos As WorldPos
Dim FuturePos As WorldPos
    
    FuturePos.map = 49
    FuturePos.x = 50
    FuturePos.Y = 50
    
    Bandera_estado = False
    MapInfo(Bandera_mapa).Pk = False

    Bandera_impe_cant = 0
    Bandera_repu_cant = 0

    'Bandera_time = 0
    Bandera_ronda_impe = 0
    Bandera_ronda_repu = 0
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Evento> El evento Captura la Bandera terminó.", FontTypeNames.FONTTYPE_SERVER))
    
    For i = 1 To MAX_BANDO_BANDERA
        
        If Bandera_impe(i) <> 0 Then
            Call WriteConsoleMsg(1, Bandera_impe(i), "El evento Captura la Bandera finalizó.", FontTypeNames.FONTTYPE_TALK)
            UserList(Bandera_impe(i)).bandera = 0
            
            Call ClosestLegalPos(FuturePos, NuevaPos)
            If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then _
                Call WarpUserChar(Bandera_impe(i), NuevaPos.map, NuevaPos.x, NuevaPos.Y, True)
            
            Bandera_impe(i) = 0
        End If
        
        If Bandera_repu(i) <> 0 Then
            Call WriteConsoleMsg(1, Bandera_repu(i), "El evento Captura la Bandera finalizó.", FontTypeNames.FONTTYPE_TALK)
            UserList(Bandera_repu(i)).bandera = 0
            
            Call ClosestLegalPos(FuturePos, NuevaPos)
            If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then _
                Call WarpUserChar(Bandera_repu(i), NuevaPos.map, NuevaPos.x, NuevaPos.Y, True)
            
            Bandera_impe(i) = 0
        End If
        
    Next

End Sub

Sub Bandera_Entra(ByVal UserIndex As Integer)
On Error Resume Next

Dim i As Integer
Dim Slot As Integer
Dim NuevaPos As WorldPos
Dim FuturePos As WorldPos

    If Not Bandera_estado Then
        Call WriteConsoleMsg(1, UserIndex, "El evento Campura la Bandera esta cerrado!", FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    ElseIf UserList(UserIndex).Pos.map = Bandera_mapa Then
        Call WriteConsoleMsg(1, UserIndex, "Ya estas adentro del evento Campura la Bandera!", FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    End If

    If (esCiuda(UserIndex) Or esArmada(UserIndex)) Then
        If Bandera_impe_cant > 9 Then
            Call WriteConsoleMsg(1, UserIndex, "No hay mas lugar para Imperiales en Captura la Bandera!", FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        End If
        
        Slot = 0
        For i = 1 To MAX_BANDO_BANDERA
            If Bandera_impe(i) = 0 Then
                Slot = i
                Exit For
            End If
        Next
        
        If Slot <> 0 Then
            Bandera_impe(Slot) = UserIndex
            
            FuturePos.map = Bandera_mapa
            FuturePos.x = Bandera_impe_sum_x
            FuturePos.Y = Bandera_impe_sum_y
    
            Bandera_impe_cant = Bandera_impe_cant + 1
        Else
            Call WriteConsoleMsg(1, UserIndex, "Ya hay demaciados Imperiales.", FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        End If

    ElseIf (esRepu(UserIndex) Or esMili(UserIndex)) Then
        If Bandera_repu_cant > 9 Then
            Call WriteConsoleMsg(1, UserIndex, "No hay mas lugar para Republicanos en Captura la Bandera!", FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        End If
        
        Slot = 0
        For i = 1 To MAX_BANDO_BANDERA
            If Bandera_repu(i) = 0 Then
                Slot = i
                Exit For
            End If
        Next
        
        If Slot <> 0 Then
            Bandera_repu(Slot) = UserIndex
            
            FuturePos.map = Bandera_mapa
            FuturePos.x = Bandera_repu_sum_x
            FuturePos.Y = Bandera_repu_sum_y
    
            Bandera_repu_cant = Bandera_repu_cant + 1
        Else
            Call WriteConsoleMsg(1, UserIndex, "Ya hay demaciados Republicanos.", FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        End If

    Else
        Call WriteConsoleMsg(1, UserIndex, "Este evento es solo para Imperiales y Republicanos.", FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    End If
    
    UserList(UserIndex).bandera = 0
    
    'Buscamos lugar donde transportarlo
    Call ClosestLegalPos(FuturePos, NuevaPos)

    If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then _
        Call WarpUserChar(UserIndex, NuevaPos.map, NuevaPos.x, NuevaPos.Y, True)

    Call WriteConsoleMsg(1, UserIndex, "Estas dentro del evento Captura la Bandera. Para salir usa la Runa o Desconectate! Protege nuestra Bandera con tu vida.", FontTypeNames.FONTTYPE_TALK)

End Sub

Public Sub Bandera_reiniciar()
On Error Resume Next
Dim i As Integer
Dim NuevaPos As WorldPos
Dim FuturePos As WorldPos

    'Bandera_time = 0
    Bandera_imperial = True
    Bandera_republicana = True

    Call Bandera_estandartes(1, True)
    Call Bandera_estandartes(2, True)
    
    If Bandera_ronda_impe >= 3 Then
        Call SendData(SendTarget.ToImpeMap, Bandera_mapa, PrepareMessagePlayWave(105, NO_3D_SOUND, NO_3D_SOUND)) 'Ganaron las rondas
        Call SendData(SendTarget.ToRepuMap, Bandera_mapa, PrepareMessagePlayWave(93, NO_3D_SOUND, NO_3D_SOUND)) 'Perdieron las ronas
        Bandera_ronda_impe = 0
        Bandera_ronda_repu = 0
    ElseIf Bandera_ronda_repu >= 3 Then
        Call SendData(SendTarget.ToRepuMap, Bandera_mapa, PrepareMessagePlayWave(105, NO_3D_SOUND, NO_3D_SOUND)) 'Ganaron las rondas
        Call SendData(SendTarget.ToImpeMap, Bandera_mapa, PrepareMessagePlayWave(93, NO_3D_SOUND, NO_3D_SOUND)) 'Perdieron las ronas
        Bandera_ronda_impe = 0
        Bandera_ronda_repu = 0
    Else
        Call SendData(SendTarget.ToMap, Bandera_mapa, PrepareMessagePlayWave(97, NO_3D_SOUND, NO_3D_SOUND)) 'Reinicia la ronda
    End If
    
    For i = 1 To MAX_BANDO_BANDERA
        If Bandera_impe(i) <> 0 Then
            FuturePos.map = Bandera_mapa
            FuturePos.x = Bandera_impe_sum_x
            FuturePos.Y = Bandera_impe_sum_y
            
            UserList(Bandera_impe(i)).bandera = 0
            
            Call ClosestLegalPos(FuturePos, NuevaPos)
            If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then _
                Call WarpUserChar(Bandera_impe(i), NuevaPos.map, NuevaPos.x, NuevaPos.Y, False)
        End If
        
        If Bandera_repu(i) <> 0 Then
            FuturePos.map = Bandera_mapa
            FuturePos.x = Bandera_repu_sum_x
            FuturePos.Y = Bandera_repu_sum_y
            
            UserList(Bandera_repu(i)).bandera = 0
            
            Call ClosestLegalPos(FuturePos, NuevaPos)
            If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then _
                Call WarpUserChar(Bandera_repu(i), NuevaPos.map, NuevaPos.x, NuevaPos.Y, False)
        End If
    Next
    
    Call SendData(SendTarget.ToMap, Bandera_mapa, PrepareMessageConsoleMsg(1, "Comienzamos la ronda numero " & (Bandera_ronda_impe + Bandera_ronda_repu) & "/5", FontTypeNames.FONTTYPE_GUILD))
    
End Sub

Public Sub Bandera_verifica_banderas(UserIndex As Integer)
On Error Resume Next
    With UserList(UserIndex)
    
        If (esCiuda(UserIndex) Or esArmada(UserIndex)) Then
    
            'Ya tiene la bandera
            If UserList(UserIndex).bandera = 2 Then
                'Si esta dentro de la base propia
                If .Pos.x >= Bandera_impe_base_x1 And .Pos.x <= Bandera_impe_base_x2 And _
                    .Pos.Y >= Bandera_impe_base_y1 And .Pos.Y <= Bandera_impe_base_y2 _
                Then
                    'Adentro del impe
                    Call SendData(SendTarget.ToImpeMap, Bandera_mapa, PrepareMessageConsoleMsg(1, "Ganamos esta ronda!", FontTypeNames.FONTTYPE_GUILD))
                    Call SendData(SendTarget.ToRepuMap, Bandera_mapa, PrepareMessageConsoleMsg(1, "Perdimos esta ronda!", FontTypeNames.FONTTYPE_GUILD))
                    UserList(UserIndex).bandera = 0
                    Bandera_ronda_impe = Bandera_ronda_impe + 1
                   
                    Call Bandera_reiniciar
                    
                End If
    
            'Esta sobre la bandera del contrario
            ElseIf Bandera_republicana And MapData(.Pos.map, .Pos.x, .Pos.Y).ObjInfo.ObjIndex = Bandera_repu_obj_num Then
                'Tiene la bandera
                Bandera_republicana = False
                UserList(UserIndex).bandera = 2
                Call Bandera_estandartes(2, False)
                Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, False)

                Call SendData(SendTarget.ToImpeMap, Bandera_mapa, PrepareMessageConsoleMsg(1, "Capturamos la Bandera enemiga!", FontTypeNames.FONTTYPE_GUILD))
                Call SendData(SendTarget.ToRepuMap, Bandera_mapa, PrepareMessageConsoleMsg(1, "Capturaron nuestra bandera!", FontTypeNames.FONTTYPE_GUILD))
                
                Call SendData(SendTarget.ToImpeMap, Bandera_mapa, PrepareMessagePlayWave(87, NO_3D_SOUND, NO_3D_SOUND))  'Consiguió la bandera
                Call SendData(SendTarget.ToRepuMap, Bandera_mapa, PrepareMessagePlayWave(139, NO_3D_SOUND, NO_3D_SOUND)) 'Robaron su Bandera
            End If
    
    
    
        ElseIf (esRepu(UserIndex) Or esMili(UserIndex)) Then
    
            'Ya tiene la bandera
            If UserList(UserIndex).bandera = 1 Then
                'Si esta dentro de la base propia
                If .Pos.x >= Bandera_repu_base_x1 And .Pos.x <= Bandera_repu_base_x2 And _
                    .Pos.Y >= Bandera_repu_base_y1 And .Pos.Y <= Bandera_repu_base_y2 _
                Then
                    'Adentro del impe
                    Call SendData(SendTarget.ToRepuMap, Bandera_mapa, PrepareMessageConsoleMsg(1, "Ganamos esta ronda!", FontTypeNames.FONTTYPE_GUILD))
                    Call SendData(SendTarget.ToImpeMap, Bandera_mapa, PrepareMessageConsoleMsg(1, "Perdimos esta ronda!", FontTypeNames.FONTTYPE_GUILD))
                    UserList(UserIndex).bandera = 0
                    Bandera_ronda_repu = Bandera_ronda_repu + 1
                    
                    Call Bandera_reiniciar
                    
                End If
    
            'Esta sobre la bandera del contrario
            ElseIf Bandera_imperial And MapData(.Pos.map, .Pos.x, .Pos.Y).ObjInfo.ObjIndex = Bandera_impe_obj_num Then
                'Tiene la bandera
                Bandera_imperial = False
                UserList(UserIndex).bandera = 1
                Call Bandera_estandartes(1, False)
                Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, False)
                
                Call SendData(SendTarget.ToRepuMap, Bandera_mapa, PrepareMessageConsoleMsg(1, "Capturamos la Bandera enemiga!", FontTypeNames.FONTTYPE_GUILD))
                Call SendData(SendTarget.ToImpeMap, Bandera_mapa, PrepareMessageConsoleMsg(1, "Capturaron nuestra bandera!", FontTypeNames.FONTTYPE_GUILD))
                
                Call SendData(SendTarget.ToRepuMap, Bandera_mapa, PrepareMessagePlayWave(87, NO_3D_SOUND, NO_3D_SOUND)) 'Consiguió la bandera
                Call SendData(SendTarget.ToImpeMap, Bandera_mapa, PrepareMessagePlayWave(139, NO_3D_SOUND, NO_3D_SOUND)) 'Robaron su Bandera
            End If
        
        End If
    
    End With

End Sub

Public Sub Bandera_muere(ByVal UserIndex As Integer)
On Error Resume Next
Dim NuevaPos As WorldPos
Dim FuturePos As WorldPos
    
    If UserList(UserIndex).bandera = 1 Then
        Bandera_imperial = True
        Call Bandera_estandartes(1, True)
        Call SendData(SendTarget.ToImpeMap, Bandera_mapa, PrepareMessageConsoleMsg(1, "Los imperiales recuperaron su bandera!", FontTypeNames.FONTTYPE_GUILD))
    ElseIf UserList(UserIndex).bandera = 2 Then
        Bandera_republicana = True
        Call Bandera_estandartes(2, True)
        Call SendData(SendTarget.ToImpeMap, Bandera_mapa, PrepareMessageConsoleMsg(1, "Los republicanos recuperaron su bandera!", FontTypeNames.FONTTYPE_GUILD))
    End If
    
    UserList(UserIndex).bandera = 0
    
    If (esCiuda(UserIndex) Or esArmada(UserIndex)) Then
        FuturePos.map = Bandera_mapa
        FuturePos.x = Bandera_impe_sum_x
        FuturePos.Y = Bandera_impe_sum_y
        
        Call ClosestLegalPos(FuturePos, NuevaPos)
        If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then _
            Call WarpUserChar(UserIndex, NuevaPos.map, NuevaPos.x, NuevaPos.Y, False)
    End If
    
    If (esRepu(UserIndex) Or esMili(UserIndex)) Then
        FuturePos.map = Bandera_mapa
        FuturePos.x = Bandera_repu_sum_x
        FuturePos.Y = Bandera_repu_sum_y
        
        Call ClosestLegalPos(FuturePos, NuevaPos)
        If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then _
            Call WarpUserChar(UserIndex, NuevaPos.map, NuevaPos.x, NuevaPos.Y, False)
    End If
    
    
    Call RevivirUsuario(UserIndex)
    UserList(UserIndex).Counters.IntervaloRevive = UserList(UserIndex).Counters.IntervaloRevive + 5000
    
End Sub

Public Sub Bandera_estandartes(ByVal faccion As Byte, ByVal bandera As Boolean)
On Error Resume Next
Dim obj As obj
    obj.Amount = 1
    
    If faccion = 1 Then 'impe
        If bandera Then
            Call EraseObj(10000, Bandera_mapa, Bandera_impe_obj_x, Bandera_impe_obj_y)
            
            obj.ObjIndex = Bandera_impe_obj_num
            Call MakeObj(obj, Bandera_mapa, Bandera_impe_obj_x, Bandera_impe_obj_y)
        Else
            Call EraseObj(10000, Bandera_mapa, Bandera_impe_obj_x, Bandera_impe_obj_y)
            
            obj.ObjIndex = Bandera_vacio_obj_num
            Call MakeObj(obj, Bandera_mapa, Bandera_impe_obj_x, Bandera_impe_obj_y)
        End If
    ElseIf faccion = 2 Then 'repu
        If bandera Then
            Call EraseObj(10000, Bandera_mapa, Bandera_repu_obj_x, Bandera_repu_obj_y)
            
            obj.ObjIndex = Bandera_repu_obj_num
            Call MakeObj(obj, Bandera_mapa, Bandera_repu_obj_x, Bandera_repu_obj_y)
        Else
            Call EraseObj(10000, Bandera_mapa, Bandera_repu_obj_x, Bandera_repu_obj_y)
            
            obj.ObjIndex = Bandera_vacio_obj_num
            Call MakeObj(obj, Bandera_mapa, Bandera_repu_obj_x, Bandera_repu_obj_y)
        End If
    End If
End Sub
'\Add

