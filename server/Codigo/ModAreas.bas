Attribute VB_Name = "ModAreas"

Option Explicit
Public Type AreaInfo
    AreaPerteneceX As Integer
    AreaPerteneceY As Integer
    
    AreaReciveX As Integer
    AreaReciveY As Integer
    
    AreaID As Long
End Type

Public Type ConnGroup
    CountEntrys As Long
    OptValue As Long
    UserEntrys() As Long
End Type

Public Const USER_NUEVO As Byte = 255

'Cuidado:
' ¡¡¡LAS AREAS ESTÁN HARDCODEADAS!!!
Private CurDay As Byte
Private CurHour As Byte

Private AreasInfo(1 To 100, 1 To 100) As Long

Private AreasRecive(12) As Long

Public ConnGroups() As ConnGroup

Public Sub InitAreas()
    Dim loopC As Long
    Dim LoopX As Long

    For loopC = 0 To 11
        AreasRecive(loopC) = (2 ^ loopC) Or IIf(loopC <> 0, 2 ^ (loopC - 1), 0) Or IIf(loopC <> 11, 2 ^ (loopC + 1), 0)
    Next loopC
    
    For loopC = 1 To 100
        For LoopX = 1 To 100
            AreasInfo(loopC, LoopX) = (loopC \ 4 + 1) * (LoopX \ 4 + 1)
        Next LoopX
    Next loopC

    CurDay = IIf(Weekday(Date) > 6, 1, 2)
    CurHour = Fix(Hour(time) \ 3)
    
    ReDim ConnGroups(1 To NumMaps) As ConnGroup
    
    For loopC = 1 To NumMaps
        ConnGroups(loopC).OptValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & loopC, CurDay & "-" & CurHour))
        
        If ConnGroups(loopC).OptValue = 0 Then ConnGroups(loopC).OptValue = 1
        ReDim ConnGroups(loopC).UserEntrys(1 To ConnGroups(loopC).OptValue) As Long
    Next loopC
End Sub

Public Sub AreasOptimizacion()
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'Es la función de autooptimizacion.... la idea es no mandar redimensionando arrays grandes todo el tiempo
'**************************************************************
    Dim loopC As Long
    Dim tCurDay As Byte
    Dim tCurHour As Byte
    Dim EntryValue As Long
    
    If (CurDay <> IIf(Weekday(Date) > 6, 1, 2)) Or (CurHour <> Fix(Hour(time) \ 3)) Then
        
        tCurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
        tCurHour = Fix(Hour(time) \ 3) 'A ke parte de la hora pertenece
        
        For loopC = 1 To NumMaps
            EntryValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & loopC, CurDay & "-" & CurHour))
            Call WriteVar(DatPath & "AreasStats.dat", "Mapa" & loopC, CurDay & "-" & CurHour, CInt((EntryValue + ConnGroups(loopC).OptValue) \ 2))
            
            ConnGroups(loopC).OptValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & loopC, tCurDay & "-" & tCurHour))
            If ConnGroups(loopC).OptValue = 0 Then ConnGroups(loopC).OptValue = 1
            If ConnGroups(loopC).OptValue >= MapInfo(loopC).NumUsers Then ReDim Preserve ConnGroups(loopC).UserEntrys(1 To ConnGroups(loopC).OptValue) As Long
        Next loopC
        
        CurDay = tCurDay
        CurHour = tCurHour
    End If
End Sub

Public Sub CheckUpdateNeededUser(ByVal UserIndex As Integer, ByVal Head As Byte)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'Es la función clave del sistema de areas... Es llamada al mover un user
'**************************************************************
    If UserList(UserIndex).AreasInfo.AreaID = AreasInfo(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y) Then Exit Sub
    
    Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, x As Long, Y As Long
    Dim TempInt As Long, map As Long

    With UserList(UserIndex)
        If Head = eHeading.NORTH Then
            MinY = .Pos.Y - 11
            MaxY = .Pos.Y - 5
            
            MinX = .Pos.x - 9
            MaxX = .Pos.x + 9
        ElseIf Head = eHeading.SOUTH Then
            MinY = .Pos.Y + 5
            MaxY = .Pos.Y + 11
            
            MinX = .Pos.x - 9
            MaxX = .Pos.x + 9
        ElseIf Head = eHeading.WEST Then
            MinY = .Pos.Y - 6
            MaxY = .Pos.Y + 6
            
            MinX = .Pos.x - 13
            MaxX = .Pos.x - 7
        ElseIf Head = eHeading.EAST Then
            MinY = .Pos.Y - 6
            MaxY = .Pos.Y + 6
            
            MinX = .Pos.x + 7
            MaxX = .Pos.x + 13
        ElseIf Head = USER_NUEVO Then
            MinY = .Pos.Y - 10
            MaxY = .Pos.Y + 10
            
            MinX = .Pos.x - 12
            MaxX = .Pos.x + 12
        End If
        
        If MinY < 1 Then MinY = 1
        If MinX < 1 Then MinX = 1
        If MaxY > 100 Then MaxY = 100
        If MaxX > 100 Then MaxX = 100
        
        map = UserList(UserIndex).Pos.map
        
        'Esto es para ke el cliente elimine lo "fuera de area..."
        Call WriteAreaChanged(UserIndex)
        
        'Actualizamos!!!
        For x = MinX To MaxX
            For Y = MinY To MaxY
                
                '<<< User >>>
                If MapData(map, x, Y).UserIndex Then
                    
                    TempInt = MapData(map, x, Y).UserIndex
                    
                    If UserIndex <> TempInt Then
                        Call MakeUserChar(False, UserIndex, TempInt, map, x, Y)
                        Call MakeUserChar(False, TempInt, UserIndex, .Pos.map, .Pos.x, .Pos.Y)
                    
                      '  If UserList(UserIndex).flags.Meditando Then
                      '      WriteCreateCharParticle TempInt, UserList(UserIndex).Char.CharIndex, ParticleToLevel(UserIndex), -1
                      '  End If
                      '  If UserList(TempInt).flags.Meditando Then
                      '      WriteCreateCharParticle UserIndex, UserList(TempInt).Char.CharIndex, ParticleToLevel(TempInt), -1
                      '  End If

                        'Si el user estaba invisible le avisamos al nuevo cliente de eso
                        If UserList(TempInt).flags.Invisible Or UserList(TempInt).flags.Oculto Then
                            Call WriteSetInvisible(UserIndex, UserList(TempInt).Char.CharIndex, True)
                        End If
                        If UserList(UserIndex).flags.Invisible Or UserList(UserIndex).flags.Oculto Then
                            Call WriteSetInvisible(TempInt, UserList(UserIndex).Char.CharIndex, True)
                        End If
                        
                        Call FlushBuffer(TempInt)
                    
                    ElseIf Head = USER_NUEVO Then
                        Call MakeUserChar(False, UserIndex, UserIndex, map, x, Y)
                    End If
                End If
                
                '<<< Npc >>>
                If MapData(map, x, Y).NpcIndex Then
                    Call MakeNPCChar(False, UserIndex, MapData(map, x, Y).NpcIndex, map, x, Y)
                 End If
                 
                '<<< Item >>>
                If MapData(map, x, Y).ObjInfo.ObjIndex Then
                    TempInt = MapData(map, x, Y).ObjInfo.ObjIndex
                    If Not EsObjetoFijo(ObjData(TempInt).OBJType) And MapData(map, x, Y).ObjEsFijo = 0 Then
                        Call WriteObjectCreate(UserIndex, x, Y, TempInt, MapData(map, x, Y).ObjInfo.Amount)
                        
                        If ObjData(TempInt).OBJType = eOBJType.otPuertas Then
                            Call Bloquear(False, UserIndex, x, Y, MapData(map, x, Y).Blocked)
                            Call Bloquear(False, UserIndex, x - 1, Y, MapData(map, x - 1, Y).Blocked)
                        End If
                    End If
                End If
            
            Next Y
        Next x
        
        'Precalculados :P
        TempInt = .Pos.x \ 9
        .AreasInfo.AreaReciveX = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceX = 2 ^ TempInt
        
        TempInt = .Pos.Y \ 9
        .AreasInfo.AreaReciveY = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        
        If .Pos.x <> 0 And .Pos.Y <> 0 Then
            .AreasInfo.AreaID = AreasInfo(.Pos.x, .Pos.Y)
        End If
    End With
End Sub

Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal Head As Byte)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
' Se llama cuando se mueve un Npc
'**************************************************************
    If Npclist(NpcIndex).AreasInfo.AreaID = AreasInfo(Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y) Then Exit Sub
    
    Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, x As Long, Y As Long
    Dim TempInt As Long
    
    With Npclist(NpcIndex)
        MinY = .Pos.Y - 10
        MaxY = .Pos.Y + 10
            
        MinX = .Pos.x - 12
        MaxX = .Pos.x + 12
            
        If MinY < 1 Then MinY = 1
        If MinX < 1 Then MinX = 1
        If MaxY > 100 Then MaxY = 100
        If MaxX > 100 Then MaxX = 100

        'Actualizamos!!!
        If MapInfo(.Pos.map).NumUsers <> 0 Then
            For x = MinX To MaxX
                For Y = MinY To MaxY
                    If MapData(.Pos.map, x, Y).UserIndex Then _
                        Call MakeNPCChar(False, MapData(.Pos.map, x, Y).UserIndex, NpcIndex, .Pos.map, .Pos.x, .Pos.Y)
                Next Y
            Next x
        End If
        
        'Precalculados :P
        TempInt = .Pos.x \ 9
        .AreasInfo.AreaReciveX = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceX = 2 ^ TempInt
            
        TempInt = .Pos.Y \ 9
        .AreasInfo.AreaReciveY = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        
        .AreasInfo.AreaID = AreasInfo(.Pos.x, .Pos.Y)
    End With
End Sub

Public Sub QuitarUser(ByVal UserIndex As Integer, ByVal map As Integer)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim TempVal As Long
    Dim loopC As Long
    
    'Search for the user
    For loopC = 1 To ConnGroups(map).CountEntrys
        If ConnGroups(map).UserEntrys(loopC) = UserIndex Then Exit For
    Next loopC
    
    'Char not found
    If loopC > ConnGroups(map).CountEntrys Then Exit Sub
    
    'Remove from old map
    ConnGroups(map).CountEntrys = ConnGroups(map).CountEntrys - 1
    TempVal = ConnGroups(map).CountEntrys
    
    'Move list back
    For loopC = loopC To TempVal
        ConnGroups(map).UserEntrys(loopC) = ConnGroups(map).UserEntrys(loopC + 1)
    Next loopC
    
    If TempVal > ConnGroups(map).OptValue Then 'Nescesito Redim?
        ReDim Preserve ConnGroups(map).UserEntrys(1 To TempVal) As Long
    End If
End Sub

Public Sub AgregarUser(ByVal UserIndex As Integer, ByVal map As Integer)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: 04/01/2007
'Modified by Juan Martín Sotuyo Dodero (Maraxus)
'   - Now the method checks for repetead users instead of trusting parameters.
'   - If the character is new to the map, update it
'**************************************************************
    Dim TempVal As Long
    Dim EsNuevo As Boolean
    Dim i As Long
    
    If Not MapaValido(map) Then Exit Sub
    
    EsNuevo = True
    
    'Prevent adding repeated users
    For i = 1 To ConnGroups(map).CountEntrys
        If ConnGroups(map).UserEntrys(i) = UserIndex Then
            EsNuevo = False
            Exit For
        End If
    Next i
    
    If EsNuevo Then
        'Update map and connection groups data
        ConnGroups(map).CountEntrys = ConnGroups(map).CountEntrys + 1
        TempVal = ConnGroups(map).CountEntrys
        
        If TempVal > ConnGroups(map).OptValue Then 'Nescesito Redim
            ReDim Preserve ConnGroups(map).UserEntrys(1 To TempVal) As Long
        End If
        
        ConnGroups(map).UserEntrys(TempVal) = UserIndex
    End If
    
    'Update user
    UserList(UserIndex).AreasInfo.AreaID = 0
    
    UserList(UserIndex).AreasInfo.AreaPerteneceX = 0
    UserList(UserIndex).AreasInfo.AreaPerteneceY = 0
    UserList(UserIndex).AreasInfo.AreaReciveX = 0
    UserList(UserIndex).AreasInfo.AreaReciveY = 0
    
    Call CheckUpdateNeededUser(UserIndex, USER_NUEVO)
End Sub

Public Sub AgregarNpc(ByVal NpcIndex As Integer)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Npclist(NpcIndex).AreasInfo.AreaID = 0
    
    Npclist(NpcIndex).AreasInfo.AreaPerteneceX = 0
    Npclist(NpcIndex).AreasInfo.AreaPerteneceY = 0
    Npclist(NpcIndex).AreasInfo.AreaReciveX = 0
    Npclist(NpcIndex).AreasInfo.AreaReciveY = 0
    
    Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)
End Sub


