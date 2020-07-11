Attribute VB_Name = "TileEngine"
Option Explicit

'***************************
'Sinuhe - Map format .CSM
'***************************
Private Type tMapHeader
    NumeroBloqueados As Long
    NumeroLayers(2 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long
End Type

Private Type tDatosBloqueados
    X As Integer
    Y As Integer
End Type

Private Type tDatosGrh
    X As Integer
    Y As Integer
    grhindex As Long
End Type

Private Type tDatosTrigger
    X As Integer
    Y As Integer
    Trigger As Integer
End Type

Private Type tDatosLuces
    X As Integer
    Y As Integer
    color As Long
    rango As Byte
End Type

Private Type tDatosParticulas
    X As Integer
    Y As Integer
    Particula As Long
End Type

Private Type tDatosNPC
    X As Integer
    Y As Integer
    NPCIndex As Integer
End Type

Private Type tDatosObjs
    X As Integer
    Y As Integer
    OBJIndex As Integer
    ObjAmmount As Integer
End Type

Private Type tDatosTE
    X As Integer
    Y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer
End Type

Private Type tMapSize
    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer
End Type

Private Type tMapDat
    map_name As String * 64
    battle_mode As Byte
    backup_mode As Byte
    restrict_mode As String * 4
    music_number As String * 16
    zone As String * 16
    terrain As String * 16
    ambient As String * 16
    base_light As Long
    letter_grh As Long
    extra1 As Long
    extra2 As Long
    extra3 As String * 32
End Type

Private MapSize As tMapSize
Private MapDat As tMapDat

'DirectX 8
Public SurfaceDB As clsTexManager

Private Type D3D8Textures
    Texture As Direct3DTexture8
    texwidth As Long
    texheight As Long
End Type

Private dX As DirectX8
Private D3D As Direct3D8
Private D3DDevice As Direct3DDevice8
Private D3DX As D3DX8

Private Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    color As Long
    tu As Single
    tv As Single
End Type

Private Const PI As Single = 3.14159265358979

Public bRunning As Boolean
Public hola As New clsAudio
Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE

Private timerElapsedTime As Single
Private timerTicksPerFrame As Double

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer


Private AlphaY As Integer 'Techos

'Textos
Private Type tFont
    font_size As Integer
    ascii_code(0 To 255) As Integer 'indice de cada letra
End Type

Private font_types(1 To 2) As tFont
'Textos

'Luces

Private Type Light
    active As Boolean 'Do we ignore this light?
    id As Long
    map_x As Integer 'Coordinates
    map_y As Integer
    color As Long 'Start colour
    range As Byte
End Type

Dim light_list() As Light
Dim light_count As Long
Dim light_last As Long
'Luces

'******Particulas******
Private Type Particle
    friction As Single
    X As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    Grh As Grh
    alive_counter As Long
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

Private Type particle_group
    active As Boolean
    id As Long
    map_x As Integer
    map_y As Integer
    char_index As Long

    frame_counter As Single
    frame_speed As Single
    
    stream_type As Byte

    particle_stream() As Particle
    particle_count As Long
    
    grh_index_list() As Long
    grh_index_count As Long
    
    alpha_blend As Boolean
    
    alive_counter As Long
    never_die As Boolean
    
    live As Long
    liv1 As Integer
    liveend As Long
    
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    
    'Added by Juan Martín Sotuyo Dodero
    speed As Single
    life_counter As Long
    
    'Added by David Justus
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type
'Particle system
Dim particle_group_list() As particle_group
Dim particle_group_count As Long
Dim particle_group_last As Long
'******Particulas******

'***************************************************************************************************
'Meteo System
'Author: Leandro Mendoza (Mannakia)
'Last Modify Date: 19/09/2010
Private meteo_state As Integer

Private meteo_hour As Integer

Private meteo_color As D3DCOLORVALUE

Private m_Color_Dia As D3DCOLORVALUE
Private m_Color_Noche As D3DCOLORVALUE
Private m_Color_Tarde As D3DCOLORVALUE
Private m_Color_Manana As D3DCOLORVALUE

Public m_Afecta As Boolean

Private ambientLight As D3DCOLORVALUE
Private AmbientColor As Long
'***************************************************************************************************

Public scroll_on As Boolean
Dim scroll_pixels_per_frame As Single
Dim AddPosY As Integer
Dim AddPosX As Integer
Dim ScrollX As Single
Dim ScrollY As Single
 
Private map_letter_grh As Grh
Private map_letter_grh_next As Long
Private map_letter_a As Single
Private map_letter_fadestatus As Byte

Dim MinLimiteX As Integer
Dim MaxLimiteX As Integer
Dim MinLimiteY As Integer
Dim MaxLimiteY As Integer

Public under_stair As Boolean

Private ceg_vertex(0 To 3) As TLVERTEX

Private Function Engine_GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    Engine_GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function
Public Function Engine_Set_Resolution()
'**************************************************************
'Author: Leandro Mendoza (Mannakia)
'Last Modify Date: 30/10/2010
'
'**************************************************************

    Dim DevMode As DevMode
    
    If Window <> 0 Then Exit Function

    EnumDisplaySettings 0&, 0&, DevMode
    
    OldX = Screen.width / Screen.TwipsPerPixelX
    OldY = Screen.height / Screen.TwipsPerPixelY
    
    OldBit = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
    OldBit = GetDeviceCaps(OldBit, 12)
    
    With DevMode
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
        .dmPelsWidth = 800 'Ancho de pantalla
        .dmPelsHeight = 600 'Alto de pantalla
        .dmBitsPerPel = BitPixel 'Cantidad de bits por pixeles
    End With
    
    mueve = 0
    
    ChangeDisplaySettings DevMode, CDS_TEST
    
    ChangeResolution = True
End Function
Public Function Engine_Reset_Resolution()
'**************************************************************
'Author: Leandro Mendoza (Mannakia)
'Last Modify Date: 30/10/2010
'
'**************************************************************
    If ChangeResolution = False Then Exit Function
    Dim DevMode As DevMode
    
    EnumDisplaySettings 0&, 0&, DevMode
    
    With DevMode
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
        .dmPelsWidth = OldX 'Ancho de pantalla
        .dmPelsHeight = OldY 'Alto de pantalla
        .dmBitsPerPel = OldBit 'cantidad de bits por pixeles
    End With
    
    ChangeDisplaySettings DevMode, CDS_TEST
    
End Function
Public Function Engine_Init() As Boolean
    If Not (Engine_Device_Started) Then
        Engine_Init = False
        Exit Function
    End If
    
    Set SurfaceDB = New clsTexManager
    
    Call SurfaceDB.Init(D3DX, D3DDevice)

    engineBaseSpeed = 0.029
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

    UserPos.X = 50
    UserPos.Y = 50
    
    AmbientColor = -1
    
    HalfWindowTileHeight = 6
    HalfWindowTileWidth = 8
    UserMap = 1
    
    LoadGrhData
    Text_Font_Initialize
    Meteo_Init_Time
    Engine_Load_Bodies
    Engine_Load_Heads
    Engine_Load_Helmet
    Engine_Load_Fxs
    Engine_Load_Weapon
    Engine_Load_Shields
    CargarParticulas

    MinXBorder = XMinMapSize + (frmMain.MainViewPic.ScaleWidth / 64)
    MaxXBorder = XMaxMapSize - (frmMain.MainViewPic.ScaleWidth / 64)
    MinYBorder = YMinMapSize + (frmMain.MainViewPic.ScaleHeight / 64)
    MaxYBorder = YMaxMapSize - (frmMain.MainViewPic.ScaleHeight / 64)
    
    ceg_vertex(0) = Geometry_Create_TLVertex(0, 416, 0, 1, &HB0000000, 1, 0, 0)
    ceg_vertex(1) = Geometry_Create_TLVertex(0, 0, 0, 1, &HB0000000, 1, 0, 0)
    ceg_vertex(2) = Geometry_Create_TLVertex(544, 416, 0, 1, &HB0000000, 1, 0, 0)
    ceg_vertex(3) = Geometry_Create_TLVertex(544, 0, 0, 1, &HB0000000, 1, 0, 0)
    
    scroll_pixels_per_frame = 5.2
    
    Default_RGB(0) = -1
    Default_RGB(1) = -1
    Default_RGB(2) = -1
    Default_RGB(3) = -1
    
    Engine_Init = True
End Function
Public Function Engine_Device_Started() As Boolean
On Error Resume Next

    Dim DispMode As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    Dim DevType As CONST_D3DDEVTYPE, DevFlags As CONST_D3DCREATEFLAGS
    Dim i As Long, ii As Long
    
    Set dX = New DirectX8
    Set D3D = dX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    Engine_Set_Resolution
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode

    With D3DWindow
        .Windowed = True
            
        If Sinc = 1 Then
            .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        Else
            .SwapEffect = D3DSWAPEFFECT_COPY
        End If
        
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = frmMain.MainViewPic.ScaleWidth
        .BackBufferHeight = frmMain.MainViewPic.ScaleHeight
        'The HWND default
        .hDeviceWindow = frmMain.MainViewPic.hwnd
        
        .AutoDepthStencilFormat = D3DFMT_D16
        .EnableAutoDepthStencil = 1
        
    End With
    
    DevType = D3DDEVTYPE_HAL

    err.Clear
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, DevType, frmMain.MainViewPic.hwnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DWindow)

    If err.Number <> 0 Or D3DDevice Is Nothing Then
        err.Clear
        Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, DevType, frmMain.MainViewPic.hwnd, D3DCREATE_MIXED_VERTEXPROCESSING, D3DWindow)
    End If
    
    If err.Number <> 0 Or D3DDevice Is Nothing Then
        err.Clear
        Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, DevType, frmMain.MainViewPic.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
    End If
                                                                                              
    If err.Number <> 0 And D3DDevice Is Nothing Then
        MsgBox "Hubo un error al intentar crear Device. Por favor contacte al administrador del juego y notifique el error:" & err.Description & " Number:" & err.Number, , "InmortalAO Engine"
        Engine_Device_Started = False
        Exit Function
    End If
    
    With D3DDevice
        'Set the TLVERTEX
        .SetVertexShader FVF
    
        .SetRenderState D3DRS_LIGHTING, False
    
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    End With

    Engine_Device_Started = True

End Function

Public Sub Engine_End()
    Erase MapData
    Erase charlist
    Erase particle_group_list
    Erase font_types
    Erase light_list
        
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set dX = Nothing
    
    Engine_Reset_Resolution
End Sub

Public Sub Draw_Grh_Index(ByVal grh_index As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal colour As Long = -1)
    If grh_index <= 0 Then Exit Sub
    Dim rgb_list(3) As Long
    
    rgb_list(0) = colour
    rgb_list(1) = colour
    rgb_list(2) = colour
    rgb_list(3) = colour
    
    Device_Box_Textured_Render grh_index, _
        X, Y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY
End Sub


Private Sub Draw_Grh(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, ByRef color() As Long, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single, Optional ByVal Shadow As Byte = 0)

    Dim CurrentGrhIndex As Integer
    
    If Grh.grhindex = 0 Then Exit Sub
    If GrhData(Grh.grhindex).NumFrames = 0 Then Exit Sub
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.speed)
            If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
                Grh.FrameCounter = 1 + Grh.FrameCounter Mod GrhData(Grh.grhindex).NumFrames
                If Grh.Loops <> -1 Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)

    'Center Grh over X,Y pos
    If Center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
        End If

        If GrhData(Grh.grhindex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
        End If
    End If

    Device_Box_Textured_Render CurrentGrhIndex, _
        X, Y, _
        GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
        color, _
        GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
        Alpha, angle, Shadow

End Sub
Public Sub Engine_Render()

On Error Resume Next

If char_current = 0 Then Exit Sub

D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
D3DDevice.BeginScene
    
Meteo_Render
Map_Render
Sound_Render

D3DDevice.EndScene
D3DDevice.Present ByVal 0&, ByVal 0&, 0, ByVal 0&

timerElapsedTime = Engine_GetElapsedTime()
timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

Dim maxInteger As Integer
maxInteger = 32000

'If FramesPerSecCounter >= maxInteger Then
'    FramesPerSecCounter = 0
'Else
    FramesPerSecCounter = FramesPerSecCounter + 1
'End If

Exit Sub

'error:
 '   LogError "Error en Engine Render: " & err.Description

End Sub
Public Function Sound_Render()
    Sound_Fogata_Fx
End Function
Public Function Engine_Scroll_Pixels(ByVal Valor As Single)
    scroll_pixels_per_frame = Valor
End Function
Sub Map_FX_Create(ByVal X As Integer, ByVal Y As Integer, ByVal FX As Integer, ByVal Loops As Integer)
'**************************************************************
'Author: Leandro Mendoza (Mannakia)
'Last Modify Date: 20/11/2010
'Set ObjGrh in array mapdata
'**************************************************************

If X = 0 Or Y = 0 Then Exit Sub

MapData(X, Y).FX = FX

Grh_Load MapData(X, Y).fXGrh, FxData(FX).Animacion
MapData(X, Y).fXGrh.grhindex = FxData(FX).Animacion
MapData(X, Y).fXGrh.Loops = Loops - 1

End Sub

Sub Map_Obj_Create(ByVal X As Integer, ByVal Y As Integer, ByVal grhindex As Integer, ByVal OBJIndex As Integer, ByVal tipe As Byte, ByVal Amount As Integer, Optional ByVal fijo As Byte = 0)
'**************************************************************
'Author: Leandro Mendoza (Mannakia)
'Last Modify Date: 13/11/2010
'Set ObjGrh in array mapdata
'**************************************************************

If X = 0 Or Y = 0 Then Exit Sub

MapData(X, Y).ObjGrh.grhindex = grhindex
MapData(X, Y).OBJInfo.OBJIndex = OBJIndex
MapData(X, Y).OBJInfo.EsFijo = fijo
MapData(X, Y).OBJInfo.tipe = tipe
MapData(X, Y).OBJInfo.Amount = Amount


Grh_Load MapData(X, Y).ObjGrh, grhindex

Select Case OBJIndex
    Case 63
        TileEngine.Light_Create X, Y, &HFFFFFF, 3
        Light_Render_All
    Case 1093
        TileEngine.Light_Create X, Y, &HFFFFCC, 1
        Light_Render_All
    Case 740
        TileEngine.Light_Create X, Y, &HFD696A, 1
        Light_Render_All
    Case 668
        TileEngine.Light_Create X, Y, &HFFCC33, 1
        Light_Render_All
    Case 1257
        TileEngine.Light_Create X, Y, &HFFFFCC, 1
        Light_Render_All
    Case 1095
        TileEngine.Light_Create X, Y, &HC9FF, 1
        Light_Render_All
    Case 402
        TileEngine.Light_Create X, Y, &HFFFFCC, 1
        Light_Render_All
    Case 1181
        TileEngine.Light_Create X, Y, &HFFFF00, 1
        Light_Render_All
    Case 1147
        TileEngine.Light_Create X, Y, &HFFAC2C, 1
        Light_Render_All
    Case 1252
        TileEngine.Light_Create X, Y, &HCCFF33, 1
        Light_Render_All
    Case 1333
        TileEngine.Light_Create X, Y, &HFFFF99, 1
        Light_Render_All
    Case 1180
        TileEngine.Light_Create X, Y, &HFFFFCC, 1
        Light_Render_All
End Select

End Sub
Sub Map_Change_Area(ByVal X As Byte, ByVal Y As Byte)
    Dim loopX As Long, loopY As Long
    MinLimiteX = X - 12
    MaxLimiteX = X + 12
    
    MinLimiteY = Y - 10
    MaxLimiteY = Y + 10
    
    For loopX = 1 To 100
        For loopY = 1 To 100
            
            If (loopY < MinLimiteY) Or (loopY > MaxLimiteY) Or (loopX < MinLimiteX) Or (loopX > MaxLimiteX) Then
                If MapData(loopX, loopY).CharIndex > 0 Then
                    If MapData(loopX, loopY).CharIndex <> char_current Then
                        Call TileEngine.Char_Remove(MapData(loopX, loopY).CharIndex)
                    End If
                End If
                
                If MapData(loopX, loopY).ObjGrh.grhindex <> 0 Then
                    Call TileEngine.Map_Obj_Delete(loopX, loopY)
                End If
                
            End If
        Next loopY
    Next loopX
End Sub
Sub Map_Obj_Delete(ByVal X As Integer, ByVal Y As Integer, Optional ByVal fijo As Boolean = False)
'**************************************************************
'Author: Leandro Mendoza (Mannakia)
'Last Modify Date: 14/10/2010
'Set ObjGrh in array mapdata the nothing value
'**************************************************************

If X = 0 Or Y = 0 Then Exit Sub
If fijo = False Then
    If MapData(X, Y).OBJInfo.EsFijo = 1 Then
        Exit Sub
    End If
End If

MapData(X, Y).ObjGrh.grhindex = 0


Select Case MapData(X, Y).OBJInfo.OBJIndex
    Case 63
        Light_Remove_From_Pos X, Y, -1
    Case 1093
        Light_Remove_From_Pos X, Y, D3DColorXRGB(255, 255, 204)
    Case 740
        Light_Remove_From_Pos X, Y, D3DColorXRGB(253, 105, 106)
    Case 668
        Light_Remove_From_Pos X, Y, D3DColorXRGB(255, 204, 51)
    Case 1257
        Light_Remove_From_Pos X, Y, D3DColorXRGB(255, 255, 204)
    Case 1095
        Light_Remove_From_Pos X, Y, D3DColorXRGB(0, 57, 255)
    Case 402
        Light_Remove_From_Pos X, Y, D3DColorXRGB(255, 255, 204)
    Case 1181
        Light_Remove_From_Pos X, Y, D3DColorXRGB(255, 255, 0)
    Case 1147
        Light_Remove_From_Pos X, Y, D3DColorXRGB(255, 172, 44)
    Case 1252
        Light_Remove_From_Pos X, Y, D3DColorXRGB(204, 255, 51)
    Case 1333
        Light_Remove_From_Pos X, Y, D3DColorXRGB(255, 255, 153)
    Case 1180
        Light_Remove_From_Pos X, Y, D3DColorXRGB(255, 255, 204)
End Select

MapData(X, Y).OBJInfo.OBJIndex = 0

End Sub
Sub Map_Render()
    Dim Y                   As Integer
    Dim X                   As Integer
    Dim screenminY          As Integer
    Dim screenmaxY          As Integer
    Dim screenminX          As Integer
    Dim screenmaxX          As Integer
    Dim minY                As Integer
    Dim maxY                As Integer
    Dim minX                As Integer
    Dim maxX                As Integer
    Dim ScreenX             As Integer
    Dim ScreenY             As Integer
    Dim minXOffset          As Integer
    Dim minYOffset          As Integer
    Dim PixelOffsetXTemp    As Integer
    Dim PixelOffsetYTemp    As Integer
    Dim addX As Integer, addY As Integer
    Dim CurrentGrhIndex     As Integer
    Dim offx                As Integer
    Dim offy                As Integer
    Dim tilex               As Integer
    Dim tiley               As Integer
    Dim PixelOffSetX        As Integer
    Dim PixelOffSetY        As Integer
    Dim Techos(0 To 3)      As Long
    Static OffsetCounterX   As Single
    Static OffsetCounterY   As Single
    
 '   On Error GoTo ErrHandler
    
    If scroll_on Then
        If AddPosX <> 0 Then
            ScrollX = ScrollX + scroll_pixels_per_frame * Sgn(AddPosX) * timerTicksPerFrame
            If (Sgn(AddPosX) = 1 And ScrollX >= 0) Or (Sgn(AddPosX) = -1 And ScrollX <= 0) Then
                ScrollX = 0
                AddPosX = 0
                scroll_on = False
            End If
        End If
        
        '****** Move screen Up and Down if needed ******
        If AddPosY <> 0 Then
            ScrollY = ScrollY + scroll_pixels_per_frame * Sgn(AddPosY) * timerTicksPerFrame
            If (Sgn(AddPosY) = 1 And ScrollY >= 0) Or (Sgn(AddPosY) = -1 And ScrollY <= 0) Then
                ScrollY = 0
                AddPosY = 0
                scroll_on = False
            End If
        End If
    End If

    tilex = UserPos.X
    tiley = UserPos.Y
    
    PixelOffSetX = ScrollX '_offset_counter_x
    PixelOffSetY = ScrollY '_offset_counter_y
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - (TileBufferSize + 4)
    maxY = screenmaxY + (TileBufferSize + 4)
    minX = screenminX - (TileBufferSize + 4)
    maxX = screenmaxX + (TileBufferSize + 4)
    
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
        addY = 1
    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
        addX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    Engine_Alpha_Calculate Techos



    'Draw floor layer
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
            If Map_Bounds(X, Y) Then
            
                'Layer 1 **********************************
                If MapData(X, Y).Graphic(1).grhindex <> 0 Then
                  
                    Call Draw_Grh(MapData(X, Y).Graphic(1), (ScreenX - 1) * 32 + PixelOffSetX, (ScreenY - 1) * 32 + PixelOffSetY, 0, 1, MapData(X, Y).light_value, , X, Y, , 5)
                End If
                '******************************************
            End If
            ScreenX = ScreenX + 1
        Next X

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y
    
    
    ScreenY = 0
    ScreenX = 0
    
    'Draw floor layer
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
            If Map_Bounds(X, Y) Then
                'Layer 2 **********************************
                If MapData(X, Y).Graphic(2).grhindex <> 0 Then
                    Call Draw_Grh(MapData(X, Y).Graphic(2), (ScreenX - 1 + addX) * 32 + PixelOffSetX, (ScreenY - 1 + addY) * 32 + PixelOffSetY, 1, 1, MapData(X, Y).light_value, , X, Y)
                End If
                '******************************************
            End If
            ScreenX = ScreenX + 1
        Next X

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y

    ScreenY = minYOffset - (TileBufferSize + 4)
    For Y = minY To maxY
        ScreenX = minXOffset - (TileBufferSize + 4)
        For X = minX To maxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffSetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffSetY
            With MapData(X, Y)
                If Map_Bounds(X, Y) Then
                    '******************************************
                    'Object Layer **********************************
                    If .ObjGrh.grhindex <> 0 Then
                        Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value, , X, Y)
                    End If
                        
                    '***********************************************
                    'Layer 3 *****************************************
                    If .Graphic(3).grhindex <> 0 And Not .Graphic(3).grhindex = .ObjGrh.grhindex Then
                        Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value, , X, Y)
                    End If
                    '************************************************
                    'Char layer ************************************
                    If .CharIndex <> 0 Then
                        If charlist(.CharIndex).Pos.X <> X Or charlist(.CharIndex).Pos.Y <> Y Then
                            Call Char_Refresh(.CharIndex)
                            .CharIndex = 0
                        Else
                            Call Char_Render(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp, X, Y)
                        End If
                    End If
                    '*************************************************
                    
                    If .fXGrh.grhindex <> 0 Then
                        Call Draw_Grh(.fXGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, .light_value, True, X, Y)
                        If .fXGrh.Started = 0 Then .fXGrh.grhindex = 0
                    End If
        
                
                End If
            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    ScreenY = minYOffset - 5

    ScreenY = minYOffset - (TileBufferSize + 4)
    For Y = minY To maxY
        ScreenX = minXOffset - (TileBufferSize + 4)
        For X = minX To maxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffSetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffSetY
            'Particles*****************************************
            If Map_Bounds(X, Y) Then
                If MapData(X, Y).particle_group_index > 0 Then
                    Particle_Group_Render MapData(X, Y).particle_group_index, ScreenX * 32 + PixelOffSetX, ScreenY * 32 + PixelOffSetY
                End If
            End If
            '**************************************************
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    ScreenY = minYOffset - 5
    
    If AlphaY <> 0 Then
        'Draw blocked tiles and grid
        ScreenY = minYOffset - (TileBufferSize + 4)
        For Y = minY To maxY
            ScreenX = minXOffset - (TileBufferSize + 4)
            For X = minX To maxX
                If Map_Bounds(X, Y) Then
                    'Layer 4 **********************************
                    If MapData(X, Y).Graphic(4).grhindex Then
                        Call Draw_Grh(MapData(X, Y).Graphic(4), _
                            ScreenX * 32 + PixelOffSetX, _
                            ScreenY * 32 + PixelOffSetY, _
                            1, 1, Techos, , X, Y)
                    End If
                    '**********************************
                End If
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If
    
    If UserCiego Then
        D3DDevice.SetTexture 0, Nothing
        
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, ceg_vertex(0), Len(ceg_vertex(0))
    End If
    
    If map_letter_fadestatus > 0 Then
        If map_letter_fadestatus = 1 Then
            map_letter_a = map_letter_a + (timerTicksPerFrame * 3.5)
            If map_letter_a >= 255 Then
                map_letter_a = 255
                map_letter_fadestatus = 2
            End If
        Else
            map_letter_a = map_letter_a - (timerTicksPerFrame * 3.5)
            If map_letter_a <= 0 Then
                map_letter_fadestatus = 0
                map_letter_a = 0
                 
                If map_letter_grh_next > 0 Then
                    map_letter_grh.grhindex = map_letter_grh_next
                    map_letter_fadestatus = 1
                    map_letter_grh_next = 0
                End If
                
            End If
        End If
    
        Techos(0) = D3DColorARGB(CInt(map_letter_a), 255, 255, 255)
        Techos(1) = Techos(0)
        Techos(2) = Techos(0)
        Techos(3) = Techos(0)
        Grh_Render map_letter_grh, 250, 75, Techos
    End If
    
    'If FPSFLAG Then
    Text_Render FPS & " FPS", 490, 10, Default_RGB

'Exit Sub

'ErrHandler:
 '   LogError ("Map Render: " & err.Description)

End Sub

Private Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.Z = 0
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'
' * v1      * v3
' |\        |
' |  \      |
' |    \    |
' |      \  |
' |        \|
' * v0      * v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.bottom - dest.Top) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
    
    
    '0 - Bottom left vertex
    If Textures_Width And Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
    
    
    '1 - Top left vertex
    If Textures_Width And Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width, src.Top / Textures_Height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
    
    
    '2 - Bottom right vertex
    If Textures_Width And Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
    
    
    '3 - Top right vertex
    If Textures_Width And Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)
    End If

End Sub
Public Sub Device_Box_Textured_Render_Advance(ByVal grhindex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, ByVal dest_width As Integer, Optional ByVal dest_height As Integer, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/15/2003
'Copies the Textures allowing resizing
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long

    If grhindex = 0 Then Exit Sub
    
    Set d3dTextures.Texture = SurfaceDB.GetTexture(GrhData(grhindex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
    
    If (light_value(0) = 0) Then light_value(0) = AmbientColor 'ambientColor
    If (light_value(1) = 0) Then light_value(1) = AmbientColor 'ambientColor
    If (light_value(2) = 0) Then light_value(2) = AmbientColor ' ambientColor
    If (light_value(3) = 0) Then light_value(3) = AmbientColor 'ambientColor
    
    dest_rect.bottom = dest_y + dest_height
    dest_rect.Left = dest_x
    dest_rect.Right = dest_x + dest_width
    dest_rect.Top = dest_y

    If angle <> 0 Then
        With src_rect
            .bottom = src_y + src_height
            .Left = src_x
            .Right = src_x + src_width
            .Top = src_y
        End With
        
        Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
    Else
        If GrhData(grhindex).hardcor = 0 Then
            With src_rect
                .bottom = src_y + src_height
                .Left = src_x
                .Right = src_x + src_width
                .Top = src_y
            End With
            
            Grh_Load_TuTv grhindex, src_rect, d3dTextures.texheight, d3dTextures.texwidth
        End If
        
        temp_verts(0) = Geometry_Create_TLVertex(dest_rect.Left, dest_rect.bottom, 1, 1, light_value(0), 0, GrhData(grhindex).tu(0), GrhData(grhindex).tv(0))
        temp_verts(1) = Geometry_Create_TLVertex(dest_rect.Left, dest_rect.Top, 1, 1, light_value(1), 0, GrhData(grhindex).tu(1), GrhData(grhindex).tv(1))
        temp_verts(2) = Geometry_Create_TLVertex(dest_rect.Right, dest_rect.bottom, 1, 1, light_value(2), 0, GrhData(grhindex).tu(2), GrhData(grhindex).tv(2))
        temp_verts(3) = Geometry_Create_TLVertex(dest_rect.Right, dest_rect.Top, 1, 1, light_value(3), 0, GrhData(grhindex).tu(3), GrhData(grhindex).tv(3))
    End If
    
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.Texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub

Public Sub Device_Box_Textured_Render(ByVal grhindex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single, _
                                            Optional ByVal Shadow As Byte = 0)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/12/2004
'Just copies the Textures
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long
    
    If grhindex = 0 Then Exit Sub
    Set d3dTextures.Texture = SurfaceDB.GetTexture(GrhData(grhindex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)

    If (light_value(0) = 0) Then light_value(0) = AmbientColor 'ambientColor
    If (light_value(1) = 0) Then light_value(1) = AmbientColor 'ambientColor
    If (light_value(2) = 0) Then light_value(2) = AmbientColor ' ambientColor
    If (light_value(3) = 0) Then light_value(3) = AmbientColor 'ambientColor
    
    If frmMain.Visible Then
        If dest_x > 544 Or dest_x + src_width <= 0 Then Exit Sub
        If dest_y > 417 Or dest_y + src_height <= 0 Then Exit Sub
    End If
    
'    If dest_x + src_width > 544 Then src_width = src_width - (dest_x + src_width - 544)
'    If dest_y + src_height > 544 Then src_height = src_height - (dest_y + src_height - 417)
'
'    If dest_y < 0 Then
'        src_height = src_height - dest_y
'        src_y = src_y + dest_y
'        dest_y = dest_y + dest_y
'    End If
'
'    If dest_x < 0 Then
'        src_width = src_width - dest_x
'        src_x = src_x + dest_x
'        dest_x = dest_x + dest_x
'    End If

    'Set up the destination rectangle

    dest_rect.bottom = dest_y + src_height
    dest_rect.Left = dest_x
    dest_rect.Right = dest_x + src_width
    dest_rect.Top = dest_y

    If angle <> 0 Then
        With src_rect
            .bottom = src_y + src_height
            .Left = src_x
            .Right = src_x + src_width
            .Top = src_y
        End With
        
        Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
    Else
        If GrhData(grhindex).hardcor = 0 Then
            With src_rect
                .bottom = src_y + src_height
                .Left = src_x
                .Right = src_x + src_width
                .Top = src_y
            End With
            
            Grh_Load_TuTv grhindex, src_rect, d3dTextures.texheight, d3dTextures.texwidth
        End If
        
        temp_verts(0) = Geometry_Create_TLVertex(dest_rect.Left, dest_rect.bottom, 1, 1, light_value(0), 0, GrhData(grhindex).tu(0), GrhData(grhindex).tv(0))
        temp_verts(1) = Geometry_Create_TLVertex(dest_rect.Left, dest_rect.Top, 1, 1, light_value(1), 0, GrhData(grhindex).tu(1), GrhData(grhindex).tv(1))
        temp_verts(2) = Geometry_Create_TLVertex(dest_rect.Right, dest_rect.bottom, 1, 1, light_value(2), 0, GrhData(grhindex).tu(2), GrhData(grhindex).tv(2))
        temp_verts(3) = Geometry_Create_TLVertex(dest_rect.Right, dest_rect.Top, 1, 1, light_value(3), 0, GrhData(grhindex).tu(3), GrhData(grhindex).tv(3))
    End If
    
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.Texture

    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If

    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub


Public Sub Engine_MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    Dim oX As Integer
    Dim oY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
        
        Case E_Heading.EAST
            X = 1
        
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y
    oX = UserPos.X
    oY = UserPos.Y
    
    'Check to see if its out of bounds
    If tX < XMinMapSize Or tX > XMaxMapSize Or tY < YMinMapSize Or tY > YMaxMapSize Then
        Exit Sub
    Else
    
        UserPos.X = tX
        UserPos.Y = tY
        
        ScrollX = -1 * (32 * -X)
        ScrollY = -1 * (32 * -Y)
        AddPosX = Sgn(-X)
        AddPosY = Sgn(-Y)
        
        scroll_on = True

        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4 Or _
                MapData(UserPos.X, UserPos.Y).Trigger >= 20, True, False)
    End If
    Call DibujarMiniMapPos
End Sub
Sub Char_Remove(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next
    charlist(CharIndex).active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until charlist(LastChar).active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    If Not charlist(CharIndex).Pos.Y = 0 And Not charlist(CharIndex).Pos.X = 0 Then MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y).CharIndex = 0
    
    'Remove char's dialog
    'Call Dialogos.RemoveDialog(CharIndex)
    Call TileEngine.Char_Dialog_Remove(CharIndex)
    
    Call TileEngine.Char_Reset_Info(CharIndex)

    
    'Update NumChars
    NumChars = NumChars - 1
End Sub
Sub Char_Reset_Info(ByVal CharIndex As Integer)
    With charlist(CharIndex)
        Char_Particle_Group_Remove_All CharIndex
        
        .active = 0
        .fxIndex = 0
        
        .invisible = False

        .Muerto = False
        .pie = False
        .UsandoArma = False
         .nombre = ""
         .offClanX = 0
         .offNameX = 0
         .clan = ""
        
        .Pos.X = 0
        .Pos.Y = 0
    End With
End Sub
Public Sub Char_Dialog_Create(ByVal CharIndex As Long, ByRef chat As String, ByVal color As Long, Optional dialogIndex As Byte = 1)
'***************************************************
'Author: Leandro Mendoza(Mannakia)
'Last Modify Date: 6/10/10
'Enter the dialog chat string and color on charindex
'***************************************************
    If CharIndex <= 0 Then Exit Sub
    
    With charlist(CharIndex)
        'Set the string .Dialog with format for aline
        .dialog = FormatChat(chat)
        
        'Set the color of dialog chat
        .dialogColor(0) = color
        .dialogColor(1) = color
        .dialogColor(2) = color
        .dialogColor(3) = color
        
        'Set TRUE dialog
        .dl = True
        
        .dialogLife = 5000 + (100 * Len(chat))
        .dialogStart = GetTickCount()
        
        .dialogHeight = 12
        
        .dialogIndex = dialogIndex
    End With
End Sub
Sub Char_Set_Aura(ByVal CharIndex As Integer, ByVal Escudo As Byte, ByVal Arma As Byte, ByVal body As Integer)
'***************************************************
'Author: Leandro Mendoza(Mannakia)
'Last Modify Date: 21/10/10
'Put "AURA" on the charIndex
'***************************************************
With charlist(CharIndex)
    If body = 255 Then
        Grh_Load .plusGrh(2).Grh, 20206
        .plusGrh(2).color = &HFFFD7E
    Else
        .plusGrh(2).Grh.grhindex = 0
    End If

    If Escudo = 32 Then
        Grh_Load .plusGrh(1).Grh, 20128
        .plusGrh(1).color = &HFFCC33
    Else
        .plusGrh(1).Grh.grhindex = 0
    End If

    If Arma = 16 Then
        Grh_Load .plusGrh(0).Grh, 20128
        .plusGrh(0).color = &HFFCC33
    ElseIf Arma = 17 Then
        Grh_Load .plusGrh(0).Grh, 20133
        .plusGrh(0).color = &HFF3300
    ElseIf Arma = 23 Then
        Grh_Load .plusGrh(0).Grh, 20152
        .plusGrh(0).color = &HFF0000
    ElseIf Arma = 24 Then
        Grh_Load .plusGrh(0).Grh, 20185
        .plusGrh(0).color = -65536
    ElseIf Arma = 25 Then
        Grh_Load .plusGrh(0).Grh, 20155
        .plusGrh(0).color = &HFF0000
    ElseIf Arma = 26 Then
        Grh_Load .plusGrh(0).Grh, 20151
        .plusGrh(0).color = &HFFFF00
    ElseIf Arma = 27 Then
        Grh_Load .plusGrh(0).Grh, 20147
        .plusGrh(0).color = &HFF
    ElseIf Arma = 28 Then
        Grh_Load .plusGrh(0).Grh, 20146
        .plusGrh(0).color = &H6B1B
    ElseIf Arma = 29 Then
        Grh_Load .plusGrh(0).Grh, 20200
        .plusGrh(0).color = &HCCFF33
    Else
        .plusGrh(0).Grh.grhindex = 0
    End If
    
    If body = 291 Then
        .ShieldOffSetY = 30
    ElseIf body = 415 Or body = 384 Or body = 382 Then
        .ShieldOffSetY = 16
    ElseIf body = 416 Then
        .ShieldOffSetY = 32
    ElseIf body = 282 Or body = 292 Then
        .ShieldOffSetY = 20
    ElseIf body = 381 Or body = 383 Then
        .ShieldOffSetY = 24
    Else
        .ShieldOffSetY = 0
    End If
    
    If body > 0 And body < UBound(BodyData) Then
        If BodyData(body).HeadOffset.Y = -28 Then
            .ShieldOffSetY = .ShieldOffSetY - 5
        End If
    End If
End With
End Sub
Public Sub Char_Dialog_Remove(ByVal CharIndex As Long)
'***************************************************
'Author: Leandro Mendoza(Mannakia)
'Last Modify Date: 7/10/10
'Delete the dialog chat of the charIndex
'***************************************************
With charlist(CharIndex)
    'Destroit the array string dialog chat
    Erase .dialog
    
    'Set FALSE dialog
    .dl = False
    
    'Set default color , White
    .dialogColor(0) = -1
    .dialogColor(1) = -1
    .dialogColor(2) = -1
    .dialogColor(3) = -1
End With
End Sub
Public Sub Char_Dialog_Remove_All()
'***************************************************
'Author: Leandro Mendoza(Mannakia)
'Last Modify Date: 7/10/10
'Delete the all dialog chat
'***************************************************

'Simple
Dim i As Long
For i = 1 To LastChar
    Char_Dialog_Remove i
Next i
End Sub
Private Sub Char_Render(ByVal CharIndex As Long, ByVal PixelOffSetX As Integer, ByVal PixelOffSetY As Integer, ByVal X As Byte, ByVal Y As Byte)
    Dim Pos As Integer
    Dim line As String
    Dim color As Long
    Static rgb_list(0 To 3) As Long
    Dim Moved As Boolean
    
    With charlist(CharIndex)
    
        If .heading = 0 Then Exit Sub
        
        If .Moving Then
            If .ScrollDirectionX <> 0 Then
                .MoveOffSetX = .MoveOffSetX + (IIf(char_current = CharIndex, scroll_pixels_per_frame, 5.2) * Sgn(.ScrollDirectionX) * timerTicksPerFrame)
    
                'Start animation
                .body.Walk(.heading).Started = 1
                .Escudo.ShieldWalk(.heading).Started = 1
                .Arma.WeaponWalk(.heading).Started = 1
                
                'Char moved
                Moved = True
    
                'Check if we already got there
                If (Sgn(.ScrollDirectionX) = 1 And .MoveOffSetX >= 0) Or (Sgn(.ScrollDirectionX) = -1 And .MoveOffSetX <= 0) Then
                    .MoveOffSetX = 0
                    .ScrollDirectionX = 0
                End If
    
            End If
    
            'If needed, move up and down
            If .ScrollDirectionY <> 0 Then
                .MoveOffSetY = .MoveOffSetY + (IIf(char_current = CharIndex, scroll_pixels_per_frame, 5.2) * Sgn(.ScrollDirectionY) * timerTicksPerFrame)
    
                'Start animation
                .body.Walk(.heading).Started = 1
                .Escudo.ShieldWalk(.heading).Started = 1
                .Arma.WeaponWalk(.heading).Started = 1
                
                'Char moved
                Moved = True
    
                'Check if we already got there
                If (Sgn(.ScrollDirectionY) = 1 And .MoveOffSetY >= 0) Or (Sgn(.ScrollDirectionY) = -1 And .MoveOffSetY <= 0) Then
                    .MoveOffSetY = 0
                    .ScrollDirectionY = 0
                End If
    
            End If
        End If
    
        'Update movement reset timer
        If .ScrollDirectionX = 0 Or .ScrollDirectionY = 0 Then
            'If done moving stop animation
            If Not Moved Then
                If .body.Walk(.heading).Started Then
    
                    'Stop animation
                    .body.Walk(.heading).Started = 0
                    .body.Walk(.heading).FrameCounter = 1
                    
                    .Moving = 0
                End If
                
                If .Arma.WeaponWalk(.heading).Started Then
                    'Stop animation
                    .Arma.WeaponWalk(.heading).Started = 0
                    .Arma.WeaponWalk(.heading).FrameCounter = 1
                End If
                
                If .Escudo.ShieldWalk(.heading).Started Then
                     'Stop animation
                    .Escudo.ShieldWalk(.heading).Started = 0
                    .Escudo.ShieldWalk(.heading).FrameCounter = 1
                End If
            End If
        End If

        PixelOffSetX = PixelOffSetX + .MoveOffSetX
        PixelOffSetY = PixelOffSetY + .MoveOffSetY
        
        
        
        
        If .Bandera = 1 Then
            Call Draw_Grh_Index(26783, PixelOffSetX - .offNameX + 5, PixelOffSetY - 105, D3DColorXRGB(255, 255, 255))
        ElseIf .Bandera = 2 Then
            Call Draw_Grh_Index(30095, PixelOffSetX - .offNameX + 5, PixelOffSetY - 105, D3DColorXRGB(255, 255, 255))
        ElseIf .Bandera = 3 Then
            Call Draw_Grh_Index(29179, PixelOffSetX - .offNameX + 5, PixelOffSetY - 105, D3DColorXRGB(255, 255, 255))
        End If
        
        'Estandarte de la Armada Real|26783|20
        'Estandarte de las Hordas del Caos|29179|20
        'Estandarte de la República|30095|20
        'Mástil sin estandarte|29180|20
         
        
        If .Head.Head(.heading).grhindex Then
            If .plusGrh(2).Grh.grhindex <> 0 Then
                Call Engine_Long_To_RGB_List(rgb_list, .plusGrh(2).color)
                Call Grh_Render(.plusGrh(2).Grh, PixelOffSetX, PixelOffSetY + 32, rgb_list, , , True)
            End If
            
            If .plusGrh(0).Grh.grhindex <> 0 Then
                Call Engine_Long_To_RGB_List(rgb_list, .plusGrh(0).color)
                Call Grh_Render(.plusGrh(0).Grh, PixelOffSetX, PixelOffSetY + 28, rgb_list, , , True)
            End If
            
            If .plusGrh(1).Grh.grhindex <> 0 Then
                Call Engine_Long_To_RGB_List(rgb_list, .plusGrh(1).color)
                Call Grh_Render(.plusGrh(1).Grh, PixelOffSetX, PixelOffSetY + 32, rgb_list, , , True)
            End If
        End If
                
        If Not .invisible And Not .Muerto Then
            .AlphaX = 0
            rgb_list(0) = MapData(X, Y).light_value(0)
            rgb_list(1) = MapData(X, Y).light_value(1)
            rgb_list(2) = MapData(X, Y).light_value(2)
            rgb_list(3) = MapData(X, Y).light_value(3)
        Else
            Call Char_Alpha_Calculate(rgb_list(), CharIndex)
        End If
        
        If .Head.Head(.heading).grhindex Then
            If .heading = EAST Then
                If .Escudo.ShieldWalk(.heading).grhindex And Not .iBody = 84 Then
                    Call Draw_Grh(.Escudo.ShieldWalk(.heading), PixelOffSetX, PixelOffSetY - .ShieldOffSetY, 1, 1, rgb_list(), , X, Y)
                End If
            End If
        End If

        If .body.Walk(.heading).grhindex Then
            Call Draw_Grh(.body.Walk(.heading), PixelOffSetX, PixelOffSetY, 1, 1, rgb_list(), , X, Y)
        End If

        If .Head.Head(.heading).grhindex Or .iBody = 87 Or .iBody = 84 Then
            Call Draw_Grh(.Head.Head(.heading), PixelOffSetX + .body.HeadOffset.X, PixelOffSetY + .body.HeadOffset.Y, 1, 0, rgb_list(), , X, Y)
   
            If Not .iBody = 84 Then
                If .Arma.WeaponWalk(.heading).grhindex And Char_Is_Montura(.iBody) Then
                    Call Draw_Grh(.Arma.WeaponWalk(.heading), PixelOffSetX, PixelOffSetY + (.body.HeadOffset.Y + 38), 1, 1, rgb_list(), , X, Y)
                End If
                    
                If .heading <> EAST Then
                    If .Escudo.ShieldWalk(.heading).grhindex Then
                        Call Draw_Grh(.Escudo.ShieldWalk(.heading), PixelOffSetX, PixelOffSetY - .ShieldOffSetY, 1, 1, rgb_list(), , X, Y)
                    End If
                End If
        
                If .Casco.Head(.heading).grhindex Then
                    Call Draw_Grh(.Casco.Head(.heading), PixelOffSetX + .body.HeadOffset.X, PixelOffSetY + .body.HeadOffset.Y, 1, 0, rgb_list(), , X, Y)
                End If
            End If
                              
            'Draw name over head
            If Nombres And .invisible = False Then
                If Len(.nombre) > 0 Then
                    Text_Render .nombre, PixelOffSetX - .offNameX + 16, PixelOffSetY + 30, .label_color
                    
                
                    If .donador = True Then
                        ' Dibujo estrellita de donador P or jo se cas telli
                        Call Draw_Grh_Index(31981, PixelOffSetX - .offNameX + 2, PixelOffSetY + 30, D3DColorXRGB(255, 255, 255))
                    End If
                   
                   
                    If .offClanX > 0 Then
                        Text_Render .clan, PixelOffSetX - .offClanX + 18, PixelOffSetY + 45, .label_color
                    End If
                End If
            
            End If
    
        End If
        
        Dim i As Long
        If charlist(CharIndex).particle_count > 0 Then
            For i = 1 To UBound(charlist(CharIndex).particle_group)
                If charlist(CharIndex).particle_group(i) > 0 Then
                    Particle_Group_Render charlist(CharIndex).particle_group(i), PixelOffSetX, PixelOffSetY
                End If
            Next i
        End If
        
        'Draw FX
        If .fxIndex <> 0 Then
            Call Draw_Grh(.FX, PixelOffSetX + FxData(.fxIndex).OffsetX, PixelOffSetY + FxData(.fxIndex).OffsetY, 1, 1, MapData(X, Y).light_value, True)

            If .FX.Started = 0 Then
                .fxIndex = 0
            End If
            
        End If
        
        If .dl Then
            If GetTickCount() - .dialogStart >= .dialogLife Then
                Char_Dialog_Remove CharIndex
            Else
                If .dialogHeight > 0 Then .dialogHeight = .dialogHeight - 1
                
                PixelOffSetY = PixelOffSetY - 13 * UBound(.dialog) + .dialogHeight
                For i = 0 To UBound(.dialog)
                    Text_Render LTrim(.dialog(i)), PixelOffSetX + .body.HeadOffset.X - Text_Width(.dialog(i), 1) / 2 + 16, PixelOffSetY + .body.HeadOffset.Y, .dialogColor, .dialogIndex
                    PixelOffSetY = PixelOffSetY + 8 + 5
                Next i
            End If
        End If
    End With
End Sub
Function Char_Is_Montura(ByVal body As Integer) As Boolean
    Char_Is_Montura = (body <> 416 And body <> 415 And _
                 body <> 412 And body <> 381 And _
                 body <> 383 And body <> 384 And _
                 body <> 382 And body <> 413 And _
                 body <> 292 And body <> 291 And _
                 body <> 317)
End Function
Public Sub Char_SetFx(ByVal CharIndex As Integer, ByVal FX As Integer, ByVal Loops As Integer)

    If CharIndex = 0 Then Exit Sub

    With charlist(CharIndex)
        .fxIndex = FX
        
        If .fxIndex > 0 Then
            Call Grh_Load(.FX, FxData(FX).Animacion)
        
            .FX.Loops = Loops
        End If
    End With
    
End Sub
Public Sub Char_Refresh_All()
    Dim loopc As Long
    
    For loopc = 1 To LastChar
        If charlist(loopc).active = 1 Then
            MapData(charlist(loopc).Pos.X, charlist(loopc).Pos.Y).CharIndex = loopc
        End If
    Next loopc
End Sub
Public Sub Char_Refresh(ByVal CharIndex As Integer)
    If charlist(CharIndex).Pos.X = 0 Or charlist(CharIndex).Pos.Y = 0 Then Exit Sub
    MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y).CharIndex = CharIndex
End Sub
Public Sub Char_Move_Head(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
    Dim addX As Integer, addY As Integer
    Dim X As Integer, Y As Integer
    Dim nx As Integer, ny As Integer
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
       ' If x = 0 Or y = 0 Then Exit Sub
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addY = -1
        
            Case E_Heading.EAST
                addX = 1
        
            Case E_Heading.SOUTH
                addY = 1
            
            Case E_Heading.WEST
                addX = -1
        End Select
        
        nx = X + addX
        ny = Y + addY
        If Map_Bounds(nx, ny) Then
            MapData(X, Y).CharIndex = 0
            
            MapData(nx, ny).CharIndex = CharIndex
            
            .Pos.X = nx
            .Pos.Y = ny
                        
            If MapData(nx, ny).ObjGrh.grhindex = 26940 Then
                .heading = 1
                under_stair = True
            Else
                under_stair = False
                .heading = nHeading
            End If
        End If
        
        .ScrollDirectionX = Sgn(addX)
        .ScrollDirectionY = Sgn(addY)
        
        .MoveOffSetX = -1 * (32 * addX)
        .MoveOffSetY = -1 * (32 * addY)
        
        .Moving = 1

    End With
    
    If UserEstado <> 1 Then Call TileEngine.Char_Pasos_Render(CharIndex)
    
    'areas viejos
    If (ny < MinLimiteY) Or (ny > MaxLimiteY) Or (nx < MinLimiteX) Or (nx > MaxLimiteX) Then
        Call Char_Remove(CharIndex)
    End If
End Sub
Sub Char_Create(ByVal CharIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .active = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = body
        
        .Head = HeadData(Head)
        If body > LBound(BodyData) Or body < UBound(BodyData) Then
            .body = BodyData(body)
        End If
        
        If Not Arma = 29 Then .Arma = WeaponAnimData(Arma)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .heading = heading

        'Update position
        .Pos.X = X
        .Pos.Y = Y
        
        'Make active
        .active = 1
        .Muerto = (Head = CASPER_HEAD)
        Select Case .Priv
            Case 1 'Gris
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(128, 128, 128)
            Case 2 'Azul
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(0, 80, 200)
            Case 3 'Rojo
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(197, 0, 5)
            Case 4 'Naranja
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(243, 147, 1)
            
            Case 9 'Lider Impe
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(0, 80, 200)
            Case 10 'Lider Repu
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(243, 147, 1)
            Case 11 'Lider Caos
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(197, 0, 5)
            
            Case 5 'Verde - conse
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(0, 128, 0)
            Case 6 'Verde - semi
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(0, 128, 0)
            Case 7 'Verde - dios
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(0, 128, 0)
            Case 8 'Verde - admin
                TileEngine.Engine_Long_To_RGB_List .label_color, D3DColorXRGB(10, 10, 10)
        End Select
    End With
    
    Call TileEngine.Char_Set_Aura(CharIndex, Escudo, Arma, body)
    
    MapData(X, Y).CharIndex = CharIndex
    

End Sub
Public Sub Map_Clean()
Dim X As Byte
Dim Y As Byte
For X = 1 To 100
    For Y = 1 To 100
        If MapData(X, Y).CharIndex Then
            Char_Remove MapData(X, Y).CharIndex
        End If
        If MapData(X, Y).ObjGrh.grhindex Then
            TileEngine.Map_Obj_Delete X, Y, True
        End If
    Next Y
Next X
End Sub
Sub Map_Load(ByVal MapRoute As String, Optional Client_Mode As Boolean = False)

TileEngine.Map_Clean
TileEngine.Particle_Group_Remove_All
TileEngine.Light_Remove_All

On Error GoTo ErrorHandler

Dim fh As Integer
Dim MH As tMapHeader
Dim Blqs() As tDatosBloqueados
Dim L1() As Long
Dim L2() As tDatosGrh
Dim L3() As tDatosGrh
Dim L4() As tDatosGrh
Dim Triggers() As tDatosTrigger
Dim Luces() As tDatosLuces
Dim Particulas() As tDatosParticulas
Dim Objetos() As tDatosObjs
Dim npcs() As tDatosNPC
Dim TEs() As tDatosTE

Dim i As Long
Dim j As Long

'Add Marius Agregamos mas mapas de arena
If MapRoute > 851 Then
    MapRoute = 238
End If
'\Add

Extract_File Maps, "mapa" & MapRoute & ".csm", resource_path

fh = FreeFile
Open resource_path & "mapa" & MapRoute & ".csm" For Binary Access Read As fh
    Get #fh, , MH
    Get #fh, , MapSize
    Get #fh, , MapDat
    
    MapDat.map_name = Trim$(MapDat.map_name)
    MapName = MapDat.map_name
    
    If Client_Mode = True Then
        Close #fh
        'Delete_File resource_path & "mapa" & MapRoute & ".csm"
        If FileExist(resource_path & "mapa" & MapRoute & ".csm", vbNormal) Then Kill resource_path & "mapa" & MapRoute & ".csm"
        Exit Sub
    End If
    
    ReDim MapData(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As MapBlock
    ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As Long
    
    Get #fh, , L1
    
    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax
            If L1(i, j) > 0 Then
                Grh_Load MapData(i, j).Graphic(1), L1(i, j)
                TileEngine.SurfaceDB.GetTexture GrhData(L1(i, j)).FileNum, 0, 0
            End If
        Next i
    Next j
    
    
    With MH
        If .NumeroBloqueados > 0 Then
            ReDim Blqs(1 To .NumeroBloqueados)
            Get #fh, , Blqs
            For i = 1 To .NumeroBloqueados
                MapData(Blqs(i).X, Blqs(i).Y).Blocked = 1
            Next i
        End If
        
        If .NumeroLayers(2) > 0 Then
            ReDim L2(1 To .NumeroLayers(2))
            Get #fh, , L2
            For i = 1 To .NumeroLayers(2)
                Grh_Load MapData(L2(i).X, L2(i).Y).Graphic(2), L2(i).grhindex
                TileEngine.SurfaceDB.GetTexture GrhData(L2(i).grhindex).FileNum, 0, 0
            Next i
        End If
        
        If .NumeroLayers(3) > 0 Then
            ReDim L3(1 To .NumeroLayers(3))
            Get #fh, , L3
            For i = 1 To .NumeroLayers(3)
                Grh_Load MapData(L3(i).X, L3(i).Y).Graphic(3), L3(i).grhindex
                TileEngine.SurfaceDB.GetTexture GrhData(L3(i).grhindex).FileNum, 0, 0
            Next i
        End If
        
        If .NumeroLayers(4) > 0 Then
            ReDim L4(1 To .NumeroLayers(4))
            Get #fh, , L4
            For i = 1 To .NumeroLayers(4)
                Grh_Load MapData(L4(i).X, L4(i).Y).Graphic(4), L4(i).grhindex
                TileEngine.SurfaceDB.GetTexture GrhData(L4(i).grhindex).FileNum, 0, 0
            Next i
        End If
        
        If .NumeroTriggers > 0 Then
            ReDim Triggers(1 To .NumeroTriggers)
            Get #fh, , Triggers
            For i = 1 To .NumeroTriggers
                MapData(Triggers(i).X, Triggers(i).Y).Trigger = Triggers(i).Trigger
            Next i
        End If
        
        If .NumeroParticulas > 0 Then
            ReDim Particulas(1 To .NumeroParticulas)
            Get #fh, , Particulas
            For i = 1 To .NumeroParticulas
                MapData(Particulas(i).X, Particulas(i).Y).particle_group_index = General_Particle_Create(Particulas(i).Particula, Particulas(i).X, Particulas(i).Y)
            Next i
        End If
        
        If .NumeroLuces > 0 Then
            ReDim Luces(1 To .NumeroLuces)
            Get #fh, , Luces
            For i = 1 To .NumeroLuces
                If Not Luces(i).color = 0 Then
                    Call TileEngine.Light_Create(Luces(i).X, Luces(i).Y, Luces(i).color, Luces(i).rango)
                End If
            Next i
        End If
        
        If .NumeroOBJs > 0 Then
            ReDim Objetos(1 To .NumeroOBJs)
            Get #fh, , Objetos
            For i = 1 To .NumeroOBJs
                'Add Marius Captura la bandera. Agregé el if solamente.
                If UserMap <> 848 Then ' And ((Objetos(i).X = 12 And Objetos(i).Y = 14) Or (Objetos(i).X = 86 And Objetos(i).Y = 78)) Then
                    Map_Obj_Create Objetos(i).X, Objetos(i).Y, objs(Objetos(i).OBJIndex).Grh, Objetos(i).OBJIndex, objs(Objetos(i).OBJIndex).tipe, Objetos(i).ObjAmmount, 1
                End If
                '\Add
            Next i
        End If
                
        If .NumeroNPCs > 0 Then
            ReDim npcs(1 To .NumeroNPCs)
            Get #fh, , npcs
            For i = 1 To .NumeroNPCs
                MapData(npcs(i).X, npcs(i).Y).NPCIndex = npcs(i).NPCIndex
            Next
                
        End If
        
    End With

Close fh
    
    'Delete_File resource_path & "mapa" & MapRoute & ".csm"
    If FileExist(resource_path & "mapa" & MapRoute & ".csm", vbNormal) Then Kill resource_path & "mapa" & MapRoute & ".csm"
    


    Dim r As Integer, g As Integer, b As Integer
    'Common light value verify
    If MapDat.base_light = 0 Then
        TileEngine.m_Afecta = True
        meteo_hour = 65
        TileEngine.Meteo_Change_Time
    Else
        General_Long_Color_to_RGB MapDat.base_light, r, g, b
        AmbientColor = D3DColorXRGB(r, g, b)
        ambientLight.r = r
        ambientLight.g = g
        ambientLight.b = b
        TileEngine.m_Afecta = False
    End If

    If MapDat.letter_grh <> 0 Then
        Map_Letter_Fade_Set MapDat.letter_grh
    Else
        Map_Letter_UnSet
    End If

    '*******************************
    'Render lights
    TileEngine.Light_Render_All
    '*******************************

    frmMain.Minimap.Cls

    Dim map_x As Long
    Dim map_y As Long
    Dim screen_x As Long
    Dim screen_y As Long
    Dim grh_index As Long
    
    For map_y = MapSize.XMin To MapSize.XMax
        For map_x = MapSize.YMin To MapSize.YMax
            screen_x = (map_x - 1) * 2
            screen_y = (map_y - 1) * 2
            '*** Start Layer 1 ***
            If MapData(map_x, map_y).Graphic(1).grhindex Then
                grh_index = MapData(map_x, map_y).Graphic(1).grhindex
                SetPixel frmMain.Minimap.hdc, map_x - 1, map_y - 1, GrhData(grh_index).mini_map_color
            End If
            If MapData(map_x, map_y).Graphic(2).grhindex Then
                grh_index = MapData(map_x, map_y).Graphic(2).grhindex
                SetPixel frmMain.Minimap.hdc, map_x - 1, map_y - 1, GrhData(grh_index).mini_map_color
            End If
            If MapData(map_x, map_y).Graphic(4).grhindex Then
                grh_index = MapData(map_x, map_y).Graphic(4).grhindex
                SetPixel frmMain.Minimap.hdc, map_x - 1, map_y - 1, GrhData(grh_index).mini_map_color
            End If
        Next map_x
    Next map_y



Exit Sub

ErrorHandler:
    If fh <> 0 Then Close fh
End Sub



Public Sub Char_Move_Pos(ByVal CharIndex As Integer, ByVal nx As Integer, ByVal ny As Integer)
On Error Resume Next
    Dim X As Integer
    Dim Y As Integer
    Dim addX As Integer
    Dim addY As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        If X < 1 Or Y < 1 Or X > 100 Or Y > 100 Then
            .Pos.X = nx
            .Pos.Y = ny
            Exit Sub
        End If
        
        addX = nx - X
        addY = ny - Y
        
        If Sgn(addX) = 1 Then
            nHeading = E_Heading.EAST
        ElseIf Sgn(addX) = -1 Then
            nHeading = E_Heading.WEST
        ElseIf Sgn(addY) = -1 Then
            nHeading = E_Heading.NORTH
        ElseIf Sgn(addY) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        If Map_Bounds(X, Y) Then
            MapData(X, Y).CharIndex = 0
        End If
        
        MapData(nx, ny).CharIndex = CharIndex
        .Pos.X = nx
        .Pos.Y = ny
        
        
        .ScrollDirectionX = Sgn(addX)
        .ScrollDirectionY = Sgn(addY)
        
        .MoveOffSetX = -1 * (32 * addX)
        .MoveOffSetY = -1 * (32 * addY)
        
        .Moving = 1

        If MapData(nx, ny).ObjGrh.grhindex = 26940 Then
            .heading = 1
        Else
            .heading = nHeading
        End If
    End With
    
    
End Sub

Private Function Engine_FToDW(f As Single) As Long
' single > long
Dim buf As D3DXBuffer
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, Engine_FToDW
End Function


Public Sub DecideIfRenderInventory()

    If RenderInv Or activateTempDrawInv Then
        TileEngine.Inventory_Render
    End If
    
End Sub

Public Sub Inventory_Render()
On Error Resume Next
    Static re As RECT
    re.Left = 0
    re.Top = 0
    re.bottom = 160
    re.Right = 160
    
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene
        Inventario.DrawInventory
    D3DDevice.EndScene
    D3DDevice.Present re, ByVal 0, frmMain.picInv.hwnd, ByVal 0
End Sub
Public Sub Draw_Grh_Hdc(ByVal desthDC As Long, ByVal Grh As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
On Error GoTo err
    Dim file_path As String
    Dim src_x As Integer
    Dim src_y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hdcsrc As Long
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long
    Dim grh_index As Integer
    
    grh_index = Grh

    If grh_index <= 0 Then Exit Sub
    If GrhData(grh_index).NumFrames = 0 Then Exit Sub

    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)
    End If

    src_x = GrhData(grh_index).sX
    src_y = GrhData(grh_index).sY
    src_width = GrhData(grh_index).pixelWidth
    src_height = GrhData(grh_index).pixelHeight
            
    hdcsrc = CreateCompatibleDC(desthDC)
    file_path = App.Path & "\" & GrhData(grh_index).FileNum & ".bmp"
    
    Extract_File Graphics, GrhData(grh_index).FileNum & ".bmp", App.Path & "\"
    PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
    Delete_File file_path
    
    'BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
    TransparentBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, src_width, src_height, RGB(0, 0, 0)
    
    Call DeleteObject(SelectObject(hdcsrc, PrevObj))
    DeleteDC hdcsrc
    
err:
If err.Number = 481 Then MsgBox "Imposible cargar recurso error: " & "481"

End Sub

Private Sub Text_Font_Initialize()

Dim a As Integer

font_types(1).font_size = 9
font_types(1).ascii_code(48) = 21452
font_types(1).ascii_code(49) = 21453
font_types(1).ascii_code(50) = 21454
font_types(1).ascii_code(51) = 21455
font_types(1).ascii_code(52) = 21456
font_types(1).ascii_code(53) = 21457
font_types(1).ascii_code(54) = 21458
font_types(1).ascii_code(55) = 21459
font_types(1).ascii_code(56) = 21460
font_types(1).ascii_code(57) = 21461
For a = 0 To 25
font_types(1).ascii_code(a + 97) = 21400 + a
Next a

For a = 0 To 25
font_types(1).ascii_code(a + 65) = 21426 + a
Next a
font_types(1).ascii_code(33) = 21462
font_types(1).ascii_code(161) = 21463
font_types(1).ascii_code(34) = 21464
font_types(1).ascii_code(36) = 21465
font_types(1).ascii_code(191) = 21466
font_types(1).ascii_code(35) = 21467
font_types(1).ascii_code(36) = 21468
font_types(1).ascii_code(37) = 21469
font_types(1).ascii_code(38) = 21470
font_types(1).ascii_code(47) = 21471
font_types(1).ascii_code(92) = 21472
font_types(1).ascii_code(40) = 21473
font_types(1).ascii_code(41) = 21474
font_types(1).ascii_code(61) = 21475
font_types(1).ascii_code(39) = 21476
font_types(1).ascii_code(123) = 21477
font_types(1).ascii_code(125) = 21478
font_types(1).ascii_code(95) = 21479
font_types(1).ascii_code(45) = 21480
font_types(1).ascii_code(63) = 21465
font_types(1).ascii_code(64) = 21481
font_types(1).ascii_code(94) = 21482
font_types(1).ascii_code(91) = 21483
font_types(1).ascii_code(93) = 21484
font_types(1).ascii_code(60) = 21485
font_types(1).ascii_code(62) = 21486
font_types(1).ascii_code(42) = 21487
font_types(1).ascii_code(43) = 21488
font_types(1).ascii_code(46) = 21489
font_types(1).ascii_code(44) = 21490
font_types(1).ascii_code(58) = 21491
font_types(1).ascii_code(59) = 21492
font_types(1).ascii_code(124) = 21493
font_types(1).ascii_code(252) = 21800
font_types(1).ascii_code(220) = 21801
font_types(1).ascii_code(225) = 21802
font_types(1).ascii_code(233) = 21803
font_types(1).ascii_code(237) = 21804
font_types(1).ascii_code(243) = 21805
font_types(1).ascii_code(250) = 21806
font_types(1).ascii_code(253) = 21807
font_types(1).ascii_code(193) = 21808
font_types(1).ascii_code(201) = 21809
font_types(1).ascii_code(205) = 21810
font_types(1).ascii_code(211) = 21811
font_types(1).ascii_code(218) = 21812
font_types(1).ascii_code(221) = 21813
font_types(1).ascii_code(224) = 21814
font_types(1).ascii_code(232) = 21815
font_types(1).ascii_code(236) = 21816
font_types(1).ascii_code(242) = 21817
font_types(1).ascii_code(249) = 21818
font_types(1).ascii_code(192) = 21819
font_types(1).ascii_code(200) = 21820
font_types(1).ascii_code(204) = 21821
font_types(1).ascii_code(210) = 21822
font_types(1).ascii_code(217) = 21823
font_types(1).ascii_code(241) = 21824
font_types(1).ascii_code(209) = 21825
font_types(1).ascii_code(196) = 25238
font_types(1).ascii_code(194) = 25239
font_types(1).ascii_code(203) = 25240
font_types(1).ascii_code(207) = 25241
font_types(1).ascii_code(214) = 25242
font_types(1).ascii_code(212) = 25243

font_types(2).font_size = 9
font_types(2).ascii_code(97) = 21936
font_types(2).ascii_code(108) = 21937
font_types(2).ascii_code(115) = 21938
font_types(2).ascii_code(70) = 21939
font_types(2).ascii_code(48) = 21940
font_types(2).ascii_code(49) = 21941
font_types(2).ascii_code(50) = 21942
font_types(2).ascii_code(51) = 21943
font_types(2).ascii_code(52) = 21944
font_types(2).ascii_code(53) = 21945
font_types(2).ascii_code(54) = 21946
font_types(2).ascii_code(55) = 21947
font_types(2).ascii_code(56) = 21948
font_types(2).ascii_code(57) = 21949
font_types(2).ascii_code(33) = 21950
font_types(2).ascii_code(161) = 21951
font_types(2).ascii_code(42) = 21952

End Sub
Sub Text_Render(Texto As String, X As Integer, Y As Integer, ByRef text_color() As Long, Optional ByVal font_index As Integer = 1, Optional multi_line As Boolean = False)

Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer, f As Integer, g As Integer
Dim graf As Grh

Dim temp_array(3) As Long 'Si le queres dar color a la letra pasa este parametro dsp xD
temp_array(0) = text_color(0)
temp_array(1) = text_color(1)
temp_array(2) = text_color(2)
temp_array(3) = text_color(3)

If (Len(Texto) = 0) Then Exit Sub

d = 0
If multi_line = False Then
    For a = 1 To Len(Texto)
        b = Asc(mid(Texto, a, 1))
        graf.grhindex = font_types(font_index).ascii_code(b)
        If b <> 32 Then
            If graf.grhindex <> 0 Then
                'mega sombra O-matica
                graf.grhindex = font_types(font_index).ascii_code(b) + 100
                Grh_Render graf, (X + d) + 1, Y + 1, temp_array, False, False, False
                graf.grhindex = font_types(font_index).ascii_code(b)
                Grh_Render graf, (X + d), Y, temp_array, False, False, False
                d = d + GrhData(GrhData(graf.grhindex).Frames(1)).pixelWidth '+ 1
            End If
        Else
            d = d + 4
        End If
    Next a
Else
    e = 0
    f = 0
    For a = 1 To Len(Texto)
        b = Asc(mid(Texto, a, 1))
        graf.grhindex = font_types(font_index).ascii_code(b)
        If b = 32 Or b = 13 Then
            If e >= 20 Then 'reemplazar por lo que os plazca
                f = f + 1
                e = 0
                d = 0
            Else
                If b = 32 Then d = d + 4
            End If
        Else
            If graf.grhindex > 12 Then
                'mega sombra O-matica
                graf.grhindex = font_types(font_index).ascii_code(b) + 100
                Grh_Render graf, (X + d) + 1, Y + 1 + f * 14, temp_array, False, False, False
                graf.grhindex = font_types(font_index).ascii_code(b)
                Grh_Render graf, (X + d), Y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                d = d + GrhData(GrhData(graf.grhindex).Frames(1)).pixelWidth '+ 1
            End If
       End If
       e = e + 1
    Next a
End If

End Sub


Private Function Particle_Group_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    If particle_group_last = 0 Then
        Particle_Group_Next_Open = 1
        Exit Function
    End If
    
    loopc = 1
    Do Until particle_group_list(loopc).active = False
        If loopc = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Next_Open = loopc
Exit Function
ErrorHandler:
    Particle_Group_Next_Open = 1
End Function

Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).active Then
            Particle_Group_Check = True
        End If
    End If
End Function

Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/14/2003
'Returns the particle_group_index if successful, else 0
'Modified by Juan Martín Sotuyo Dodero
'Modified by Augusto José Rando
'**************************************************************
    
    If (map_x <> -1) And (map_y <> -1) Then
        If Particle_Map_Group_Get(map_x, map_y) = 0 Then
            Particle_Group_Create = Particle_Group_Next_Open
            Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
        End If
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    End If

End Function

Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True
    End If
End Function

Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index
        End If
    Next Index
    
    Particle_Group_Remove_All = True
End Function

Public Function Particle_Group_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until particle_group_list(loopc).id = id
        If loopc = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Find = loopc
Exit Function
ErrorHandler:
    Particle_Group_Find = 0
End Function
Public Function Particle_Get_Type(ByVal particle_group_index As Long) As Byte
On Error GoTo ErrorHandler:
    Particle_Get_Type = particle_group_list(particle_group_index).stream_type
Exit Function
ErrorHandler:
    Particle_Get_Type = 0
End Function
Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
On Error Resume Next
    Dim temp As particle_group
    Dim i As Integer
    
    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then
        MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group_index = 0
    ElseIf particle_group_list(particle_group_index).char_index Then
        If Char_Check(particle_group_list(particle_group_index).char_index) Then
            For i = 1 To charlist(particle_group_list(particle_group_index).char_index).particle_count
                If charlist(particle_group_list(particle_group_index).char_index).particle_group(i) = particle_group_index Then
                    charlist(particle_group_list(particle_group_index).char_index).particle_group(i) = 0
                    Exit For
                End If
            Next i
        End If
    ElseIf particle_group_index = meteo_particle Then
        meteo_particle = 0
    End If
    
    particle_group_list(particle_group_index) = temp
    
    'Update array size
    If particle_group_index = particle_group_last Then
        Do Until particle_group_list(particle_group_last).active
            particle_group_last = particle_group_last - 1
            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve particle_group_list(1 To particle_group_last) As particle_group
    End If
    particle_group_count = particle_group_count - 1
End Sub
Private Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'Renders a particle stream at a paticular screen point
'*****************************************************************
    Dim loopc As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
    Dim detroy As Byte
    
    If particle_group_index > UBound(particle_group_list) Then Exit Sub
    
    If GetTickCount - particle_group_list(particle_group_index).live > (particle_group_list(particle_group_index).liv1 * 25) And Not particle_group_list(particle_group_index).liv1 = -1 Then detroy = 1
    If detroy = 1 Then
        Particle_Group_Destroy particle_group_index
        Exit Sub
    End If
        
    With particle_group_list(particle_group_index)
        'Set colors
        temp_rgb(0) = .rgb_list(0)
        temp_rgb(1) = .rgb_list(1)
        temp_rgb(2) = .rgb_list(2)
        temp_rgb(3) = .rgb_list(3)
        
        'See if it is time to move a particle
        .frame_counter = .frame_counter + timerTicksPerFrame
        If .frame_counter > .frame_speed Then
            .frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
            
        'If it's still alive render all the particles inside
        For loopc = 1 To .particle_count
                
        'Render particle
            Particle_Render .particle_stream(loopc), _
                        screen_x, screen_y, _
                        .grh_index_list(Round(General_Random_Number(1, .grh_index_count), 0)), _
                        temp_rgb(), _
                        .alpha_blend, no_move, _
                        .x1, .y1, .angle, _
                        .vecx1, .vecx2, _
                        .vecy1, .vecy2, _
                        .life1, .life2, _
                        .fric, .spin_speedL, _
                        .gravity, .grav_strength, _
                        .bounce_strength, .x2, _
                        .y2, .XMove, _
                        .move_x1, .move_x2, _
                        .move_y1, .move_y2, _
                        .YMove, .spin_speedH, _
                        .spin, .grh_resize, .grh_resizex, .grh_resizey
        Next loopc
                
    If no_move = False Then
        'Update the group alive counter
        If .never_die = False Then
            .alive_counter = .alive_counter - 1
        End If
    End If
    
        
    End With
End Sub


Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_y As Integer, _
                            ByVal grh_index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                            Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                            Optional grh_resizex As Integer, Optional grh_resizey As Integer)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'**************************************************************
    If no_move = False Then
        If temp_particle.alive_counter = 0 Then
            'Start new particle
            Grh_Load temp_particle.Grh, grh_index, alpha_blend
            temp_particle.X = General_Random_Number(x1, x2)
            temp_particle.Y = General_Random_Number(y1, y2)
            temp_particle.vector_x = General_Random_Number(vecx1, vecx2)
            temp_particle.vector_y = General_Random_Number(vecy1, vecy2)
            temp_particle.angle = angle
            temp_particle.alive_counter = General_Random_Number(life1, life2)
            temp_particle.friction = fric
        Else
            'Continue old particle
            'Do gravity
            If gravity = True Then
                temp_particle.vector_y = temp_particle.vector_y + grav_strength
                If temp_particle.Y > 0 Then
                    'bounce
                    temp_particle.vector_y = bounce_strength
                End If
            End If
            'Do rotation
            If spin = True Then temp_particle.Grh.angle = temp_particle.Grh.angle + (General_Random_Number(spin_speedL, spin_speedH) / 100)
            If temp_particle.angle >= 360 Then
                temp_particle.angle = 0
            End If
            
            If XMove = True Then temp_particle.vector_x = General_Random_Number(move_x1, move_x2)
            If YMove = True Then temp_particle.vector_y = General_Random_Number(move_y1, move_y2)
        End If
        
        'Add in vector
        temp_particle.X = temp_particle.X + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.vector_y \ temp_particle.friction)
    
        'decrement counter
         temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If
    
    'Draw it
    If grh_resize = True Then
        If temp_particle.Grh.grhindex Then
            Grh_Render_Advance temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, grh_resizex, grh_resizey, rgb_list(), True, True, alpha_blend
            Exit Sub
        End If
    End If
    If temp_particle.Grh.grhindex Then
        Grh_Render temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, rgb_list(), True, True, alpha_blend
    End If
End Sub
Private Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
                                
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(particle_group_index).active = True
    
    'Map pos
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(particle_group_index).map_x = map_x
        particle_group_list(particle_group_index).map_y = map_y
    End If
    
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).liv1 = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).liv1 = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
    
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
    
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
    
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
    
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
    
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
    
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
    
    'handle
    particle_group_list(particle_group_index).id = id
    
    particle_group_list(particle_group_index).live = GetTickCount()
    
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'plot particle group on map
    If (map_x <> -1) And (map_y <> -1) Then
        MapData(map_x, map_y).particle_group_index = particle_group_index
    End If
    
End Sub
Public Function Particle_Map_Group_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a particle_group_index and return it
'*****************************************************************
    If Map_Bounds(map_x, map_y) Then
        Particle_Map_Group_Get = MapData(map_x, map_y).particle_group_index
    Else
        Particle_Map_Group_Get = 0
    End If
End Function
Public Function General_Char_Particle_Create(ByVal ParticulaInd As Long, ByVal char_index As Integer, Optional ByVal particle_life As Long = 0) As Long

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

General_Char_Particle_Create = TileEngine.Char_Particle_Group_Create(char_index, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, _
    StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey)

End Function

Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal particle_life As Long = 0) As Long

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

General_Particle_Create = TileEngine.Particle_Group_Create(X, Y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, _
    StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey)

End Function
Public Function Char_Particle_Group_Create(ByVal char_index As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim char_part_free_index As Integer
    
    'If Char_Particle_Group_Find(char_index, stream_type) Then Exit Function ' hay que ver si dejar o sacar esto...
    If Not Char_Check(char_index) Then Exit Function
    char_part_free_index = Char_Particle_Group_Next_Open(char_index)
    
    If char_part_free_index > 0 Then
        Char_Particle_Group_Create = Particle_Group_Next_Open
        Char_Particle_Group_Make Char_Particle_Group_Create, char_index, char_part_free_index, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    End If

End Function

Public Function Char_Particle_Group_Remove(ByVal char_index As Integer, ByVal stream_type As Long)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim char_part_index As Integer
    
    If Char_Check(char_index) Then
        char_part_index = Char_Particle_Group_Find(char_index, stream_type)
        If char_part_index = -1 Then Exit Function
        Call Particle_Group_Remove(char_part_index)
    End If

End Function

Public Function Char_Particle_Group_Remove_All(ByVal char_index As Integer)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim i As Integer
    
    If Char_Check(char_index) And Not charlist(char_index).particle_count = 0 Then
        For i = 1 To UBound(charlist(char_index).particle_group)
            If charlist(char_index).particle_group(i) <> 0 Then Call Particle_Group_Remove(charlist(char_index).particle_group(i))
        Next i
        Erase charlist(char_index).particle_group
        charlist(char_index).particle_count = 0
    End If
    
End Function

Private Function Char_Particle_Group_Find(ByVal char_index As Integer, ByVal stream_type As Long) As Integer
'*****************************************************************
'Author: Augusto José Rando
'Modified: returns slot or -1
'*****************************************************************

Dim i As Integer

For i = 1 To charlist(char_index).particle_count
    If charlist(char_index).particle_group(i) <> 0 Then
        If particle_group_list(charlist(char_index).particle_group(i)).stream_type = stream_type Then
            Char_Particle_Group_Find = charlist(char_index).particle_group(i)
            Exit Function
        End If
    End If
Next i

Char_Particle_Group_Find = -1

End Function

Private Function Char_Particle_Group_Next_Open(ByVal char_index As Integer) As Integer
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    If charlist(char_index).particle_count = 0 Then
        Char_Particle_Group_Next_Open = charlist(char_index).particle_count + 1
        charlist(char_index).particle_count = Char_Particle_Group_Next_Open
        ReDim Preserve charlist(char_index).particle_group(1 To Char_Particle_Group_Next_Open) As Long
        Exit Function
    End If
    
    loopc = 1
    Do Until charlist(char_index).particle_group(loopc) = 0
        If loopc = charlist(char_index).particle_count Then
            Char_Particle_Group_Next_Open = charlist(char_index).particle_count + 1
            charlist(char_index).particle_count = Char_Particle_Group_Next_Open
            ReDim Preserve charlist(char_index).particle_group(1 To Char_Particle_Group_Next_Open) As Long
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Char_Particle_Group_Next_Open = loopc

Exit Function

ErrorHandler:
    charlist(char_index).particle_count = 1
    ReDim charlist(char_index).particle_group(1 To 1) As Long
    Char_Particle_Group_Next_Open = 1

End Function
Private Function Char_Check(ByVal char_index As Integer) As Boolean
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (charlist(char_index).heading > 0)
    End If
End Function
Private Sub Char_Particle_Group_Make(ByVal particle_group_index As Long, ByVal char_index As Integer, ByVal particle_char_index As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
                                
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(particle_group_index).active = True
    
    'Char index
    particle_group_list(particle_group_index).char_index = char_index
    
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).liv1 = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).liv1 = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
    
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
    
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
    
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
    
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
    
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
    
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
    
    'handle
    particle_group_list(particle_group_index).id = id
    particle_group_list(particle_group_index).live = GetTickCount()
    
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'plot particle group on char
    charlist(char_index).particle_group(particle_char_index) = particle_group_index
End Sub

Private Sub Grh_Render_Advance(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal height As Integer, ByVal width As Integer, ByRef rgb_list() As Long, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 11/19/2003
'Similar to Grh_Render, but let´s you resize the Grh
'**************************************************************
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Long
    
    'Animation
    If Grh.Started Then
        Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.speed)
        If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
            Grh.FrameCounter = 1
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    grh_index = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)
    
    'Draw it to device
    Device_Box_Textured_Render_Advance grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        width, height, alpha_blend, Grh.angle
End Sub
Private Sub Grh_Render(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'Modified by Juan Martín Sotuyo Dodero
'Added centering
'**************************************************************
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Long
    
    If Grh.grhindex = 0 Then Exit Sub
        
    'Animation
    If Grh.Started Then
        Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.speed)
        If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
            Grh.FrameCounter = 1
        End If
    End If


    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    
    grh_index = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)
    
    If grh_index <= 0 Then Exit Sub
    If GrhData(grh_index).FileNum = 0 Then Exit Sub
        
    'Modified by Augusto José Rando
    'Simplier function - according to basic ORE engine
    If h_centered Then
        If GrhData(Grh.grhindex).TileWidth <> 1 Then
            screen_x = screen_x - Int(GrhData(Grh.grhindex).TileWidth * (32 \ 2)) + 32 \ 2
        End If
    End If
    
    If v_centered Then
        If GrhData(Grh.grhindex).TileHeight <> 1 Then
            screen_y = screen_y - Int(GrhData(Grh.grhindex).TileHeight * 32) + 32
        End If
    End If
    
    'Draw it to device
    Device_Box_Textured_Render grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list(), _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        alpha_blend, _
        Grh.angle

End Sub
Public Sub Char_Account_Render(ByVal Index As Byte)
    Dim cColor As Long
    
    Index = Index - 1
    If cPJ(Index).nombre <> "" Then
        frmPanelAccount.lblAccData(Index + 1).Caption = cPJ(Index).nombre
        Select Case cPJ(Index).color
            Case 1 'Gris
                cColor = RGB(175, 175, 175)
            Case 2 'Azul
                cColor = RGB(39, 131, 243)
            Case 3 'Rojo
                cColor = RGB(217, 0, 5)
            Case 4 'Naranja
                cColor = RGB(243, 147, 1)
            Case 5 'Verde
                cColor = RGB(0, 142, 72)
        End Select
        frmPanelAccount.lblAccData(Index + 1).ForeColor = cColor
    End If
    If cPJ(Index).nombre = "" Then Exit Sub
    
    On Error Resume Next
    
    Dim i As Integer
    
    Dim init_x As Integer
    Dim init_y As Integer
    Dim tempito(3) As Long
    Dim grhtemp As Grh
    Static re As RECT
    
    re.Left = 0
    re.Top = 0
    re.bottom = 80
    re.Right = 76
    
    tempito(0) = &HFFFFFFFF
    tempito(1) = &HFFFFFFFF
    tempito(2) = &HFFFFFFFF
    tempito(3) = &HFFFFFFFF
    
    If cPJ(Index).body = 8 Then
        cPJ(Index).Head = 500
    End If
    
    If cPJ(Index).tipPet > 0 Then
        init_x = 10
        init_y = 40
    Else
        init_x = 25
        init_y = 40
    End If
   
    D3DDevice.BeginScene
    D3DDevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, 0, 0, 0

        If cPJ(Index).body = 255 Then
            Draw_Grh_Index 20206, init_x - 48, init_y - 70, D3DColorXRGB(255, &HFD, &H7E)
        End If
    
If cPJ(Index).body <> iBarca And cPJ(Index).body <> iGalera And cPJ(Index).body <> iGaleon And cPJ(Index).body <> iFragataFantasmal Then
        If cPJ(Index).Shield = 32 Then
            Draw_Grh_Index 20128, init_x - 48, init_y - 70, D3DColorXRGB(255, 0, 0)
        End If

        If cPJ(Index).Weapon = 16 Then
            Draw_Grh_Index 20128, init_x - 48, init_y - 70, D3DColorXRGB(255, &HCC, 51) ' &HFFCC33
        ElseIf cPJ(Index).Weapon = 17 Then
            Draw_Grh_Index 20133, init_x - 45, init_y - 70, D3DColorXRGB(255, 51, 0)   '&HFF3300
        ElseIf cPJ(Index).Weapon = 23 Then
            Draw_Grh_Index 20152, init_x - 48, init_y - 70, D3DColorXRGB(255, 0, 0)
        ElseIf cPJ(Index).Weapon = 24 Then
            Draw_Grh_Index 20185, init_x - 48, init_y - 70, -65536
        ElseIf cPJ(Index).Weapon = 25 Then
            Draw_Grh_Index 20155, init_x - 48, init_y - 70, D3DColorXRGB(255, 0, 0)
        ElseIf cPJ(Index).Weapon = 26 Then
            Draw_Grh_Index 20151, init_x - 48, init_y - 70, D3DColorXRGB(255, 255, 0)
        ElseIf cPJ(Index).Weapon = 27 Then
            Draw_Grh_Index 20147, init_x - 48, init_y - 70, D3DColorXRGB(0, 0, 255)
        ElseIf cPJ(Index).Weapon = 28 Then
            Draw_Grh_Index 20146, init_x - 48, init_y - 70, D3DColorXRGB(0, &H6B, &H1B)
        ElseIf cPJ(Index).Weapon = 29 Then
            Draw_Grh_Index 20200, init_x - 48, init_y - 70, D3DColorXRGB(&HCC, 255, 51)
        End If
End If
        If cPJ(Index).tipPet > 0 Then
            grhtemp.grhindex = getGrhPet(cPJ(Index).tipPet)
            
            Grh_Render grhtemp, init_x + 30, init_y, tempito
        End If

        If cPJ(Index).body <> 0 Then
            grhtemp.grhindex = BodyData(cPJ(Index).body).Walk(3).grhindex
            Grh_Render grhtemp, init_x, init_y, tempito
        End If



If cPJ(Index).body <> iBarca And cPJ(Index).body <> iGalera And cPJ(Index).body <> iGaleon And cPJ(Index).body <> iFragataFantasmal Then
        If cPJ(Index).Head <> 0 Then
            grhtemp.grhindex = HeadData(cPJ(Index).Head).Head(3).grhindex
            Grh_Render grhtemp, init_x + BodyData(cPJ(Index).body).HeadOffset.X, init_y + BodyData(cPJ(Index).body).HeadOffset.Y, tempito
        End If
        
        If cPJ(Index).Casco <> 0 Then
            grhtemp.grhindex = CascoAnimData(cPJ(Index).Casco).Head(3).grhindex
            Grh_Render grhtemp, init_x + BodyData(cPJ(Index).body).HeadOffset.X, init_y + BodyData(cPJ(Index).body).HeadOffset.Y, tempito
        End If
        
        If cPJ(Index).Weapon <> 0 Then
            grhtemp.grhindex = WeaponAnimData(cPJ(Index).Weapon).WeaponWalk(3).grhindex
            Grh_Render grhtemp, init_x, init_y, tempito
        End If
        
        If cPJ(Index).Shield <> 0 Then
            grhtemp.grhindex = ShieldAnimData(cPJ(Index).Shield).ShieldWalk(3).grhindex
            Grh_Render grhtemp, init_x, init_y, tempito
        End If
End If

    D3DDevice.EndScene
    D3DDevice.Present re, ByVal 0&, frmPanelAccount.picChar(Index).hwnd, ByVal 0&
    
End Sub

Private Sub Engine_Alpha_Calculate(roofrgb_list() As Long)
    Dim color As D3DCOLORVALUE
    Static last_tick As Long
    If UserPos.X = 0 Or UserPos.Y = 0 Then Exit Sub
        If IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 7 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4 Or _
                MapData(UserPos.X, UserPos.Y).Trigger >= 20, True, False) Then
            If AlphaY > 0 Then
                If GetTickCount - last_tick >= 18 Then
                    AlphaY = AlphaY - 5
                    last_tick = GetTickCount
                End If
            End If
        Else
            If AlphaY < 255 Then
                If GetTickCount - last_tick >= 18 Then
                    AlphaY = AlphaY + 5
                    last_tick = GetTickCount
                End If
            End If
        End If
        
        color = ambientLight
        color.a = AlphaY
        Engine_D3DColorToRgbList roofrgb_list(), color
End Sub
Private Sub Char_Alpha_Calculate(roofrgb_list() As Long, ByVal CharIndex As Integer)
    Dim color As D3DCOLORVALUE
    Dim suma As Double
    suma = (timerTicksPerFrame * 6.7)
    
    If charlist(CharIndex).AlphaX < 30 Then
        If suma > 0 Then
            suma = suma / 10
        End If
    End If
    
    With charlist(CharIndex)
        If .state Then
            If .AlphaX > 0 Then
                If .AlphaX - suma <= 0 Then
                    .AlphaX = 0
                Else
                    .AlphaX = .AlphaX - suma
                End If
                .last_tick = GetTickCount
            Else
                .state = 0
                .AlphaX = 0
            End If
        Else
            If .AlphaX < 255 Then
                If .AlphaX + suma > 255 Then
                    .AlphaX = 255
                Else
                    .AlphaX = .AlphaX + suma
                End If
                .last_tick = GetTickCount
            Else
                .state = 1
                .AlphaX = 255
            End If
        End If
    End With
    
    color = ambientLight
    color.a = charlist(CharIndex).AlphaX
    Engine_D3DColorToRgbList roofrgb_list(), color
End Sub
Public Sub Engine_D3DColorToRgbList(rgb_list() As Long, color As D3DCOLORVALUE)
    rgb_list(0) = D3DColorARGB(color.a, color.r, color.g, color.b)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Public Sub Engine_Long_To_RGB_List(rgb_list() As Long, long_color As Long)
    rgb_list(0) = long_color
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub
Function Text_Width(Texto As String, ByVal font_index As Integer, Optional multi As Boolean = False) As Integer
On Error Resume Next
Dim a As Integer, b As Integer, d As Integer, e As Integer, f As Integer
Dim graf As Grh

If multi = False Then
    For a = 1 To Len(Texto)
        b = Asc(mid(Texto, a, 1))
        graf.grhindex = font_types(1).ascii_code(b)
        If (b <> 32) And (b <> 5) And (b <> 129) And (b <> 9) And (b <> 4) And (b <> 255) And (b <> 2) And graf.grhindex <> 0 Then
            Text_Width = Text_Width + GrhData(GrhData(graf.grhindex).Frames(1)).pixelWidth '+ 1
        Else
            Text_Width = Text_Width + 4
        End If
    Next a
Else
    e = 0
    f = 0
    For a = 1 To Len(Texto)
        b = Asc(mid(Texto, a, 1))
        graf.grhindex = font_types(1).ascii_code(b)
        If b = 32 Or b = 13 Then
            If e >= 20 Then 'reemplazar por lo que os plazca
                f = f + 1
                e = 0
                d = 0
            Else
                If b = 32 Then d = d + 4
            End If
        Else
            If graf.grhindex > 12 Then
                d = d + GrhData(GrhData(graf.grhindex).Frames(1)).pixelWidth '+ 1
                If d > Text_Width Then Text_Width = d
            End If
        End If
        e = e + 1
    Next a
End If

End Function
Public Function Light_Remove(ByVal light_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        Light_Destroy light_index
        Light_Remove = True
    End If
End Function
Public Function Light_Remove_From_Pos(ByVal posX As Byte, ByVal posY As Byte, Optional vColor As Long = 0) As Boolean
    Dim i As Long
    
    For i = 1 To light_last
        If light_list(i).map_x = posX Then
            If light_list(i).map_y = posY Then
                If vColor <> 0 Then
                    If light_list(i).color = vColor Then
                        If Light_Check(i) Then
                            Light_Destroy i
                        End If
                    End If
                Else
                    If Light_Check(i) Then
                        Light_Destroy i
                    End If
                End If
            End If
        End If
    Next i
    
    Light_Render_All
End Function

Public Function Light_Color_Value_Get(ByVal light_index As Long, ByRef color_value As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        color_value = light_list(light_index).color
        Light_Color_Value_Get = True
    End If
End Function

Public Function Light_Create(ByVal map_x As Integer, ByVal map_y As Integer, Optional ByVal color_value As Long = &HFFFFFF, _
                            Optional ByVal range As Byte = 1, Optional ByVal id As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns the light_index if successful, else 0
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
    If Map_Bounds(map_x, map_y) Then
        'Make sure there is no light in the given map pos
       ' If Light_Map_Get(map_x, map_y) <> 0 Then
        '    Light_Create = 0
       '     Exit Function
       ' End If
        Light_Create = Light_Next_Open
        
        Dim r As Integer, g As Integer, b As Integer
        General_Long_Color_to_RGB color_value, r, g, b
        color_value = D3DColorXRGB(r, g, b)
        
        Light_Make Light_Create, map_x, map_y, color_value, range, id
    End If
End Function
Public Function Light_Map_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a light_index and return it
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    'We start from the back, to get the last light to be placed on the tile first
    If light_last = 0 Then Exit Function
    
    loopc = light_last
    Do Until light_list(loopc).map_x = map_x And light_list(loopc).map_y = map_y
        If loopc = 0 Then
            Light_Map_Get = 0
            Exit Function
        End If
        loopc = loopc - 1
        If loopc = 0 Then Exit Function
    Loop
    
    Light_Map_Get = loopc
Exit Function
ErrorHandler:
    Light_Map_Get = 0
End Function
Public Function Light_Move(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
        'Make sure it's a legal move
        If Map_Bounds(map_x, map_y) Then
        
            'Move it
            Light_Erase light_index
            light_list(light_index).map_x = map_x
            light_list(light_index).map_y = map_y
    
            Light_Move = True
            
        End If
    End If
End Function

Public Function Light_Move_By_Head(ByVal light_index As Long, ByVal heading As Byte) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 15/05/2002
'Returns true if successful, else false
'**************************************************************
    Dim map_x As Integer
    Dim map_y As Integer
    Dim nx As Integer
    Dim ny As Integer
    
    'Check for valid heading
    If heading < 1 Or heading > 8 Then
        Light_Move_By_Head = False
        Exit Function
    End If

    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
    
        map_x = light_list(light_index).map_x
        map_y = light_list(light_index).map_y
        
        nx = map_x
        ny = map_y
        
        'Convert_Heading_to_Direction heading, nX, nY
        
        'Make sure it's a legal move
        If Map_Bounds(nx, ny) Then
        
            'Move it
            Light_Erase light_index

            light_list(light_index).map_x = nx
            light_list(light_index).map_y = ny
    
            Light_Move_By_Head = True
            
        End If
    End If
End Function

Private Sub Light_Make(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, ByVal rgb_value As Long, _
                        ByVal range As Long, Optional ByVal id As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    'Update array size
    If light_index > light_last Then
        light_last = light_index
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count + 1
    
    'Make active
    light_list(light_index).active = True
    
    light_list(light_index).map_x = map_x
    light_list(light_index).map_y = map_y
    light_list(light_index).color = rgb_value
    light_list(light_index).range = range
    light_list(light_index).id = id
End Sub

Private Function Light_Check(ByVal light_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check light_index
    If light_index > 0 And light_index <= light_last Then
        If light_list(light_index).active Then
            Light_Check = True
        End If
    End If
End Function

Public Sub Light_Render_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim loop_counter As Long
            
    For loop_counter = 1 To light_count
        
        If light_list(loop_counter).active Then
            Light_Render loop_counter
        End If
    
    Next loop_counter
End Sub

Private Sub Light_Render(ByVal light_index As Long)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 11/11/2004
'Renders a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim color As Long
    Dim light_trigger As Integer
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    light_trigger = MapData(light_list(light_index).map_x, light_list(light_index).map_y).Trigger
    
    'Set color
    color = light_list(light_index).color
    
    'Arrange corners
    'NE
    If Map_Bounds(min_x, min_y) Then
        If MapData(min_x, min_y).Trigger = light_trigger Then _
            MapData(min_x, min_y).light_value(2) = color
    End If
    'NW
    If Map_Bounds(max_x, min_y) Then
        If MapData(max_x, min_y).Trigger = light_trigger Then _
            MapData(max_x, min_y).light_value(0) = color
    End If
    'SW
    If Map_Bounds(max_x, max_y) Then
        If MapData(max_x, max_y).Trigger = light_trigger Then _
            MapData(max_x, max_y).light_value(1) = color
    End If
    'SE
    If Map_Bounds(min_x, max_y) Then
        If MapData(min_x, max_y).Trigger = light_trigger Then _
            MapData(min_x, max_y).light_value(3) = color
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If Map_Bounds(X, min_y) Then
            If MapData(X, min_y).Trigger = light_trigger Then
                MapData(X, min_y).light_value(0) = color
                MapData(X, min_y).light_value(2) = color
            End If
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If Map_Bounds(X, max_y) Then
            If MapData(X, max_y).Trigger = light_trigger Then
                MapData(X, max_y).light_value(1) = color
                MapData(X, max_y).light_value(3) = color
            End If
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If Map_Bounds(min_x, Y) Then
            If MapData(min_x, Y).Trigger = light_trigger Then
                MapData(min_x, Y).light_value(2) = color
                MapData(min_x, Y).light_value(3) = color
            End If
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If Map_Bounds(max_x, Y) Then
            If MapData(max_x, Y).Trigger = light_trigger Then
                MapData(max_x, Y).light_value(0) = color
                MapData(max_x, Y).light_value(1) = color
            End If
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If Map_Bounds(X, Y) Then
                If MapData(X, Y).Trigger = light_trigger Then
                    MapData(X, Y).light_value(0) = color
                    MapData(X, Y).light_value(1) = color
                    MapData(X, Y).light_value(2) = color
                    MapData(X, Y).light_value(3) = color
                End If
            End If
        Next Y
    Next X
    
End Sub

Private Function Light_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    If light_last = 0 Then
        Light_Next_Open = 1
        Exit Function
    End If
    
    loopc = 1
    Do Until light_list(loopc).active = False
        If loopc = light_last Then
            Light_Next_Open = light_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Next_Open = loopc

Exit Function

ErrorHandler:

End Function

Public Function Light_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).id = id
        If loopc = light_last Then
            Light_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Find = loopc
Exit Function
ErrorHandler:
    Light_Find = 0
End Function

Public Function Light_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To light_last
        'Make sure it's a legal index
        If Light_Check(Index) Then
            Light_Destroy Index
        End If
    Next Index
    
    Light_Remove_All = True
End Function

Private Sub Light_Destroy(ByVal light_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As Light
    
    Light_Erase light_index
    
    light_list(light_index) = temp
    
    'Update array size
    If light_index = light_last Then
        Do Until light_list(light_last).active
            light_last = light_last - 1
            If light_last = 0 Then
                light_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count - 1
End Sub

Private Sub Light_Erase(ByVal light_index As Long)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 3/31/2003
'Correctly erases a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim light_trigger As Integer
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    light_trigger = MapData(light_list(light_index).map_x, light_list(light_index).map_y).Trigger
    
    'Arrange corners
    'NE
    If Map_Bounds(min_x, min_y) Then
        If MapData(min_x, min_y).Trigger = light_trigger Then _
            MapData(min_x, min_y).light_value(2) = AmbientColor
    End If
    'NW
    If Map_Bounds(max_x, min_y) Then
        If MapData(max_x, min_y).Trigger = light_trigger Then _
            MapData(max_x, min_y).light_value(0) = AmbientColor
    End If
    'SW
    If Map_Bounds(max_x, max_y) Then
        If MapData(max_x, max_y).Trigger = light_trigger Then _
            MapData(max_x, max_y).light_value(1) = AmbientColor
    End If
    'SE
    If Map_Bounds(min_x, max_y) Then
        If MapData(min_x, max_y).Trigger = light_trigger Then _
            MapData(min_x, max_y).light_value(3) = AmbientColor
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If Map_Bounds(X, min_y) Then
            If MapData(X, min_y).Trigger = light_trigger Then
                MapData(X, min_y).light_value(0) = AmbientColor
                MapData(X, min_y).light_value(2) = AmbientColor
            End If
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If Map_Bounds(X, max_y) Then
            If MapData(X, max_y).Trigger = light_trigger Then
                MapData(X, max_y).light_value(1) = AmbientColor
                MapData(X, max_y).light_value(3) = AmbientColor
            End If
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If Map_Bounds(min_x, Y) Then
            If MapData(min_x, Y).Trigger = light_trigger Then
                MapData(min_x, Y).light_value(2) = AmbientColor
                MapData(min_x, Y).light_value(3) = AmbientColor
            End If
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If Map_Bounds(max_x, Y) Then
            If MapData(max_x, Y).Trigger = light_trigger Then
                MapData(max_x, Y).light_value(0) = AmbientColor
                MapData(max_x, Y).light_value(1) = AmbientColor
            End If
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If Map_Bounds(X, Y) Then
                If MapData(X, Y).Trigger = light_trigger Then
                    MapData(X, Y).light_value(0) = AmbientColor
                    MapData(X, Y).light_value(1) = AmbientColor
                    MapData(X, Y).light_value(2) = AmbientColor
                    MapData(X, Y).light_value(3) = AmbientColor
                End If
            End If
        Next Y
    Next X
End Sub
Public Function Meteo_Change_Time()
'**************************************************************
'Author: Leandro Mendoza (Mannakia)
'Last Modify Date: 19/09/2010
'Change the meteo time for start the animation with alphacolor
'**************************************************************
Dim tmpCartel As Byte

If meteo_hour = tHora Then
    If Not AmbientColor = -1 Then Exit Function
End If

meteo_hour = tHora
If tHora >= 5 And tHora <= 7 Then
    meteo_color = m_Color_Manana
ElseIf tHora >= 8 And tHora <= 17 Then
    meteo_color = m_Color_Dia
ElseIf tHora = 18 Or tHora = 19 Then
    meteo_color = m_Color_Tarde
Else
    meteo_color = m_Color_Noche
End If


frmMain.imgHora.Picture = LoadInterface(IIf(tHora <= 9, CStr("0" & tHora), CStr(tHora)))

If ambientLight.r > meteo_color.r Then
    meteo_state = 1 'Animacion Desendiente
Else
    meteo_state = 2 'Animacion Asendente
End If
End Function
Public Function Get_Time_String() As String

    Get_Time_String = tHora & ":" & Format(tMinuto, "00")
    
    Select Case tHora
        Case 5, 6, 7
            Get_Time_String = Get_Time_String & " el sol se asoma lentamente en el horizonte"
        Case 8, 9, 10, 11, 12, 13, 14, 15, 16, 17
            Get_Time_String = Get_Time_String & " ¡no pierdas el tiempo!"
        Case 18, 19
            Get_Time_String = Get_Time_String & " lentamente el dia termina"
        Case Else
            Get_Time_String = Get_Time_String & " ¿despierto a estas horas? ¡no olvides visitar El Mesón Hostigado!"
    End Select
End Function

Private Function Meteo_Render()
'**************************************************************
'Author: Leandro Mendoza (Mannakia)
'Last Modify Date: 19/09/2010
'Rendering the animation with desvan
'**************************************************************
Dim change As Boolean

If meteo_state = 0 Then Exit Function
If m_Afecta = False Then Exit Function

Select Case meteo_state
    Case 1
        If ambientLight.r > meteo_color.r Then
            ambientLight.r = ambientLight.r - 1
            change = True
        End If
        
        If ambientLight.g > meteo_color.g Then
            ambientLight.g = ambientLight.g - 1
            change = True
        End If
        
        If ambientLight.b > meteo_color.b Then
            ambientLight.b = ambientLight.b - 1
            change = True
        End If
        
        If change = False Then meteo_state = 0
        AmbientColor = D3DColorXRGB(ambientLight.r, ambientLight.g, ambientLight.b)
        
    Case 2
        If ambientLight.r < meteo_color.r Then
            ambientLight.r = ambientLight.r + 1
            change = True
        End If
        
        If ambientLight.g < meteo_color.g Then
            ambientLight.g = ambientLight.g + 1
            change = True
        End If
        
        If ambientLight.b < meteo_color.b Then
            ambientLight.b = ambientLight.b + 1
            change = True
        End If
        
        If change = False Then meteo_state = 0
        AmbientColor = D3DColorXRGB(ambientLight.r, ambientLight.g, ambientLight.b)
        
        
End Select
End Function
Public Function Meteo_Init_Time()
'**************************************************************
'Author: Leandro Mendoza (Mannakia)
'Last Modify Date: 19/09/2010
'Start all shades of the day
'**************************************************************
With m_Color_Dia
    .a = 255
    .b = 255
    .r = 255
    .g = 255
End With

With m_Color_Noche
    .a = 255
    .b = 170
    .r = 170
    .g = 170
End With

With m_Color_Tarde
    .a = 255
    .b = 200
    .r = 230
    .g = 200
End With

With m_Color_Manana
    .a = 255
    .b = 230
    .r = 200
    .g = 200
End With

meteo_hour = -1
meteo_state = -1
tCartel = -1
End Function
Public Function Meteo_Clean()
    meteo_hour = -1
    meteo_state = -1
    tCartel = -1
End Function
Sub Char_Pasos_Render(ByVal CharIndex As Integer)
    Dim paso As Byte
    
    If Not UserNavegando Then
        With charlist(CharIndex)
            If .Muerto = False And Char_Is_Area(CharIndex) = True And (.Priv <> 7) And (.Priv <> 8) Then
                .pie = Not .pie
                paso = Map_GetTerrenoDePaso(GrhData(MapData(.Pos.X, .Pos.Y).Graphic(1).grhindex).FileNum)
                
                If paso = 1 Then
                    If .pie Then
                        Call Audio.PlayWave(201, .Pos.X, .Pos.Y)
                    Else
                        Call Audio.PlayWave(202, .Pos.X, .Pos.Y)
                    End If
                ElseIf paso = 2 Or paso = 5 Then
                    If .pie Then
                        Call Audio.PlayWave(23, .Pos.X, .Pos.Y)
                    Else
                        Call Audio.PlayWave(24, .Pos.X, .Pos.Y)
                    End If
                ElseIf paso = 3 Then
                    If .pie Then
                        Call Audio.PlayWave(199, .Pos.X, .Pos.Y)
                    Else
                        Call Audio.PlayWave(200, .Pos.X, .Pos.Y)
                    End If
                ElseIf paso = 4 Then
                    If .pie Then
                        Call Audio.PlayWave(197, .Pos.X, .Pos.Y)
                    Else
                        Call Audio.PlayWave(198, .Pos.X, .Pos.Y)
                    End If
                End If
            End If
        End With
    ElseIf UserMontando = True Then
        Call Audio.PlayWave(23, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y)
    ElseIf UserNavegando = True Then
        'Call Audio.PlayWave(50, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y)
    End If
End Sub
Private Function Map_GetTerrenoDePaso(ByVal TerrainFileNum As Integer) As Byte
    If (TerrainFileNum >= 6000 And TerrainFileNum <= 6004) Or (TerrainFileNum >= 550 And TerrainFileNum <= 552) Or (TerrainFileNum >= 6018 And TerrainFileNum <= 6020) Then
        Map_GetTerrenoDePaso = 1
        Exit Function
    ElseIf (TerrainFileNum >= 7501 And TerrainFileNum <= 7507) Or (TerrainFileNum = 7500 Or TerrainFileNum = 7508 Or TerrainFileNum = 1533 Or TerrainFileNum = 2508) Then
        Map_GetTerrenoDePaso = 2
        Exit Function
    ElseIf (TerrainFileNum >= 5000 And TerrainFileNum <= 5004) Then
        Map_GetTerrenoDePaso = 3
        Exit Function
    ElseIf TerrainFileNum = 6021 Then
        Map_GetTerrenoDePaso = 4
        Exit Function
    Else
        Map_GetTerrenoDePaso = 5
    End If
End Function
Public Sub Engine_Convert_CP_To_TP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tX = UserPos.X + viewPortX \ 32 - frmMain.MainViewPic.ScaleWidth \ 64
    tY = UserPos.Y + viewPortY \ 32 - frmMain.MainViewPic.ScaleHeight \ 64

End Sub

Private Sub Grh_Load(ByRef Grh As Grh, ByVal grhindex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.grhindex = grhindex
    
    If GrhData(Grh.grhindex).NumFrames > 1 Then
        Grh.Started = 1
        Grh.speed = 0.4
    Else
        Grh.Started = 0
    End If

    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    
End Sub
Private Sub Grh_Load_TuTv(ByVal grhindex As Integer, ByRef src As RECT, ByVal H As Integer, ByVal W As Integer)
    With GrhData(grhindex)
        If H <> 0 Then
            .tu(0) = src.Left / W
            .tv(0) = (src.bottom + 1) / H
            .tu(1) = src.Left / W
            .tv(1) = src.Top / H
            .tu(2) = (src.Right + 1) / W
            .tv(2) = (src.bottom + 1) / H
            .tu(3) = (src.Right + 1) / W
            .tv(3) = src.Top / H
            .hardcor = 1
        End If
    End With
End Sub
Sub Engine_Load_Bodies()
    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    Extract_File Scripts, "personajes.ind", resource_path
    
    N = FreeFile()
    Open resource_path & "personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).body(1) Then
            Grh_Load BodyData(i).Walk(1), MisCuerpos(i).body(1), 0
            Grh_Load BodyData(i).Walk(2), MisCuerpos(i).body(2), 0
            Grh_Load BodyData(i).Walk(3), MisCuerpos(i).body(3), 0
            Grh_Load BodyData(i).Walk(4), MisCuerpos(i).body(4), 0
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
    
    'Delete_File resource_path & "Personajes.ind"
    If FileExist(resource_path & "Personajes.ind", vbNormal) Then Kill resource_path & "Personajes.ind"
End Sub
Sub Engine_Load_Heads()
    Dim N As Integer
    Dim i As Long
    Dim NumHeads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    'Delete_File Resource_path
    'If FileExist(Resource_path, vbNormal) Then Kill Resource_path
    
    Extract_File Scripts, "cabezas.ind", resource_path
    N = FreeFile()
    Open resource_path & "cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumHeads
    
    'Resize array
    ReDim HeadData(0 To NumHeads) As HeadData
    ReDim Miscabezas(0 To NumHeads) As tIndiceCabeza
    
    For i = 1 To NumHeads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call Grh_Load(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call Grh_Load(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call Grh_Load(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call Grh_Load(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
    
    'Delete_File resource_path & "Cabezas.ind"
    If FileExist(resource_path & "Cabezas.ind", vbNormal) Then Kill resource_path & "Cabezas.ind"
End Sub
Sub Engine_Load_Helmet()
    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    Extract_File Scripts, "cascos.ind", resource_path
    
    N = FreeFile()
    Open resource_path & "Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call Grh_Load(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call Grh_Load(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call Grh_Load(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call Grh_Load(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
    
    'Delete_File resource_path & "Cascos.ind"
    If FileExist(resource_path & "Cascos.ind", vbNormal) Then Kill resource_path & "Cascos.ind"
    
End Sub
Sub Engine_Load_Fxs()
    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer
    
    Extract_File Scripts, "fxs.ind", resource_path
    
    N = FreeFile()
    Open resource_path & "fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
    
    'Delete_File resource_path & "Fxs.ind"
    If FileExist(resource_path & "Fx.ind", vbNormal) Then Kill resource_path & "Fxs.ind"
    
End Sub
Sub Engine_Load_Weapon()
On Error Resume Next

    Dim loopc As Long
    Dim Leer As New clsIniReader
    
    Extract_File Scripts, "armas.dat", resource_path
    
    Leer.Initialize resource_path & "armas.dat"
    
    NumWeaponAnims = val(Leer.GetValue("INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        Grh_Load WeaponAnimData(loopc).WeaponWalk(1), val(Leer.GetValue("ARMA" & loopc, "Dir1")), 0
        Grh_Load WeaponAnimData(loopc).WeaponWalk(2), val(Leer.GetValue("ARMA" & loopc, "Dir2")), 0
        Grh_Load WeaponAnimData(loopc).WeaponWalk(3), val(Leer.GetValue("ARMA" & loopc, "Dir3")), 0
        Grh_Load WeaponAnimData(loopc).WeaponWalk(4), val(Leer.GetValue("ARMA" & loopc, "Dir4")), 0
        Dim a As Integer
        Dim Index As Integer
        Index = loopc - 1
        a = WeaponAnimData(loopc).WeaponWalk(1).grhindex
        
    Next loopc
    
    Set Leer = Nothing
    
    'Delete_File resource_path & "armas.dat"
    If FileExist(resource_path & "armas.dat", vbNormal) Then Kill resource_path & "armas.dat"
    
End Sub
Sub Engine_Load_Shields()
On Error Resume Next

    Dim loopc As Long
    Dim Leer As New clsIniReader

    Extract_File Scripts, "escudos.dat", resource_path
    
    Leer.Initialize resource_path & "escudos.dat"
    
    NumEscudosAnims = val(Leer.GetValue("INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        Grh_Load ShieldAnimData(loopc).ShieldWalk(1), val(Leer.GetValue("ESC" & loopc, "Dir1")), 0
        Grh_Load ShieldAnimData(loopc).ShieldWalk(2), val(Leer.GetValue("ESC" & loopc, "Dir2")), 0
        Grh_Load ShieldAnimData(loopc).ShieldWalk(3), val(Leer.GetValue("ESC" & loopc, "Dir3")), 0
        Grh_Load ShieldAnimData(loopc).ShieldWalk(4), val(Leer.GetValue("ESC" & loopc, "Dir4")), 0
    Next loopc
    
    Set Leer = Nothing
    
    'Delete_File resource_path & "escudos.dat"
    If FileExist(resource_path & "escudos.dat", vbNormal) Then Kill resource_path & "escudos.dat"
    
End Sub

Public Sub Sound_Fogata_Fx()
    Dim location As Position
    
    If bFogata Then
        bFogata = Map_Have_Fogata(location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = Map_Have_Fogata(location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave(SND_FUEGO, location.X, location.Y, LoopStyle.Enabled)
    End If
End Sub

Public Function Char_Is_Area(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        Char_Is_Area = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function

Function Map_Bounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    Map_Bounds = True
End Function

Private Function Map_Have_Fogata(ByRef location As Position) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6
            If Map_Bounds(j, k) Then
                If MapData(j, k).ObjGrh.grhindex = 1521 Then
                    location.X = j
                    location.Y = k
                    
                    Map_Have_Fogata = True
                    Exit Function
                End If
            End If
        Next k
    Next j
End Function

Function Map_Legal_Pos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
    'Limites del mapa
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    If MapData(X, Y).CharIndex > 0 Then
        If Not charlist(MapData(X, Y).CharIndex).Muerto Then
            Exit Function
        End If
    End If
   
    If UserNavegando <> Map_Have_Water(X, Y) Then
        Exit Function
    End If
    
    If UserMontando = True Then
        If MapData(X, Y).Trigger = 1 Or MapData(X, Y).Trigger = 2 Or MapData(X, Y).Trigger = 4 Or MapData(X, Y).Trigger >= 20 Then
            Exit Function
        End If
    End If
    
    Map_Legal_Pos = True
End Function

Function Map_Have_Water(ByVal X As Integer, ByVal Y As Integer) As Boolean
    Map_Have_Water = ((MapData(X, Y).Graphic(1).grhindex >= 1505 And MapData(X, Y).Graphic(1).grhindex <= 1520) Or _
            (MapData(X, Y).Graphic(1).grhindex >= 5665 And MapData(X, Y).Graphic(1).grhindex <= 5680) Or _
            (MapData(X, Y).Graphic(1).grhindex >= 13547 And MapData(X, Y).Graphic(1).grhindex <= 13562)) And _
                MapData(X, Y).Graphic(2).grhindex = 0
                
End Function
Sub Engine_Move(ByVal Direccion As E_Heading, Optional refresh As Boolean = True)

    Dim LegalOk As Boolean
    
    If refresh Then
        If BloqMov Then
            BloqDir = Direccion
        End If
    End If
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = Map_Legal_Pos(UserPos.X, UserPos.Y - 1)
        Case E_Heading.EAST
            LegalOk = Map_Legal_Pos(UserPos.X + 1, UserPos.Y)
        Case E_Heading.SOUTH
            LegalOk = Map_Legal_Pos(UserPos.X, UserPos.Y + 1)
        Case E_Heading.WEST
            LegalOk = Map_Legal_Pos(UserPos.X - 1, UserPos.Y)
    End Select
    
    LegalOk = LegalOk And (IIf(under_stair, Direccion = NORTH Or Direccion = SOUTH, True))

    If LegalOk And Not UserParalizado Then

        If estaHabilitadoParaCaminar Then
            estaHabilitadoParaCaminar = False
            Call WriteWalk(Direccion)
            
            If Not UserDescansar Then
                TileEngine.Char_Move_Head char_current, Direccion
                TileEngine.Engine_MoveScreen Direccion
            End If
        End If
        
    Else
        If charlist(char_current).heading <> Direccion Then
            Call WriteChangeHeading(Direccion)
        End If
    End If
    
    ' Update 3D sounds!
    Call Audio.MoveListener(UserPos.X, UserPos.Y)
    
End Sub


Public Function Map_Letter_Fade_Set(ByVal grh_index As Long, Optional ByVal after_grh As Long = -1) As Boolean
    If grh_index <= 0 Or grh_index = map_letter_grh.grhindex Then Exit Function
        
    If after_grh = -1 Then
        map_letter_grh.grhindex = grh_index
        map_letter_fadestatus = 1
        map_letter_a = 0
        map_letter_grh_next = 0
    Else
        map_letter_grh.grhindex = after_grh
        map_letter_fadestatus = 1
        map_letter_a = 0
        map_letter_grh_next = grh_index
    End If
    
    Map_Letter_Fade_Set = True
End Function

Public Function Map_Letter_UnSet() As Boolean
    map_letter_grh.grhindex = 0
    map_letter_fadestatus = 0
    map_letter_a = 0
    map_letter_grh_next = 0
    Map_Letter_UnSet = True
End Function


