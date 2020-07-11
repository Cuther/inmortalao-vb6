VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDX8Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public bRunning As Boolean

Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public FPS As Integer
Private FramesPerSecCounter As Integer
Private timerElapsedTime As Single
Private timerTicksPerFrame As Double
Private particletimer As Single
Public engineBaseSpeed As Single
Private lFrameTimer As Long

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

Private MainViewTop As Integer
Private MainViewLeft As Integer
Private MainDestRect   As RECT
Private MainViewRect   As RECT
Private BackBufferRect As RECT

Private MainViewWidth As Integer
Private MainViewHeight As Integer


Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer


Private Type Particle
    friction As Single
    x As Single
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
 
Private Type Stream
    Name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
    
    speed As Single
    life_counter As Long
End Type
 
'Modified by: Ryan Cain (Onezero)
'Last modify date: 5/14/2003
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
 
'Dim StreamData() As particle_group
Dim TotalStreams As Long
Dim particle_group_list() As particle_group
Dim particle_group_count As Long
Dim particle_group_last As Long


Private char_current As Integer

Dim valoreBlur As Long
Dim dimeTex As Long
Dim tex As Direct3DTexture8
Dim D3DbackBuffer As Direct3DSurface8
Dim zTarget As Direct3DSurface8
Dim stencil As Direct3DSurface8
Dim superTex As Direct3DSurface8
Dim blur As Boolean
Public blur_factor As Byte

'Sets a Grh animation to loop indefinitely.

Private Function GetElapsedTime() As Single
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
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Public Sub Engine_Init()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    Dim DispMode As D3DDISPLAYMODE
    Dim DispModeBK As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    Dim ColorKeyVal As Long
    
    Set SurfaceDB = New clsSurfaceManDynDX8
    
    Set dX = New DirectX8
    Set D3D = dX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK
    
    
    With D3DWindow
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = frmMain.renderer.ScaleWidth
        .BackBufferHeight = frmMain.renderer.ScaleHeight
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.renderer.hWnd
    End With
    DispMode.Format = D3DFMT_X8R8G8B8
    If D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, DispMode.Format, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then
        Dim Caps8 As D3DCAPS8
        D3D.GetDeviceCaps 0, D3DDEVTYPE_HAL, Caps8


            DispMode.Format = DispModeBK.Format

    Else

        DispMode.Format = DispModeBK.Format
    End If
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.renderer.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                                            D3DWindow)
                                                            
                                                            
    HalfWindowTileHeight = (frmMain.renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.renderer.ScaleWidth / 32) \ 2
    
    TileBufferSize = 9
    TileBufferPixelOffsetX = (TileBufferSize - 1) * 32
    TileBufferPixelOffsetY = (TileBufferSize - 1) * 32
    
    D3DDevice.SetVertexShader FVF
    
    '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    
    Call SurfaceDB.Init(D3DX, D3DDevice, General_Get_Free_Ram_Bytes)

    engineBaseSpeed = 0.029
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set FPS value to 60 for startup
    FPS = 1000
    FramesPerSecCounter = 60
    
    
    MinXBorder = XMinMapSize + (frmMain.renderer.ScaleWidth / 64)
    MaxXBorder = XMaxMapSize - (frmMain.renderer.ScaleWidth / 64)
    MinYBorder = YMinMapSize + (frmMain.renderer.ScaleHeight / 64)
    MaxYBorder = YMaxMapSize - (frmMain.renderer.ScaleHeight / 64)

    'partículas
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0

    
    bRunning = True
    Exit Sub

ErrHandler:
Debug.Print "Error Number Returned: " & Err.Number
bRunning = False
End Sub

Public Sub Engine_Deinit()
Erase MapData
Set D3DDevice = Nothing
Set D3D = Nothing
Set dX = Nothing
End
End Sub


Private Function CreateTLVertex(x As Single, Y As Single, Z As Single, rhw As Single, color As Long, Specular As Long, tu As Single, tv As Single) As TLVERTEX
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    CreateTLVertex.x = x
    CreateTLVertex.Y = Y
    CreateTLVertex.Z = Z
    CreateTLVertex.rhw = rhw
    CreateTLVertex.color = color
    CreateTLVertex.Specular = Specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv
End Function

Public Sub Render()

    D3DDevice.BeginScene
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0


    Call RenderScreen(50, 50)

    
    frmMain.Label1.Caption = "FPS: " & FPS

    
    D3DDevice.EndScene
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

    If GetTickCount - lFrameTimer > 1000 Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 0
        lFrameTimer = GetTickCount
    Else
        FramesPerSecCounter = FramesPerSecCounter + 1
    End If
    
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

End Sub

Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
    Dim Y                   As Integer     'Keeps track of where on map we are
    Dim x                   As Integer     'Keeps track of where on map we are
    Dim screenminY          As Integer  'Start Y pos on current screen
    Dim screenmaxY          As Integer  'End Y pos on current screen
    Dim screenminX          As Integer  'Start X pos on current screen
    Dim screenmaxX          As Integer  'End X pos on current screen
    Dim ScreenX             As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY             As Integer  'Keeps track of where to place tile on screen
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
    Dim PixelOffsetX As Integer, PixelOffsetY As Integer
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    PixelOffsetX = OffsetCounterX
    PixelOffsetY = OffsetCounterY

    For Y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            With MapData(x, Y)
                '***********************************************
  
           
               If .particle_group > 0 Then
           
                Particle_Group_Render .particle_group, (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY
                
                End If
            End With
            ScreenX = ScreenX + 1
        Next x
        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next Y
        
End Sub

Private Function Geometry_Create_TLVertex(ByVal x As Single, ByVal Y As Single, ByVal Z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.x = x
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.Specular = Specular
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

Private Sub Device_Box_Textured_Render_Advance(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
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

    If GrhIndex = 0 Then Exit Sub
    Set d3dTextures.texture = SurfaceDB.GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
        
    light_value(0) = &HFFFFFFFF
    light_value(1) = &HFFFFFFFF
    light_value(2) = &HFFFFFFFF
    light_value(3) = &HFFFFFFFF

    
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
        
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + dest_height
        .Left = dest_x
        .Right = dest_x + dest_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
        
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.texture
    
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
        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    End If

End Sub

Public Sub Device_Box_Textured_Render(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
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
    
    If GrhIndex = 0 Then Exit Sub
    Set d3dTextures.texture = SurfaceDB.GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
    
    If (light_value(0) = 0) Then light_value(0) = base_light
    If (light_value(1) = 0) Then light_value(1) = base_light
    If (light_value(2) = 0) Then light_value(2) = base_light
    If (light_value(3) = 0) Then light_value(3) = base_light
        
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
                
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + src_height
        .Left = dest_x
        .Right = dest_x + src_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
    
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    If alpha_blend Then
        'Sets up properties for transparency.
        D3DDevice.SetRenderState D3DRS_ALPHAREF, 255
        D3DDevice.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL
    End If
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    End If

End Sub

Public Sub Start()
On Error Resume Next

Do While prgRun

If frmMain.WindowState <> vbMinimized And frmMain.Visible = True Then
    engine.Render
Else
    Sleep 10&
End If
DoEvents

Loop

engine.Engine_Deinit

    EngineRun = False
    frmCargando.Show

    'Destruimos los objetos públicos creados
    Set SurfaceDB = Nothing
    
    Call UnloadAllForms
    
End
End Sub

Private Function Engine_FToDW(f As Single) As Long
' single > long
Dim buf As D3DXBuffer
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, Engine_FToDW
End Function

Public Sub GrhRenderToHdc(ByVal grh_index As Long, desthDC As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional transparent As Boolean = False)
On Error Resume Next


    Dim file_path As String
    Dim src_x As Integer
    Dim src_y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hdcsrc As Long
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long

    If grh_index <= 0 Then Exit Sub
    If GrhData(grh_index).NumFrames = 0 Then Exit Sub
    'If it's animated switch grh_index to first frame
    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)
    End If


        file_path = App.Path & "\" & GrhData(grh_index).FileNum & ".bmp"
    
        Extract_File Graphics, GrhData(grh_index).FileNum & ".bmp", App.Path & "\"
        
        src_x = GrhData(grh_index).sX
        src_y = GrhData(grh_index).sY
        src_width = GrhData(grh_index).pixelWidth
        src_height = GrhData(grh_index).pixelHeight
            
        hdcsrc = CreateCompatibleDC(desthDC)
        PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
        
        Delete_File file_path
        
        If transparent = False Then
            BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
        Else
            MaskDC = CreateCompatibleDC(desthDC)
            
            PrevObj2 = SelectObject(MaskDC, LoadPicture(file_path))
            
            Grh_Create_Mask hdcsrc, MaskDC, src_x, src_y, src_width, src_height
            
            'Render tranparently
            BitBlt desthDC, screen_x, screen_y, src_width, src_height, MaskDC, src_x, src_y, vbSrcAnd
            BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcPaint
            
            Call DeleteObject(SelectObject(MaskDC, PrevObj2))
            
            DeleteDC MaskDC
        End If
        
        Call DeleteObject(SelectObject(hdcsrc, PrevObj))
        DeleteDC hdcsrc

    Exit Sub
End Sub

Private Sub Grh_Create_Mask(ByRef hdcsrc As Long, ByRef MaskDC As Long, ByVal src_x As Integer, ByVal src_y As Integer, ByVal src_width As Integer, ByVal src_height As Integer)
    Dim x As Integer
    Dim Y As Integer
    Dim TransColor As Long
    Dim ColorKey As String
    

    ColorKey = "10"
    TransColor = &HF

    'Make it a mask (set background to black and foreground to white)
    'And set the sprite's background white
    For Y = src_y To src_height + src_y
        For x = src_x To src_width + src_x
            If GetPixel(hdcsrc, x, Y) = TransColor Then
                SetPixel MaskDC, x, Y, vbWhite
                SetPixel hdcsrc, x, Y, vbBlack
            Else
                SetPixel MaskDC, x, Y, vbBlack
            End If
        Next x
    Next Y
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
    If Grh.Started = 1 Then
        Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.speed)
        If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
            Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
            If Grh.Loops <> -1 Then
                If Grh.Loops > 0 Then
                    Grh.Loops = Grh.Loops - 1
                Else
                    Grh.Started = 0
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    grh_index = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    'Center Grh over X, Y pos
    If GrhData(Grh.GrhIndex).TileWidth <> 1 Then
        screen_x = screen_x - Int(GrhData(Grh.GrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
    End If
    
    If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
        screen_y = screen_y - Int(GrhData(Grh.GrhIndex).TileHeight * 32) + 32
    End If
    
    'Draw it to device
    Device_Box_Textured_Render_Advance grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        width, height, alpha_blend, Grh.angle
End Sub

Private Sub Grh_Render(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False, Optional ByVal KillAnim As Boolean = 0)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'Modified by Juan Martín Sotuyo Dodero
'Added centering
'**************************************************************
On Error Resume Next
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Long
   
    If Grh.GrhIndex = 0 Then Exit Sub
       
    'Animation
    If Grh.Started = 1 Then
        Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.speed)
        If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
            Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
            If Grh.Loops <> -1 Then
                If Grh.Loops > 0 Then
                    Grh.Loops = Grh.Loops - 1
                Else
                    Grh.Started = 0
                End If
            End If
        End If
    End If
 
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    grh_index = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    If grh_index <= 0 Then Exit Sub
    If GrhData(grh_index).FileNum = 0 Then Exit Sub
       
    'Modified by Augusto José Rando
    'Simplier function - according to basic ORE engine
    If h_centered Then
        If GrhData(Grh.GrhIndex).TileWidth <> 1 Then
            screen_x = screen_x - Int(GrhData(Grh.GrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
        End If
    End If
   
    If v_centered Then
        If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
            screen_y = screen_y - Int(GrhData(Grh.GrhIndex).TileHeight * 32) + 32
        End If
    End If
   
    'Draw it to device
    Device_Box_Textured_Render grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list(), _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        alpha_blend, Grh.angle
 
End Sub
 
Private Function Particle_Group_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
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
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer) As Long
    
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 12/15/2002
'Returns the particle_group_index if successful, else 0
'**************************************************************
    If (map_x <> -1) And (map_y <> -1) Then
    If Map_Particle_Group_Get(map_x, map_y) = 0 Then
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    End If
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
'Last Modify Date: 10/07/2002
'Makes a new particle effect
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
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
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
    
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'plot particle group on map
    MapData(map_x, map_y).particle_group = particle_group_index

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
'Last Modify Date: 4/24/2003
'
'**************************************************************

    If no_move = False Then
                If temp_particle.alive_counter = 0 Then
                    Grh_Initialize temp_particle.Grh, grh_index, , alpha_blend
                    temp_particle.x = RandomNumber(x1, x2)
                    temp_particle.Y = RandomNumber(y1, y2)
                    temp_particle.vector_x = RandomNumber(vecx1, vecx2)
                    temp_particle.vector_y = RandomNumber(vecy1, vecy2)
                    temp_particle.angle = angle
                    temp_particle.alive_counter = RandomNumber(life1, life2)
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
                   If spin = True Then temp_particle.Grh.angle = temp_particle.Grh.angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
                    If temp_particle.angle >= 360 Then
                        temp_particle.angle = 0
                    End If
                                
                    If XMove = True Then temp_particle.vector_x = RandomNumber(move_x1, move_x2)
                    If YMove = True Then temp_particle.vector_y = RandomNumber(move_y1, move_y2)
                End If

        'Add in vector
        temp_particle.x = temp_particle.x + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.vector_y \ temp_particle.friction)
    
        'decrement counter
         temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If
    
    'Draw it
    If grh_resize = True Then
        If temp_particle.Grh.GrhIndex Then
             Grh_Render_Advance temp_particle.Grh, temp_particle.x + screen_x, temp_particle.Y + screen_y, grh_resizex, grh_resizey, rgb_list(), True, True, alpha_blend
            Exit Sub
        End If
    End If
 
    'Draw it
    If temp_particle.Grh.GrhIndex Then
        Grh_Render temp_particle.Grh, _
                temp_particle.x + screen_x, _
                temp_particle.Y + screen_y, _
                rgb_list(), 1, False, True _
                , True
    End If
End Sub

Private Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 12/15/2002
'Renders a particle stream at a paticular screen point
'*****************************************************************
 On Error Resume Next
    Dim loopc As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
    
    'Set colors
        temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
        temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
        temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
        temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)
        
    If particle_group_list(particle_group_index).alive_counter Then
    
        'See if it is time to move a particle
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timerTicksPerFrame
        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
    
    
        'If it's still alive render all the particles inside
        For loopc = 1 To particle_group_list(particle_group_index).particle_count
        
            'Render particle
            Particle_Render particle_group_list(particle_group_index).particle_stream(loopc), _
                            screen_x, screen_y, _
                            particle_group_list(particle_group_index).grh_index_list(Round(RandomNumber(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
                            temp_rgb(), _
                            particle_group_list(particle_group_index).alpha_blend, no_move, _
                            particle_group_list(particle_group_index).x1, particle_group_list(particle_group_index).y1, particle_group_list(particle_group_index).angle, _
                            particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
                            particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
                            particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
                            particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
                            particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
                            particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).x2, _
                            particle_group_list(particle_group_index).y2, particle_group_list(particle_group_index).XMove, _
                            particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
                            particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
                            particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
                            particle_group_list(particle_group_index).spin, particle_group_list(particle_group_index).grh_resize, particle_group_list(particle_group_index).grh_resizex, particle_group_list(particle_group_index).grh_resizey
                            
        Next loopc
        
        If no_move = False Then
            'Update the group alive counter
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1
            End If
        End If
    
    Else
        'If it's dead destroy it
        particle_group_list(particle_group_index).particle_count = particle_group_list(particle_group_index).particle_count - 1
        If particle_group_list(particle_group_index).particle_count <= 0 Then Particle_Group_Destroy particle_group_index
    End If
End Sub

Public Function Particle_Type_Get(ByVal particle_index As Long) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 8/27/2003
'Returns the stream type of a particle stream
'*****************************************************************
    If Particle_Group_Check(particle_index) Then
        Particle_Type_Get = particle_group_list(particle_index).stream_type
    End If
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

Public Function Particle_Group_Map_Pos_Set(ByVal particle_group_index As Long, ByVal map_x As Long, ByVal map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/27/2003
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        'Make sure it's a legal move
        If InMapBounds(map_x, map_y) Then
            'Move it
            particle_group_list(particle_group_index).map_x = map_x
            particle_group_list(particle_group_index).map_y = map_y
    
            Particle_Group_Map_Pos_Set = True
        End If
    End If
End Function

Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As particle_group
    
    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then
       MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group = 0
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
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count - 1
End Sub

Public Function Map_Particle_Group_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long

    If InMapBounds(map_x, map_y) Then
        Map_Particle_Group_Get = MapData(map_x, map_y).particle_group
    Else
        Map_Particle_Group_Get = 0
    End If
End Function

Private Sub Grh_Initialize(ByRef Grh As Grh, ByVal grh_index As Long, Optional ByVal Started As Byte = 2, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan José Lozano (juanjo_22)
'Last Modify Date: 10/07/2002
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    On Error Resume Next
    
    'Copy of parameters
    If grh_index = 0 Then Exit Sub
    
    Grh.GrhIndex = grh_index
    Grh.alpha_blend = alpha_blend
    Grh.angle = angle
    
    If Started = 2 Then
        'Start it if it's a animated grh
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    'Set frame counters
    Grh.FrameCounter = 1
    Grh.speed = GrhData(Grh.GrhIndex).speed
End Sub
