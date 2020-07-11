Attribute VB_Name = "Mod_TileEngine"
Option Explicit

'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

''
'Sets a Grh animation to loop indefinitely.
Public Const INFINITE_LOOPS As Integer = -1

'Posicion en un mapa
Public Type Position
    x As Long
    Y As Long
End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
    Public Type GrhData
        sX As Integer
        sY As Integer
        FileNum As Integer
        pixelWidth As Integer
        pixelHeight As Integer
        TileWidth As Single
        TileHeight As Single
        NumFrames As Integer
        Frames() As Integer
        speed As Single
        mini_map_color As Long
        
        tu(3) As Single
        tv(3) As Single
        hardcor As Byte
    End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    speed As Single
    Started As Byte
    Loops As Integer
    alpha_blend As Boolean
    angle As Single
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    particle_group As Integer
End Type



Public IniPath As String
Public MapPath As String


'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public EngineRun As Boolean

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Function InMapBounds(ByVal x As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If x < XMinMapSize Or x > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function
