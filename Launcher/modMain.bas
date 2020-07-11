Attribute VB_Name = "modMain"
Option Explicit
'********************************************
'*************Configuracion******************
'********************************************
Public Sound As Byte
Public Music As Byte
Public AmbientSound As Byte
Public RepitMusic As Byte
Public InvertCanal As Byte
Public VolumeSound As Integer
Public VolumeMusic As Integer
Public TileBufferSize As Byte

Public Window As Byte
Public Sinc As Byte

Public BitPixel As Byte

'New
Public CursoresStandar As Byte
Public Cursores As Byte

Public ChatGlobal As Byte
Public ChatFaccionario As Byte
'********************************************
'*************/Configuracion*****************
'********************************************

Public DirInterfaces As String
'*********************************************************************
'*********************************************************************

'Cargar interfaces desde memoria
Public Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Public Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Public Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)


Sub Main()

On Error GoTo ErrHandler

    Dim lStr As String
     
    Call Shell("dxdiag " & App.Path & "\Resources\info.txt") 'c:\dxdialogt.txt")
        
    DirInterfaces = App.Path & "\Resources\Interface\"
    Load frmMain
    
    frmMain.Visible = True
    
    frmMain.cmbRes.AddItem "800 x 600 x 32"
    frmMain.cmbRes.AddItem "800 x 600 x 16"

    LoadConfig
    
   Do While Not FileExist(App.Path & "\Resources\info.txt", vbNormal)
        DoEvents
    Loop
    
    Open App.Path & "\Resources\info.txt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, lStr
            If InStr(1, lStr, "ard name") Then
                frmMain.cmbDevice.ListIndex = -1
                frmMain.cmbDevice.List(0) = LTrim$(RTrim$(ReadField(2, lStr, Asc(":"))))
                frmMain.cmbDevice.ListIndex = 0
  
                Close #1
                Kill App.Path & "\Resources\info.txt"
                Exit Sub
            End If
        Loop
    Close #1
    

    DoEvents
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description
    
    
End Sub
'*********************************************************************
Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function
'*********************************************************************
Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
    End If
End Function
Public Sub SaveConfig()
    Dim f As Integer
    
    If FileExist(App.Path & "\init\config.tao", vbNormal) Then Kill App.Path & "\init\config.tao"
    
    f = FreeFile
    Open App.Path & "\init\config.tao" For Binary Access Write As #f
        Put #f, , Music
        Put #f, , Sound
        Put #f, , AmbientSound
        Put #f, , VolumeMusic
        Put #f, , VolumeSound
        Put #f, , RepitMusic
        Put #f, , InvertCanal
        Put #f, , TileBufferSize
        Put #f, , Window
        Put #f, , Sinc
        
        Put #f, , BitPixel
        
        Put #f, , Cursores
        
        Put #f, , ChatFaccionario
        Put #f, , ChatGlobal
    Close #f
End Sub
'*********************************************************************
Public Sub LoadConfigDefault()
    Music = 1
    Sound = 1
    AmbientSound = 1
    VolumeMusic = 0
    VolumeSound = 100
    RepitMusic = 0
    InvertCanal = 0
    TileBufferSize = 4
    Window = 0
    
    BitPixel = 32
    
    Sinc = 0
    
    Cursores = 0
    
    ChatFaccionario = 1
    ChatGlobal = 1
End Sub
Public Sub LoadConfig()
    Dim f As Integer
    
    If Not FileExist(App.Path & "\init\config.tao", vbNormal) Then
        LoadConfigDefault
        SaveConfig
        Exit Sub
    End If
    
    f = FreeFile
    Open App.Path & "\init\config.tao" For Binary Access Read As #f
        Get #f, , Music
        Get #f, , Sound
        Get #f, , AmbientSound
        Get #f, , VolumeMusic
        Get #f, , VolumeSound
        Get #f, , RepitMusic
        Get #f, , InvertCanal
        Get #f, , TileBufferSize
        Get #f, , Window
        Get #f, , Sinc
        
        Get #f, , BitPixel
        
        Get #f, , Cursores
        
        Get #f, , ChatFaccionario
        Get #f, , ChatGlobal
    Close #f
    
    frmMain.lblBuffer.Caption = TileBufferSize
    
    If Window = 0 Then
       ' frmMain.cmbDevice.Enabled = True
        frmMain.cmbRes.Enabled = True
                    
        frmMain.imgOp(0).Picture = LoadInterface("vsyncoff")
    Else
        frmMain.cmbDevice.Enabled = False
        frmMain.cmbRes.Enabled = False
        frmMain.imgOp(0).Picture = LoadInterface("correron")
    End If
    
    If Sinc = 0 Then
        frmMain.imgOp(1).Picture = LoadInterface("vsyncoff")
    Else
        frmMain.imgOp(1).Picture = LoadInterface("correron")
    End If
    
    If Music = 0 Then
        frmMain.imgOp(2).Picture = LoadInterface("vsyncoff")
        'frmMain.imgCmb.Picture = LoadInterface("ambient-launcher-off")
    Else
        frmMain.imgOp(2).Picture = LoadInterface("correron")
    End If
    
    If Sound = 0 Then
        frmMain.imgOp(3).Picture = LoadInterface("vsyncoff")
    Else
        frmMain.imgOp(3).Picture = LoadInterface("correron")
    End If
    
    If BitPixel = 16 Then
        frmMain.cmbRes.ListIndex = 1
    Else
        frmMain.cmbRes.ListIndex = 0
    End If
End Sub
