VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''
' Set to true once the object is created to prevent multiple initializations of the same object.
Dim inited As Boolean

Private NumSoundBuffers As Byte
Private Const NumSoundBuffersMax As Byte = 30
 
Private Const MAX_DISTANCE_TO_SOURCE As Integer = 150
Private Const DELTA_FQ As Integer = 600

Private Type SoundBuffer
    filename As String
    looping As Boolean
    X As Byte
    Y As Byte
    normalFq As Long
    buffer As DirectSoundSecondaryBuffer8
End Type

Public Enum LoopStyle
    Default = 0
    Disabled = 1
    Enabled = 2
End Enum

Dim lastPosX As Integer
Dim lastPosY As Integer

Dim SoundPath As String
Dim SndVolume As Long

'**** Direct Sound ********
 
Private dX As DirectX8
Private Sound_Buffer() As DirectSoundSecondaryBuffer8
Private DS As DirectSound8
Private DS_Enum As DirectSoundEnum8
 
'Array of all existing sound buffers
Dim DSBuffers(1 To NumSoundBuffersMax) As SoundBuffer

'**** MP3 ********
Dim IMC   As IMediaControl
Dim IBA   As IBasicAudio
Dim IME   As IMediaEvent
Dim IMPos As IMediaPosition

Dim MusicPath As String
Dim CurrentMusicFile As String
Dim MusicPlaying As Boolean

Private Sub Class_Terminate()
On Local Error Resume Next
    Dim loopc As Long
   
    'Stop every channel being used and destroy the buffer
    For loopc = 1 To NumSoundBuffers
        Call DSBuffers(loopc).buffer.Stop
        Set DSBuffers(loopc).buffer = Nothing
    Next loopc
   
    'Stop music
    'If Not Segment Is Nothing Then Call Performance.StopEx(Segment, 0, DMUS_SEGF_DEFAULT)

    inited = False
End Sub
 
''
'Initializes the clsAudio object.
Public Function Initialize(ByVal hwnd As Long, ByVal sound_path As String, ByVal music_path As String) As Boolean
On Error GoTo ErrHandler
   
    ' Prevent multiple creations
    If inited Then Exit Function
   
     'Cargamos
    Set dX = New DirectX8
   
    'Validate DX object
    If dX Is Nothing Then GoTo ErrHandler
   
    'Set sound path
    SoundPath = sound_path
   
    'Set default volume
    SndVolume = 0
   
    Set DS_Enum = dX.GetDSEnum
    Set DS = dX.DirectSoundCreate(DS_Enum.GetGuid(1))
 
    DS.SetCooperativeLevel hwnd, DSSCL_PRIORITY
   
    'Set music path
    MusicPath = music_path
   
    CurrentMusicFile = vbNullString
    
    inited = True
   
    Initialize = True
Exit Function
 
ErrHandler:
End Function
 
 
''
' Plays a new WAVE file.
 
Public Function PlayWave(ByVal filename As Long, Optional ByVal SrcX As Byte = 0, Optional ByVal SrcY As Byte = 0, Optional ByVal LoopSound As LoopStyle = LoopStyle.Default) As Long
On Error GoTo ErrHandler
    Dim bufferIndex As Long
   
    If modDeclaraciones.Sound = 0 Then Exit Function
    
    'Get the buffer index were wave was loaded
    bufferIndex = LoadWave(filename & ".wav", LoopSound)
    
    If bufferIndex = 0 Or filename = 0 Then Exit Function   'If an error ocurred abort

    With DSBuffers(bufferIndex)
        If .looping Then
            .buffer.Play DSBPLAY_LOOPING
        Else
            .buffer.Play DSBPLAY_DEFAULT
        End If
       
        'Apply volume
        Call .buffer.SetVolume(SndVolume)
       
        'Store position
        .X = SrcX
        .Y = SrcY
       
        'Get standard frequency
        .normalFq = .buffer.GetFrequency()
    End With
   
    PlayWave = bufferIndex
   
    If SrcX <> 0 And SrcY <> 0 Then
        Call Update3DSound(bufferIndex, 0, 0)
    End If

Exit Function
 
ErrHandler:
End Function
 
Private Sub Update3DSound(ByVal bufferIndex As Long, ByVal deltaX As Integer, ByVal deltaY As Integer)
    Dim linearDistanceOld As Single
    Dim linearDistanceNew As Single
    Dim distanceXOld As Long 'Integer
    Dim distanceYOld As Long 'Integer
    Dim distanceXNew As Long ' Integer
    Dim distanceYNew As Long

    With DSBuffers(bufferIndex)
        distanceXOld = .X - lastPosX
        distanceYOld = .Y - lastPosY
       
        distanceXNew = distanceXOld + deltaX
        distanceYNew = distanceYOld + deltaY
       
        linearDistanceOld = Sqr(distanceXOld * distanceXOld + distanceYOld * distanceYOld)
        linearDistanceNew = Sqr(distanceXNew * distanceXNew + distanceYNew * distanceYNew)
        If linearDistanceNew > MAX_DISTANCE_TO_SOURCE Then linearDistanceNew = MAX_DISTANCE_TO_SOURCE

        'Set volumen amortiguation according to distance
        Call .buffer.SetVolume(SndVolume * (1 - linearDistanceNew / MAX_DISTANCE_TO_SOURCE))
       
        'Prevent division by zero
        If linearDistanceNew = 0 Then linearDistanceNew = 1
       
        'Set panning according to relative position of the source to the listener
'        Call .Buffer.SetPan((distanceXNew / linearDistanceNew) * DSBPAN_LEFT)
    End With
End Sub
 
''
' Updates 3D sounds based on the movement of the listener.
 
Public Sub MoveListener(ByVal X As Integer, ByVal Y As Integer)
    Dim i As Long
    Dim deltaX As Integer
    Dim deltaY As Integer
   
    deltaX = X - lastPosX
    deltaY = Y - lastPosY
   
    For i = 1 To NumSoundBuffersMax
        If Not DSBuffers(i).buffer Is Nothing Then
            If DSBuffers(i).buffer.GetStatus And DSBSTATUS_PLAYING Then
                If DSBuffers(i).X <> 0 And DSBuffers(i).Y <> 0 Then
                    Call Update3DSound(i, deltaX, deltaY)
                End If
            End If
        End If
    Next i
   
    lastPosX = X
    lastPosY = Y
End Sub
 
''
' Loads a new WAVE file.
 
Private Function LoadWave(ByVal filename As String, ByVal looping As LoopStyle) As Long

On Error GoTo ErrHandler
    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX
    Dim i As Long
   
  '  filename = UCase$(filename)
   
    'Check if the buffer is in memory and not playing
    For i = 1 To NumSoundBuffersMax
        If DSBuffers(i).filename = filename Then
            If DSBuffers(i).buffer.GetStatus = 0 Or DSBuffers(i).buffer.GetStatus() = DSBSTATUS_BUFFERLOST Then
                'Found it!!! We just play this one :)
                LoadWave = i
               
                'Set looping if needed
                If looping <> LoopStyle.Default Then DSBuffers(i).looping = (looping = LoopStyle.Enabled)
               
                Exit Function
            End If
        End If
    Next i
   
    'Not in memory, search for an empty buffer
    For i = 1 To NumSoundBuffersMax
        If DSBuffers(i).buffer Is Nothing Then
            NumSoundBuffers = i
            Exit For
        End If
    Next i
   
    'No empty buffers?? Search for one not being played and use that slot
    If i > NumSoundBuffersMax Then
        For i = 1 To NumSoundBuffersMax
            If DSBuffers(i).buffer.GetStatus() = DSBSTATUS_BUFFERLOST Or DSBuffers(i).buffer.GetStatus() = 0 Then
                Exit For
            End If
        Next i
    End If
   
    'Still no free buffers?? Get the first non-looping sound
    If i > NumSoundBuffersMax Then
        For i = 1 To NumSoundBuffersMax
            If Not (DSBuffers(i).buffer.GetStatus() And DSBSTATUS_LOOPING) Then
                Call DSBuffers(i).buffer.Stop
                Exit For
            End If
        Next i
    End If
   
    'If every sound is looping, just ignore new sound unless it's looping too (in which case we set it at the first slot by default)
    If i > NumSoundBuffersMax Then
        If Not looping Then Exit Function
        i = 1
    End If
   
    'Create new sound
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
   
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 32
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
   
    With DSBuffers(i)
        Set .buffer = Nothing   'Get rid of any previous data
       
        .filename = filename
        .looping = (looping = LoopStyle.Enabled)
               
    
            Extract_File wav, filename, resource_path

            Set .buffer = DS.CreateSoundBufferFromFile(resource_path & filename, bufferDesc)  ' waveFormat
            Delete_File resource_path & filename
           If FileExist(resource_path & filename, vbNormal) Then Kill resource_path & filename
     
    End With
   
    LoadWave = i
Exit Function
 
ErrHandler:
End Function
 
''
' Stops a given sound or all of them.
 
Public Sub StopWave(Optional ByVal bufferIndex As Long = 0)
    If bufferIndex > 0 And bufferIndex <= NumSoundBuffersMax Then
        If DSBuffers(bufferIndex).buffer.GetStatus() And DSBSTATUS_PLAYING Then
            Call DSBuffers(bufferIndex).buffer.Stop
        End If
    ElseIf bufferIndex = 0 Then
        Dim i As Long
        For i = 1 To NumSoundBuffersMax
            If Not DSBuffers(i).buffer Is Nothing Then
                If DSBuffers(i).buffer.GetStatus() <> 0 And DSBuffers(i).buffer.GetStatus() <> DSBSTATUS_BUFFERLOST Then
                    Call DSBuffers(i).buffer.Stop
                End If
            End If
        Next i
    End If
End Sub
 
''
' Retrieves wether there are sounds currentyl playing or not.
 
Public Property Get PlayingSound() As Boolean
    Dim i As Long
   
    For i = 1 To NumSoundBuffersMax
        If DSBuffers(i).buffer.GetStatus() And DSBSTATUS_PLAYING Then
            PlayingSound = True
            Exit Property
        End If
    Next i
End Property
 
''
' Sets the volume of sound.
 
Public Property Let SoundVolume(ByVal Volume As Byte)
    Dim i As Long
   
    If Volume > 100 Then Exit Property
   
    VolumeSound = Volume
    
    'Take percentage to actual value
    SndVolume = Volume * (DSBVOLUME_MAX - DSBVOLUME_MIN) / 100 + DSBVOLUME_MIN
   
    For i = 1 To NumSoundBuffersMax
        If Not DSBuffers(i).buffer Is Nothing Then
            If DSBuffers(i).buffer.GetStatus And DSBSTATUS_PLAYING Then
                Call DSBuffers(i).buffer.SetVolume(SndVolume)
            End If
        End If
    Next i
End Property
 
''
' Gets the volume of sound.
 
Public Property Get SoundVolume() As Byte
    SoundVolume = CByte((SndVolume - DSBVOLUME_MIN) * 100 / (DSBVOLUME_MAX - DSBVOLUME_MIN))
End Property
 

Public Function PlayMusic(Optional ByVal file As String = "", Optional ByVal Loops As Long = -1) As Boolean

On Error GoTo ErrHandler
    If LenB(file) > 0 Then CurrentMusicFile = file
   
    If modDeclaraciones.Music = 0 Then Exit Function

    If LenB(file) > 0 Then
        If Not LoadMusic(file) Then
            Exit Function
        Else
        
            MusicPlaying = True
        End If
    End If
   
    IMC.Run
    
    PlayMusic = True
Exit Function
 
ErrHandler:
End Function
Private Function LoadMusic(ByVal file As String) As Boolean
On Error GoTo ErrHandler
        
    If MusicPlaying Then Call StopMusic
  
    If MIDI_ACTIVATE = 1 Then
        If Not FileExist(resource_path & "Music\" & file & ".mid", vbNormal) Then GoTo ErrHandler
    Else
        If Not FileExist(resource_path & "Music\" & file & ".mp3", vbNormal) Then GoTo ErrHandler
    End If
    
  
    
    Set IMC = New FilgraphManager
    
    If UCase$(Right$(file, 4)) = "" Then file = Left(file, Len(file) - 4)
    
    If MIDI_ACTIVATE = 1 Then
        IMC.RenderFile resource_path & "Music\" & file & ".id"
    Else
        IMC.RenderFile resource_path & "Music\" & file & ".mp3"
    End If
    
    Set IBA = IMC
    IBA.Volume = VolumeMusic
        
    Set IME = IMC
        
    Set IMPos = IMC
    If ObjPtr(IMPos) Then IMPos.Rate = 1#
    IMPos.CurrentPosition = 0
    
    LoadMusic = True
    Exit Function
    
ErrHandler:
    LoadMusic = False
    Debug.Print "LoadMusic : " & Err.Description

End Function
Private Function StopMusic() As Boolean
    If Not IMC Is Nothing Then IMC.Stop
    MusicPlaying = False
End Function
Public Property Let MusicVolume(ByVal Volume As Long)
    modDeclaraciones.VolumeMusic = Volume
    If MusicPlaying Then
        IBA.Volume = Volume
    End If
End Property
Public Property Get MusicVolume() As Long
    If Music = 0 Then Exit Property
   
    MusicVolume = modDeclaraciones.VolumeMusic
End Property
 
Public Sub MusicRefresh()
    If Music Then
        Call PlayMusic(CurrentMusicFile)
    Else
        Call StopMusic
    End If
End Sub

Public Sub SoundRefresh()
    If Not Sound Then
        Call StopWave
    End If
End Sub
 
