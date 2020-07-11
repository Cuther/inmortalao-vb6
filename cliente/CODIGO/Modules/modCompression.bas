Attribute VB_Name = "modCompression"

Option Explicit
'Cargar interfaces desde memoria
Public Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Public Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Public Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

'This structure will describe our binary file's
'size and number of contained files
Public Type FILEHEADER
    lngFileSize As Long                 'How big is this file? (Used to check integrity)
    intNumFiles As Integer              'How many files are inside?
End Type

'This structure will describe each file contained
'in our binary file
Public Type INFOHEADER
    lngFileStart As Long            'Where does the chunk start?
    lngFileSize As Long             'How big is this chunk of stored data?
    strFileName As String * 32      'What's the name of the file this data came from?
    lngFileSizeUncompressed As Long 'How big is the file compressed
End Type

Public Enum resource_file_type
    Graphics
    Midi
    MP3
    wav
    Scripts
    Patch
    Interface
    Maps
End Enum

Public resource_path As String
Public init_path As String
Public graph_path As String
Public inter_path As String
Public map_path As String
Public wav_path As String


Private Declare Function Compress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function UnCompress Lib "zlib.dll" Alias "uncompress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
'To get free bytes in drive
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, BytesTotal As Currency, FreeBytesTotal As Currency) As Long

'To get free bytes in RAM



Public Sub Compress_Data(ByRef data() As Byte)

    Dim Dimensions As Long
    Dim DimBuffer As Long
    Dim BufTemp() As Byte
    Dim BufTemp2() As Byte
    Dim loopc As Long
    
    'Get size of the uncompressed data
    Dimensions = UBound(data)
    
    'Prepare a buffer 1.06 times as big as the original size
    DimBuffer = Dimensions * 1.06
    ReDim BufTemp(DimBuffer)
    
    'Compress data using zlib
    Compress BufTemp(0), DimBuffer, data(0), Dimensions
    
    'Deallocate memory used by uncompressed data
    Erase data
    
    'Get rid of unused bytes in the compressed data buffer
    ReDim Preserve BufTemp(DimBuffer - 1)
    
    'Copy the compressed data buffer to the original data array so it will return to caller
    data = BufTemp
    
    'Deallocate memory used by the temp buffer
    Erase BufTemp
    
    'Encrypt the first byte of the compressed data for extra security
    data(0) = data(0) Xor 189
End Sub

Public Sub Decompress_Data(ByRef data() As Byte, ByVal OrigSize As Long)

    Dim BufTemp() As Byte
    
    ReDim BufTemp(OrigSize - 1)
    
    'Des-encrypt the first byte of the compressed data
    data(0) = data(0) Xor 189
    
    UnCompress BufTemp(0), OrigSize, data(0), UBound(data) + 1
    
    ReDim data(OrigSize - 1)
    
    data = BufTemp
    
    Erase BufTemp
End Sub

Public Function Extract_All_Files(ByVal file_type As resource_file_type, ByVal resource_path As String) As Boolean

    Dim loopc As Long
    Dim SourceFilePath As String
    Dim OutputFilePath As String
    Dim SourceFile As Integer
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim handle As Integer
    
'Set up the error handler
On Local Error GoTo ErrHandler
    
    Select Case file_type
        Case Graphics

            SourceFilePath = resource_path & "\Graficos.iao"
            OutputFilePath = graph_path
            
        Case Midi

            SourceFilePath = resource_path & "\MIDI.iao"
            OutputFilePath = graph_path
        
        Case MP3
                
            SourceFilePath = resource_path & "\MP3.iao"
            OutputFilePath = graph_path
        
        Case wav

            SourceFilePath = resource_path & "\Sounds.iao"
            OutputFilePath = wav_path
        
        Case Scripts

            SourceFilePath = resource_path & "\Init.iao"
            OutputFilePath = init_path
        
        Case Interface

            SourceFilePath = resource_path & "\Interface.iao"
            OutputFilePath = inter_path
            
        Case Maps
        
            SourceFilePath = resource_path & "\Mapas.iao"
            OutputFilePath = map_path

        
        Case Else
            Exit Function
    End Select
    
    'Open the binary file
    SourceFile = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As SourceFile
    
    'Extract the FILEHEADER
    Get SourceFile, 1, FileHead
        
      '  Desencrypt File Header
    Encrypt_File_Header FileHead
        
        
    'Check the file for validity
    If LOF(SourceFile) <> FileHead.lngFileSize Then
        MsgBox "Resource file " & SourceFilePath & " seems to be corrupted.", , "Error"
        Close SourceFile
        Erase InfoHead
        Exit Function
    End If
    
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
    
    'Extract the INFOHEADER
    Get SourceFile, , InfoHead
        
    'Extract all of the files from the binary file
    For loopc = 0 To UBound(InfoHead)
        
        'Check if there is enough memory
        If InfoHead(loopc).lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left(App.Path, 3)) Then
            MsgBox "There is not enough free memory to continue extracting files."
            Exit Function
        End If
        
       ' If Not Form1.ProgressBar1 Is Nothing Then Form1.ProgressBar1.max = FileHead.intNumFiles
        
        'Resize the byte data array
        ReDim SourceData(InfoHead(loopc).lngFileSize - 1)
        
        'Get the data
        Get SourceFile, InfoHead(loopc).lngFileStart, SourceData
        
        'Decompress all data
        Decompress_Data SourceData, InfoHead(loopc).lngFileSizeUncompressed
        
        'Get a free handler
        handle = FreeFile
        
        'Create a new file and put in the data
        Open OutputFilePath & InfoHead(loopc).strFileName For Binary As handle
        
        Put handle, , SourceData
        
        Close handle
        
        Erase SourceData
        
      '  If Not Form1.ProgressBar1 Is Nothing Then Form1.ProgressBar1.value = Form1.ProgressBar1.value + 1
        
        DoEvents
    Next loopc
    
    'Close the binary file
    Close SourceFile
    
    Erase InfoHead
    
    Extract_All_Files = True
Exit Function

ErrHandler:
    Close SourceFile
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "Unable to decode binary file. Reason: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error"
End Function



Public Function Compress_Files(ByVal file_type As resource_file_type, ByVal resource_path As String, ByVal dest_path As String) As Boolean


    Dim SourceFilePath As String
    Dim SourceFileExtension As String
    Dim OutputFilePath As String
    Dim SourceFile As Long
    Dim OutputFile As Long
    Dim SourceFileName As String
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim FileNames() As String
    Dim lngFileStart As Long
    Dim loopc As Long
    
'Set up the error handler
On Local Error GoTo ErrHandler
    
    Select Case file_type
        Case Graphics
            SourceFilePath = graph_path
            SourceFileExtension = ".bmp"
            OutputFilePath = dest_path & "Graficos.iao"
        
        Case Midi
            SourceFilePath = graph_path
            SourceFileExtension = ".mid"
            OutputFilePath = dest_path & "MIDI.iao"
        
        Case MP3
            SourceFilePath = graph_path
            SourceFileExtension = ".mp3"
            OutputFilePath = dest_path & "MP3.iao"
        
        Case wav
            SourceFilePath = wav_path
            SourceFileExtension = ".wav"
            OutputFilePath = dest_path & "Sounds.iao"
                
        Case Scripts
            SourceFilePath = init_path
            SourceFileExtension = ".*"
            OutputFilePath = dest_path & "Init.iao"
        
        Case Maps
            SourceFilePath = map_path
            SourceFileExtension = ".*"
            OutputFilePath = dest_path & "mapas.iao"
    
        Case Interface
            SourceFilePath = inter_path
            SourceFileExtension = ".bmp"
            OutputFilePath = dest_path & "Interface.iao"
        

    
    End Select
    
    'Get first file in the directoy
    SourceFileName = Dir$(SourceFilePath & "*" & SourceFileExtension, vbNormal)
    
   
    
    SourceFile = FreeFile
    
    'Get all other files i nthe directory
    While SourceFileName <> ""
        FileHead.intNumFiles = FileHead.intNumFiles + 1
        
        ReDim Preserve FileNames(FileHead.intNumFiles - 1)
        FileNames(FileHead.intNumFiles - 1) = LCase(SourceFileName)
        
        'Search new file
        SourceFileName = Dir$()
    Wend
    
    'If we found none, be can't compress a thing, so we exit
    If FileHead.intNumFiles = 0 Then
        MsgBox "There are no files of extension " & SourceFileExtension & " in " & SourceFilePath & ".", , "Error"
        Exit Function
    End If
    
   ' Form1.ProgressBar1.max = FileHead.intNumFiles
   ' Form1.ProgressBar1.value = 0
    
    'Sort file names alphabetically (this will make patching much easier).
    General_Quick_Sort FileNames(), 0, UBound(FileNames)
    
    'Resize InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
        
    'Destroy file if it previuosly existed
    If Dir(OutputFilePath, vbNormal) <> "" Then
        Kill OutputFilePath
    End If
    
    'Open a new file
    OutputFile = FreeFile
    Open OutputFilePath For Binary Access Read Write As OutputFile
    
    For loopc = 0 To FileHead.intNumFiles - 1
        'Find a free file number to use and open the file
        SourceFile = FreeFile
        Open SourceFilePath & FileNames(loopc) For Binary Access Read Lock Write As SourceFile
        
        'Store file name
        InfoHead(loopc).strFileName = FileNames(loopc)
        
        'Find out how large the file is and resize the data array appropriately
        ReDim SourceData(LOF(SourceFile) - 1)
        
        'Store the value so we can decompress it later on
        InfoHead(loopc).lngFileSizeUncompressed = LOF(SourceFile)
        
        'Get the data from the file
        Get SourceFile, , SourceData
        
        'Compress it
        Compress_Data SourceData
        
        'Save it to a temp file
        Put OutputFile, , SourceData
        
        'Set up the file header
        FileHead.lngFileSize = FileHead.lngFileSize + UBound(SourceData) + 1
        
        'Set up the info headers
        InfoHead(loopc).lngFileSize = UBound(SourceData) + 1
        
        Erase SourceData
        
        'Close temp file
        Close SourceFile
        
      '  If Not Form1.ProgressBar1 Is Nothing Then Form1.ProgressBar1.value = Form1.ProgressBar1.value + 1
        
        DoEvents
    Next loopc
    
    'Finish setting the FileHeader data
    FileHead.lngFileSize = FileHead.lngFileSize + CLng(FileHead.intNumFiles) * Len(InfoHead(0)) + Len(FileHead)
    
    'Set InfoHead data
    lngFileStart = Len(FileHead) + CLng(FileHead.intNumFiles) * Len(InfoHead(0)) + 1
    For loopc = 0 To FileHead.intNumFiles - 1
        InfoHead(loopc).lngFileStart = lngFileStart
        lngFileStart = lngFileStart + InfoHead(loopc).lngFileSize
        Encrypt_Info_Header InfoHead(loopc)
    Next loopc
    
        'Encrypt the FileHeader
    Encrypt_File_Header FileHead
        
    '************ Write Data
    
    'Get all data stored so far
    ReDim SourceData(LOF(OutputFile) - 1)
    Seek OutputFile, 1
    Get OutputFile, , SourceData
    
    Seek OutputFile, 1
    
    'Store the data in the file
    Put OutputFile, , FileHead
    Put OutputFile, , InfoHead
    Put OutputFile, , SourceData
    
    'Close the file
    Close OutputFile
    
    Erase InfoHead
    Erase SourceData
Exit Function

ErrHandler:
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "Unable to create binary file. Reason: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Function Extract_File(ByVal file_type As resource_file_type, ByVal file_name As String, ByVal OutputFilePath As String) As Boolean
    Dim loopc As Long
    Dim SourceFilePath As String
    Dim SourceData() As Byte
    Dim InfoHead As INFOHEADER
    Dim handle As Integer
    
'Set up the error handler
On Local Error GoTo ErrHandler


    Select Case file_type
        Case Graphics
            SourceFilePath = resource_path & "\Graficos.iao"

        Case Midi
            SourceFilePath = resource_path & "\MIDI.iao"

        Case MP3
            SourceFilePath = resource_path & "\MP3.iao"
        
        Case wav
 
            SourceFilePath = resource_path & "\Sounds.iao"

        Case Scripts
            SourceFilePath = resource_path & "\Init.iao"

        Case Interface
            SourceFilePath = resource_path & "\Interface.iao"
            
        Case Maps
            SourceFilePath = resource_path & "\Mapas.iao"
        
        Case Else
            Exit Function
    End Select
    
    'Find the Info Head of the desired file
    InfoHead = File_Find(SourceFilePath, file_name)
    
    If InfoHead.strFileName = "" Or InfoHead.lngFileSize = 0 Then Exit Function

    'Open the binary file
    handle = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As handle
    
    'Check the file for validity
    'If LOF(handle) <> InfoHead.lngFileSize Then
    '    Close handle
    '    MsgBox "Resource file " & SourceFilePath & " seems to be corrupted.", , "Error"
    '    Exit Function
    'End If
    
    'Make sure there is enough space in the HD
    If InfoHead.lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left$(App.Path, 3)) Then
        Close handle
        MsgBox "There is not enough drive space to extract the compressed file.", , "Error"
        Exit Function
    End If
    
    'Extract file from the binary file
    
    'Resize the byte data array
    ReDim SourceData(InfoHead.lngFileSize - 1)
    

    
    'Get the data
    Get handle, InfoHead.lngFileStart, SourceData
    
    'Decompress all data
    Decompress_Data SourceData, InfoHead.lngFileSizeUncompressed
    
    'Close the binary file
    Close handle
    
    'Get a free handler
    handle = FreeFile
    
    Open OutputFilePath & InfoHead.strFileName For Binary As handle
    
    Put handle, 1, SourceData
    
    Close handle
   
    Erase SourceData
        
 
    Extract_File = True
Exit Function

ErrHandler:
    Close handle
    Erase SourceData
    'Display an error message if it didn't work
    'MsgBox "Unable to decode binary file. Reason: " & Err.number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Sub Delete_File(ByVal file_path As String)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 3/03/2005
'Deletes a resource files
'*****************************************************************
    Dim handle As Integer
    Dim data() As Byte
    
    On Error GoTo Error_Handler
    
    'We open the file to delete
    handle = FreeFile
    Open file_path For Binary Access Write Lock Read As handle
    
    'We replace all the bytes in it with 0s
    ReDim data(LOF(handle) - 1)
    Put handle, 1, data
    
    'We close the file
    Close handle
    
    'Now we delete it, knowing that if they retrieve it (some antivirus may create backup copies of deleted files), it will be useless
    Kill file_path
    
    Exit Sub
    
Error_Handler:
    Kill file_path
        
End Sub

Private Function File_Find(ByVal resource_file_path As String, ByVal file_name As String) As INFOHEADER
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Looks for a compressed file in a resource file. Uses binary search ;)
'**************************************************************
On Error GoTo ErrHandler
    Dim max As Integer  'Max index
    Dim min As Integer  'Min index
    Dim mid As Integer  'Middle index
    Dim file_handler As Integer
    Dim file_head As FILEHEADER
    Dim info_head As INFOHEADER
    
    'Fill file name with spaces for compatibility
    If Len(file_name) < Len(info_head.strFileName) Then _
        file_name = file_name & Space$(Len(info_head.strFileName) - Len(file_name))
    
    'Open resource file
    file_handler = FreeFile
    Open resource_file_path For Binary Access Read Lock Write As file_handler
    
    'Get file head
    Get file_handler, 1, file_head
    
    min = 1
    max = file_head.intNumFiles
    
    Do While min <= max
        mid = (min + max) / 2
        
        'Get the info header of the appropiate compressed file
        Get file_handler, CLng(Len(file_head) + CLng(Len(info_head)) * CLng((mid - 1)) + 1), info_head
                
        If file_name < info_head.strFileName Then
            If max = mid Then
                max = max - 1
            Else
                max = mid
            End If
        ElseIf file_name > info_head.strFileName Then
            If min = mid Then
                min = min + 1
            Else
                min = mid
            End If
        Else
            'Copy info head
            File_Find = info_head
            
            'Close file and exit
            Close file_handler
            Exit Function
        End If
    Loop
    
ErrHandler:
    'Close file
    Close file_handler
    File_Find.strFileName = ""
    File_Find.lngFileSize = 0
End Function
Public Sub General_Quick_Sort(ByRef SortArray As Variant, ByVal first As Long, ByVal last As Long)
'**************************************************************
'Author: juan Martín Sotuyo Dodero
'Last Modify Date: 3/03/2005
'Good old QuickSort algorithm :)
'**************************************************************
    Dim Low As Long, High As Long
    Dim temp As Variant
    Dim List_Separator As Variant
    
    Low = first
    High = last
    List_Separator = SortArray((first + last) / 2)
    Do While (Low <= High)
        Do While SortArray(Low) < List_Separator
            Low = Low + 1
        Loop
        Do While SortArray(High) > List_Separator
            High = High - 1
        Loop
        If Low <= High Then
            temp = SortArray(Low)
            SortArray(Low) = SortArray(High)
            SortArray(High) = temp
            Low = Low + 1
            High = High - 1
        End If
    Loop
    If first < High Then General_Quick_Sort SortArray, first, High
    If Low < last Then General_Quick_Sort SortArray, Low, last
End Sub


Public Function General_Drive_Get_Free_Bytes(ByVal DriveName As String) As Currency


    Dim RetVal As Long
    Dim FB As Currency
    Dim BT As Currency
    Dim FBT As Currency
    
    RetVal = GetDiskFreeSpace(Left(DriveName, 2), FB, BT, FBT)
    
    General_Drive_Get_Free_Bytes = FB * 10000 'convert result to actual size in bytes
End Function




Public Function LoadInterface(ByVal filename As String) As IPictureDisp

On Error GoTo ErrorHandler

Dim LowerBound As Long
Dim ByteCount As Long
Dim hMem As Long
Dim lpMem As Long
Dim IID_IPicture(15) As Long
Dim istm As stdole.IUnknown
Dim bytArr() As Byte


Extract_BMP_Memory LCase$(filename) & ".bmp", bytArr()


LowerBound = LBound(bytArr)
ByteCount = (UBound(bytArr) - LowerBound) + 1
hMem = GlobalAlloc(&H2, ByteCount)



If hMem <> 0 Then
    lpMem = GlobalLock(hMem)
    If lpMem <> 0 Then
        MoveMemory ByVal lpMem, bytArr(LowerBound), ByteCount
        Call GlobalUnlock(hMem)
        If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then
            If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                Call OleLoadPicture(ByVal ObjPtr(istm), ByteCount, 0, IID_IPicture(0), LoadInterface)
            End If
        End If
    End If
End If

Exit Function

ErrorHandler:
    'MsgBox Err.Description
    

End Function

Public Function Extract_BMP_Memory(ByVal file_name As String, SourceData() As Byte) As Boolean
Dim loopc As Long
Dim InfoHead As INFOHEADER
Dim handle As Integer

'Set up the error handler
On Local Error GoTo ErrHandler

    
    'Find the Info Head of the desired file
    InfoHead = File_Find(resource_path & "Interface.iao", file_name)
 

    If InfoHead.strFileName = "" Or InfoHead.lngFileSize = 0 Then Exit Function

    'Open the binary file
    handle = FreeFile
    Open resource_path & "Interface.iao" For Binary Access Read Lock Write As handle
    
        'Resize the byte data array
        ReDim SourceData(InfoHead.lngFileSize - 1)
        
        'Get the data
        Get handle, InfoHead.lngFileStart, SourceData
      
                '  Desencrypt File Header
        

        
        'Decompress all data
        Decompress_Data SourceData, InfoHead.lngFileSizeUncompressed
        

    'Close the binary file
    Close handle

    Extract_BMP_Memory = True
Exit Function

ErrHandler:
    Close handle
    Erase SourceData
End Function

Public Sub Encrypt_File_Header(ByRef FileHead As FILEHEADER)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Encrypts normal data or turns encrypted data back to normal
'*****************************************************************
    'Each different variable is encrypted with a different key for extra security
    With FileHead
        .intNumFiles = .intNumFiles Xor 4725
        .lngFileSize = .lngFileSize Xor 2435437
    End With
End Sub

Public Sub Encrypt_Info_Header(ByRef InfoHead As INFOHEADER)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Encrypts normal data or turns encrypted data back to normal
'*****************************************************************
    Dim EncryptedFileName As String
    Dim loopc As Long
    
    For loopc = 1 To Len(InfoHead.strFileName)
        If loopc Mod 2 = 0 Then
            EncryptedFileName = EncryptedFileName & Chr$(Asc(mid$(InfoHead.strFileName, loopc, 1)) Xor 161)
        Else
            EncryptedFileName = EncryptedFileName & Chr$(Asc(mid$(InfoHead.strFileName, loopc, 1)) Xor 47)
        End If
    Next loopc
    
    'Each different variable is encrypted with a different key for extra security
    With InfoHead
        .lngFileSize = .lngFileSize Xor 689947
        .lngFileSizeUncompressed = .lngFileSizeUncompressed Xor 445275
        .lngFileStart = .lngFileStart Xor 87540
        .strFileName = EncryptedFileName
    End With
End Sub


Public Function Extract_BMP_Memory2(ByVal file_name As String, SourceData() As Byte) As Boolean
Dim loopc As Long
Dim InfoHead As INFOHEADER
Dim handle As Integer

'Set up the error handler
On Local Error GoTo ErrHandler

    'Find the Info Head of the desired file
    InfoHead = File_Find(resource_path & "Graficos.iao", file_name)

    
    If InfoHead.strFileName = "" Or InfoHead.lngFileSize = 0 Then Exit Function

    'Open the binary file
    handle = FreeFile
    Open resource_path & "Graficos.iao" For Binary Access Read Lock Write As handle
    
        'Resize the byte data array
        ReDim SourceData(InfoHead.lngFileSize - 1)
        
        'Get the data
        Get handle, InfoHead.lngFileStart, SourceData



        'Decompress all data
        Decompress_Data SourceData, InfoHead.lngFileSizeUncompressed
        

    'Close the binary file
    Close handle

    Extract_BMP_Memory2 = True
Exit Function

ErrHandler:
    Close handle
    Erase SourceData
End Function

