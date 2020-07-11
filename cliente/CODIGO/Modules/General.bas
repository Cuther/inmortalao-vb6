Attribute VB_Name = "modGeneral"
Option Explicit
Public cCursores As clsCursor

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
'Declaración del Api SetLayeredWindowAttributes que establece _
 la transparencia al form
 
Public ActualSecond As Long
Public LastSecond As Long



Private Declare Function foo Lib "InmortalDLL.dll" () As Integer

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
                
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    
Public bFogata As Boolean

Type localeObj
    name As String
    tipe As Byte
    Grh As Long
End Type
Public objs() As localeObj
Public numObjs As Long
'***********************************************

'***********************************************
Type localeNpc
    name As String
    Desc As String
End Type
Public npcs() As localeNpc
Public numNpcs As Long
'***********************************************

'***********************************************
Type localeSpl
    name As String
    t1 As String
    t2 As String
    t3 As String
    Desc As String
End Type
Public spells() As localeSpl
Public numSpells As Long
'***********************************************

'***********************************************
Public MapNames(1 To 851) As String
Public MapTable(1 To 30, 1 To 23) As Integer
'***********************************************

Sub AddToChat(ByRef Text As String, Optional ByVal Font As Byte = 11, Optional ByVal Console As Byte = 1)
    With FontTypes(Font)
        If Console = 1 Then 'Chat
            AddtoRichTextBox frmMain.RecChat, Text, .red, .green, .blue, .bold, .italic
        ElseIf Console = 2 Then 'Combate
            AddtoRichTextBox frmMain.RecCombat, Text, .red, .green, .blue, .bold, .italic
        Else 'Global
            AddtoRichTextBox frmMain.RecGlobal, Text, .red, .green, .blue, .bold, .italic
        End If
    End With
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)

    With RichTextBox
        If Len(.Text) > 900 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)

    End With
End Sub
Sub UnloadAllForms()
    On Error Resume Next
    Dim mifrm As Form
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    'Unload the connect form
    Unload frmCrearPersonaje
    Unload frmConnect
    Unload frmPanelAccount
    
    frmMain.lblNick = UserName
    LoadMacros
    Dim i As Integer
    For i = 1 To 11
        frmBindKey.DibujarMenuMacros i
    Next i
    
    'Load main form
    frmMain.Visible = True

End Sub




Private Sub CheckKeys()
    If Not IsAppActive() Then Exit Sub
    
    If Comerciando Then Exit Sub

    If frmForo.Visible Or frmMap.Visible Or frmCorreo.Visible Then Exit Sub

    If Pausa Then Exit Sub
    
    'Don't allow any these keys during movement..
    If Not TileEngine.scroll_on Then
        If Not UserEstupido Then
            'Move Up
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                Call TileEngine.Engine_Move(NORTH)
                Exit Sub
            End If
            
            'Move Right
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                Call TileEngine.Engine_Move(EAST)
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                Call TileEngine.Engine_Move(SOUTH)
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                Call TileEngine.Engine_Move(WEST)
                Exit Sub
            End If
            
            ' We haven't moved - Update 3D sounds!
            Call Audio.MoveListener(UserPos.X, UserPos.Y)
            If BloqMov And BloqDir <> 0 Then
                Call TileEngine.Engine_Move(BloqDir, False)
            End If
        Else
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            
            If kp Then
                Call TileEngine.Engine_Move(General_Random_Number(NORTH, WEST))
                Call Audio.MoveListener(UserPos.X, UserPos.Y)
                
                Call DibujarMiniMapPos
            End If
        End If
    End If
End Sub

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        count = count + 1
    Loop While curPos <> 0
    
    FieldCount = count
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Sub Main()
On Error GoTo Err


    
    InicializarVariables

    frmMain.Winsock1.Close
    frmMain.Winsock1.Protocol = sckUDPProtocol
   '' frmMain.Winsock1.Bind 5000
    
    frmMain.Winsock1.RemoteHost = "192.168.0.3"
    frmMain.Winsock1.RemotePort = 5000

    'AoDefAntiShInitialize
    
    'If AoDefDebugger Then
    '    Call AoDefAntiDebugger
    '    End
    'End If
    
    'Activar IMPORTANTE
    'If AoDefMultiClient Then
        'Call AoDefMultiClientOn
        'End
    'End If
    
    'Add Marius Actualizamos el Luncher
    'If FileExist(App.Path & "\Launcher2.exe", vbNormal) Then
    '    If FileExist(App.Path & "\Launcher.exe", vbNormal) Then
    '        Call Kill(App.Path & "\Launcher.exe")
    '    End If
    '    'Renombar Launcher2.exe a Launcher.exe
    '    Call FileCopy(App.Path + "\Launcher2.exe", App.Path + "\Launcher.exe")
    '    Call Kill(App.Path & "\Launcher2.exe")
    'End If
    '\Add
    
    
    Load frmCargando
    frmCargando.Visible = True
    DoEvents

    
    Protocol.InitFonts
    LoadConfig

    cCursores.Init
    cCursores.Parse_Form frmCargando
    cCursores.Parse_Form frmCargando, E_WAIT

    If Not (TileEngine.Engine_Init) Then End


    frmCargando.SetAddWith 30
    
    
    LoadLocales

    Call Audio.Initialize(frmMain.hwnd, App.Path & "\WAV\", App.Path & "\MIDI\")
    Audio.MusicVolume = VolumeMusic
    Audio.SoundVolume = VolumeSound
    Call Audio.PlayMusic("01")
        
    frmCargando.SetAddWith 25
        
        Call Inventario.Initialize(frmMain.picInv)
        'frmMain.Socket1.Startup
    
        'Inicialización de variables globales
        perm = True
        prgRun = True
        Pausa = False
        
    frmCargando.SetAddWith 20
        
        'Set the intervals of timers
        Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
        Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
        Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
        Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
        Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
        Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
        Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
        Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
        
    frmCargando.SetAddWith 15
    
        
         
        'Init timers
         Call MainTimer.Start(TimersIndex.Attack)
         Call MainTimer.Start(TimersIndex.Work)
         Call MainTimer.Start(TimersIndex.UseItemWithU)
         Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
         Call MainTimer.Start(TimersIndex.SendRPU)
         Call MainTimer.Start(TimersIndex.CastSpell)
         Call MainTimer.Start(TimersIndex.Arrows)
         Call MainTimer.Start(TimersIndex.CastAttack)
         
    frmCargando.SetAddWith 10
    frmCargando.Picture = LoadInterface("iniciando")
    frmCargando.picLoad.Visible = False
    DoEvents
    
    frmCargando.SetAddWith 5
    
    Sleep 1000
    
    cCursores.Parse_Form frmCargando
    
    Unload frmCargando
    
    cCursores.Parse_Form frmConnect
    frmCargando.SetAddWith 0
    
    frmConnect.Visible = True

    thFPSAndHour = SetTimer(0, 0, 1000, AddressOf Timer_HoraAndFPS)
    
    
    Call TileEngine.Map_Load(1)
    
    Call Timer_HoraAndFPS
    Do While prgRun
        If IsAppActive Then
            If frmMain.Visible Then
                Call TileEngine.Engine_Render
                Call CheckKeys
                
                If RenderInv Then
                    TileEngine.Inventory_Render
                End If
            End If
        Else
            If frmMain.Visible Then
                RenderInv = True
            End If
            
            Sleep 100
        End If
'Call outgoingData.ReadASCIIStringFixed(outgoingData.length)
                FlushBuffer
        DoEvents
    Loop
    
    Call CloseClient
    
    Exit Sub
Err:
   ' LogError Err.Description & " Numero " & Err.Number
    Resume Next
    
End Sub

Private Sub InicializarVariables()

    'Set TileEngine = New clsTileEngineX
    Set Audio = New clsAudio
    Set Inventario = New clsGrapchicalInventory
    Set incomingData = New clsByteQueue
    Set outgoingData = New clsByteQueue
    Set MainTimer = New clsTimer
    Set CustomKeys = New clsCustomKeys
    Set cCursores = New clsCursor
    
    Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
    Ciudades(eCiudad.cNix) = "Nix"
    Ciudades(eCiudad.cBanderbill) = "Banderbill"
    Ciudades(eCiudad.cLindos) = "Lindos"
    Ciudades(eCiudad.cArghal) = "Arghâl"
    Ciudades(eCiudad.cIlliandor) = "Illiandor"
    Ciudades(eCiudad.cNuevaEsperanza) = "Nueva Esperanza"
    Ciudades(eCiudad.cOrac) = "Orac"
    Ciudades(eCiudad.cRinkel) = "Rinkel"
    Ciudades(eCiudad.cSuramei) = "Suramei"
    Ciudades(eCiudad.cTiama) = "Tiama"
    
    ListaRazas(eRaza.HUMANO) = "Humano"
    ListaRazas(eRaza.ELFO) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Drow"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    ListaRazas(eRaza.Orco) = "Orco"

    ListaClases(eClass.Mago) = "Mago"
    ListaClases(eClass.Clerigo) = "Clerigo"
    ListaClases(eClass.Guerrero) = "Guerrero"
    ListaClases(eClass.Asesino) = "Asesino"
    ListaClases(eClass.Ladron) = "Ladron"
    ListaClases(eClass.Bardo) = "Bardo"
    ListaClases(eClass.Druida) = "Druida"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Cazador) = "Cazador"
    ListaClases(eClass.Pescador) = "Pescador"
    ListaClases(eClass.Herrero) = "Herrero"
    ListaClases(eClass.Leñador) = "Leñador"
    ListaClases(eClass.Minero) = "Minero"
    ListaClases(eClass.Carpintero) = "Carpintero"
    ListaClases(eClass.Mercenario) = "Mercenario"
    ListaClases(eClass.Nigromante) = "Nigromante"
    ListaClases(eClass.Sastre) = "Sastre"
    ListaClases(eClass.Gladiador) = "Gladiador"
    
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Tacticas de combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar arboles"
    SkillsNames(eSkill.Comercio) = "Comercio"
    SkillsNames(eSkill.DefensaEscudos) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Armas de proyectiles"
    SkillsNames(eSkill.Artes) = "Artes Marciales"
    SkillsNames(eSkill.Navegacion) = "Navegacion"
    SkillsNames(eSkill.Alquimia) = "Alquimia"
    SkillsNames(eSkill.Arrojadizas) = "Armas Arrojadizas"
    SkillsNames(eSkill.Botanica) = "Botanica"
    SkillsNames(eSkill.Equitacion) = "Equitacion"
    SkillsNames(eSkill.Musica) = "Musica"
    SkillsNames(eSkill.Resistencia) = "Resistencia Magica"
    SkillsNames(eSkill.Sastreria) = "Sastreria"

    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Carisma) = "Carisma"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"
    
    ReDim Head_Range(1 To NUMRAZAS) As tHeadRange

    Head_Range(HUMANO).mStart = 1
    Head_Range(HUMANO).mEnd = 30
    Head_Range(Enano).mStart = 301
    Head_Range(Enano).mEnd = 315
    Head_Range(ELFO).mStart = 101
    Head_Range(ELFO).mEnd = 121
    Head_Range(ElfoOscuro).mStart = 202
    Head_Range(ElfoOscuro).mEnd = 212
    Head_Range(Gnomo).mStart = 401
    Head_Range(Gnomo).mEnd = 409
    Head_Range(Orco).mStart = 501
    Head_Range(Orco).mEnd = 514

    Head_Range(HUMANO).fStart = 70
    Head_Range(HUMANO).fEnd = 80
    Head_Range(Enano).fStart = 370
    Head_Range(Enano).fEnd = 373
    Head_Range(ELFO).fStart = 170
    Head_Range(ELFO).fEnd = 189
    Head_Range(ElfoOscuro).fStart = 270
    Head_Range(ElfoOscuro).fEnd = 278
    Head_Range(Gnomo).fStart = 470
    Head_Range(Gnomo).fEnd = 481
    Head_Range(Orco).fStart = 570
    Head_Range(Orco).fEnd = 573
    
    MouseS = General_Get_Mouse_Speed


    mueve = 1
    resource_path = App.Path & "\Resources\"
    
    'Musica inicial mp3
    MIDI_ACTIVATE = IIf(FileExist(resource_path & "Music\01.mp3", vbNormal), 0, 1)
    
End Sub

''
' Removes all text from the console and dialogs
Public Sub Auto_Drag(ByVal hwnd As Long)
    Call ReleaseCapture
    Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
End Sub
Public Sub CloseClient()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 8/14/2007
'Frees all used resources, cleans up and leaves
'**************************************************************
    'Stop tile engine
    Call TileEngine.Engine_End
    
    'Destruimos los objetos públicos creados
    Set CustomKeys = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    
    Call UnloadAllForms
    
    KillTimer 0, thFPSAndHour
    
    End
End Sub
Public Function Map_NameLoad(ByVal map_num As Integer) As String
On Error GoTo ErrorHandler
    
    Map_NameLoad = MapNames(map_num)
    Exit Function

ErrorHandler:
    Map_NameLoad = "Mapa Desconocido"

End Function
Public Sub General_Long_Color_to_RGB(ByVal long_color As Long, ByRef red As Integer, ByRef green As Integer, ByRef blue As Integer)
'***********************************
'Coded by Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modified: 2/19/03
'Takes a long value and separates RGB values to the given variables
'***********************************
    Dim temp_color As String
    
    temp_color = Hex(long_color)
    If Len(temp_color) < 6 Then
        'Give is 6 digits for easy RGB conversion.
        temp_color = String(6 - Len(temp_color), "0") + temp_color
    End If
    
    red = CLng("&H" + mid$(temp_color, 1, 2))
    green = CLng("&H" + mid$(temp_color, 3, 2))
    blue = CLng("&H" + mid$(temp_color, 5, 2))

End Sub
Function getGrhPet(ByVal tipe As Byte) As Long
    Select Case tipe
        Case eMascota.Ely
            getGrhPet = 851
        Case eMascota.fuego
            getGrhPet = 12220
        Case eMascota.Agua
            getGrhPet = 12256
        Case eMascota.Tierra
            getGrhPet = 12184
        Case eMascota.Fatuo
            getGrhPet = 266
        Case eMascota.Tigre
            getGrhPet = 12032
        Case eMascota.Lobo
            getGrhPet = 2262
        Case eMascota.Oso
            getGrhPet = 4404
        Case eMascota.Ent
            getGrhPet = 16730
    End Select
End Function

''
' Splits a text into several lines to make it comply with the MAX_LENGTH unless it's impossible (a single word longer than MAX_LENGTH).
'
' @param    chat    The text to be formated.
'
' @return   The array of lines into which the text is splitted.
'
' @see      MAX_LENGTH

Public Function FormatChat(ByRef chat As String) As String()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Formats a dialog into different text lines.
'**************************************************************
    Dim word As String
    Dim curPos As Long
    Dim length As Long
    Dim acumLength As Long
    Dim lineLength As Long
    Dim wordLength As Long
    Dim curLine As Long
    Dim chatLines() As String
    
    'Initialize variables
    curLine = 0
    curPos = 1
    length = Len(chat)
    acumLength = 0
    lineLength = -1
    ReDim chatLines(FieldCount(chat, 32)) As String
    
    'Start formating
    Do While acumLength < length
        word = ReadField(curPos, chat, 32)
        
        wordLength = Len(word)
        
        ' Is the first word of the first line? (it's the only that can start at -1)
        If lineLength = -1 Then
            chatLines(curLine) = word
            
            lineLength = wordLength
            acumLength = wordLength
        Else
            ' Is the word too long to fit in this line?
            If lineLength + wordLength + 1 > 18 Then
                'Put it in the next line
                curLine = curLine + 1
                chatLines(curLine) = word
                
                lineLength = wordLength
            Else
                'Add it to this line
                chatLines(curLine) = chatLines(curLine) & " " & word
                
                lineLength = lineLength + wordLength + 1
            End If
            
            acumLength = acumLength + wordLength + 1
        End If
        
        'Increase to search for next word
        curPos = curPos + 1
    Loop
    
    ' If it's only one line, center text
    If curLine = 0 And length < 18 Then
        chatLines(curLine) = String((18 - length) \ 2 + 1, " ") & chatLines(curLine)
        chatLines(curLine) = RTrim$(LTrim$(chatLines(curLine)))
    End If
    
    'Resize array to fit
    ReDim Preserve chatLines(curLine) As String
    
    FormatChat = chatLines
End Function
Public Function General_Get_Mouse_Speed() As Long
    SystemParametersInfo SPI_GETMOUSESPEED, 0, General_Get_Mouse_Speed, 0
End Function

Public Sub General_Set_Mouse_Speed(ByVal lngSpeed As Long)
    SystemParametersInfo SPI_SETMOUSESPEED, 0, ByVal lngSpeed, 0
End Sub
Public Function General_Locale_Name_Spell(ByVal num As Long) As String

    If num = 0 Then
        General_Locale_Name_Spell = "(None)"
        Exit Function
    End If
    
    General_Locale_Name_Spell = spells(num).name
End Function
Public Function General_Locale_Name_Obj(ByVal num As Long) As String
    If num = 0 Then
        General_Locale_Name_Obj = ""
        Exit Function
    End If
    
    General_Locale_Name_Obj = objs(num).name
End Function
Public Sub LoadLocales()
    On Error Resume Next
    Dim f As Integer
    Dim i As Long
    Dim tmpStr As String
    
    '***********************************************

    Extract_File Scripts, "locale_spl_es.ind", resource_path
    
    f = FreeFile
    Open resource_path & "locale_spl_es.ind" For Input As #f
        
        ReDim spells(1 To General_Get_Line_Count(resource_path & "locale_spl_es.ind")) As localeSpl
        i = 0
        Do While Not EOF(f)
            i = i + 1
            Line Input #f, tmpStr
            spells(i).name = ReadField(1, tmpStr, Asc("|"))
            spells(i).Desc = ReadField(2, tmpStr, Asc("|"))
            spells(i).t1 = ReadField(3, tmpStr, Asc("|"))
            spells(i).t2 = ReadField(4, tmpStr, Asc("|"))
            spells(i).t3 = ReadField(5, tmpStr, Asc("|"))
        Loop
    Close #f
    
    tmpStr = ""
    Delete_File resource_path & "locale_spl_es.ind"
    '***********************************************

    '***********************************************
    Extract_File Scripts, "locale_obj_es.ind", resource_path
    
    f = FreeFile
    Open resource_path & "locale_obj_es.ind" For Input As #f
        
        ReDim objs(1 To General_Get_Line_Count(resource_path & "locale_obj_es.ind")) As localeObj
        i = 0
        Do While Not EOF(f)
            i = i + 1
            Line Input #f, tmpStr
            objs(i).name = ReadField(1, tmpStr, Asc("|"))
            objs(i).Grh = val(ReadField(2, tmpStr, Asc("|")))
            objs(i).tipe = val(ReadField(3, tmpStr, Asc("|")))
        Loop
    Close #f
    
    tmpStr = ""
    Delete_File resource_path & "locale_obj_es.ind"
    '***********************************************

    '***********************************************
    Extract_File Scripts, "locale_npc_es.ind", resource_path
    
    f = FreeFile
    Open resource_path & "locale_npc_es.ind" For Input As #f
        
        ReDim npcs(1 To General_Get_Line_Count(resource_path & "locale_npc_es.ind")) As localeNpc
        i = 0
        Do While Not EOF(f)
            i = i + 1
            Line Input #f, tmpStr
            npcs(i).name = ReadField(1, tmpStr, Asc("|"))
            npcs(i).Desc = ReadField(2, tmpStr, Asc("|"))
        Loop
    Close #f
    
    Delete_File resource_path & "locale_npc_es.ind"
    '***********************************************

    '***********************************************
    Extract_File Scripts, "table.ind", resource_path
    
    f = FreeFile
    Open resource_path & "table.ind" For Binary As #f
        Get #f, , MapTable
    Close #f
    
    Delete_File resource_path & "table.ind"
    '***********************************************

    '***********************************************
    Extract_File Scripts, "mapa.as", resource_path
    
    f = FreeFile
    Open resource_path & "mapa.as" For Input As #f
        For i = 1 To 851
            Line Input #f, MapNames(i)
            MapNames(i) = RTrim$(MapNames(i))
        Next i
    Close #f
    
    Delete_File resource_path & "mapa.as"
    
End Sub
Public Function CleanClient()
    Dim i As Long
    BloqDir = 0
    BloqMov = False
    
    Pausa = False
    UserMeditar = False
    UserNavegando = False
    UserMontando = False
    UserDescansar = False
    UserParalizado = False
    UserCiego = False
    UserMeditar = False
    bFogata = False
    IScombate = False
    
    'Add Marius
    frmMain.modocombate.Visible = False
    frmMain.nomodocombate.Visible = True
    
    UserCiego = False
    UserEstupido = False
    '\Add
    
    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    SkillPoints = 0
    Alocados = 0
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i
    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    
    UserEmail = ""
    TileEngine.Engine_Scroll_Pixels 5.2

    Call Audio.StopWave

    frmMain.modocombate.Visible = False
    frmMain.nomodocombate.Visible = True
    frmMain.Visible = False

    For i = 1 To LastChar
        Call TileEngine.Char_Reset_Info(i)
        Call TileEngine.Char_Dialog_Remove(i)
    Next i
    
End Function
Public Function General_Get_Line_Count(ByVal filename As String) As Long
On Error GoTo ErrorHandler
    Dim N As Integer, tmpStr As String
    If LenB(filename) Then
        N = FreeFile()
        
        Open filename For Input As #N
            Do While Not EOF(N)
                General_Get_Line_Count = General_Get_Line_Count + 1
                Line Input #N, tmpStr
            Loop
        Close N
    End If
    Exit Function

ErrorHandler:
    Resume Next
    
End Function
Function GetElapsedTime() As Single
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
    GetElapsedTime = (start_time - end_time) / timer_freq
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Public Function MostrarCantidad(ByVal i As Integer) As Boolean
    MostrarCantidad = objs(i).tipe <> eObjType.otPuertas And _
            objs(i).tipe <> eObjType.otForos And _
            objs(i).tipe <> eObjType.otCarteles And _
            objs(i).tipe <> eObjType.otArboles And _
            objs(i).tipe <> eObjType.otYacimiento And _
            objs(i).tipe <> eObjType.otTeleport And _
            objs(i).tipe <> eObjType.otCorreo
End Function
Public Function DirInt() As String
    DirInt = resource_path & "Interface\"
End Function

Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
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
Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function
Public Function General_Random_Number(ByVal LowerBound As Long, ByVal UpperBound As Long) As Single
    Randomize Timer
    General_Random_Number = (UpperBound - LowerBound) * Rnd + LowerBound
End Function



Public Sub DibujarMiniMapPos()

    frmMain.UserP.Left = UserPos.X - 1
    frmMain.UserP.Top = UserPos.Y - 1
    frmMain.Minimap.refresh
    
    frmMain.Label2(0).Caption = "Posición: " & UserMap & ", " & UserPos.X & ", " & UserPos.Y
    
End Sub

Public Function IsAppActive() As Boolean
    IsAppActive = GetActiveWindow
End Function
Public Function LogError(Desc As String)
On Error Resume Next
    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\errores.log" For Append Shared As #nfile
        Print #nfile, Date & " " & time & " " & Desc
    Close #nfile
End Function


'Configuracion personalizada
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
    
    'SACAR DESPUES es para testear la sincronizacion vertical para FPS mas altos
    ' JOSE CASTELLI
 'Sinc = 0
    
End Sub
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
    CursoresStandar = 0
    
    ChatGlobal = 1
    ChatFaccionario = 1
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
        CursoresStandar = Cursores

        Get #f, , ChatFaccionario
        Get #f, , ChatGlobal
    Close #f
    
    Sinc = 0
 
End Sub
Public Sub LoadMacros()
    Dim LC As Byte
    Dim Leer As clsIniReader: Set Leer = New clsIniReader
    
    ReDim Preserve MacroKeys(1 To 11) As tBoton
    
    If Not FileExist(App.Path & "\Init\" & UserName & ".dat", vbNormal) Then
        Open App.Path & "\Init\" & UserName & ".dat" For Append As #1
                Print #1, "[" & UserName & "]"
            For LC = 1 To 11
            MacroKeys(LC).TipoAccion = 0
            MacroKeys(LC).invslot = 0
            MacroKeys(LC).SendString = ""
            MacroKeys(LC).hlist = 0
                Print #1, "Accion" & LC & "=" & MacroKeys(LC).TipoAccion
                Print #1, "InvSlot" & LC & "=" & MacroKeys(LC).invslot
                Print #1, "SndString" & LC & "=" & MacroKeys(LC).SendString
                Print #1, "Hlist" & LC & "=" & MacroKeys(LC).hlist
                    
            Next LC
            
            Print #1, "" 'Separacion entre macro y macro
        Close #1
    End If
    
    Leer.Initialize App.Path & "\Init\" & UserName & ".dat"
    For LC = 1 To 11
        MacroKeys(LC).TipoAccion = val(Leer.GetValue(UserName, "Accion" & LC))
        MacroKeys(LC).hlist = val(Leer.GetValue(UserName, "hlist" & LC))
        MacroKeys(LC).invslot = val(Leer.GetValue(UserName, "invslot" & LC))
        MacroKeys(LC).SendString = Leer.GetValue(UserName, "SndString" & LC)
    Next LC
    Set Leer = Nothing
End Sub
Public Sub LoadGrhData()
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim tempInt As Integer
Dim f As Integer

ReDim GrhData(0 To CANT_GRH_INDEX) As GrhData

Extract_File Scripts, "graficos.ind", resource_path

f = FreeFile()
Open resource_path & "graficos.ind" For Binary Access Read As #f
    
    Seek #f, 1
    
    Get #f, , tempInt
    Get #f, , tempInt
    Get #f, , tempInt
    Get #f, , tempInt
    Get #f, , tempInt

    'Get first Grh Number
    Get #f, , Grh
    
    Do Until Grh <= 0
        'Get number of frames
        Get #f, , GrhData(Grh).NumFrames
        
        If GrhData(Grh).NumFrames <= 0 Then
            GoTo ErrorHandler
        End If
        
        ReDim GrhData(Grh).Frames(1 To GrhData(Grh).NumFrames)
  
        If GrhData(Grh).NumFrames > 1 Then
        
            'Read a animation GRH set
            For Frame = 1 To GrhData(Grh).NumFrames
                Get #f, , GrhData(Grh).Frames(Frame)
                If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > CANT_GRH_INDEX Then GoTo ErrorHandler
            Next Frame
        
            Get #f, , tempInt
            
            If tempInt <= 0 Then GoTo ErrorHandler
            GrhData(Grh).speed = tempInt

            'Compute width and height
            GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
            If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
            
            GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
            If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler

            GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
            If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler

            GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
            If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
        Else
            'Read in normal GRH data
            Get #f, , GrhData(Grh).FileNum
            If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler

            Get #f, , GrhData(Grh).sX
            If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
            
            Get #f, , GrhData(Grh).sY
            If GrhData(Grh).sY < 0 Then GoTo ErrorHandler

            Get #f, , GrhData(Grh).pixelWidth
            If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler

            Get #f, , GrhData(Grh).pixelHeight
            If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler

            'Compute width and height
            GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / 32
            GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / 32
            
            GrhData(Grh).Frames(1) = Grh
        End If
        'Get Next Grh Number
        Get #f, , Grh
    Loop
    
Close #f

Delete_File resource_path & "Graficos.ind"
If FileExist(resource_path & "Graficos.ind", vbNormal) Then Kill resource_path & "Graficos.ind"

Extract_File Scripts, "minimap.dat", resource_path

Dim count As Long
f = FreeFile
Open resource_path & "minimap.dat" For Binary As #f
    Seek #1, 1
    For count = 1 To CANT_GRH_INDEX
        If Grh_Check(count) Then
            Get #f, , GrhData(count).mini_map_color
        End If
    Next count
Close #f

Delete_File resource_path & "minimap.dat"
If FileExist(resource_path & "minimap.dat", vbNormal) Then Kill resource_path & "minimap.dat"

Exit Sub

ErrorHandler:
    Close #1
    MsgBox "Error al cargar el recurso de índice de gráficos: " & Err.Description & " (" & Grh & ")", vbCritical, "Error al cargar"

End Sub


Public Sub CargarParticulas()
    Dim loopc As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
    Dim Leer As New clsIniReader

    Dim StreamFile As String

    Extract_File Scripts, "particulas.ini", resource_path

    StreamFile = resource_path & "particulas.ini"
    
    Leer.Initialize StreamFile

    TotalStreams = val(Leer.GetValue("INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
    
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams
        StreamData(loopc).name = Leer.GetValue(val(loopc), "Name")
        StreamData(loopc).NumOfParticles = Leer.GetValue(val(loopc), "NumOfParticles")
        StreamData(loopc).x1 = Leer.GetValue(val(loopc), "X1")
        StreamData(loopc).y1 = Leer.GetValue(val(loopc), "Y1")
        StreamData(loopc).x2 = Leer.GetValue(val(loopc), "X2")
        StreamData(loopc).y2 = Leer.GetValue(val(loopc), "Y2")
        StreamData(loopc).angle = Leer.GetValue(val(loopc), "Angle")
        StreamData(loopc).vecx1 = Leer.GetValue(val(loopc), "VecX1")
        StreamData(loopc).vecx2 = Leer.GetValue(val(loopc), "VecX2")
        StreamData(loopc).vecy1 = Leer.GetValue(val(loopc), "VecY1")
        StreamData(loopc).vecy2 = Leer.GetValue(val(loopc), "VecY2")
        StreamData(loopc).life1 = Leer.GetValue(val(loopc), "Life1")
        StreamData(loopc).life2 = Leer.GetValue(val(loopc), "Life2")
        StreamData(loopc).friction = Leer.GetValue(val(loopc), "Friction")
        StreamData(loopc).spin = Leer.GetValue(val(loopc), "Spin")
        StreamData(loopc).spin_speedL = Leer.GetValue(val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = Leer.GetValue(val(loopc), "Spin_SpeedH")
        StreamData(loopc).AlphaBlend = Leer.GetValue(val(loopc), "AlphaBlend")
        StreamData(loopc).gravity = Leer.GetValue(val(loopc), "Gravity")
        StreamData(loopc).grav_strength = Leer.GetValue(val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = Leer.GetValue(val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = Leer.GetValue(val(loopc), "XMove")
        StreamData(loopc).YMove = Leer.GetValue(val(loopc), "YMove")
        StreamData(loopc).move_x1 = Leer.GetValue(val(loopc), "move_x1")
        StreamData(loopc).move_x2 = Leer.GetValue(val(loopc), "move_x2")
        StreamData(loopc).move_y1 = Leer.GetValue(val(loopc), "move_y1")
        StreamData(loopc).move_y2 = Leer.GetValue(val(loopc), "move_y2")
        StreamData(loopc).life_counter = Leer.GetValue(val(loopc), "life_counter")
        StreamData(loopc).speed = val(Leer.GetValue(val(loopc), "Speed"))
        
        Dim temp As Integer
        temp = Leer.GetValue(val(loopc), "resize")
        
        StreamData(loopc).grh_resize = IIf((temp = -1), True, False)
        StreamData(loopc).grh_resizex = Leer.GetValue(val(loopc), "rx")
        StreamData(loopc).grh_resizey = Leer.GetValue(val(loopc), "ry")
        
        
        StreamData(loopc).NumGrhs = Leer.GetValue(val(loopc), "NumGrhs")
        
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = Leer.GetValue(val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = ReadField(str(i), GrhListing, 44)
        Next i
        
        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        
        For ColorSet = 1 To 4
            TempSet = Leer.GetValue(val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).r = ReadField(1, TempSet, 44)
            StreamData(loopc).colortint(ColorSet - 1).g = ReadField(2, TempSet, 44)
            StreamData(loopc).colortint(ColorSet - 1).b = ReadField(3, TempSet, 44)
        Next ColorSet
                
    Next loopc
    
    Set Leer = Nothing
    Delete_File resource_path & "particulas.ini"
    If FileExist(resource_path & "particulas.ini", vbNormal) Then Kill resource_path & "particulas.ini"
    

End Sub
Private Function Grh_Check(ByVal grh_index As Long) As Boolean
    If grh_index > 0 And grh_index <= CANT_GRH_INDEX Then
        Grh_Check = (GrhData(grh_index).NumFrames > 0)
    End If
End Function

Public Sub Timer_HoraAndFPS()
    If Connected Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 0

        tSeg = tSeg + 1
        If tSeg >= 15 Then
            tMinuto = tMinuto + 1
            tSeg = 0
            frmMain.imgHora.ToolTipText = TileEngine.Get_Time_String
            If tMinuto >= 60 Then
                tMinuto = 0
                tHora = tHora + 1
                If tHora = 24 Then tHora = 0
            End If
        End If
    End If
End Sub
Function Generate_Char_Status(ByVal PercVida As Long, ByVal Paralizado As Byte, ByVal Inmovilizado As Byte, ByVal Envenenado As Byte, Optional ByVal Trabajando As Byte = 0, Optional ByVal Silenciado As Byte = 0, Optional ByVal Ciego As Byte = 0, Optional ByVal Incinerado As Byte = 0, Optional ByVal Transformado As Byte = 0, Optional ByVal Comerciando As Byte = 0, Optional ByVal Inactivo As Byte = 0) As String

If PercVida <> -1 Then
    If PercVida = 100 Then
        Generate_Char_Status = "| Intacto "
    ElseIf PercVida >= 80 Then
        Generate_Char_Status = "| Levemente Herido "
    ElseIf PercVida >= 50 Then
        Generate_Char_Status = "| Herido "
    ElseIf PercVida >= 30 Then
        Generate_Char_Status = "| Malherido "
    ElseIf PercVida <> 0 Then
        Generate_Char_Status = "| Casi Muerto "
    Else
        Generate_Char_Status = "| Muerto "
    End If
End If

If Paralizado = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| Paralizado "
End If

If Inmovilizado = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| Inmovilizado "
End If

If Envenenado > 0 Then
    Generate_Char_Status = Generate_Char_Status & "| Envenenado "
End If

If Trabajando = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| Trabajando "
End If

If Silenciado = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| Silenciado "
End If

If Ciego = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| Ciego "
End If

If Incinerado = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| Incinerado "
End If

If Transformado = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| Transformado "
End If

If Comerciando = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| Comerciando "
End If

If Inactivo = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| Inactivo "
End If

End Function
Public Function General_Tittle_Caos(ByVal rango As Byte) As String
    Select Case rango
        Case 1
            General_Tittle_Caos = "Miembro de las Hordas"
        Case 2
            General_Tittle_Caos = "Guerrero del Caos"
        Case 3
            General_Tittle_Caos = "Teniente del Caos"
        Case 4
            General_Tittle_Caos = "Comandante del Caos"
        Case 5
            General_Tittle_Caos = "General del Caos"
        Case 6
            General_Tittle_Caos = "Elite del Caos"
        Case 7
            General_Tittle_Caos = "Asolador de las Sombras"
        Case 8
            General_Tittle_Caos = "Caballero Negro"
        Case 9
            General_Tittle_Caos = "Emisario de las Sombras"
        Case 10
            General_Tittle_Caos = "Avatar del Apocalipsis"
        Case 200
            General_Tittle_Caos = "Lider Caótico"
    End Select
End Function
Public Function General_Tittle_Real(ByVal rango As Byte) As String
    Select Case rango
        Case 1
            General_Tittle_Real = "Legionario"
        Case 2
            General_Tittle_Real = "Soldado Real"
        Case 3
            General_Tittle_Real = "Teniente Real"
        Case 4
            General_Tittle_Real = "Comandante Real"
        Case 5
            General_Tittle_Real = "General Real"
        Case 6
            General_Tittle_Real = "Elite Real"
        Case 7
            General_Tittle_Real = "Guardian del Bien"
        Case 8
            General_Tittle_Real = "Caballero Imperial"
        Case 9
            General_Tittle_Real = "Justiciero"
        Case 10
            General_Tittle_Real = "Guardia Imperial"
        Case 200
            General_Tittle_Real = "Lider Imperial"
    End Select
End Function
Public Function General_Tittle_Milicia(ByVal rango As Byte) As String
    Select Case rango
        Case 1
            General_Tittle_Milicia = "Milicia de Reserva"
        Case 2
            General_Tittle_Milicia = "Miliciano"
        Case 3
            General_Tittle_Milicia = "Miliciano Elite"
        Case 4
            General_Tittle_Milicia = "Soldado de la República"
        Case 5
            General_Tittle_Milicia = "Soldado Raso"
        Case 6
            General_Tittle_Milicia = "Soldado Elite"
        Case 7
            General_Tittle_Milicia = "Comandante de la República"
        Case 200
            General_Tittle_Milicia = "Lider Republicano"
    End Select
End Function
Public Function BoolToByte(ByVal val As Boolean) As Byte
    BoolToByte = IIf(val = True, 1, 0)
End Function
'Función que aplica la transparencia, se le pasa el hwnd del form y un valor de 0 a 255
Public Function Aplicar_Transparencia(ByVal hwnd As Long, Valor As Integer) As Long
On Error Resume Next
    Dim msg As Long

    If Not (Valor < 0 Or Valor > 255) Then
        msg = GetWindowLong(hwnd, GWL_EXSTYLE)
        msg = msg Or WS_EX_LAYERED
    
        SetWindowLong hwnd, GWL_EXSTYLE, msg
    
        'Establece la transparencia
        SetLayeredWindowAttributes hwnd, 0, Valor, LWA_ALPHA
    
        Aplicar_Transparencia = 0
    End If

End Function

Public Sub msgbox_ok(ByVal msg As String)
    ' Nod Kopfnickend
    ' Se hizo mas lindo con el frmMensaje en vez de con el msgbox
    frmMensaje.msg.Caption = msg
    frmMensaje.Show
End Sub


