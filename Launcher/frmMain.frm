VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "InmortalAO Launcher"
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00808080&
   Icon            =   "frmMain.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   460
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   514
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet mInet 
      Left            =   120
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ComboBox cmbRes 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmMain.frx":030A
      Left            =   480
      List            =   "frmMain.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3600
      Width           =   2625
   End
   Begin VB.ComboBox cmbDevice 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmMain.frx":030E
      Left            =   480
      List            =   "frmMain.frx":0310
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2910
      Width           =   4455
   End
   Begin VB.Image cmdMus 
      Enabled         =   0   'False
      Height          =   225
      Index           =   1
      Left            =   5730
      Top             =   3390
      Width           =   225
   End
   Begin VB.Image cmdMus 
      Enabled         =   0   'False
      Height          =   225
      Index           =   0
      Left            =   5730
      Top             =   3150
      Width           =   225
   End
   Begin VB.Image cmdBuffer 
      Height          =   135
      Index           =   1
      Left            =   3840
      Top             =   3780
      Width           =   150
   End
   Begin VB.Image cmdBuffer 
      Height          =   135
      Index           =   0
      Left            =   3840
      Top             =   3630
      Width           =   150
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¡Bienvenido a Inmortal AO!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   525
      Left            =   2220
      TabIndex        =   3
      Top             =   6150
      Width           =   3255
   End
   Begin VB.Image imgCmb 
      Enabled         =   0   'False
      Height          =   435
      Left            =   5715
      Top             =   3165
      Width           =   1125
   End
   Begin VB.Image cmdSolucion 
      Height          =   600
      Left            =   5175
      Top             =   3765
      Width           =   2250
   End
   Begin VB.Label lblBuffer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4020
      TabIndex        =   0
      Top             =   3645
      Width           =   195
   End
   Begin VB.Image imgOp 
      Height          =   255
      Index           =   3
      Left            =   5310
      Top             =   2595
      Width           =   300
   End
   Begin VB.Image imgOp 
      Height          =   255
      Index           =   2
      Left            =   5310
      Top             =   2895
      Width           =   300
   End
   Begin VB.Image cmdPlay 
      Height          =   615
      Left            =   5670
      Top             =   6120
      Width           =   1755
   End
   Begin VB.Image cmdExit 
      Height          =   615
      Left            =   285
      Top             =   6120
      Width           =   1755
   End
   Begin VB.Image cmdNote 
      Height          =   600
      Left            =   6075
      Top             =   5190
      Width           =   1350
   End
   Begin VB.Image cmdPreg 
      Height          =   600
      Left            =   4635
      Top             =   5190
      Width           =   1350
   End
   Begin VB.Image cmdManual 
      Height          =   600
      Left            =   3180
      Top             =   5190
      Width           =   1350
   End
   Begin VB.Image cmdForo 
      Height          =   600
      Left            =   1740
      Top             =   5190
      Width           =   1350
   End
   Begin VB.Image cmdWeb 
      Height          =   600
      Left            =   285
      Top             =   5190
      Width           =   1350
   End
   Begin VB.Image imgOp 
      Height          =   255
      Index           =   1
      Left            =   3030
      Top             =   4020
      Width           =   300
   End
   Begin VB.Image imgOp 
      Height          =   255
      Index           =   0
      Left            =   750
      Top             =   4005
      Width           =   300
   End
   Begin VB.Image imgIniciando 
      Height          =   3015
      Left            =   150
      Top             =   1485
      Width           =   7410
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private botones(7) As New clsButtonFlash
Public LastPressed As New clsButtonFlash
Private descIndex As Integer
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Changed As Boolean
Dim Valaa As Boolean
Private Sub cmbDevice_Click()
    Changed = True
End Sub

Private Sub cmbRes_Click()
Changed = True

Select Case cmbRes.ListIndex
    Case 0
        BitPixel = 32
    Case 1
        BitPixel = 16
End Select
End Sub

Private Sub cmdBuffer_Click(Index As Integer)
    If Index = 0 Then
        If TileBufferSize < 4 Then
            TileBufferSize = TileBufferSize + 1
        End If
    Else
        If TileBufferSize > 1 Then
            TileBufferSize = TileBufferSize - 1
        End If
    End If
    
    lblBuffer.Caption = TileBufferSize
        
    Changed = True
End Sub

Private Sub cmdExit_Click()
SaveConfig
End
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set_Description 1
End Sub

Private Sub cmdForo_Click()
    Call ShellExecute(0, "Open", "http://inmortalao.com.ar/comunidad/", "", App.Path, 0)
End Sub

Private Sub cmdForo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set_Description 5
End Sub

Private Sub cmdManual_Click()
    Call ShellExecute(0, "Open", "http://inmortalao.com.ar/wiki/", "", App.Path, 0)
End Sub

Private Sub cmdManual_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set_Description 6
End Sub


Private Sub cmdNote_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set_Description 8
End Sub

Private Sub cmdPlay_Click()
    If Valaa = True Then Exit Sub
    
    Valaa = True
    
    Call SaveConfig
    
    Call Analizar
    
    If FileExist(App.Path & "\InmortalAO.exe", vbNormal) Then _
        Call Shell(App.Path & "\InmortalAO.exe", vbNormalFocus)
    End
End Sub

Private Sub cmdPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set_Description 2
End Sub

Private Sub cmdPreg_Click()

    Call ShellExecute(0, "Open", "http://inmortalao.com.ar/wiki/", "", App.Path, 0)
End Sub

Private Sub cmdPreg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set_Description 7
End Sub

Private Sub cmdSolucion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set_Description 3
End Sub

Private Sub cmdWeb_Click()
    Call ShellExecute(0, "Open", "http://inmortalao.com.ar/", "", App.Path, 0)
End Sub

Private Sub cmdWeb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set_Description 4
End Sub

Private Sub Form_Load()
    resource_path = App.Path & "\Resources\"
'    Me.cmbDevice.AddItem "Buscando dispositivos ..."
'    Me.cmbDevice.ListIndex = 0
    
    botones(0).Initialize cmdWeb, _
                "sitioover", _
                "sitiodown", _
                Me, False

    botones(1).Initialize cmdForo, _
                "foroover", _
                "forodown", _
                Me, False

    botones(2).Initialize cmdPlay, _
                "iniciarover", _
                "iniciardown", _
                Me, False

    botones(3).Initialize cmdManual, _
                "manualover", _
                "manualdown", _
                Me, False

    botones(4).Initialize cmdSolucion, _
                "solucover", _
                "solucdown", _
                Me, False

    botones(5).Initialize cmdPreg, _
                "faqover", _
                "faqdown", _
                Me, False

    botones(6).Initialize cmdExit, _
                "salirover", _
                "salirdown", _
                Me, False

    botones(7).Initialize cmdNote, _
                "notasover", _
                "notasdown", _
                Me, False
                
                
                

    Me.Picture = LoadInterface("launchear")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
    Set_Description 0
End Sub

Private Sub Set_Description(ByVal Index As Integer)
    If descIndex = Index Then Exit Sub
    
    If Valaa = True Then Exit Sub
    
    Select Case Index
        Case 0
            lblEstado.Caption = "¡Bienvenido a Inmortal AO! Haz click en Iniciar Juego para jugar"
        Case 1
            lblEstado.Caption = "Volver al escritorio de Windows®"
        Case 2
            lblEstado.Caption = "¡Inicia Inmortal AO!"
        Case 3
            lblEstado.Caption = "En caso de tener problemas al jugar este programa verificará la integridad del sistema"
        Case 4
            lblEstado.Caption = "Visita http://inmortalao.com.ar/"
        Case 5
            lblEstado.Caption = "Visita los foros de discusión donde podrás opinar, pedir ayuda o simplemente relajarte"
        Case 6
            lblEstado.Caption = "Manual del juego: seguramente la mayoría de tus dudas pueden ser respondidas aquí"
        Case 7
            lblEstado.Caption = "Accede a la sección de preguntas frecuentes del sitio"
        Case 8
            lblEstado.Caption = "No habilitado" 'Información de desarrollo"
        Case 9
            lblEstado.Caption = "Deja habilitada la sincronización vertical (avanzado), pero en caso de tener problemas, deshabilítala"
        Case 10
            lblEstado.Caption = "Habilita la música de ambiente del juego"
        Case 11
            lblEstado.Caption = "Habilita los efectos sonoros, de clima, ambiente, etc."
    End Select
    
    If Index = 0 Then
        lblEstado.FontBold = True
        lblEstado.FontSize = 8
    ElseIf Index = 9 Then
        lblEstado.FontSize = 7
        lblEstado.FontBold = False
    Else
        lblEstado.FontBold = False
        lblEstado.FontSize = 8
    End If
    
    descIndex = Index
End Sub


Private Sub imgIniciando_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set_Description 0
    LastPressed.ToggleToNormal
End Sub

Private Sub imgOp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 1 'Sinc
            Set_Description 9
                
        Case 2 'Musica
            Set_Description 10
                
        Case 3 'Sonido
            Set_Description 11
                
    End Select
End Sub

Private Sub imgOp_Click(Index As Integer)
    With imgOp(Index)
        Select Case Index
            Case 0 'Ventana
                If Window Then
                    Window = 0
                    
                 '   cmbDevice.Enabled = True
                    cmbRes.Enabled = True
                    
                    .Picture = LoadInterface("vsyncoff")
                Else
                    Window = 1
                    
                    cmbDevice.Enabled = False
                    cmbRes.Enabled = False
                    .Picture = LoadInterface("correron")
                End If
            Case 1 'Sinc
                If Sinc Then
                    Sinc = 0
                    
                    .Picture = LoadInterface("vsyncoff")
                Else
                    Sinc = 1
                    
                    .Picture = LoadInterface("correron")
                End If
            Case 2 'Musica
                If Music Then
                    Music = 0
                    
                    .Picture = LoadInterface("vsyncoff")
                    imgCmb.Picture = LoadInterface("ambient-launcher-off")
                Else
                    Music = 1
                    
                    Set imgCmb.Picture = Nothing
                    
                    .Picture = LoadInterface("correron")
                End If
            Case 3 'Sonido
                If Sound Then
                    Sound = 0
                    
                    .Picture = LoadInterface("vsyncoff")
                Else
                    Sound = 1
                
                
                    .Picture = LoadInterface("correron")
                End If
        End Select
    End With
    Changed = True
End Sub

Private Sub mInet_StateChanged(ByVal State As Integer)
On Error GoTo Err_Sub
    
    Dim tempArray() As Byte, bDone As Boolean, filesize As Long, vtData As Variant, lenf As Long

    Select Case State
        Case icResponseCompleted
            bDone = False
            filesize = mInet.GetHeader("Content-length")
            Open Directory For Binary As #1
                vtData = mInet.GetChunk(1024, icByteArray)
                DoEvents
             
                If Len(vtData) = 0 Then bDone = True
                
                Do While Not bDone
                    tempArray = vtData
                    Put #1, , tempArray
                    
                    vtData = mInet.GetChunk(1024, icByteArray)
                    DoEvents
                    
                    lenf = lenf + Len(vtData)
                    
                    If Len(vtData) = 0 Then bDone = True
                    
                    lblEstado.Caption = "Descargando Parche " & dNum & " " & Round(lenf / 1024, 2) & "/" & Round(filesize / 1024, 2) & " kbs"
                Loop
            Close #1
            
            dMain = True
    End Select

    Exit Sub

Err_Sub:
    MsgBox Err.Description, vbCritical
    On Error Resume Next
    mInet.Cancel
    dMain = True
End Sub
