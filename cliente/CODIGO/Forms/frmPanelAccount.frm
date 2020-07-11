VERSION 5.00
Begin VB.Form frmPanelAccount 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Panel de Cuenta"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   Icon            =   "frmPanelAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   360
      Top             =   3000
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   0
      Left            =   2415
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   9
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   1
      Left            =   3915
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   8
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   2
      Left            =   5415
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   3
      Left            =   6915
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   6
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   4
      Left            =   8415
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   5
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   5
      Left            =   2415
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   4
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   6
      Left            =   3915
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   3
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   7
      Left            =   5415
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   2
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   8
      Left            =   6915
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   9
      Left            =   8415
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   5730
      Width           =   1140
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   0
      Left            =   2250
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   1
      Left            =   3750
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   2
      Left            =   5250
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   3
      Left            =   6750
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   4
      Left            =   8250
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   5
      Left            =   2250
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   6
      Left            =   3750
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   7
      Left            =   5250
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   8
      Left            =   6750
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   9
      Left            =   8250
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image cmdChange 
      Height          =   615
      Left            =   6180
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   2085
      Width           =   1755
   End
   Begin VB.Image cmdCrear 
      Height          =   615
      Left            =   2235
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   7575
      Width           =   1755
   End
   Begin VB.Image cmdExit 
      Height          =   615
      Left            =   8025
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   2085
      Width           =   1755
   End
   Begin VB.Image cmdDelete 
      Height          =   615
      Left            =   4155
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   7575
      Width           =   1755
   End
   Begin VB.Image cmdConnect 
      Height          =   615
      Left            =   8025
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   7575
      Width           =   1755
   End
   Begin VB.Image WebLink 
      Height          =   345
      Left            =   8460
      MousePointer    =   99  'Custom
      Top             =   8550
      Width           =   3405
   End
   Begin VB.Label lblAccData 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2250
      TabIndex        =   23
      Top             =   2370
      Width           =   3705
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   2295
      TabIndex        =   22
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   3810
      TabIndex        =   21
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   5295
      TabIndex        =   20
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   6810
      TabIndex        =   19
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   5
      Left            =   8310
      TabIndex        =   18
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   6
      Left            =   2295
      TabIndex        =   17
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   7
      Left            =   3810
      TabIndex        =   16
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   8
      Left            =   5295
      TabIndex        =   15
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   9
      Left            =   6810
      TabIndex        =   14
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   10
      Left            =   8310
      TabIndex        =   13
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   6210
      TabIndex        =   12
      Top             =   7650
      Width           =   345
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   6210
      TabIndex        =   11
      Top             =   7800
      Width           =   675
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   6180
      TabIndex        =   10
      Top             =   7920
      Width           =   45
   End
End
Attribute VB_Name = "frmPanelAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Seleccionado As Byte
Public LastPressed As clsButton
Dim Botones(1 To 5) As clsButton
Private Sub cmdChange_Click()
    msgbox_ok "No habilitado"
End Sub

Private Sub cmdConnect_Click()
    UserName = lblAccData(1 + Seleccionado).Caption
    WriteLoginExistingChar
End Sub

Private Sub cmdCrear_Click()
    Dim i As Long
    For i = 0 To 9
        If lblAccData(i + 1).Caption = "" Then
            frmPanelAccount.Visible = False
            frmCrearPersonaje.Show
            Exit Sub
        End If
    Next i
    
    frmMensaje.msg.Caption = "No puedes crear mas personajes."
    frmMensaje.Show
End Sub

Private Sub cmdDelete_Click()
    ' Nod Kopfnickend
    ' Cosas que faltan hacer
    ShellExecute Me.hwnd, "open", "http://inmortalao.com.ar/panel-de-usuario.php?ac=mis-personajes", "", "", 1

End Sub

Private Sub cmdExit_Click()

    Unload frmConnect
    perm = True
    frmMain.Socket1.Disconnect
    frmConnect.Visible = True
    frmPanelAccount.Visible = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Me.Picture = LoadInterface("cuenta")
    Me.Icon = frmMain.Icon
    
    Dim i As Byte
    For i = 1 To 10
        lblAccData(i).Caption = ""
    Next i
    
    Set Botones(1) = New clsButton
    Botones(1).Initialize cmdConnect, _
        "conectar-cuenta-over", _
        "conectar-cuenta-down", _
        Me, True
        
    Set Botones(2) = New clsButton
    Botones(2).Initialize cmdExit, _
        "salir-cuenta-over", _
        "salir-cuenta-down", _
        Me, True
        
    Set Botones(3) = New clsButton
    Botones(3).Initialize cmdCrear, _
        "crear-cuenta-over", _
        "crear-cuenta-down", _
        Me, True
        
    Set Botones(4) = New clsButton
    Botones(4).Initialize cmdDelete, _
        "borrar-cuenta-over", _
        "borrar-cuenta-down", _
        Me, True
        
    Set Botones(5) = New clsButton
    Botones(5).Initialize cmdChange, _
        "cambiar-cuenta-over", _
        "cambiar-cuenta-down", _
        Me, True
        
    Set LastPressed = New clsButton
    
End Sub
Private Sub lblName_Click(Index As Integer)
    Seleccionado = Index
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) And mueve = 1 Then Call Auto_Drag(Me.hwnd)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub picChar_Click(Index As Integer)

On Error Resume Next
    Seleccionado = Index
    If cPJ(Index).nombre <> "" Then
        lblCharData(0) = ListaClases(cPJ(Index).Clase) & " Nivel " & cPJ(Index).Nivel
        lblCharData(1) = Map_NameLoad(cPJ(Index).Mapa)  'Ubicacion
        lblCharData(2) = ""
    Else
        lblCharData(0) = ""
        lblCharData(1) = ""
        lblCharData(2) = ""
    End If
End Sub

Private Sub picChar_DblClick(Index As Integer)
    Seleccionado = Index
    If Not lblAccData(Index + 1).Caption = "" Then
        UserName = lblAccData(1 + Index).Caption
        WriteLoginExistingChar
    Else
        frmCrearPersonaje.Show
    End If
End Sub

Private Sub Timer1_Timer()
    Dim i As Byte
    For i = 1 To 10
        TileEngine.Char_Account_Render i
    Next i
End Sub
