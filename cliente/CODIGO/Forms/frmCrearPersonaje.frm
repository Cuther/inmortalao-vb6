VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Inmortal AO"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox lstFamiliar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      ItemData        =   "frmCrearPersonaje.frx":0000
      Left            =   8520
      List            =   "frmCrearPersonaje.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   1860
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.TextBox txtFamiliar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9300
      MaxLength       =   30
      TabIndex        =   48
      Top             =   1020
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox picFamiliar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   10365
      ScaleHeight     =   1185
      ScaleWidth      =   840
      TabIndex        =   47
      Top             =   1560
      Width           =   870
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0004
      Left            =   870
      List            =   "frmCrearPersonaje.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   3810
      Width           =   2055
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0008
      Left            =   840
      List            =   "frmCrearPersonaje.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   3150
      Width           =   2055
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":000C
      Left            =   870
      List            =   "frmCrearPersonaje.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   2490
      Width           =   2055
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2130
      MaxLength       =   20
      TabIndex        =   41
      Top             =   1110
      Width           =   5865
   End
   Begin VB.PictureBox HeadView 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1770
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   40
      Top             =   4545
      Width           =   375
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0010
      Left            =   8520
      List            =   "frmCrearPersonaje.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3600
      Width           =   2820
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   2760
      TabIndex        =   51
      Top             =   8160
      Width           =   6795
   End
   Begin VB.Image imgNoDisp 
      Height          =   2145
      Left            =   8370
      Top             =   780
      Width           =   3045
   End
   Begin VB.Label lblFamiInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion del familiar."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8535
      TabIndex        =   50
      Top             =   2280
      Width           =   1635
   End
   Begin VB.Image seguir 
      Height          =   615
      Left            =   9600
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   8160
      Width           =   1755
   End
   Begin VB.Image salir 
      Height          =   615
      Left            =   660
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   8175
      Width           =   1755
   End
   Begin VB.Image Image1 
      Height          =   3570
      Left            =   8520
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   2835
   End
   Begin VB.Image MenosHead 
      Height          =   600
      Left            =   1440
      Tag             =   "0"
      Top             =   4440
      Width           =   270
   End
   Begin VB.Image MasHead 
      Height          =   600
      Left            =   2160
      Tag             =   "0"
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Height          =   195
      Index           =   0
      Left            =   2460
      TabIndex        =   39
      Top             =   5700
      Width           =   240
   End
   Begin VB.Label lbAtributos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "40"
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
      Left            =   2535
      TabIndex        =   38
      Top             =   7500
      Width           =   255
   End
   Begin VB.Label lbBonificador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   0
      Left            =   2160
      TabIndex        =   37
      Top             =   5700
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   0
      Left            =   2715
      Top             =   5640
      Width           =   195
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   0
      Left            =   2715
      Top             =   5790
      Width           =   195
   End
   Begin VB.Label lbBonificador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   36
      Top             =   6060
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lbBonificador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   2
      Left            =   2160
      TabIndex        =   35
      Top             =   6420
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lbBonificador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   3
      Left            =   2160
      TabIndex        =   34
      Top             =   6780
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lbBonificador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   4
      Left            =   2160
      TabIndex        =   33
      Top             =   7140
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Height          =   195
      Index           =   1
      Left            =   2460
      TabIndex        =   32
      Top             =   6060
      Width           =   240
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Height          =   195
      Index           =   2
      Left            =   2460
      TabIndex        =   31
      Top             =   6420
      Width           =   240
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Height          =   195
      Index           =   3
      Left            =   2460
      TabIndex        =   30
      Top             =   6780
      Width           =   240
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Height          =   195
      Index           =   4
      Left            =   2460
      TabIndex        =   29
      Top             =   7140
      Width           =   240
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   1
      Left            =   2715
      Top             =   6030
      Width           =   195
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   2
      Left            =   2715
      Top             =   6360
      Width           =   195
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   3
      Left            =   2715
      Top             =   6750
      Width           =   195
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   4
      Left            =   2715
      Top             =   7080
      Width           =   195
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   1
      Left            =   2715
      Top             =   6150
      Width           =   195
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   2
      Left            =   2715
      Top             =   6510
      Width           =   195
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   3
      Left            =   2715
      Top             =   6870
      Width           =   195
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   4
      Left            =   2715
      Top             =   7230
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   5310
      TabIndex        =   28
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   5310
      TabIndex        =   27
      Top             =   2340
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   5310
      TabIndex        =   26
      Top             =   3090
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   5310
      TabIndex        =   25
      Top             =   3450
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   5310
      TabIndex        =   24
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   5310
      TabIndex        =   23
      Top             =   4200
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   5310
      TabIndex        =   22
      Top             =   4590
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   5310
      TabIndex        =   21
      Top             =   4950
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   5310
      TabIndex        =   20
      Top             =   5340
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   5310
      TabIndex        =   19
      Top             =   5700
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   5310
      TabIndex        =   18
      Top             =   6090
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   5310
      TabIndex        =   17
      Top             =   6450
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   5310
      TabIndex        =   16
      Top             =   6840
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   5310
      TabIndex        =   15
      Top             =   7200
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   7365
      TabIndex        =   14
      Top             =   2340
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   7365
      TabIndex        =   13
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   7365
      TabIndex        =   12
      Top             =   3090
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   7365
      TabIndex        =   11
      Top             =   3450
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":0014
      Top             =   4680
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   40
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":0166
      Top             =   4560
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   39
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":02B8
      Top             =   4320
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   38
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":040A
      Top             =   4170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   37
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":055C
      Top             =   3930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   36
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":06AE
      Top             =   3780
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   35
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":0800
      Top             =   3570
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":0952
      Top             =   3420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":0AA4
      Top             =   3180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":0BF6
      Top             =   3060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   31
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":0D48
      Top             =   2820
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":0E9A
      Top             =   2670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   29
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":0FEC
      Top             =   2430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   28
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":113E
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   26
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":1290
      Top             =   7170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":13E2
      Top             =   6810
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   22
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":1534
      Top             =   6420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":1686
      Top             =   6060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   18
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":17D8
      Top             =   5670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   16
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":192A
      Top             =   5280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   14
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":1A7C
      Top             =   4920
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":1BCE
      Top             =   4560
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":1D20
      Top             =   4170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   8
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":1E72
      Top             =   3780
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   6
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":1FC4
      Top             =   3420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":2116
      Top             =   3060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   2
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":2268
      Top             =   2670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":23BA
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   1
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":250C
      Top             =   2430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   27
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":265E
      Top             =   7290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   25
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":27B0
      Top             =   6930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   23
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":2902
      Top             =   6540
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   21
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":2A54
      Top             =   6180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   19
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":2BA6
      Top             =   5790
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   17
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":2CF8
      Top             =   5430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   15
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":2E4A
      Top             =   5040
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   13
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":2F9C
      Top             =   4680
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   11
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":30EE
      Top             =   4290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   9
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":3240
      Top             =   3930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   7
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":3392
      Top             =   3540
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   5
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":34E4
      Top             =   3180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   3
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":3636
      Top             =   2790
      Width           =   195
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
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
      Left            =   6795
      TabIndex        =   10
      Top             =   7260
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   7365
      TabIndex        =   9
      Top             =   4590
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   7365
      TabIndex        =   8
      Top             =   4200
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   7365
      TabIndex        =   7
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   7365
      TabIndex        =   6
      Top             =   4950
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   42
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":3788
      Top             =   4920
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   43
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":38DA
      Top             =   5040
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   7365
      TabIndex        =   5
      Top             =   5340
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   44
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":3A2C
      Top             =   5280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   45
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":3B7E
      Top             =   5430
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   7365
      TabIndex        =   4
      Top             =   5700
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   46
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":3CD0
      Top             =   5670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   47
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":3E22
      Top             =   5790
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   7365
      TabIndex        =   3
      Top             =   6090
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   48
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":3F74
      Top             =   6060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   49
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":40C6
      Top             =   6180
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   7365
      TabIndex        =   2
      Top             =   6450
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   50
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":4218
      Top             =   6420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   51
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":436A
      Top             =   6540
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   26
      Left            =   7365
      TabIndex        =   1
      Top             =   6840
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   52
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":44BC
      Top             =   6810
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   53
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":460E
      Top             =   6930
      Width           =   195
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1440
      TabIndex        =   45
      Top             =   4440
      Width           =   435
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   2160
      TabIndex        =   46
      Top             =   4440
      Width           =   315
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LastPressed As clsButton
Dim Botones(1 To 2) As clsButton

Public SkillPoints As Byte
Public Actual As Integer
Private MaxEleccion As Integer, MinEleccion As Integer

Function CheckData() As Boolean

If lstHogar.ListIndex = -1 Then
    lblInfo.Caption = "Seleccione el hogar del personaje."
    CheckData = False
    Exit Function
End If

If txtNombre.Text = vbNullString Then
    lblInfo.Caption = "Seleccione el nombre al personaje."
    CheckData = False
    Exit Function
End If

If lstRaza.ListIndex = -1 Then
    lblInfo.Caption = "Seleccione la raza del personaje."
    CheckData = False
    Exit Function
End If

If lstGenero.ListIndex = -1 Then
    lblInfo.Caption = "Seleccione el sexo del personaje."
    CheckData = False
    Exit Function
End If

If lstProfesion.ListIndex = -1 Then
    lblInfo.Caption = "Seleccione la clase del personaje."
    CheckData = False
    Exit Function
End If


If (lstProfesion.ListIndex + 1) = eClass.Mago Or (lstProfesion.ListIndex + 1) = eClass.Cazador Or (lstProfesion.ListIndex + 1) = eClass.Druida Then
    If PetType = 0 Then
        lblInfo.Caption = "Seleccione su familiar o mascota."
        CheckData = False
        Exit Function
    ElseIf Len(PetName) = 0 Then
        lblInfo.Caption = "Seleccione un nombre al su familiar o mascota."
        CheckData = False
        Exit Function
    ElseIf Len(PetName) > 30 Then
        lblInfo.Caption = "El nombre de su familiar o mascota debe tener menos de 30 letras."
        CheckData = False
        Exit Function
    End If
End If

If SkillPoints > 0 Then
    lblInfo.Caption = "Asigne los skillpoints del personaje."
    CheckData = False
    Exit Function
End If

If val(lbAtributos.Caption) > 0 Then
    lblInfo.Caption = "Asigne los atributos del personaje."
    CheckData = False
    Exit Function
ElseIf val(lbAtributos) < 0 Then
    lblInfo.Caption = "Los atributos del personaje son inválidos."
    CheckData = False
    Exit Function
End If

CheckData = True

End Function

Private Sub Command1_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)

Dim indice
If (Index And &H1) = 0 Then
    If SkillPoints > 0 Then
        indice = Index \ 2
        Skill(indice).Caption = val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
    End If
Else
    If SkillPoints < 10 Then
        
        indice = Index \ 2
        If val(Skill(indice).Caption) > 0 Then
            Skill(indice).Caption = val(Skill(indice).Caption) - 1
            SkillPoints = SkillPoints + 1
        End If
    End If
End If

Puntos.Caption = SkillPoints
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    If frmMain.Socket1.Connected = False Then Unload frmConnect

    SkillPoints = 10
    Puntos.Caption = SkillPoints

    Me.Picture = LoadInterface("crearpj")
    Me.Icon = frmMain.Icon

    lstProfesion.Clear
    For i = LBound(ListaClases) To UBound(ListaClases)
        lstProfesion.AddItem ListaClases(i)
    Next i

    lstHogar.Clear
    lstHogar.AddItem "Nix (Imperio)"
    lstHogar.AddItem "Suramei (República)"
    

    lstRaza.Clear
    For i = LBound(ListaRazas()) To UBound(ListaRazas())
        lstRaza.AddItem ListaRazas(i)
    Next i


    lstProfesion.Clear
    For i = LBound(ListaClases()) To UBound(ListaClases())
        lstProfesion.AddItem ListaClases(i)
    Next i

    lstProfesion.ListIndex = 1

    lstGenero.AddItem "Hombre"
    lstGenero.AddItem "Mujer"
    Image1.Picture = LoadInterface(lstProfesion.Text & "")

    Set Botones(1) = New clsButton
    Botones(1).Initialize salir, _
            "volver-crearpj-over", _
            "volver-crearpj-down", _
            Me, True
            
    Set Botones(2) = New clsButton
    Botones(2).Initialize seguir, _
            "crear-crearpj-over", _
            "crear-crearpj-down", _
            Me, True
            
    Set LastPressed = New clsButton
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub


Private Sub ImgAtributoMas_Click(Index As Integer)

If val(lbAtt(Index).Caption) >= 18 Or val(lbAtributos.Caption) <= 0 Then Exit Sub
    
lbAtt(Index).Caption = val(lbAtt(Index).Caption) + 1
lbAtributos.Caption = lbAtributos.Caption - 1

End Sub

Private Sub ImgAtributoMenos_Click(Index As Integer)

If val(lbAtt(Index).Caption) <= 6 Then Exit Sub

lbAtt(Index).Caption = val(lbAtt(Index).Caption) - 1
lbAtributos.Caption = lbAtributos.Caption + 1

End Sub



Private Sub lstFamiliar_Click()
If lstFamiliar.ListIndex > 0 Then
    lblFamiInfo.Caption = ListaFamiliares(lstFamiliar.ListIndex).Desc
    picFamiliar.Picture = LoadInterface(ListaFamiliares(lstFamiliar.ListIndex).Imagen)
    PetType = ListaFamiliares(lstFamiliar.ListIndex).tipe
Else
    lblFamiInfo.Caption = "Descripcion del familiar."
    picFamiliar.Picture = Nothing
End If
End Sub

Private Sub lstProfesion_Click()
On Error Resume Next
    Image1.Picture = LoadInterface(lstProfesion.Text)
    
    If lstProfesion.Text = "Mago" Then
        frmCrearPersonaje.txtFamiliar.Visible = True
        frmCrearPersonaje.lstFamiliar.Visible = True
        imgNoDisp.Picture = Nothing
        lblFamiInfo.Visible = True
        picFamiliar.Visible = True
        Call CambioFamiliar(5)
    ElseIf lstProfesion.Text = "Cazador" Or lstProfesion.Text = "Druida" Then
        frmCrearPersonaje.txtFamiliar.Visible = True
        frmCrearPersonaje.lstFamiliar.Visible = True
        imgNoDisp.Picture = Nothing
        lblFamiInfo.Visible = True
        picFamiliar.Visible = True
        Call CambioFamiliar(4)
    Else
        frmCrearPersonaje.txtFamiliar.Visible = False
        frmCrearPersonaje.lstFamiliar.Visible = False
        imgNoDisp.Picture = LoadInterface("crearpj-nodisp")
        picFamiliar.Visible = False
        lblFamiInfo.Visible = False
    End If
    
End Sub
Private Sub CambioFamiliar(ByVal NumFamiliares As Integer)

If NumFamiliares = 5 Then

    ReDim ListaFamiliares(1 To NumFamiliares) As tListaFamiliares
    ListaFamiliares(1).name = "Elemental De Fuego"
    ListaFamiliares(1).Desc = "Hecho de puro fuego, lanzará tormentas sobre tus contrincantes."
    ListaFamiliares(1).Imagen = "elefuego"
    ListaFamiliares(1).tipe = eMascota.fuego
    
    ListaFamiliares(2).name = "Elemental De Agua"
    ListaFamiliares(2).Desc = "Con su cuerpo acuoso paralizará a tus enemigos."
    ListaFamiliares(2).Imagen = "eleagua"
    ListaFamiliares(2).tipe = eMascota.Agua
    
    ListaFamiliares(3).name = "Elemental De Tierra"
    ListaFamiliares(3).Desc = "Sus fuertes brazos inmovilizarán cualquier criatura viviente."
    ListaFamiliares(3).Imagen = "eletierra"
    ListaFamiliares(3).tipe = eMascota.Tierra
    
    ListaFamiliares(4).name = "Ely"
    ListaFamiliares(4).Desc = "Te protegerá constantemente con sus conjuros defensivos."
    ListaFamiliares(4).Imagen = "ely"
    ListaFamiliares(4).tipe = eMascota.Ely
    
    ListaFamiliares(5).name = "Fuego Fatuo"
    ListaFamiliares(5).Desc = "Débil pero con grán poder mágico, siempre estará a tu lado."
    ListaFamiliares(5).Imagen = "fatuo"
    ListaFamiliares(5).tipe = eMascota.Fatuo
Else

    ReDim ListaFamiliares(1 To NumFamiliares) As tListaFamiliares
    ListaFamiliares(1).name = "Tigre"
    ListaFamiliares(1).Desc = "Poseen grandes y filosas garras para atacar a tus oponentes."
    ListaFamiliares(1).Imagen = "tigre"
    ListaFamiliares(1).tipe = eMascota.Tigre
    
    ListaFamiliares(2).name = "Lobo"
    ListaFamiliares(2).Desc = "Astutos y arrogantes, su mordedura causa estragos en sus víctimas."
    ListaFamiliares(2).Imagen = "lobo"
    ListaFamiliares(2).tipe = eMascota.Lobo
    
    ListaFamiliares(3).name = "Oso Pardo"
    ListaFamiliares(3).Desc = "Se caracterizan por ser territoriales y muy resistentes."
    ListaFamiliares(3).Imagen = "oso"
    ListaFamiliares(3).tipe = eMascota.Oso
    
    ListaFamiliares(4).name = "Ent"
    ListaFamiliares(4).Desc = "¡Esta robusta criatura te defenderá cual muro de piedra!"
    ListaFamiliares(4).Imagen = "ent"
    ListaFamiliares(4).tipe = eMascota.Ent
End If

Dim i As Integer
lstFamiliar.Clear
lstFamiliar.AddItem vbNullString
For i = 1 To UBound(ListaFamiliares)
    lstFamiliar.AddItem ListaFamiliares(i).name
Next i

lstFamiliar.ListIndex = 0
End Sub

Private Sub lstGenero_Click()
    Call DameOpciones
End Sub
Private Sub lstRaza_Click()
    Call DameOpciones
    Dim i As Integer, tmpInt As Integer
    For i = 1 To NUMATRIBUTOS
        tmpInt = BonificadorRaza(i, lstRaza.ListIndex + 1)
        
        lbBonificador(i - 1).Caption = IIf(tmpInt > 0, "+" & CStr(tmpInt), CStr(tmpInt))
        If val(lbBonificador(i - 1)) = 0 Then
            lbBonificador(i - 1).Visible = False
        Else
            lbBonificador(i - 1).Visible = True
        End If
    Next i
End Sub
Private Sub MenosHead_Click()
    Call Audio.PlayWave(SND_CLICK)
    Actual = Actual - 1
    If Actual > MaxEleccion Then
       Actual = MaxEleccion
    ElseIf Actual < MinEleccion Then
       Actual = MinEleccion
    End If
    HeadView.Cls
    Call TileEngine.Draw_Grh_Hdc(HeadView.hdc, HeadData(Actual).Head(3).grhindex, 8, 5)
    HeadView.refresh
End Sub
Private Sub MasHead_Click()
    Call Audio.PlayWave(SND_CLICK)
    Actual = Actual + 1
    If Actual > MaxEleccion Then
       Actual = MaxEleccion
    ElseIf Actual < MinEleccion Then
       Actual = MinEleccion
    End If
    HeadView.Cls
    Call TileEngine.Draw_Grh_Hdc(HeadView.hdc, HeadData(Actual).Head(3).grhindex, 5, 5)
    HeadView.refresh
End Sub
Sub DameOpciones()
Dim i As Integer

If frmCrearPersonaje.lstGenero.ListIndex < 0 Then Exit Sub
If frmCrearPersonaje.lstRaza.ListIndex < 0 Then Exit Sub

Select Case frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex)
   Case "Hombre"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
            Case "Humano"
                Actual = Head_Range(HUMANO).mStart
                MaxEleccion = Head_Range(HUMANO).mEnd
                MinEleccion = Head_Range(HUMANO).mStart
            Case "Elfo"
                Actual = Head_Range(ELFO).mStart
                MaxEleccion = Head_Range(ELFO).mEnd
                MinEleccion = Head_Range(ELFO).mStart
            Case "Elfo Drow"
                Actual = Head_Range(ElfoOscuro).mStart
                MaxEleccion = Head_Range(ElfoOscuro).mEnd
                MinEleccion = Head_Range(ElfoOscuro).mStart
            Case "Enano"
                Actual = Head_Range(Enano).mStart
                MaxEleccion = Head_Range(Enano).mEnd
                MinEleccion = Head_Range(Enano).mStart
            Case "Gnomo"
                Actual = Head_Range(Gnomo).mStart
                MaxEleccion = Head_Range(Gnomo).mEnd
                MinEleccion = Head_Range(Gnomo).mStart
            Case "Orco"
                Actual = Head_Range(Orco).mStart
                MaxEleccion = Head_Range(Orco).mEnd
                MinEleccion = Head_Range(Orco).mStart
            Case Else
                Actual = 30
                MaxEleccion = 30
                MinEleccion = 30
        End Select
   Case "Mujer"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
            Case "Humano"
                Actual = Head_Range(HUMANO).fStart
                MaxEleccion = Head_Range(HUMANO).fEnd
                MinEleccion = Head_Range(HUMANO).fStart
            Case "Elfo"
                Actual = Head_Range(ELFO).fStart
                MaxEleccion = Head_Range(ELFO).fEnd
                MinEleccion = Head_Range(ELFO).fStart
            Case "Elfo Drow"
                Actual = Head_Range(ElfoOscuro).fStart
                MaxEleccion = Head_Range(ElfoOscuro).fEnd
                MinEleccion = Head_Range(ElfoOscuro).fStart
            Case "Enano"
                Actual = Head_Range(Enano).fStart
                MaxEleccion = Head_Range(Enano).fEnd
                MinEleccion = Head_Range(Enano).fStart
            Case "Gnomo"
                Actual = Head_Range(Gnomo).fStart
                MaxEleccion = Head_Range(Gnomo).fEnd
                MinEleccion = Head_Range(Gnomo).fStart
            Case "Orco"
                Actual = Head_Range(Orco).fStart
                MaxEleccion = Head_Range(Orco).fEnd
                MinEleccion = Head_Range(Orco).fStart
            Case Else
                Actual = 30
                MaxEleccion = 30
                MinEleccion = 30
        End Select
End Select
 
HeadView.Cls
Call TileEngine.Draw_Grh_Hdc(HeadView.hdc, HeadData(Actual).Head(3).grhindex, 5, 5)
HeadView.refresh
End Sub
Public Function BonificadorRaza(ByVal Atributo As Integer, ByVal Raza As Byte) As Integer

Select Case Atributo
    Case Fuerza
        If Raza = HUMANO Then BonificadorRaza = 1
        If Raza = ElfoOscuro Then BonificadorRaza = 2
        If Raza = Enano Then BonificadorRaza = 3
        If Raza = ELFO Then BonificadorRaza = 0
        If Raza = Orco Then BonificadorRaza = 5
        If Raza = Gnomo Then BonificadorRaza = -5
    Case Agilidad
        If Raza = HUMANO Then BonificadorRaza = 1
        If Raza = ElfoOscuro Then BonificadorRaza = 0
        If Raza = Enano Then BonificadorRaza = -1
        If Raza = ELFO Then BonificadorRaza = 2
        If Raza = Orco Then BonificadorRaza = -2
        If Raza = Gnomo Then BonificadorRaza = 3
    Case Inteligencia
        If Raza = HUMANO Then BonificadorRaza = 1
        If Raza = ElfoOscuro Then BonificadorRaza = 2
        If Raza = Enano Then BonificadorRaza = -7
        If Raza = ELFO Then BonificadorRaza = 3
        If Raza = Orco Then BonificadorRaza = -6
        If Raza = Gnomo Then BonificadorRaza = 4
    Case Carisma
        If Raza = HUMANO Then BonificadorRaza = 0
        If Raza = ElfoOscuro Then BonificadorRaza = -1
        If Raza = Enano Then BonificadorRaza = -1
        If Raza = ELFO Then BonificadorRaza = 2
        If Raza = Orco Then BonificadorRaza = -4
        If Raza = Gnomo Then BonificadorRaza = 0
    Case Constitucion
        If Raza = HUMANO Then BonificadorRaza = 2
        If Raza = ElfoOscuro Then BonificadorRaza = 1
        If Raza = Enano Then BonificadorRaza = 4
        If Raza = ELFO Then BonificadorRaza = 0
        If Raza = Orco Then BonificadorRaza = 4
        If Raza = Gnomo Then BonificadorRaza = -1
End Select

End Function

Private Sub ResetAtributos()
lbAtributos.Caption = 40

Dim i As Integer

For i = 1 To NUMATRIBUTOS
    lbAtt(i - 1).Caption = "6"
    UserAtributos(i) = 6
Next i

End Sub

Private Sub salir_Click()
    frmPanelAccount.Show
    Unload Me
End Sub

Private Sub seguir_Click()
    
    If Trim(txtNombre.Text) = "" Then
    frmMensaje.msg.Caption = "Nombre invalido."
    frmMensaje.Show
    Exit Sub
    End If
    
    
    Dim i As Integer
    Dim k As Object
    i = 1
    For Each k In Skill
        UserSkills(i) = k.Caption
        i = i + 1
    Next
    


    If CheckData() Then
        EstadoLogin = CrearNuevoPj
        
        If frmMain.Socket1.Connected = False Then
            msgbox_ok "Error: Se ha perdido la conexion con el server."
            Unload Me
        Else
        
        ' LLENAMOS LAS VARIABLES NI BIEN VAMOS A PROCEDER
        ' SINO NO TIENE SENTIDO HACERLO (CASTELLI)

    UserName = Trim$(txtNombre.Text)

    UserRaza = lstRaza.ListIndex + 1
    UserSexo = lstGenero.ListIndex + 1
    UserClase = lstProfesion.ListIndex + 1
    
    UserAtributos(1) = val(lbAtt(0).Caption)
    UserAtributos(2) = val(lbAtt(1).Caption)
    UserAtributos(3) = val(lbAtt(2).Caption)
    UserAtributos(4) = val(lbAtt(3).Caption)
    UserAtributos(5) = val(lbAtt(4).Caption)
    
    UserHogar = lstHogar.ListIndex

        
            Call Login
        End If
    End If
            
End Sub

Private Sub txtFamiliar_Change()
    PetName = txtFamiliar.Text
End Sub
Private Sub txtNombre_Change()
    If Len(txtNombre.Text) > 20 Then
        txtNombre.Text = mid$(txtNombre.Text, 1, 20)
        
        txtNombre.SelStart = 20
    End If
End Sub
