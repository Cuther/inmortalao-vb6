VERSION 5.00
Begin VB.Form frmEstadisticas 
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   Caption         =   "Estadisticas"
   ClientHeight    =   6795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6510
   ClipControls    =   0   'False
   Icon            =   "frmEstadisticas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   434
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   5910
      TabIndex        =   49
      Top             =   5370
      Width           =   225
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   4200
      TabIndex        =   48
      Top             =   4920
      Width           =   2220
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   165
      Index           =   4
      Left            =   4800
      TabIndex        =   47
      Top             =   5790
      Width           =   435
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Acá van las habilidades especiales del familiar"
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
      Height          =   345
      Index           =   5
      Left            =   4230
      TabIndex        =   46
      Top             =   6240
      Width           =   2160
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   2
      Left            =   4365
      TabIndex        =   45
      Top             =   5430
      Width           =   1260
   End
   Begin VB.Label Fami 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Index           =   3
      Left            =   5610
      TabIndex        =   44
      Top             =   5880
      Width           =   645
   End
   Begin VB.Image imgFami 
      Height          =   1680
      Left            =   4155
      Top             =   4980
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Image iEx 
      Height          =   450
      Left            =   6120
      Top             =   0
      Width           =   390
   End
   Begin VB.Image imgStat 
      Height          =   315
      Left            =   525
      Top             =   6315
      Width           =   1080
   End
   Begin VB.Label est 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Criaturas matadas"
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
      Height          =   195
      Index           =   1
      Left            =   2250
      TabIndex        =   43
      Top             =   6180
      Width           =   1665
   End
   Begin VB.Label est 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Veces muerto"
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
      Height          =   195
      Index           =   4
      Left            =   2280
      TabIndex        =   42
      Top             =   5550
      Width           =   1665
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   6
      Left            =   1230
      TabIndex        =   41
      Top             =   5700
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   5
      Left            =   1230
      TabIndex        =   40
      Top             =   5490
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   4
      Left            =   1230
      TabIndex        =   39
      Top             =   5250
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   3
      Left            =   1230
      TabIndex        =   38
      Top             =   4890
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   2
      Left            =   1200
      TabIndex        =   37
      Top             =   4680
      Width           =   630
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   1
      Left            =   1230
      TabIndex        =   36
      Top             =   4440
      Width           =   600
   End
   Begin VB.Label est 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Clase"
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
      Height          =   135
      Index           =   0
      Left            =   960
      TabIndex        =   35
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label est 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Género"
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
      Height          =   135
      Index           =   5
      Left            =   960
      TabIndex        =   34
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label est 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Raza"
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
      Height          =   135
      Index           =   6
      Left            =   960
      TabIndex        =   33
      Top             =   3330
      Width           =   975
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   165
      Index           =   1
      Left            =   1590
      TabIndex        =   32
      Top             =   720
      Width           =   105
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   165
      Index           =   2
      Left            =   1590
      TabIndex        =   31
      Top             =   975
      Width           =   105
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   165
      Index           =   3
      Left            =   1590
      TabIndex        =   30
      Top             =   1260
      Width           =   105
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   165
      Index           =   4
      Left            =   1590
      TabIndex        =   29
      Top             =   1530
      Width           =   105
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   165
      Index           =   5
      Left            =   1590
      TabIndex        =   28
      Top             =   1800
      Width           =   105
   End
   Begin VB.Image cmdGuardar 
      Height          =   450
      Left            =   3780
      Tag             =   "1"
      Top             =   3900
      Width           =   1050
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   195
      Index           =   26
      Left            =   5700
      TabIndex        =   27
      Top             =   3400
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   5700
      TabIndex        =   26
      Top             =   3180
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   5700
      TabIndex        =   25
      Top             =   2960
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   5700
      TabIndex        =   24
      Top             =   2720
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   5700
      TabIndex        =   23
      Top             =   2500
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   5700
      TabIndex        =   22
      Top             =   2280
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   5700
      TabIndex        =   21
      Top             =   1600
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   5700
      TabIndex        =   20
      Top             =   1820
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   5700
      TabIndex        =   19
      Top             =   2040
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   5700
      TabIndex        =   18
      Top             =   1380
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   5700
      TabIndex        =   17
      Top             =   1150
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   5700
      TabIndex        =   16
      Top             =   930
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   5700
      TabIndex        =   15
      Top             =   710
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   4050
      TabIndex        =   14
      Top             =   3620
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   4050
      TabIndex        =   13
      Top             =   3400
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   4050
      TabIndex        =   12
      Top             =   3180
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   4050
      TabIndex        =   11
      Top             =   2960
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   4050
      TabIndex        =   10
      Top             =   2720
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   4050
      TabIndex        =   9
      Top             =   2500
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   4050
      TabIndex        =   8
      Top             =   2280
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   4050
      TabIndex        =   7
      Top             =   2040
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   4050
      TabIndex        =   6
      Top             =   1810
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   4050
      TabIndex        =   5
      Top             =   1590
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   4050
      TabIndex        =   4
      Top             =   1370
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   4050
      TabIndex        =   3
      Top             =   1150
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   4050
      TabIndex        =   2
      Top             =   710
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4050
      TabIndex        =   1
      Top             =   930
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   3
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":000C
      Top             =   975
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   5
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":015E
      Top             =   1215
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   7
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":02B0
      Top             =   1455
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   9
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":0402
      Top             =   1680
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   11
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":0554
      Top             =   1890
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   13
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":06A6
      Top             =   2115
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   15
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":07F8
      Top             =   2355
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   17
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":094A
      Top             =   2565
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   19
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":0A9C
      Top             =   2790
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   21
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":0BEE
      Top             =   3015
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   23
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":0D40
      Top             =   3240
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   25
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":0E92
      Top             =   3465
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   27
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":0FE4
      Top             =   3705
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   1
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":1136
      Top             =   765
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   0
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":1288
      Top             =   675
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   2
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":13DA
      Top             =   885
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   4
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":152C
      Top             =   1125
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   6
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":167E
      Top             =   1350
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   8
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":17D0
      Top             =   1590
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   10
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":1922
      Top             =   1815
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   12
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":1A74
      Top             =   2040
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   14
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":1BC6
      Top             =   2265
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   16
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":1D18
      Top             =   2475
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   18
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":1E6A
      Top             =   2715
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   20
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":1FBC
      Top             =   2925
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   22
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":210E
      Top             =   3150
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   24
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":2260
      Top             =   3390
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   26
      Left            =   4320
      MouseIcon       =   "frmEstadisticas.frx":23B2
      Top             =   3615
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   53
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":2504
      Top             =   3465
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   52
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":2656
      Top             =   3375
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   51
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":27A8
      Top             =   3255
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   50
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":28FA
      Top             =   3165
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   49
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":2A4C
      Top             =   3015
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   48
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":2B9E
      Top             =   2925
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   47
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":2CF0
      Top             =   2790
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   46
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":2E42
      Top             =   2715
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   45
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":2F94
      Top             =   2580
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   44
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":30E6
      Top             =   2475
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   43
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":3238
      Top             =   2355
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   42
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":338A
      Top             =   2265
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   28
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":34DC
      Top             =   675
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   29
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":362E
      Top             =   765
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   30
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":3780
      Top             =   900
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   31
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":38D2
      Top             =   990
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   32
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":3A24
      Top             =   1125
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   33
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":3B76
      Top             =   1215
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   34
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":3CC8
      Top             =   1365
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   35
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":3E1A
      Top             =   1440
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   36
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":3F6C
      Top             =   1575
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   37
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":40BE
      Top             =   1650
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   38
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":4210
      Top             =   1800
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   39
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":4362
      Top             =   1890
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   40
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":44B4
      Top             =   2025
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   41
      Left            =   5970
      MouseIcon       =   "frmEstadisticas.frx":4606
      Top             =   2115
      Width           =   195
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   195
      Left            =   5850
      TabIndex        =   0
      Top             =   3750
      Width           =   285
   End
   Begin VB.Shape fHPShp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Left            =   5610
      Top             =   5910
      Width           =   645
   End
   Begin VB.Shape fExpShp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Left            =   4335
      Top             =   5460
      Width           =   1275
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit
Private LibresOrig As Integer
Private RealizoCambios As Boolean
Private NewSkills(1 To NUMSKILLS) As Byte
Private Sub cmdGuardar_Click()
Dim i As Byte
If RealizoCambios Then
    If MsgBox("Los skills han sido modificados. Deseas guardarlos?", vbYesNo) = vbYes Then
        For i = 1 To NUMSKILLS
            NewSkills(i) = CByte(Skill(i - 1).Caption) - UserSkills(i)
            UserSkills(i) = Val(Skill(i - 1).Caption)
        Next i
        
        Call WriteModifySkills(NewSkills())
    End If
End If
Unload Me
End Sub


Public Sub Iniciar_Labels()
'Iniciamos los labels con los valores de los atributos y los skills
Dim i As Integer
For i = 1 To NUMATRIBUTOS
    Atri(i).Caption = UserAtributos(i)
Next

For i = 1 To NUMSKILLS
    Skill(i - 1).Caption = UserSkills(i)
Next

With UserEstadisticas
    Label4(1).Caption = .CiudadanosMatados
    Label4(2).Caption = .RepublicanosMatados
    Label4(3).Caption = .RenegadosMatados
    Label4(4).Caption = .ArmadaMatados
    Label4(5).Caption = .MiliciaMatados
    Label4(6).Caption = .CaosMatados
    
    est(0).Caption = ListaClases(.Clase)
    
    est(4).Caption = .UsuariosMatados
    est(1).Caption = .NpcMatados
    
    est(5).Caption = IIf(.Genero = 1, "Masculino", "Femenino")
    est(6).Caption = ListaRazas(.Raza)
    
    Select Case UserStat
        Case 1
            imgStat.Picture = LoadInterface("renegado")
        Case 2
            imgStat.Picture = LoadInterface("armada")
        Case 3
            imgStat.Picture = LoadInterface("caos")
        Case 4
            imgStat.Picture = LoadInterface("miliciano")
        Case 5
            imgStat.Picture = LoadInterface("republicano")
        Case 6
            imgStat.Picture = LoadInterface("imperial")
    End Select
End With

'Ponemos las estadisticas del familiar en pantalla
If UserPet.TieneFamiliar <> 0 Then
    imgFami.Picture = Nothing
    Fami(0).Visible = True
    Fami(1).Visible = True
    Fami(2).Visible = True
    Fami(3).Visible = True
    Fami(4).Visible = True
    Fami(5).Visible = True
    fHPShp.Visible = True
    fExpShp.Visible = True
    
    Fami(0).Caption = UserPet.nombre
    Fami(1).Caption = UserPet.ELV
    
    Call PetExpPerc
    
    If PetPercExp <> 0 Then
        fExpShp.width = Round((PetPercExp / 100) * 85)
    Else
        fExpShp.Visible = False
        fExpShp.width = 0
    End If
    
    Fami(2).Caption = PetPercExp & "%"
    
    If UserPet.MinHP <= 0 Then
        Fami(3).Caption = "Tu familiar esta muerto"
        Fami(3).ForeColor = vbWhite
        fHPShp.width = 0
    Else
        fHPShp.width = (((UserPet.MinHP / 100) / (UserPet.MaxHP / 100)) * 43)
        Fami(3).Caption = UserPet.MinHP & "/" & UserPet.MaxHP
    End If
    
    Fami(4).Caption = UserPet.MinHit & "/" & UserPet.MaxHit
    Fami(5).Caption = IIf(UserPet.Habilidad = vbNullString, "Ninguna", UserPet.Habilidad)
Else
    imgFami.Picture = LoadInterface("fmnodisp")
    Fami(0).Visible = False
    Fami(1).Visible = False
    Fami(2).Visible = False
    Fami(3).Visible = False
    Fami(4).Visible = False
    Fami(5).Visible = False
    fHPShp.Visible = False
    fExpShp.Visible = False
End If

LibresOrig = SkillPoints

Puntos.Caption = SkillPoints
RealizoCambios = False

End Sub

Private Sub cmdGuardar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdGuardar.Tag <> "2" Then
        cmdGuardar.Picture = LoadInterface("guardar-down")
        cmdGuardar.Tag = "2"
    End If
End Sub

Private Sub cmdGuardar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then
        If cmdGuardar.Tag <> "1" Then
            cmdGuardar.Picture = LoadInterface("guardar-over")
            cmdGuardar.Tag = "1"
        End If
    Else
        If cmdGuardar.Tag <> "2" Then
            cmdGuardar.Picture = LoadInterface("guardar-down")
            cmdGuardar.Tag = "2"
        End If
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim indice As Byte
If (Index And &H1) = 0 Then
    indice = Index \ 2
    If SkillPoints > 0 And Val(Skill(indice).Caption) < 100 Then
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
        Call Audio.PlayWave(SND_CLICK)
    End If
Else
    indice = Index \ 2
    If Val(Skill(indice).Caption) > 0 And Not (Val(Skill(indice).Caption) = SkillsOrig(indice + 1)) Then
        Skill(indice).Caption = Val(Skill(indice).Caption) - 1
        SkillPoints = SkillPoints + 1
        Call Audio.PlayWave(SND_CLICK)
    End If
End If

Puntos.Caption = SkillPoints
RealizoCambios = (SkillPoints <> LibresOrig)
Skill(indice).ForeColor = IIf(Val(Skill(indice).Caption) = SkillsOrig(indice + 1), vbWhite, vbRed)

End Sub

Private Sub Form_Load()
    Me.Picture = LoadInterface("estisticas")
    
    imgFami.Picture = LoadInterface("fmnodisp")
    imgFami.Visible = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If iEx.Tag <> "0" Then
        Set iEx.Picture = Nothing
        iEx.Tag = "0"
    End If
    If cmdGuardar.Tag <> "0" Then
        Set cmdGuardar.Picture = Nothing
        cmdGuardar.Tag = "0"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub s_Click()

End Sub

Public Function SkillRealToIndex(ByVal SkillIndex As Integer) As Integer

Select Case SkillIndex
    Case 1
        SkillRealToIndex = 4
    Case 2
        SkillRealToIndex = 5
    Case 3
        SkillRealToIndex = 20
    Case 4
        SkillRealToIndex = 7
    Case 5
        SkillRealToIndex = 23
    Case 6
        SkillRealToIndex = 19
    Case 7
        SkillRealToIndex = 12
    Case 8
        SkillRealToIndex = 2
    Case 9
        SkillRealToIndex = 22
    Case 10
        SkillRealToIndex = 6
    Case 11
        SkillRealToIndex = 8
    Case 12
        SkillRealToIndex = 18
    Case 13
        SkillRealToIndex = 1
    Case 14
        SkillRealToIndex = 3
    Case 15
        SkillRealToIndex = 11
    Case 16
        SkillRealToIndex = 9
    Case 17
        SkillRealToIndex = 17
    Case 18
        SkillRealToIndex = 13
    Case 19
        SkillRealToIndex = 14
    Case 20
        SkillRealToIndex = 10
    Case 21
        SkillRealToIndex = 26
    Case 22
        SkillRealToIndex = 16
    Case 23
        SkillRealToIndex = 15
    Case 24
        SkillRealToIndex = 24
    Case 25
        SkillRealToIndex = 25
    Case 26
        SkillRealToIndex = 21
    Case 27
        SkillRealToIndex = 27
End Select

End Function
Public Function RealSkillToIndex(ByVal Skill As Integer) As Integer

Select Case Skill
    Case 4
        RealSkillToIndex = 1
    Case 5
        RealSkillToIndex = 2
    Case 20
        RealSkillToIndex = 3
    Case 7
        RealSkillToIndex = 4
    Case 23
        RealSkillToIndex = 5
    Case 19
        RealSkillToIndex = 6
    Case 12
        RealSkillToIndex = 7
    Case 2
        RealSkillToIndex = 8
    Case 22
        RealSkillToIndex = 9
    Case 6
        RealSkillToIndex = 10
    Case 8
        RealSkillToIndex = 11
    Case 18
        RealSkillToIndex = 12
    Case 1
        RealSkillToIndex = 13
    Case 3
        RealSkillToIndex = 14
    Case 11
        RealSkillToIndex = 15
    Case 9
        RealSkillToIndex = 16
    Case 17
        RealSkillToIndex = 17
    Case 13
        RealSkillToIndex = 18
    Case 14
        RealSkillToIndex = 19
    Case 10
        RealSkillToIndex = 20
    Case 26
        RealSkillToIndex = 21
    Case 16
        RealSkillToIndex = 22
    Case 15
        RealSkillToIndex = 23
    Case 24
        RealSkillToIndex = 24
    Case 25
        RealSkillToIndex = 25
    Case 21
        RealSkillToIndex = 26
    Case 27
        RealSkillToIndex = 27
End Select

End Function

Private Sub iEx_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If iEx.Tag <> "2" Then
        iEx.Picture = LoadInterface("cerrar-est-down")
        iEx.Tag = "2"
    End If
End Sub

Private Sub iEx_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then
        If iEx.Tag <> "1" Then
            iEx.Picture = LoadInterface("cerrar-est-over")
            iEx.Tag = "1"
        End If
    Else
        If iEx.Tag <> "2" Then
            iEx.Picture = LoadInterface("cerrar-est-down")
            iEx.Tag = "2"
        End If
    End If
End Sub

Private Sub iEx_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub
Sub PetExpPerc()

    If UserPet.EXP > 0 And UserPet.ELU > 0 Then
        PetPercExp = CLng((UserPet.EXP * 100) / UserPet.ELU)
        If PetPercExp = 100 Then PetPercExp = 99
    Else
        PetPercExp = 0
    End If

End Sub
