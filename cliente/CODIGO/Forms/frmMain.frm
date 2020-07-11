VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Inmortal AO"
   ClientHeight    =   9000
   ClientLeft      =   345
   ClientTop       =   360
   ClientWidth     =   12000
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   NegotiateMenus  =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   600
      Top             =   0
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   10240
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   -1  'True
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   10000
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   8400
      TabIndex        =   35
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   210
      MousePointer    =   99  'Custom
      ScaleHeight     =   417
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   544
      TabIndex        =   34
      Top             =   2055
      Width           =   8160
   End
   Begin RichTextLib.RichTextBox RecCombat 
      Height          =   1500
      Left            =   210
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   165
      Visible         =   0   'False
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   2646
      _Version        =   393217
      BackColor       =   255
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RecGlobal 
      Height          =   1500
      Left            =   210
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   165
      Visible         =   0   'False
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   2646
      _Version        =   393217
      BackColor       =   12632319
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":0387
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   0
      Left            =   240
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   27
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   1
      Left            =   825
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   26
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   3
      Left            =   1995
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   25
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   2
      Left            =   1410
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   24
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   4
      Left            =   2580
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   23
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   5
      Left            =   3165
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   22
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   6
      Left            =   3750
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   21
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   7
      Left            =   4335
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   20
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   8
      Left            =   4920
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   19
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   9
      Left            =   5505
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   18
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   10
      Left            =   6090
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   17
      Top             =   8430
      Width           =   480
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      Left            =   240
      MaxLength       =   500
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1725
      Visible         =   0   'False
      Width           =   7470
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2565
      Left            =   8850
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2100
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox Minimap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1410
      Left            =   10200
      ScaleHeight     =   94
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   7
      Top             =   7380
      Width           =   1455
      Begin VB.Shape UserP 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   45
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   45
      End
   End
   Begin RichTextLib.RichTextBox RecChat 
      Height          =   1500
      Left            =   210
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   165
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   2646
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":0404
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   2400
      Left            =   9030
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   1
      Top             =   2175
      Width           =   2400
   End
   Begin VB.Image nuevocorreo 
      Height          =   255
      Left            =   9630
      Picture         =   "frmMain.frx":0481
      ToolTipText     =   "Nuevo correo"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdCerrar 
      Height          =   225
      Left            =   11580
      Top             =   180
      Width           =   255
   End
   Begin VB.Image cmdMinimizar 
      Height          =   225
      Left            =   11280
      Top             =   180
      Width           =   225
   End
   Begin VB.Image imgMiniCerra 
      Enabled         =   0   'False
      Height          =   315
      Left            =   11340
      Top             =   150
      Width           =   510
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   5
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   4905
      Width           =   1890
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   3
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   3735
      Width           =   1890
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   2
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   3150
      Width           =   1890
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   1
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   2565
      Width           =   1890
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   0
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   1980
      Width           =   1890
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   4
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   4320
      Width           =   1890
   End
   Begin VB.Image imgTalk 
      Height          =   300
      Left            =   7800
      Top             =   1725
      Width           =   600
   End
   Begin VB.Label lblTxtCombat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Combate"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   720
      TabIndex        =   33
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label lblTxtDefault 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chat"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   32
      Top             =   1800
      Width           =   390
   End
   Begin VB.Label lblTxtGlobal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Global"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1560
      TabIndex        =   31
      Top             =   1800
      Width           =   525
   End
   Begin VB.Image cmdMen 
      Height          =   540
      Left            =   10740
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Image CmdLanzar 
      Height          =   435
      Left            =   8775
      MousePointer    =   99  'Custom
      Top             =   4905
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa desconocido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   8640
      TabIndex        =   28
      Top             =   7020
      Width           =   3105
   End
   Begin VB.Image modocombate 
      Height          =   255
      Left            =   8820
      Picture         =   "frmMain.frx":38C0
      ToolTipText     =   "Modo Combate"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image modoseguro 
      Height          =   255
      Left            =   9240
      Picture         =   "frmMain.frx":3CFE
      ToolTipText     =   "Seguro"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Image nomodocombate 
      Height          =   255
      Left            =   8820
      Picture         =   "frmMain.frx":413C
      ToolTipText     =   "Modo Combate"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Image nomodoseguro 
      Height          =   255
      Left            =   9240
      Picture         =   "frmMain.frx":457A
      ToolTipText     =   "Seguro"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgHora 
      Height          =   480
      Left            =   6675
      Top             =   8430
      Width           =   1695
   End
   Begin VB.Label lblHAM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   135
      Left            =   10320
      TabIndex        =   16
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label lblSED 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   135
      Left            =   10320
      TabIndex        =   15
      Top             =   6660
      Width           =   1365
   End
   Begin VB.Label lblFU 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9330
      TabIndex        =   14
      Top             =   8340
      Width           =   345
   End
   Begin VB.Label lblAG 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9330
      TabIndex        =   13
      Top             =   8550
      Width           =   345
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   135
      Left            =   8745
      TabIndex        =   12
      Top             =   5850
      Width           =   1350
   End
   Begin VB.Label lblMP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   135
      Left            =   8745
      TabIndex        =   11
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label lblST 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   135
      Left            =   8745
      TabIndex        =   10
      Top             =   6660
      Width           =   1365
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0% (0/0)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   8880
      TabIndex        =   6
      Top             =   870
      Width           =   1815
   End
   Begin VB.Label lblNick 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NickDelPersonaje"
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
      Height          =   375
      Left            =   8640
      TabIndex        =   5
      Top             =   240
      Width           =   2625
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   420
      Index           =   0
      Left            =   11475
      MousePointer    =   99  'Custom
      Top             =   3390
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label GldLbl 
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
      Height          =   135
      Left            =   10620
      TabIndex        =   4
      Top             =   5745
      Width           =   1110
   End
   Begin VB.Image cmdDropGold 
      Height          =   300
      Left            =   10260
      MousePointer    =   99  'Custom
      Top             =   5670
      Width           =   300
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   420
      Index           =   1
      Left            =   11475
      MousePointer    =   99  'Custom
      Top             =   2880
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdInfo 
      Height          =   435
      Left            =   10650
      MousePointer    =   99  'Custom
      Top             =   4905
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Image cmdCon 
      Height          =   540
      Left            =   9660
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Image cmdInv 
      Height          =   540
      Left            =   8580
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10995
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape ExpShp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8835
      Top             =   900
      Width           =   1800
   End
   Begin VB.Shape MANShp 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8745
      Top             =   6285
      Width           =   1365
   End
   Begin VB.Shape Hpshp 
      BackColor       =   &H00000080&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8745
      Top             =   5880
      Width           =   1365
   End
   Begin VB.Shape STAShp 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   135
      Left            =   8745
      Top             =   6690
      Width           =   1365
   End
   Begin VB.Shape AGUAsp 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      Height          =   135
      Left            =   10320
      Top             =   6690
      Width           =   1365
   End
   Begin VB.Shape COMIDAsp 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   135
      Left            =   10320
      Top             =   6285
      Width           =   1365
   End
   Begin VB.Label lblInvInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   9000
      TabIndex        =   8
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Image InvEqu 
      Height          =   4275
      Left            =   8580
      Top             =   1200
      Width           =   3240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
