VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Tempest AO"
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
   LockControls    =   -1  'True
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
      Left            =   0
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
      KeepAlive       =   0   'False
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
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   210
      MousePointer    =   99  'Custom
      ScaleHeight     =   417
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   544
      TabIndex        =   35
      Top             =   2055
      Width           =   8160
   End
   Begin RichTextLib.RichTextBox Desarrollo 
      Height          =   780
      Left            =   240
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   9120
      Visible         =   0   'False
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   1376
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":08CA
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
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":0947
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
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":09C4
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
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":0A41
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
      Picture         =   "frmMain.frx":0ABE
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
      Picture         =   "frmMain.frx":3EFD
      ToolTipText     =   "Modo Combate"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image modoseguro 
      Height          =   255
      Left            =   9240
      Picture         =   "frmMain.frx":433B
      ToolTipText     =   "Seguro"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Image nomodocombate 
      Height          =   255
      Left            =   8820
      Picture         =   "frmMain.frx":4779
      ToolTipText     =   "Modo Combate"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Image nomodoseguro 
      Height          =   255
      Left            =   9240
      Picture         =   "frmMain.frx":4BB7
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
Option Explicit

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long

Public MouseShift As Long
Private clicX As Long
Private clicY As Long

Private Sub cmdCon_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadInterface("Hechizos")

    picInv.Visible = False

    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    
    cmdMoverHechi(0).Enabled = True
    cmdMoverHechi(1).Enabled = True
    
    cmdMenu(0).Visible = False
    cmdMenu(1).Visible = False
    cmdMenu(2).Visible = False
    cmdMenu(3).Visible = False
    cmdMenu(4).Visible = False
    cmdMenu(5).Visible = False
    
    lblInvInfo.Visible = False
End Sub



Private Sub cmdInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdInfo.Tag <> "2" Then
        cmdInfo.Picture = LoadInterface("[hechizos]info-down")
        cmdInfo.Tag = "2"
    End If
End Sub

Private Sub cmdInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then
        If cmdInfo.Tag <> "1" Then
            cmdInfo.Picture = LoadInterface("[hechizos]info-over")
            cmdInfo.Tag = "1"
        End If
    ElseIf Button <> 0 Then
        If cmdInfo.Tag <> "2" Then
            cmdInfo.Picture = LoadInterface("[hechizos]info-down")
            cmdInfo.Tag = "2"
        End If
    End If
End Sub

Private Sub cmdINV_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadInterface("Inventory")
    picInv.Visible = True

    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    
    cmdMoverHechi(0).Visible = False
    cmdMoverHechi(1).Visible = False
    
    cmdMoverHechi(0).Enabled = False
    cmdMoverHechi(1).Enabled = False
    lblInvInfo.Visible = True
    
    cmdMenu(0).Visible = False
    cmdMenu(1).Visible = False
    cmdMenu(2).Visible = False
    cmdMenu(3).Visible = False
    cmdMenu(4).Visible = False
    cmdMenu(5).Visible = False
    
    RenderInv = True
End Sub
Private Sub cmdMen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then
        If cmdMen.Tag <> "1" Then
            cmdMen.Tag = "1"
            cmdMen.Picture = LoadInterface("[solapas]menu-over")
        End If
    End If
    If cmdInv.Tag <> "0" Then
        cmdInv.Tag = "0"
        Set cmdInv.Picture = Nothing
    End If

    If cmdCon.Tag <> "0" Then
        cmdCon.Tag = "0"
        Set cmdCon.Picture = Nothing
    End If
End Sub

Private Sub cmdInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then
        If cmdInv.Tag <> "1" Then
            cmdInv.Tag = "1"
            cmdInv.Picture = LoadInterface("[solapas]inventario-over")
        End If
    End If
    
    If cmdMen.Tag <> "0" Then
        cmdMen.Tag = "0"
        Set cmdMen.Picture = Nothing
    End If

    If cmdCon.Tag <> "0" Then
        cmdCon.Tag = "0"
        Set cmdCon.Picture = Nothing
    End If
End Sub

Private Sub cmdCon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then
        If cmdCon.Tag <> "1" Then
            cmdCon.Tag = "1"
            cmdCon.Picture = LoadInterface("[solapas]hechizos-over")
        End If
    End If
    
    If cmdMen.Tag <> "0" Then
        cmdMen.Tag = "0"
        Set cmdMen.Picture = Nothing
    End If

    If cmdInv.Tag <> "0" Then
        cmdInv.Tag = "0"
        Set cmdInv.Picture = Nothing
    End If
End Sub

Private Sub CmdLanzar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If CmdLanzar.Tag <> "2" Then
        CmdLanzar.Picture = LoadInterface("[hechizos]lanzar-down")
        CmdLanzar.Tag = "2"
    End If
End Sub

Private Sub cmdMen_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadInterface("Menu")
    picInv.Visible = False

    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    
    cmdMoverHechi(0).Visible = False
    cmdMoverHechi(1).Visible = False
    
    cmdMoverHechi(0).Enabled = False
    cmdMoverHechi(1).Enabled = False
    
    lblInvInfo.Visible = False
    
    cmdMenu(0).Visible = True
    cmdMenu(1).Visible = True
    cmdMenu(2).Visible = True
    cmdMenu(3).Visible = True
    cmdMenu(4).Visible = True
    cmdMenu(5).Visible = True
End Sub



Private Sub cmdMenu_Click(Index As Integer)
    Call Audio.PlayWave(SND_CLICK)

    Select Case Index
        Case 0
            If GrupoIndex <> 0 Then
                Call WriteRequestGrupo
                FlushBuffer
            Else
                UsingSkill = 250
                cCursores.Parse_Form Me, E_SHOOT
            End If
            
        Case 5
            Call frmOpciones.Show(vbModeless, frmMain)
            
        Case 1
            LlegaronEstadisticas = False
            Call WriteRequestEstadisticas
            Call FlushBuffer
            
            Do While Not LlegaronEstadisticas
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            LlegaronEstadisticas = False
        
        Case 2
            If frmGuildLeader.Visible Then Unload frmGuildLeader
            
            Call WriteRequestGuildLeaderInfo
    End Select
End Sub

Private Sub cmdMinimizar_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)
    If hlst.ListIndex = -1 Then Exit Sub
    Dim sTemp As String

    Select Case Index
        Case 1 'subir
            If hlst.ListIndex = 0 Then Exit Sub
        Case 0 'bajar
            If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
    End Select

    Call WriteMoveSpell(Index, hlst.ListIndex + 1)
    
    Select Case Index
        Case 1 'subir
            sTemp = hlst.List(hlst.ListIndex - 1)
            hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = sTemp
            hlst.ListIndex = hlst.ListIndex - 1
        Case 0 'bajar
            sTemp = hlst.List(hlst.ListIndex + 1)
            hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = sTemp
            hlst.ListIndex = hlst.ListIndex + 1
    End Select
End Sub


Public Sub ControlSeguroResu(ByVal Mostrar As Boolean)
If Mostrar Then
    'If Not PicResu.Visible Then
    '    PicResu.Visible = True
    'End If
Else
    'If PicResu.Visible Then
    '    PicResu.Visible = False
    'End If
End If
End Sub


Public Sub DibujarSeguro()
modoseguro.Visible = True
nomodoseguro.Visible = False
End Sub

Public Sub DesDibujarSeguro()
modoseguro.Visible = False
nomodoseguro.Visible = True
End Sub


Private Sub cmdMoverHechi_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            If cmdMoverHechi(0).Tag <> "2" Then
                cmdMoverHechi(0).Picture = LoadInterface("[hechizos]flechaabajo-down")
                cmdMoverHechi(0).Tag = "2"
            End If
        Case 1
            If cmdMoverHechi(1).Tag <> "2" Then
                cmdMoverHechi(1).Picture = LoadInterface("[hechizos]flechaarriba-down")
                cmdMoverHechi(1).Tag = "2"
            End If
    End Select
End Sub

Private Sub cmdMoverHechi_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            If Button = 0 Then
                If cmdMoverHechi(0).Tag <> "1" Then
                    cmdMoverHechi(0).Picture = LoadInterface("[hechizos]flechaabajo-over")
                    cmdMoverHechi(0).Tag = "1"
                End If
            End If
        Case 1
            If Button = 0 Then
                If cmdMoverHechi(1).Tag <> "1" Then
                    cmdMoverHechi(1).Picture = LoadInterface("[hechizos]flechaarriba-over")
                    cmdMoverHechi(1).Tag = "1"
                End If
            End If
    End Select
End Sub

Private Sub cmdCerrar_Click()
'    If MsgBox("Está seguro que desea salir?", vbYesNo + vbCritical, "Salir") = vbYes Then
        Call WriteQuit
 '       End
'    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If (Not SendTxt.Visible) Then
        
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleCombatMode)
                    Call WriteCombatModeToggle
                    IScombate = Not IScombate
                    If IScombate = True Then
                        modocombate.Visible = True
                        nomodocombate.Visible = False
                    Else
                        modocombate.Visible = False
                        nomodocombate.Visible = True
                    End If
                    
                
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    Call WriteWork(eSkill.Domar)
                
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    Call WriteWork(eSkill.Robar)
                            
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    Call WriteWork(eSkill.Ocultarse)
                
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)

                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                    Call WriteSafeToggle
                Case CustomKeys.BindedKey(eKeyType.mKeyBloqMov)
                    BloqMov = Not BloqMov
            End Select
        End If
    End If
    

    Select Case KeyCode
        Case vbKeyF1 To vbKeyF11
            Call frmBindKey.Bind_Accion(KeyCode - vbKeyF1 + 1)
            
        Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
            Call ScreenCapture
        
        Case CustomKeys.BindedKey(eKeyType.mKeyToggleFPS)
            FPSFLAG = Not FPSFLAG

        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
            If Shift <> 0 Then Exit Sub
            
            If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
            If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
            Else
                If Not MainTimer.Check(TimersIndex.Attack) Or UserDescansar Or UserMeditar Then Exit Sub
            End If
            
            Call WriteAttack
        
        Case KeyCodeConstants.vbKeyReturn
            If (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And _
              (Not frmBancoObj.Visible) And _
              (Not frmMSG.Visible) And (Not frmForo.Visible) And _
              (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                If SendTxt.Visible = False Then

                    Select Case CurrentUser.SendingType
                        Case 1 'Normal
                            SendTxt.Text = ""
                        Case 2 'Gritar
                            SendTxt.Text = "-"
                        Case 3 'Privado
                            SendTxt.Text = "\" & CurrentUser.sndPrivateTo & " "
                        Case 4 'Clan
                            SendTxt.Text = "/CMSG "
                        Case 5 'Grupo
                            SendTxt.Text = "/PMSG "
                        Case 6 'Global
                            SendTxt.Text = ";"
                        Case 7 'Faccion
                            SendTxt.Text = "/FMSG "
                        Case Else
                            SendTxt.Text = ""
                    End Select
                    
                    SendTxt.SelStart = Len(SendTxt.Text)
                End If
                SendTxt.Visible = True
                SendTxt.SetFocus
            End If
            
    End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton And mueve = 1 Then Auto_Drag Me.hwnd
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If UserPasarNivel = 0 Then
        lblExp.Caption = "¡Nivel máximo!"
    Else
        frmMain.lblExp.Caption = CStr(Round(UserExp * 100 / UserPasarNivel) & "%") & " (" & UserExp & "/" & UserPasarNivel & ")"
    End If
    
    If CmdLanzar.Tag <> "0" Then
        Set CmdLanzar.Picture = Nothing
        CmdLanzar.Tag = "0"
    End If
    
    If cmdInfo.Tag <> "0" Then
        Set cmdInfo.Picture = Nothing
        cmdInfo.Tag = "0"
    End If
    
    If cmdMoverHechi(0).Tag <> "0" Then
        Set cmdMoverHechi(0).Picture = Nothing
        cmdMoverHechi(0).Tag = "0"
    End If
    
    If cmdMoverHechi(1).Tag <> "0" Then
        Set cmdMoverHechi(1).Picture = Nothing
        cmdMoverHechi(1).Tag = "0"
    End If
    
    If cmdMen.Tag <> "0" Then
        cmdMen.Tag = "0"
        Set cmdMen.Picture = Nothing
    End If
    
    If imgTalk.Tag <> "0" Then
        imgTalk.Tag = "0"
        Set imgTalk.Picture = Nothing
    End If

    If cmdInv.Tag <> "0" Then
        cmdInv.Tag = "0"
        Set cmdInv.Picture = Nothing
    End If

    If cmdCon.Tag <> "0" Then
        cmdCon.Tag = "0"
        Set cmdCon.Picture = Nothing
    End If

    frmMain.Label2(0).Caption = "Posición: " & UserMap & ", " & UserPos.X & ", " & UserPos.Y

    Dim i As Byte
    For i = 0 To 5
        If cmdMenu(i).Tag <> "0" Then
            cmdMenu(i).Tag = "0"
            Set cmdMenu(i).Picture = Nothing
        End If
    Next i
End Sub





Private Sub imgTalk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgTalk.Tag <> "2" Then
        imgTalk.Tag = "2"
        imgTalk.Picture = LoadInterface("modotextodown")
    End If
End Sub

Private Sub imgTalk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then
        If imgTalk.Tag <> "1" Then
            imgTalk.Tag = "1"
            imgTalk.Picture = LoadInterface("modotextoover")
        End If
    End If
End Sub

Private Sub imgTalk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call imgTalk_MouseMove(Button, Shift, X, Y)
frmCustomKeys.PopupMenuMensaje
End Sub

Private Sub InvEqu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton And mueve = 1 Then Auto_Drag Me.hwnd
End Sub

Private Sub InvEqu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If CmdLanzar.Tag <> "0" Then
        Set CmdLanzar.Picture = Nothing
        CmdLanzar.Tag = "0"
    End If
    
    If cmdInfo.Tag <> "0" Then
        Set cmdInfo.Picture = Nothing
        cmdInfo.Tag = "0"
    End If
    
    If cmdMoverHechi(0).Tag <> "0" Then
        Set cmdMoverHechi(0).Picture = Nothing
        cmdMoverHechi(0).Tag = "0"
    End If
    
    If cmdMoverHechi(1).Tag <> "0" Then
        Set cmdMoverHechi(1).Picture = Nothing
        cmdMoverHechi(1).Tag = "0"
    End If
    
        If cmdMen.Tag <> "0" Then
        cmdMen.Tag = "0"
        Set cmdMen.Picture = Nothing
    End If

    If cmdInv.Tag <> "0" Then
        cmdInv.Tag = "0"
        Set cmdInv.Picture = Nothing
    End If

    If cmdCon.Tag <> "0" Then
        cmdCon.Tag = "0"
        Set cmdCon.Picture = Nothing
    End If
    
    frmMain.Label2(0).Caption = "Posición: " & UserMap & ", " & UserPos.X & ", " & UserPos.Y
    
        Dim i As Byte
    For i = 0 To 5
        If cmdMenu(i).Tag <> "0" Then
            cmdMenu(i).Tag = "0"
            Set cmdMenu(i).Picture = Nothing
        End If
    Next i
End Sub


Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2(0).Caption = RTrim$(MapName)
End Sub

Private Sub lblFU_Click()
#If Coneccion Then
    Debug.Print Accomont.PeekASCIIStringFixed(Accomont.length)
#End If
End Sub

Private Sub lblTxtCombat_Click()
    RecChat.Visible = False
    RecCombat.Visible = True
    RecGlobal.Visible = False
    
    lblTxtGlobal.Font.bold = False
    lblTxtDefault.Font.bold = False
    lblTxtCombat.Font.bold = True
End Sub
Private Sub lblTxtDefault_Click()
    RecChat.Visible = True
    RecCombat.Visible = False
    RecGlobal.Visible = False
    
    lblTxtGlobal.Font.bold = False
    lblTxtDefault.Font.bold = True
    lblTxtCombat.Font.bold = False
End Sub
Private Sub lblTxtGlobal_Click()
    RecChat.Visible = False
    RecCombat.Visible = False
    RecGlobal.Visible = True
    
    lblTxtGlobal.Font.bold = True
    lblTxtDefault.Font.bold = False
    lblTxtCombat.Font.bold = False
End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub MainViewPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub


Private Sub Minimap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call WriteWarpChar("YO", UserMap, IIf(X < 1, 1, X), IIf(Y < 1, 1, Y))
    DibujarMiniMapPos
End Sub


Private Sub modocombate_Click()
    Call WriteCombatModeToggle
    IScombate = Not IScombate
    If IScombate = True Then
        modocombate.Visible = True
        nomodocombate.Visible = False
    Else
        modocombate.Visible = False
        nomodocombate.Visible = True
    End If
End Sub

Private Sub modoseguro_Click()
    Call WriteSafeToggle
End Sub

Private Sub nomodocombate_Click()
    Call WriteCombatModeToggle
    IScombate = Not IScombate
    If IScombate = True Then
        modocombate.Visible = True
        nomodocombate.Visible = False
    Else
        modocombate.Visible = False
        nomodocombate.Visible = True
    End If
End Sub

Private Sub nomodoseguro_Click()
Call WriteSafeToggle
End Sub

Private Sub picInv_Paint()
    RenderInv = True
End Sub

Private Sub picMacro_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
BotonElegido = Index + 1

If MacroKeys(BotonElegido).TipoAccion = 0 Or Button = vbRightButton Then
    frmBindKey.Show vbModeless, frmMain
Else
    Call frmBindKey.Bind_Accion(Index + 1)
End If
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
    End If
End Sub



'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            Call WriteDrop(Inventario.SelectedItem, 1)
        Else
           If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                frmCantidad.Show , frmMain
           End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    Call WritePickUp
End Sub

Private Sub UsarItem()

    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
            Call WriteUseItem(Inventario.SelectedItem)

End Sub

Private Sub EquiparItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteEquipItem(Inventario.SelectedItem)
End Sub

''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''


Private Sub cmdLanzar_Click()
    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.Work, False) Then
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.Magia)
        End If
    End If
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then
        If CmdLanzar.Tag <> "1" Then
            CmdLanzar.Picture = LoadInterface("[hechizos]lanzar-over")
            CmdLanzar.Tag = "1"
        End If
    ElseIf Button <> 0 Then
        If CmdLanzar.Tag <> "2" Then
            CmdLanzar.Picture = LoadInterface("[hechizos]lanzar-down")
            CmdLanzar.Tag = "2"
        End If
    End If
End Sub

Private Sub cmdINFO_Click()
    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)
    End If
End Sub



Private Sub MainViewPic_Click()

    If Not Comerciando Then
        Call TileEngine.Engine_Convert_CP_To_TP(MouseX, MouseY, tX, tY)
    
        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                If UsingSkill = 0 Then
                    Call WriteLeftClick(tX, tY)
  
                    
                    If MapData(tX, tY).NPCIndex <> 0 Then
                        If npcs(MapData(tX, tY).NPCIndex).Desc <> "" Then _
                            Call TileEngine.Char_Dialog_Create(MapData(tX, tY).CharIndex, npcs(MapData(tX, tY).NPCIndex).Desc, -1)
                    End If
                    
                    If MapData(tX, tY + 1).NPCIndex <> 0 Then
                        If npcs(MapData(tX, tY + 1).NPCIndex).Desc <> "" Then _
                            Call TileEngine.Char_Dialog_Create(MapData(tX, tY + 1).CharIndex, npcs(MapData(tX, tY + 1).NPCIndex).Desc, -1)
                    End If
                    
                    If MapData(tX, tY).OBJInfo.OBJIndex <> 0 Then
                        If MostrarCantidad(MapData(tX, tY).OBJInfo.OBJIndex) Then
                            Call AddToChat(objs(MapData(tX, tY).OBJInfo.OBJIndex).name & " - " & MapData(tX, tY).OBJInfo.Amount)
                        Else
                            Call AddToChat(objs(MapData(tX, tY).OBJInfo.OBJIndex).name)
                        End If
                    End If
                Else

                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        cCursores.Parse_Form Me
                        UsingSkill = 0
                        Exit Sub
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Or UsingSkill = Arrojadizas Then
                        If Not MainTimer.Check(TimersIndex.Arrows) Then
                            cCursores.Parse_Form Me
                            UsingSkill = 0
                            Exit Sub
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Magia Then
                        If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
                            If Not MainTimer.Check(TimersIndex.CastAttack) Then 'Corto intervalo de Golpe-Magia
                                cCursores.Parse_Form Me
                                UsingSkill = 0
                                Exit Sub
                            End If
                        Else
                            If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                cCursores.Parse_Form Me
                                UsingSkill = 0
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                        If Not MainTimer.Check(TimersIndex.Work) Then
                            cCursores.Parse_Form Me
                            UsingSkill = 0
                            Exit Sub
                        End If
                    End If

                    cCursores.Parse_Form Me
                    
                    Call WriteWorkLeftClick(tX, tY, UsingSkill)
                    UsingSkill = 0
                End If
            Else
                If Not frmForo.Visible And Not frmComerciar.Visible And Not frmBancoObj.Visible Then
                    Call WriteDoubleClick(tX, tY)
                End If
            End If
        ElseIf MouseBoton = vbRightButton Then
            Call WriteDoubleClick(tX, tY)
        End If
    End If
End Sub

Private Sub MainViewPic_DblClick()
    #If ClienteGM = 1 Then
        If MouseBoton = vbRightButton Then
            Call WriteWarpChar("YO", UserMap, tX, tY)
        End If
    #End If
End Sub

Private Sub Form_Load()

    SetWindowLong RecChat.hwnd, -20, &H20& 'Consola Transparente
    SetWindowLong RecGlobal.hwnd, -20, &H20&
    SetWindowLong RecCombat.hwnd, -20, &H20&
    
    frmMain.Caption = "Inmortal AO"
    
    Me.Picture = LoadInterface("Main")
    InvEqu.Picture = LoadInterface("Inventory")
    
    Me.Left = 0
    Me.Top = 0

    lblTxtGlobal.Font.bold = False
    lblTxtDefault.Font.bold = True
    lblTxtCombat.Font.bold = False
    
    cmdMenu(0).Visible = False
    cmdMenu(1).Visible = False
    cmdMenu(2).Visible = False
    cmdMenu(3).Visible = False
    cmdMenu(4).Visible = False
    cmdMenu(5).Visible = False
    
    #If Coneccion Then
        Desarrollo.Visible = True
        Me.height = 10000
    #Else
        Me.height = 9000
    #End If
    
    RecChat.Visible = True
    
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
    
    If CmdLanzar.Tag <> "0" Then
        Set CmdLanzar.Picture = Nothing
        CmdLanzar.Tag = "0"
    End If
    
    If cmdInfo.Tag <> "0" Then
        Set cmdInfo.Picture = Nothing
        cmdInfo.Tag = "0"
    End If
    
    If imgTalk.Tag <> "0" Then
        imgTalk.Tag = "0"
        Set imgTalk.Picture = Nothing
    End If
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub


Private Sub cmdDropGold_Click()
    Inventario.SelectGold
    If UserGLD > 0 Then
        frmCantidad.Show , frmMain
    End If
End Sub


Private Sub picInv_DblClick()
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    Dim tipo As Byte
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
        If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otMapas Then
            frmMap.Show , frmMain
            frmMap.Top = frmMain.Top
            frmMap.Left = frmMain.Left
            frmMap.SetMapPoint
        Else
            Call WriteUseItem(Inventario.SelectedItem)
        End If
    End If

End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub RecTxt_Change()
On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If Not IsAppActive() Then Exit Sub
    
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And _
      (Not frmBancoObj.Visible) And _
      (Not frmMSG.Visible) And (Not frmForo.Visible) And _
      (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) And (picInv.Visible) Then
        picInv.SetFocus
    End If
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    
    If Len(SendTxt.Text) > 1600 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub


Sub ScreenCapture()

    Clipboard.Clear
    TileEngine.Inventory_Render
    keybd_event &H2C, 1, 0, 0
    
    DoEvents
    If Clipboard.GetFormat(vbCFBitmap) Then
        SavePicture Clipboard.GetData(vbCFBitmap), App.Path & "\Screenshots\" & Format(Now, "DD-MM-YYYY hh-mm-ss") & ".bmp"
        Call AddtoRichTextBox(Me.RecChat, "Imagen capturada correctamente!! Se ha almacenado en " & App.Path & "\Screenshots\" & Format(Now, "DD-MM-YYYY hh-mm-ss") & ".bmp", FontTypes(3).red, FontTypes(3).green, FontTypes(3).blue)
    Else
        MsgBox " Error ", vbCritical
    End If
    
End Sub
Private Sub cmdMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdMenu(Index).Tag = "2"
Select Case Index
    Case 0 'Grupo
        cmdMenu(0).Picture = LoadInterface("[menu]grupo-down")
    Case 1 'Estadisticas
        cmdMenu(1).Picture = LoadInterface("[menu]estadisticas-down")
    Case 2 'Guild
        cmdMenu(2).Picture = LoadInterface("[menu]clanes-down")
    Case 3 'Quest
        cmdMenu(3).Picture = LoadInterface("[menu]quests-down")
    Case 4 'Torneos
        cmdMenu(4).Picture = LoadInterface("[menu]torneos-down")
    Case 5 'Opciones
        cmdMenu(5).Picture = LoadInterface("[menu]opciones-down")
End Select

End Sub

Private Sub cmdMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If cmdMenu(Index).Tag = "1" Then
    Exit Sub
Else
    cmdMenu(Index).Tag = "1"
End If

Select Case Index

    Case 0 'Grupo
        cmdMenu(0).Picture = LoadInterface("[menu]grupo-over")
    Case 1 'Estadisticas
        cmdMenu(1).Picture = LoadInterface("[menu]estadisticas-over")
    Case 2 'Guild
        cmdMenu(2).Picture = LoadInterface("[menu]clanes-over")
    Case 3 'Quest
        cmdMenu(3).Picture = LoadInterface("[menu]quests-over")
    Case 4 'Torneos
        cmdMenu(4).Picture = LoadInterface("[menu]torneos-over")
    Case 5 'Opciones
        cmdMenu(5).Picture = LoadInterface("[menu]opciones-over")
End Select

End Sub
Private Sub cmdCerrar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMiniCerra.Picture = LoadInterface("cerrardown")
End Sub
Private Sub cmdMinimizar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMiniCerra.Picture = LoadInterface("minimizardown")
End Sub
Private Sub cmdMinimizar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMiniCerra.Picture = LoadInterface("minimizarover")
End Sub

Private Sub cmdCerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMiniCerra.Picture = LoadInterface("cerrarover")
End Sub






Private Sub Socket1_Connect()
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.length)

    Select Case EstadoLogin
    
    
        Case E_MODO.CrearNuevoPj
            Call Login
        
        Case E_MODO.Normal
            Call Login
            
        Case E_MODO.ConectarCuenta
            Call Login
        
        Case E_MODO.CrearNuevaCuenta
            Call Login

    End Select
End Sub

Private Sub Socket1_Disconnect()
    Dim i As Long
    

    Connected = False
    
    Socket1.Cleanup

    
    cCursores.Parse_Form frmConnect
    

    Do While i < Forms.count - 1
        i = i + 1
        
        If Forms(i).name <> Me.name And Forms(i).name <> frmConnect.name And Forms(i).name <> frmCrearPersonaje.name Then
            Unload Forms(i)
        End If
    Loop
    
    On Local Error GoTo 0
    
    If Not frmCrearPersonaje.Visible Then
        frmConnect.Visible = True
    End If
    
    frmMain.Visible = False
    
    Pausa = False
    UserMeditar = False
    

    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    
    For i = 1 To MAX_INVENTORY_SLOTS
        
    Next i


    SkillPoints = 0
    Alocados = 0
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        frmMensaje.msg.Caption = "Por favor espere, intentando completar conexion."
        frmMensaje.Show , frmConnect
    ElseIf ErrorCode = 24061 Then
        frmMensaje.msg.Caption = "No hay coneccion con el servidor. Verifique su coneccion de internet, o bien fijese que el servidor este online en www.inmortalao.com.ar"
        frmMensaje.Show , frmConnect
    Else
        frmMensaje.msg.Caption = ErrorString
        frmMensaje.Show , frmConnect
    End If
    
    frmMain.Socket1.Disconnect

    Err.Clear

   
    
End Sub


Private Sub Socket1_Read(dataLength As Integer, IsUrgent As Integer)
    Dim RD As String
    Dim data() As Byte
    
    Call Socket1.Read(RD, dataLength)
    data = StrConv(RD, vbFromUnicode)
    
    If RD = vbNullString Then Exit Sub

    'Put data in the buffer
    Call incomingData.WriteBlock(data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub
