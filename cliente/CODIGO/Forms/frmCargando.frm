VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmCargando 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLoad 
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   2250
      ScaleHeight     =   570
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   8070
      Width           =   15
   End
   Begin InetCtlsObjects.Inet mInet 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const WidthReal As Long = 500
Private Sub Form_Load()
    Me.picLoad.Picture = LoadInterface("barra-cargando")
    Me.picLoad.width = 0
    
    Me.Picture = LoadInterface("cargando")
    DoEvents
End Sub
Public Sub SetAddWith(ByVal porc As Long)
    If porc > 100 Then porc = 100
    
    Me.picLoad.width = Me.picLoad.width + CInt(porc * WidthReal / 100)
    If Me.picLoad.width > WidthReal Then _
        Me.picLoad.width = WidthReal
        
    Me.picLoad.refresh
    DoEvents
    
End Sub


