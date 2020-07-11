VERSION 5.00
Begin VB.Form frmMensaje 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2685
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   179
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   261
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label msg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   1785
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3675
      WordWrap        =   -1  'True
   End
   Begin VB.Image cmdAceptar 
      Height          =   465
      Left            =   1365
      Tag             =   "1"
      Top             =   2070
      Width           =   1200
   End
   Begin VB.Menu a 
      Caption         =   "Hola"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Boton As clsButton
Public LastPressed As clsButton

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub Form_Deactivate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    Set Boton = New clsButton
    Set LastPressed = New clsButton
    
    Boton.Initialize cmdAceptar, _
            "infookover", _
            "infookdown", _
            Me, True

    Me.Picture = LoadInterface("info")
    
    Aplicar_Transparencia Me.hwnd, 200
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Auto_Drag Me.hwnd
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub msg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Auto_Drag Me.hwnd
End Sub
