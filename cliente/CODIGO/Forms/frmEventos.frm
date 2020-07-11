VERSION 5.00
Begin VB.Form frmEventos2 
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   Caption         =   "Eventos"
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7635
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Carrera 
      Height          =   420
      Left            =   2770
      Top             =   2480
      Width           =   2175
   End
   Begin VB.Image Bandera 
      Height          =   420
      Left            =   2770
      Top             =   1300
      Width           =   2175
   End
   Begin VB.Image Torneo2vs2 
      Height          =   420
      Left            =   5360
      Top             =   2250
      Width           =   1890
   End
   Begin VB.Image Torneo1vs1 
      Height          =   420
      Left            =   5360
      Top             =   1440
      Width           =   1890
   End
   Begin VB.Image Arena2vs2 
      Height          =   420
      Left            =   480
      Top             =   2250
      Width           =   1890
   End
   Begin VB.Image Arena1vs1 
      Height          =   420
      Left            =   480
      Top             =   1440
      Width           =   1890
   End
   Begin VB.Image BotSalir 
      Height          =   743
      Left            =   130
      Top             =   3240
      Width           =   7410
   End
End
Attribute VB_Name = "frmEventos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Add Marius
Option Explicit

Dim Botones(1 To 7) As clsButton
Public LastPressed As clsButton

Private Sub Form_Load()

    Me.Picture = LoadInterface("eventos")
    Me.Icon = frmMain.Icon
    
    Set Botones(1) = New clsButton
    Set Botones(2) = New clsButton
    Set Botones(3) = New clsButton
    Set Botones(4) = New clsButton
    Set Botones(5) = New clsButton
    Set Botones(6) = New clsButton
    Set Botones(7) = New clsButton
    Set LastPressed = New clsButton
    
    Botones(1).Initialize Arena1vs1, _
            "[eventos]arena-1vs1-over", _
            "[eventos]arena-1vs1-down", _
            Me, True
    
    Botones(2).Initialize Arena2vs2, _
            "[eventos]arena-2vs2-over", _
            "[eventos]arena-2vs2-down", _
            Me, True
            
    Botones(3).Initialize Bandera, _
            "[eventos]bandera-entrar-over", _
            "[eventos]bandera-entrar-down", _
            Me, True
    
    Botones(4).Initialize Carrera, _
            "[eventos]carrera-participar-over", _
            "[eventos]carrera-participar-down", _
            Me, True
    
    Botones(5).Initialize Torneo1vs1, _
            "[eventos]torneo-1vs1-over", _
            "[eventos]torneo-1vs1-down", _
            Me, True
            
    Botones(6).Initialize Torneo2vs2, _
            "[eventos]torneo-2vs2-over", _
            "[eventos]torneo-2vs2-down", _
            Me, True
    
    Botones(7).Initialize BotSalir, _
            "[eventos]salir-over", _
            "[eventos]salir-down", _
            Me, True
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton And mueve = 1 Then Auto_Drag Me.hwnd
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Private Sub Arena1vs1_Click()
    If UserEstado = 1 Then 'Muerto
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
        Exit Sub
    End If
    Dim str As String
    str = "A1VS1"
    Call WriteParticipar(str)
    Unload Me
End Sub

Private Sub Arena2vs2_Click()
    If UserEstado = 1 Then 'Muerto
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
        Exit Sub
    End If
    Dim str As String
    str = "A2VS2"
    Call WriteParticipar(str)
    Unload Me
End Sub

Private Sub Bandera_Click()
    If UserEstado = 1 Then 'Muerto
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
        Exit Sub
    End If
    Dim str As String
    str = "BANDERA"
    Call WriteParticipar(str)
    Unload Me
End Sub
Private Sub Carrera_Click()
    If UserEstado = 1 Then 'Muerto
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
        Exit Sub
    End If
    Dim str As String
    str = "CARRERA"
    Call WriteParticipar(str)
    Unload Me
End Sub

Private Sub Torneo1vs1_Click()
    If UserEstado = 1 Then 'Muerto
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
        Exit Sub
    End If
    Dim str As String
    str = "T1VS1"
    Call WriteParticipar(str)
    Unload Me
End Sub

Private Sub Torneo2vs2_Click()
    If UserEstado = 1 Then 'Muerto
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
        Exit Sub
    End If
    Dim str As String
    str = "T2VS2"
    Call WriteParticipar(str)
    Unload Me
End Sub
Private Sub BotSalir_Click()
    Unload Me
End Sub
