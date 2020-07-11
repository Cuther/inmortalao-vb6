VERSION 5.00
Begin VB.Form frmEventos 
   Caption         =   "Eventos"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   6255
   End
   Begin VB.CommandButton Torneo2vs2 
      Caption         =   "2 vs 2"
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Torneo1vs1 
      Caption         =   "1 vs 1"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Carrera 
      Caption         =   "Participar"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton Bandera 
      Caption         =   "Entrar"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   360
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Caption         =   "Torneo"
      Height          =   1575
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Carrera"
      Height          =   735
      Left            =   2280
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Captura la Bandera"
      Height          =   735
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Arena2vs2 
      Caption         =   "2 vs 2"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Arenas"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton Arena1vs1 
         Caption         =   "1 vs 1"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cerrar_Click()
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
