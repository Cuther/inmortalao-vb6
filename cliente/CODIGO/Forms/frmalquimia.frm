VERSION 5.00
Begin VB.Form frmAlquimia 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alquimia"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstPociones 
      Height          =   2205
      Left            =   120
      TabIndex        =   3
      Top             =   90
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtCantidad 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "1"
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1380
   End
End
Attribute VB_Name = "frmAlquimia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command3_Click()
    On Error Resume Next

    Call WriteCraftalquimia(Objalquimia(lstPociones.ListIndex + 1), Val(txtCantidad.Text))

    Unload Me
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub txtCantidad_Change()
If Val(txtCantidad.Text) < 0 Then
    txtCantidad.Text = 1
End If

If Val(txtCantidad.Text) > 1000 Then
    txtCantidad.Text = 1
End If

End Sub

