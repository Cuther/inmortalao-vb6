VERSION 5.00
Begin VB.Form frmPregunta 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   179
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   261
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Image img 
      Height          =   465
      Index           =   1
      Left            =   2010
      Top             =   2070
      Width           =   1215
   End
   Begin VB.Image img 
      Height          =   465
      Index           =   0
      Left            =   720
      Top             =   2070
      Width           =   1200
   End
End
Attribute VB_Name = "frmPregunta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Enum eAction
    Retirar = 1
    Duelo = 2
    Hogar = 3
End Enum
Public Accion As Byte
Public Sub SetAccion(ByVal action As Byte, ByRef a As String)
    Accion = action

    If action = eAction.Retirar Then
        Label1.Caption = "Los renegados no pueden formar parte de grupos y pierden 10% de la experiencia ganada, para recuperar su ciudadania debera pagar un precio luego. En caso de entender esto presione Aceptar."
    ElseIf action = eAction.Duelo Then
        Label1.Caption = a & " te esta retando a un duelo."
    End If
End Sub

Private Sub Form_Load()
    Me.Picture = LoadInterface("preg")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If img(0).Tag <> "0" Then
        Set img(0).Picture = Nothing
        img(0).Tag = "0"
    End If
    
    If img(1).Tag <> "0" Then
        Set img(1).Picture = Nothing
        
        img(1).Tag = "0"
    End If
End Sub

Private Sub img_Click(Index As Integer)
    Select Case Index
        Case 0
            If Accion = eAction.Retirar Then
                WriteLeaveFaction 1
            ElseIf Accion = eAction.Duelo Then
                WriteDuelo 2
            End If
            Unload Me
            
        Case 1
            If Accion = eAction.Duelo Then
                WriteDuelo 3
            End If
            Unload Me
            
    End Select
End Sub

Private Sub img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            If img(0).Tag <> "1" Then
                img(0).Picture = LoadInterface("preg-aceptar-down")
                img(0).Tag = "1"
            End If
            
        Case 1
            If img(1).Tag <> 1 Then
                img(1).Picture = LoadInterface("preg-cancelar-down")
                img(1).Tag = "1"
            End If
    End Select
End Sub

Private Sub img_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then Exit Sub
    
    Select Case Index
        Case 0
            If img(0).Tag <> "2" Then
                img(0).Picture = LoadInterface("preg-aceptar-over")
                img(0).Tag = "2"
            End If
            
        Case 1
            If img(1).Tag <> "2" Then
                img(1).Picture = LoadInterface("preg-cancelar-over")
                img(1).Tag = "2"
            End If
    End Select
End Sub

