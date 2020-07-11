VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   1335
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   2220
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   89
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   148
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCant 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   210
      Left            =   300
      TabIndex        =   0
      Top             =   540
      Width           =   1470
   End
   Begin VB.Image imgAceptar 
      Height          =   405
      Left            =   150
      Tag             =   "0"
      Top             =   840
      Width           =   975
   End
   Begin VB.Image imgTodo 
      Height          =   405
      Left            =   1125
      Tag             =   "0"
      Top             =   840
      Width           =   945
   End
   Begin VB.Image imgCerrar 
      Height          =   330
      Left            =   1920
      Tag             =   "0"
      Top             =   0
      Width           =   315
   End
   Begin VB.Image imgMas 
      Height          =   135
      Left            =   1800
      Top             =   510
      Width           =   195
   End
   Begin VB.Image imgMenos 
      Height          =   135
      Left            =   1800
      Top             =   630
      Width           =   195
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Picture = LoadInterface("Cantidad")
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Auto_Drag Me.hWnd
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If imgAceptar.Tag = "1" Then
    imgAceptar.Picture = Nothing
    imgAceptar.Tag = "0"
End If

If imgTodo.Tag = "1" Then
    imgTodo.Picture = Nothing
    imgTodo.Tag = "0"
End If

If imgCerrar.Tag = "1" Then
    imgCerrar.Picture = Nothing
    imgCerrar.Tag = "0"
End If
End Sub

Private Sub imgAceptar_Click()
    If LenB(frmCantidad.txtCant.Text) > 0 Then
        If Not IsNumeric(frmCantidad.txtCant.Text) Then Exit Sub  'Should never happen
        Call WriteDrop(Inventario.SelectedItem, frmCantidad.txtCant.Text)
        frmCantidad.txtCant.Text = ""
    End If
    
    Unload Me
End Sub

Private Sub imgCerrar_Click()
Unload Me
End Sub
Private Sub imgCerrar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCerrar.Picture = LoadInterface("cerrarcantdown")
End Sub

Private Sub imgCerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If imgAceptar.Tag = "1" Then
    imgAceptar.Picture = Nothing
    imgAceptar.Tag = "0"
End If

If imgTodo.Tag = "1" Then
    imgTodo.Picture = Nothing
    imgTodo.Tag = "0"
End If

If imgCerrar.Tag = "0" Then
    imgCerrar.Picture = LoadInterface("cerrarcantover")
    imgCerrar.Tag = "1"
End If

End Sub

Private Sub imgMas_Click()
txtCant.Text = Val(txtCant.Text) + 1
End Sub
Private Sub imgMenos_Click()

If Val(txtCant.Text) > 0 Then _
    txtCant.Text = Val(txtCant.Text) - 1

End Sub
Private Sub imgTodo_Click()
    If Inventario.SelectedItem = 0 Then Exit Sub
    
    If Inventario.SelectedItem <> FLAGORO Then
        Call WriteDrop(Inventario.SelectedItem, Inventario.Amount(Inventario.SelectedItem))
        Unload Me
    Else
        If UserGLD > 100000 Then
            Dim i As Long
            For i = 1 To 10
                Call WriteDrop(Inventario.SelectedItem, 10000)
            Next i
            Unload Me
        Else
            Call WriteDrop(Inventario.SelectedItem, UserGLD)
            Unload Me
        End If
    End If

    frmCantidad.txtCant.Text = ""
End Sub

Private Sub imgTodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgTodo.Picture = LoadInterface("dejartododown")
End Sub

Private Sub imgTodo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If imgTodo.Tag = "0" Then
    imgTodo.Picture = LoadInterface("dejartodoover")
    imgTodo.Tag = "1"
End If

If imgAceptar.Tag = "1" Then
    imgAceptar.Picture = Nothing
    imgAceptar.Tag = "0"
End If

End Sub
Private Sub imgAceptar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgAceptar.Picture = LoadInterface("dejardown")
End Sub

Private Sub imgAceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If imgTodo.Tag = "1" Then
    imgTodo.Picture = Nothing
    imgTodo.Tag = "0"
End If

If imgAceptar.Tag = "0" Then
    imgAceptar.Picture = LoadInterface("dejarover")
    imgAceptar.Tag = "1"
End If

End Sub
Private Sub txtCant_Change()
On Error GoTo ErrHandler
    If Val(txtCant.Text) < 0 Then
        txtCant.Text = "1"
    End If
    
    If Val(txtCant.Text) > MAX_INVENTORY_OBJS Then
        txtCant.Text = "10000"
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    txtCant.Text = "1"
End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub
