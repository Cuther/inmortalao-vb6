VERSION 5.00
Begin VB.Form frmCustomKeys 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuraración de controles"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7995
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   397
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   533
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      TabIndex        =   46
      Text            =   "10"
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   420
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1140
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1860
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2580
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3300
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   5
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   4020
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   6
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4740
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   7
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   420
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   8
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1140
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   9
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1860
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   10
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2580
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   11
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3300
      Width           =   2415
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Guardar"
      Height          =   315
      Index           =   0
      Left            =   5430
      TabIndex        =   12
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Cargar defaults"
      Height          =   315
      Index           =   1
      Left            =   5400
      TabIndex        =   11
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Salir"
      Height          =   315
      Index           =   2
      Left            =   5430
      TabIndex        =   10
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   13
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   420
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   14
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1140
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   15
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1860
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   16
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2580
      Width           =   2415
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reconfigurar macros"
      Height          =   315
      Left            =   5400
      TabIndex        =   5
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Enabled         =   0   'False
      Height          =   315
      Index           =   17
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Impr. Pant."
      Top             =   5490
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Enabled         =   0   'False
      Height          =   315
      Index           =   18
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "*"
      Top             =   4020
      Width           =   2415
   End
   Begin VB.TextBox txConfig 
      Enabled         =   0   'False
      Height          =   315
      Index           =   19
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "ALT1"
      Top             =   4770
      Width           =   2415
   End
   Begin VB.HScrollBar scrSens 
      Height          =   345
      LargeChange     =   15
      Left            =   2760
      Max             =   20
      Min             =   1
      TabIndex        =   1
      Top             =   5550
      Value           =   1
      Width           =   2025
   End
   Begin VB.TextBox txConfig 
      Enabled         =   0   'False
      Height          =   315
      Index           =   12
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Bloq. Num."
      Top             =   3300
      Width           =   2415
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atacar"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   45
      Top             =   120
      Width           =   465
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tomar objeto"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   44
      Top             =   840
      Width           =   930
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tirar objeto"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   43
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usar objeto"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   42
      Top             =   2280
      Width           =   810
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Equipar objeto"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   41
      Top             =   3000
      Width           =   1020
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Activar / Desactivar Seguro"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   40
      Top             =   3720
      Width           =   1980
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mostrar / Ocultar Nicknames"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   39
      Top             =   4440
      Width           =   2040
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Domar"
      Height          =   195
      Index           =   7
      Left            =   2760
      TabIndex        =   38
      Top             =   120
      Width           =   465
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Robar"
      Height          =   195
      Index           =   8
      Left            =   2760
      TabIndex        =   37
      Top             =   840
      Width           =   435
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizar posicion"
      Height          =   195
      Index           =   9
      Left            =   2760
      TabIndex        =   36
      Top             =   1560
      Width           =   1320
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ocultarse"
      Height          =   195
      Index           =   10
      Left            =   2760
      TabIndex        =   35
      Top             =   2280
      Width           =   675
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modo Combate"
      Height          =   195
      Index           =   11
      Left            =   2760
      TabIndex        =   34
      Top             =   3000
      Width           =   1080
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moverse hacia arriba"
      Height          =   195
      Index           =   13
      Left            =   5400
      TabIndex        =   33
      Top             =   120
      Width           =   1485
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moverse hacia abajo"
      Height          =   195
      Index           =   14
      Left            =   5400
      TabIndex        =   32
      Top             =   840
      Width           =   1485
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moverse hacia la izquierda"
      DataField       =   "Moverse hacia la izquierda"
      Height          =   195
      Index           =   15
      Left            =   5400
      TabIndex        =   31
      Top             =   1560
      Width           =   1890
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moverse hacia la derecha"
      Height          =   195
      Index           =   16
      Left            =   5400
      TabIndex        =   30
      Top             =   2280
      Width           =   1845
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tomar screenshot"
      Height          =   195
      Index           =   17
      Left            =   120
      TabIndex        =   29
      Top             =   5190
      Width           =   1275
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ver FPS"
      Height          =   195
      Index           =   18
      Left            =   2760
      TabIndex        =   28
      Top             =   3720
      Width           =   585
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comenzar/parar de grabar un video"
      Height          =   195
      Index           =   19
      Left            =   2760
      TabIndex        =   27
      Top             =   4470
      Width           =   2520
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sensibilidad del mouse"
      Height          =   195
      Index           =   20
      Left            =   2760
      TabIndex        =   26
      Top             =   5250
      Width           =   1605
   End
   Begin VB.Label lbNames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bloque de movimiento"
      Height          =   195
      Index           =   21
      Left            =   5400
      TabIndex        =   25
      Top             =   3000
      Width           =   1560
   End
   Begin VB.Menu mnuMensaje 
      Caption         =   "Mensajes"
      Visible         =   0   'False
      Begin VB.Menu mnuNormal 
         Caption         =   "Normal"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuGlobal 
         Caption         =   "Global"
      End
      Begin VB.Menu mnuPrivado 
         Caption         =   "Privado"
      End
      Begin VB.Menu mnuClan 
         Caption         =   "Clan"
      End
      Begin VB.Menu mnuGritar 
         Caption         =   "Gritar"
      End
      Begin VB.Menu mnuGrupo 
         Caption         =   "Grupo"
      End
      Begin VB.Menu mnuFaccion 
         Caption         =   "Faccion"
      End
   End
End
Attribute VB_Name = "frmCustomKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command2_Click()
    Dim i As Long
    
    For i = 0 To CustomKeys.count
        If LenB(txConfig(i).Text) = 0 Then
            msgbox_ok ("Hay una o mas teclas no validas, por favor verifique.")
            Exit Sub
        End If
    Next i

    Unload Me
End Sub

Private Sub cmdAccion_Click(Index As Integer)
    Dim i As Long
    Select Case Index
        Case 0
            For i = 0 To CustomKeys.count
                If LenB(txConfig(i).Text) = 0 Then
                    msgbox_ok ("Hay una o mas teclas no validas, por favor verifique.")
                    Exit Sub
                End If
            Next i
            CustomKeys.SaveCustomKeys
            
        Case 1
            Call CustomKeys.LoadDefaults
            
            For i = 0 To CustomKeys.count
                If Not i > 16 And Not i = 12 Then txConfig(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
            Next i
            
        Case 2
            Unload Me
            
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    For i = 0 To CustomKeys.count
        If Not i > 16 And Not i = 12 Then txConfig(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
    Next i
    
    scrSens.value = MouseS
    Me.Text1.Text = MouseS
End Sub

Private Sub scrSens_Change()
    MouseS = scrSens.value
    Call General_Set_Mouse_Speed(MouseS)
    Me.Text1.Text = scrSens.value
End Sub

Private Sub txConfig_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If LenB(CustomKeys.ReadableName(KeyCode)) = 0 Then Exit Sub
    'If key is not valid, we exit
    
    txConfig(Index).Text = CustomKeys.ReadableName(KeyCode)
    txConfig(Index).SelStart = Len(txConfig(Index).Text)
    
    For i = 0 To CustomKeys.count
        If i <> Index Then
            If CustomKeys.BindedKey(i) = KeyCode Then
                txConfig(Index).Text = "" 'If the key is already assigned, simply reject it
                Call Beep 'Alert the user
                KeyCode = 0
                Exit Sub
            End If
        End If
    Next i
    
    CustomKeys.BindedKey(Index) = KeyCode
End Sub

Private Sub txConfig_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txConfig_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call txConfig_KeyDown(Index, KeyCode, Shift)
End Sub

Public Sub PopupMenuMensaje()

Select Case CurrentUser.SendingType
    Case 1
        frmCustomKeys.mnuNormal.Checked = True
        frmCustomKeys.mnuGritar.Checked = False
        frmCustomKeys.mnuPrivado.Checked = False
        frmCustomKeys.mnuClan.Checked = False
        frmCustomKeys.mnuGrupo.Checked = False
        frmCustomKeys.mnuGlobal.Checked = False
        frmCustomKeys.mnuFaccion.Checked = False
    Case 2
        frmCustomKeys.mnuNormal.Checked = False
        frmCustomKeys.mnuGritar.Checked = True
        frmCustomKeys.mnuPrivado.Checked = False
        frmCustomKeys.mnuClan.Checked = False
        frmCustomKeys.mnuGrupo.Checked = False
        frmCustomKeys.mnuGlobal.Checked = False
        frmCustomKeys.mnuFaccion.Checked = False
    Case 3
        frmCustomKeys.mnuNormal.Checked = False
        frmCustomKeys.mnuGritar.Checked = False
        frmCustomKeys.mnuPrivado.Checked = True
        frmCustomKeys.mnuClan.Checked = False
        frmCustomKeys.mnuGrupo.Checked = False
        frmCustomKeys.mnuGlobal.Checked = False
        frmCustomKeys.mnuFaccion.Checked = False
    Case 4
        frmCustomKeys.mnuNormal.Checked = False
        frmCustomKeys.mnuGritar.Checked = False
        frmCustomKeys.mnuPrivado.Checked = False
        frmCustomKeys.mnuClan.Checked = True
        frmCustomKeys.mnuGrupo.Checked = False
        frmCustomKeys.mnuGlobal.Checked = False
        frmCustomKeys.mnuFaccion.Checked = False
    Case 5
        frmCustomKeys.mnuNormal.Checked = False
        frmCustomKeys.mnuGritar.Checked = False
        frmCustomKeys.mnuPrivado.Checked = False
        frmCustomKeys.mnuClan.Checked = False
        frmCustomKeys.mnuGrupo.Checked = True
        frmCustomKeys.mnuGlobal.Checked = False
        frmCustomKeys.mnuFaccion.Checked = False
    Case 6
        frmCustomKeys.mnuNormal.Checked = False
        frmCustomKeys.mnuGritar.Checked = False
        frmCustomKeys.mnuPrivado.Checked = False
        frmCustomKeys.mnuClan.Checked = False
        frmCustomKeys.mnuGrupo.Checked = False
        frmCustomKeys.mnuGlobal.Checked = True
        frmCustomKeys.mnuFaccion.Checked = False
    Case 7
        frmCustomKeys.mnuNormal.Checked = False
        frmCustomKeys.mnuGritar.Checked = False
        frmCustomKeys.mnuPrivado.Checked = False
        frmCustomKeys.mnuClan.Checked = False
        frmCustomKeys.mnuGrupo.Checked = False
        frmCustomKeys.mnuGlobal.Checked = False
        frmCustomKeys.mnuFaccion.Checked = True
End Select

PopupMenu mnuMensaje

End Sub

Private Sub mnuNormal_Click()

CurrentUser.SendingType = 1
If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

Private Sub mnuGritar_click()

CurrentUser.SendingType = 2
If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

Private Sub mnuPrivado_click()

CurrentUser.sndPrivateTo = InputBox("Escriba el usuario con el que desea iniciar una conversación privada", vbNullString)

If CurrentUser.sndPrivateTo <> vbNullString Then
    CurrentUser.SendingType = 3
    
    'Add Marius Agregamos el + por que si tiene espacios el nick no manda el mensaje o se lo manda a cualquiera.
    CurrentUser.sndPrivateTo = Replace(CurrentUser.sndPrivateTo, " ", "+")
    '\Add
    
    If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
Else
    msgbox_ok ("¡Ingrese un nombre válido!")
End If

End Sub

Private Sub mnuClan_click()

CurrentUser.SendingType = 4
If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub
Private Sub mnuGrupo_click()

CurrentUser.SendingType = 5
If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

Private Sub mnuGlobal_Click()

CurrentUser.SendingType = 6
If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

Private Sub mnuFaccion_Click()

CurrentUser.SendingType = 7
If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub
