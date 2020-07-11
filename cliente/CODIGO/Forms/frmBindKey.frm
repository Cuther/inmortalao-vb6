VERSION 5.00
Begin VB.Form frmBindKey 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asignar acción"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBindKey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3270
      Width           =   1455
   End
   Begin VB.CommandButton cmdAccept 
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
      Left            =   1800
      TabIndex        =   7
      Top             =   3270
      Width           =   1455
   End
   Begin VB.TextBox txtComandoEnvio 
      Enabled         =   0   'False
      Height          =   285
      Left            =   390
      TabIndex        =   4
      Top             =   2070
      Width           =   2655
   End
   Begin VB.OptionButton optAccion 
      Caption         =   "Equipar ítem elegido"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   2970
      Width           =   3135
   End
   Begin VB.OptionButton optAccion 
      Caption         =   "Usar ítem elegido"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   2700
      Width           =   3135
   End
   Begin VB.OptionButton optAccion 
      Caption         =   "Lanzar hechizo elegido"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2430
      Width           =   3135
   End
   Begin VB.OptionButton optAccion 
      Caption         =   "Enviar Comando"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label lblTecla 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tecla:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   2775
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3240
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3240
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "$7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "/"
      Height          =   255
      Left            =   270
      TabIndex        =   5
      Top             =   2070
      Width           =   105
   End
End
Attribute VB_Name = "frmBindKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccept_Click()

On Error Resume Next

Dim i As Integer

For i = optAccion.LBound To optAccion.UBound
    If optAccion(i).value = True Then
        MacroKeys(BotonElegido).TipoAccion = i + 1
        Exit For
    End If
Next i

Select Case MacroKeys(BotonElegido).TipoAccion
    
    Case 1
        If LenB(txtComandoEnvio.Text) = 0 Then
            msgbox_ok "Debes escribir un comando válido a enviar."
            Exit Sub
        End If
        
        MacroKeys(BotonElegido).SendString = UCase$(txtComandoEnvio.Text)
        MacroKeys(BotonElegido).hlist = 0
        MacroKeys(BotonElegido).invslot = 0
    
    Case 2
        MacroKeys(BotonElegido).hlist = frmMain.hlst.ListIndex + 1
        MacroKeys(BotonElegido).SendString = vbNullString
        MacroKeys(BotonElegido).invslot = 0
    
    Case 3
        MacroKeys(BotonElegido).hlist = 0
        MacroKeys(BotonElegido).SendString = vbNullString
        MacroKeys(BotonElegido).invslot = Inventario.SelectedItem
    
    Case 4
        MacroKeys(BotonElegido).hlist = 0
        MacroKeys(BotonElegido).SendString = vbNullString
        MacroKeys(BotonElegido).invslot = Inventario.SelectedItem

End Select

Call DibujarMenuMacros(BotonElegido)

Dim LC As Long

If FileExist(App.Path & "\Init\" & UserName & ".dat", vbNormal) Then _
    Kill App.Path & "\Init\" & UserName & ".dat"
  
Open App.Path & "\Init\" & UserName & ".dat" For Append As #1
                Print #1, "[" & UserName & "]"
            For LC = 1 To 11
                Print #1, "Accion" & LC & "=" & MacroKeys(LC).TipoAccion
                Print #1, "InvSlot" & LC & "=" & MacroKeys(LC).invslot
                Print #1, "SndString" & LC & "=" & MacroKeys(LC).SendString
                Print #1, "Hlist" & LC & "=" & MacroKeys(LC).hlist
            
    Next LC
    Print #1, "" 'Separacion entre macro y macro
Close #1

Unload Me
End Sub

Private Sub cmdCancel_Click()

MacroKeys(BotonElegido).TipoAccion = 0
Unload Me

End Sub



Private Sub optAccion_Click(Index As Integer)

If Index = 0 Then
    txtComandoEnvio.Enabled = True
Else
    txtComandoEnvio.Enabled = False
End If

End Sub

Private Sub Form_Load()

lblTecla.Caption = "Tecla: F" & BotonElegido
Label2.Caption = "Adevertencia: el uso incorrecto de este sitema puede terminar en severas penas, entre ellas la prohibicion de ingreso al juego. Recomendamos leer el reglamento antes de utilizarlos."

If MacroKeys(BotonElegido).TipoAccion <> 0 Then

    Select Case MacroKeys(BotonElegido).TipoAccion
        Case 1 'Envia comando
            optAccion(0).value = True
            txtComandoEnvio.Text = MacroKeys(BotonElegido).SendString
            txtComandoEnvio.Enabled = True
        Case 2 'Lanza hechizo
            optAccion(1).value = True
        Case 3 'Equipa
            optAccion(2).value = True
        Case 4 'Usa
            optAccion(3).value = True
    End Select
    
End If

End Sub
Public Sub DibujarMenuMacros(Optional ActualizarCual As Integer = 0, Optional AlphaEffect As Byte = 0)

Dim i As Integer

If ActualizarCual <= 0 Then
    For i = 1 To 11
        frmMain.picMacro(i - 1).Cls
        Select Case MacroKeys(i).TipoAccion
            Case 1 'Envia comando
                Call TileEngine.Draw_Grh_Hdc(frmMain.picMacro(i - 1).hdc, 17506, 0, 0)
                frmMain.picMacro(i - 1).ToolTipText = "Enviar comando: " & MacroKeys(i).SendString
            Case 2 'Lanza hechizo
                Call TileEngine.Draw_Grh_Hdc(frmMain.picMacro(i - 1).hdc, 609, 0, 0)
                frmMain.picMacro(i - 1).ToolTipText = "Lanzar hechizo elegido: " & frmMain.hlst.List(MacroKeys(i).hlist - 1)
            Case 3 'Equipa
                Call TileEngine.Draw_Grh_Hdc(frmMain.picMacro(i - 1).hdc, UserInventory(MacroKeys(i).invslot).grhindex, 0, 0)
                frmMain.picMacro(i - 1).ToolTipText = "Equipar objeto: " & UserInventory(MacroKeys(i).invslot).name
            Case 4 'Usa
                Call TileEngine.Draw_Grh_Hdc(frmMain.picMacro(i - 1).hdc, UserInventory(MacroKeys(i).invslot).grhindex, 0, 0)
                frmMain.picMacro(i - 1).ToolTipText = "Usar objeto: " & UserInventory(MacroKeys(i).invslot).name
        End Select
        frmMain.picMacro(i - 1).refresh
    Next i
Else
    frmMain.picMacro(ActualizarCual - 1).Cls
    
    Select Case MacroKeys(ActualizarCual).TipoAccion
        Case 1 'Envia comando
            Call TileEngine.Draw_Grh_Hdc(frmMain.picMacro(ActualizarCual - 1).hdc, 17506, 0, 0)
            frmMain.picMacro(ActualizarCual - 1).ToolTipText = "Enviar comando: " & MacroKeys(ActualizarCual).SendString
        Case 2 'Lanza hechizo
            Call TileEngine.Draw_Grh_Hdc(frmMain.picMacro(ActualizarCual - 1).hdc, 609, 0, 0)
            frmMain.picMacro(ActualizarCual - 1).ToolTipText = "Lanzar hechizo elegido: " & frmMain.hlst.List(MacroKeys(ActualizarCual).hlist - 1)
        Case 3 'Equipa
            Call TileEngine.Draw_Grh_Hdc(frmMain.picMacro(ActualizarCual - 1).hdc, UserInventory(MacroKeys(ActualizarCual).invslot).grhindex, 0, 0)
            frmMain.picMacro(ActualizarCual - 1).ToolTipText = "Equipar objeto: " & UserInventory(MacroKeys(ActualizarCual).invslot).name
        Case 4 'Usa
            Call TileEngine.Draw_Grh_Hdc(frmMain.picMacro(ActualizarCual - 1).hdc, Inventario.grhindex(MacroKeys(ActualizarCual).invslot), 0, 0)
            frmMain.picMacro(ActualizarCual - 1).ToolTipText = "Usar objeto: " & Inventario.ItemName(MacroKeys(ActualizarCual).invslot)
    End Select

    frmMain.picMacro(ActualizarCual - 1).refresh

End If

End Sub

Public Sub Bind_Accion(ByVal FNUM As Integer)

If MacroKeys(FNUM).TipoAccion = 0 Then
    BotonElegido = FNUM
    Form_Load
    Me.Visible = True
    Exit Sub
End If

Select Case MacroKeys(FNUM).TipoAccion

    Case 1 'Envia comando
        Call ParseUserCommand("/" & MacroKeys(FNUM).SendString)
    Case 2 'Lanza hechizo
        If frmMain.hlst.List(MacroKeys(FNUM).hlist - 1) <> "(None)" And MainTimer.Check(TimersIndex.Work, False) Then
            If UserEstado = 1 Then Exit Sub
            Call WriteCastSpell(MacroKeys(FNUM).hlist)
            Call WriteWork(eSkill.Magia)
        End If
    Case 3 'Equipa
        If UserEstado = 1 Then Exit Sub
        Call WriteEquipItem(MacroKeys(FNUM).invslot)
    Case 4 'Usa
        If MainTimer.Check(TimersIndex.UseItemWithU) Then Call WriteUseItem(MacroKeys(FNUM).invslot)

End Select

End Sub
