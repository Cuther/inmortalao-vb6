VERSION 5.00
Begin VB.Form frmCorreo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Correo"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Enviar Mensaje"
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   7215
      Begin VB.TextBox txCantidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5520
         TabIndex        =   15
         Text            =   "1"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Enviar"
         Height          =   495
         Left            =   5520
         TabIndex        =   14
         Top             =   2560
         Width           =   1575
      End
      Begin VB.CommandButton cmdClean 
         Caption         =   "Limpiar"
         Height          =   495
         Left            =   5520
         TabIndex        =   13
         Top             =   1960
         Width           =   1575
      End
      Begin VB.CheckBox adjItem 
         Caption         =   "Adjuntar item"
         Height          =   375
         Left            =   2880
         TabIndex        =   10
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txSndMsg 
         Height          =   1875
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txTo 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2655
      End
      Begin VB.ListBox lstInv 
         Enabled         =   0   'False
         Height          =   2595
         ItemData        =   "frmcorreo.frx":0000
         Left            =   2880
         List            =   "frmcorreo.frx":003A
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
      Begin VB.PictureBox picInvT 
         BackColor       =   &H00000000&
         Height          =   540
         Left            =   5520
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
         Top             =   480
         Width           =   540
      End
      Begin VB.Label lbCount 
         AutoSize        =   -1  'True
         Caption         =   "cantidad"
         Height          =   195
         Left            =   6120
         TabIndex        =   16
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Costo:"
         Height          =   195
         Left            =   5520
         TabIndex        =   12
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   5520
         TabIndex        =   11
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Para"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   330
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mensajes"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.PictureBox picItem 
         BackColor       =   &H00000000&
         Height          =   540
         Left            =   2640
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   19
         Top             =   1920
         Width           =   540
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Guardar item"
         Height          =   495
         Left            =   5640
         TabIndex        =   18
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "Borrar"
         Height          =   495
         Left            =   5640
         TabIndex        =   17
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txMensaje 
         Height          =   1575
         Left            =   2640
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   4455
      End
      Begin VB.ListBox lstMsg 
         Height          =   2790
         ItemData        =   "frmcorreo.frx":00AA
         Left            =   120
         List            =   "frmcorreo.frx":00AC
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lbCant 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   3240
         TabIndex        =   20
         Top             =   1920
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmCorreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SelectMsg As Byte
Public Sub ActualizarCorreo()
 
  SelectMsg = lstMsg.ListIndex + 1
    
    picItem.Cls

    
    If lstMsg.List(lstMsg.ListIndex) <> "(Nada)" And lstMsg.List(lstMsg.ListIndex) <> "" Then
        
        txMensaje.Text = Correos(SelectMsg).mensaje
       
        If Correos(SelectMsg).item <> 0 Then
            picItem.Visible = True
            cmdSave.Enabled = True
            lbCant.Caption = "Cantidad : " & Correos(SelectMsg).Cantidad
            
            TileEngine.Draw_Grh_Hdc picItem.hdc, objs(Correos(SelectMsg).item).Grh, 0, 0
            
        Else
            picItem.Visible = False
            lbCant.Caption = ""
            cmdSave.Enabled = False
        End If
    Else
        lbCant.Caption = ""
        txMensaje.Text = ""
        picItem.Visible = False
        cmdSave.Enabled = False
    End If
End Sub

Private Sub lstInv_Click()
    picInvT.Cls
    If UserInventory(lstInv.ListIndex + 1).Amount <> 0 Then
        TileEngine.Draw_Grh_Hdc picInvT.hdc, UserInventory(lstInv.ListIndex + 1).grhindex, 0, 0
        lbCount.Caption = UserInventory(lstInv.ListIndex + 1).Amount
    Else
        lbCount.Caption = "cantidad"
    End If
End Sub

Private Sub adjItem_Click()
    If adjItem.value = vbChecked Then
        lstInv.Enabled = True
        txCantidad.Enabled = True
    Else
        txCantidad.Enabled = True
        lstInv.Enabled = False
    End If
End Sub

Private Sub cmdClean_Click()
    txTo.Text = ""
    txSndMsg.Text = ""
    txCantidad.Text = ""
End Sub

Private Sub cmdDel_Click()
    WriteBorrarMensaje SelectMsg
End Sub

Private Sub cmdSave_Click()
    WriteExtractItem SelectMsg
End Sub

Private Sub cmdSend_Click()
    If lstInv.Enabled Then
        If lstInv.ListIndex = -1 Then
            msgbox_ok "Seleciona un item"
        Else
            WriteEnviarMensaje txSndMsg.Text, lstInv.ListIndex + 1, txTo.Text, val(txCantidad.Text)
        End If
    Else
        WriteEnviarMensaje txSndMsg.Text, 0, txTo.Text, 0
    End If
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    For i = 1 To 20
        Correos(i).De = ""
        Correos(i).mensaje = ""
        Correos(i).item = 0
        Correos(i).Cantidad = 0
    Next
End Sub

Public Sub actualizar_inventario()
    Dim i As Long
    lstInv.Clear
    
    For i = 1 To MAX_INVENTORY_SLOTS
        If UserInventory(i).name <> "" Then
            lstInv.AddItem UserInventory(i).name
        Else
            lstInv.AddItem "Nada"
        End If
    Next i
End Sub


Private Sub lstMsg_Click()

    Call ActualizarCorreo

End Sub

