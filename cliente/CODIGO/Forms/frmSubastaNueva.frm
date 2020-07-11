VERSION 5.00
Begin VB.Form frmSubastaNueva 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nueva subasta"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbDuracion 
      Height          =   315
      ItemData        =   "frmSubastaNueva.frx":0000
      Left            =   2580
      List            =   "frmSubastaNueva.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2130
      Width           =   2770
   End
   Begin VB.TextBox txtFinal 
      Height          =   315
      Left            =   2580
      MaxLength       =   9
      TabIndex        =   8
      Text            =   "0"
      Top             =   1500
      Width           =   2770
   End
   Begin VB.TextBox txtInicial 
      Height          =   315
      Left            =   2580
      MaxLength       =   9
      TabIndex        =   6
      Text            =   "1"
      Top             =   870
      Width           =   2770
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   2580
      TabIndex        =   5
      Top             =   3020
      Width           =   1335
   End
   Begin VB.CommandButton cmdCrear 
      Caption         =   "Aceptar"
      Height          =   435
      Left            =   4020
      TabIndex        =   4
      Top             =   3020
      Width           =   1335
   End
   Begin VB.ListBox lstItems 
      Height          =   2790
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   2415
   End
   Begin VB.TextBox txtCant 
      Height          =   315
      Left            =   660
      MaxLength       =   9
      TabIndex        =   1
      Text            =   "1"
      Top             =   270
      Width           =   1815
   End
   Begin VB.PictureBox picItem 
      BackColor       =   &H00000000&
      Height          =   540
      Left            =   60
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   60
      Width           =   540
   End
   Begin VB.Label lblCosto 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2610
      TabIndex        =   13
      Top             =   2580
      Width           =   2745
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Duración Máxima"
      Height          =   225
      Left            =   2580
      TabIndex        =   11
      Top             =   1920
      Width           =   2835
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "El costo de la subasta depende de la duración y el precio inicial o final."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   2580
      TabIndex        =   10
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio Final (opcional)"
      Height          =   225
      Left            =   2580
      TabIndex        =   9
      Top             =   1290
      Width           =   2835
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio Inicial"
      Height          =   225
      Left            =   2580
      TabIndex        =   7
      Top             =   660
      Width           =   2835
   End
   Begin VB.Label lblCantidad 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      Height          =   225
      Left            =   660
      TabIndex        =   3
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmSubastaNueva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private OroMod As Long
Private DurMod As Long
Private VedMod As Long
Private Sub cmbDuracion_Click()
    Select Case cmbDuracion.ListIndex
        Case 0: DurMod = 30
        Case 1: DurMod = 60
        Case 2: DurMod = 120
        Case 3: DurMod = 240
        Case 4: DurMod = 480
    End Select
    If DurMod > OroMod Then
        VedMod = DurMod
        lblCosto.Caption = "Depósito: " & DurMod & " monedas"
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdCrear_Click()
    Dim slot As Byte, cant As Integer, fnlOfert As Long
    
    slot = lstItems.ListIndex + 1
    If slot <= 0 Or slot > MAX_INVENTORY_SLOTS Then Exit Sub
    cant = val(txtCant.Text)
    If cant > 10000 Or cant <= 0 Then Exit Sub
    fnlOfert = val(txtFinal.Text)
    
    Call WriteSSubastar(slot, val(txtCant.Text), _
                        DurMod / 10, val(txtInicial.Text), fnlOfert)
                        
    Unload Me
      
   
      
    Call WriteSRequest
    
End Sub

Private Sub Form_Load()
Dim i As Byte
For i = 1 To MAX_INVENTORY_SLOTS
    If UserInventory(i).name <> "" Then
        lstItems.AddItem UserInventory(i).name
    Else
        lstItems.AddItem "Nada"
    End If
Next i

cmbDuracion.ListIndex = 0
End Sub


Private Sub lstItems_Click()
    picItem.Cls
    Call TileEngine.Draw_Grh_Hdc(picItem.hdc, Inventario.grhindex(lstItems.ListIndex + 1), 3, 5)
End Sub

Private Sub txtFinal_Change()
    If val(txtFinal.Text) / 2 / 10 > DurMod Then
        OroMod = val(txtFinal.Text) / 2 / 10
        lblCosto.Caption = "Depósito: " & OroMod & " monedas"
        VedMod = OroMod
    Else
        lblCosto.Caption = "Depósito: " & DurMod & " monedas"
        OroMod = 0
    End If
End Sub

Private Sub txtInicial_Change()
    If Len(txtInicial.Text) = 0 Then txtInicial.Text = 1
End Sub
