VERSION 5.00
Begin VB.Form frmGoliath 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Operación bancaria"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4590
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGoliath.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   3090
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Height          =   345
      Left            =   3000
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtDatos 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   4335
   End
   Begin VB.ListBox lstBanco 
      Height          =   840
      ItemData        =   "frmGoliath.frx":000C
      Left            =   90
      List            =   "frmGoliath.frx":000E
      TabIndex        =   1
      Top             =   1230
      Width           =   4395
   End
   Begin VB.Label lblDatos 
      Caption         =   "¿Cuánto deseas depositar?"
      Height          =   405
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGoliath.frx":0010
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4395
   End
End
Attribute VB_Name = "frmGoliath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Oro As Long
Private Items As Long
Private CantTransferencia As Long
Private NombreTransferencia As String
Private EtapaTransferencia As Byte

Public Sub ParseBancoInfo(ByVal GLD As Long, ByVal item As Long)

On Error GoTo Error_Handler

Oro = GLD
Items = item

If val(Oro) > 0 And val(Items) > 0 Then
    lblInfo.Caption = "Bienvenido a la cadena de finanzas Goliath. Tienes " _
        & Items & " objetos en tu bóveda y " & Oro & " monedas de oro en tu cuenta... ¿Cómo te puedo ayudar?"
ElseIf val(Oro) <= 0 And val(Items) > 0 Then
    lblInfo.Caption = "Bienvenido a la cadena de finanzas Goliath. Tienes " _
        & Items & " objetos en tu bóveda y aún no has depositado oro... ¿Cómo te puedo ayudar?"
ElseIf val(Oro) > 0 And val(Items) <= 0 Then
    lblInfo.Caption = "Bienvenido a la cadena de finanzas Goliath. Tu bóveda está vacía y posees " & Oro & " monedas de oro en tu cuenta... ¿Cómo te puedo ayudar?"
ElseIf val(Oro) <= 0 And val(Items) <= 0 Then
    lblInfo.Caption = "Bienvenido a la cadena de finanzas Goliath. Tu bóveda y tu cuenta están vacías... ¿Cómo te puedo ayudar?"
End If

Me.Show vbModeless, frmMain

Exit Sub

Error_Handler:
    'Error vite'

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()

Select Case lstBanco.ListIndex
    Case 0, -1 'Depositar
    
        'Negativos y ceros
        If (val(txtDatos.Text) <= 0 And (UCase$(txtDatos.Text) <> "TODO")) Then lblDatos.Caption = "Cantidad inválida."
    
        If val(txtDatos.Text) <= UserGLD Or UCase$(txtDatos.Text) = "TODO" Then
            Call WriteBankDepositGold(IIf(val(txtDatos.Text) > 0, val(txtDatos.Text), UserGLD))
            Unload Me
        Else
            lblDatos.Caption = "No tienes esa cantidad."
        End If
    Case 1 'Retirar
    
        'Negativos y ceros
        If (val(txtDatos.Text) <= 0 And (UCase$(txtDatos.Text) <> "TODO")) Then lblDatos.Caption = "Cantidad inválida."
    
        If val(txtDatos.Text) <= Oro Or UCase$(txtDatos.Text) = "TODO" Then
            Call WriteBankExtractGold(IIf(val(txtDatos.Text) > 0, val(txtDatos.Text), Oro))
            Unload Me
        Else
            lblDatos.Caption = "No tienes esa cantidad."
        End If
    Case 2 'Bóveda
        Unload Me
    Case 3 'Transferir - Destino - Cantidad
        If EtapaTransferencia = 0 Then
        
            'Negativos y ceros
            If val(txtDatos.Text) <= 0 Then
                lblDatos.Caption = "No tienes esa cantidad."
                txtDatos.Text = vbNullString
                Exit Sub
            End If
            
            If val(txtDatos.Text) <= Oro Then
                CantTransferencia = val(txtDatos.Text)
                lblDatos.Caption = "¿A quién le deseas enviar " & CantTransferencia & " monedas de oro?"
                EtapaTransferencia = 1
                txtDatos.Text = vbNullString
            Else
                lblDatos.Caption = "No tienes esa cantidad."
                txtDatos.Text = vbNullString
            End If
        ElseIf EtapaTransferencia = 1 Then
            If LenB(txtDatos.Text) > 0 Then
                NombreTransferencia = txtDatos.Text
                lblDatos.Caption = "Se transferirán " & CantTransferencia & " monedas de oro a " & NombreTransferencia & ". Si es correcto, presione aceptar."
                EtapaTransferencia = 2
            Else
                lblDatos.Caption = "¡Nombre de destino inválido!"
                txtDatos.Text = vbNullString
            End If
        ElseIf EtapaTransferencia = 2 Then
            Call WriteBankTransferGold(CantTransferencia, NombreTransferencia)
            Unload Me
        End If
End Select

End Sub

Private Sub Form_Load()

lstBanco.AddItem "Realizar un depósito de oro"
lstBanco.AddItem "Retirar de la cuenta"
lstBanco.AddItem "Ver la bóveda"
lstBanco.AddItem "Realizar una transferencia"

End Sub

Private Sub lstBanco_Click()

Select Case lstBanco.ListIndex
    Case 0 'Depositar
        lblDatos.Caption = "¿Cuánto deseas depositar?"
    Case 1 'Retirar
        lblDatos.Caption = "¿Cuánto deseas retirar?"
    Case 2 'Bóveda
        Call WriteBankStart
        Unload Me
    Case 3 'Transferir
        EtapaTransferencia = 0
        lblDatos.Caption = "¿Qué cantidad deseas transferir?"
End Select

End Sub
