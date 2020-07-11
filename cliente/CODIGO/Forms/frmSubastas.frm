VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSubastas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Centro de Subastas"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10215
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
   ScaleHeight     =   7425
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame framSubastas 
      Caption         =   "Buscar items"
      Height          =   7245
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   10065
      Begin MSComctlLib.ListView lstItems 
         Height          =   5865
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   10345
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CheckBox chkOwner 
         Caption         =   "Mostrar solamente mis subastas"
         Height          =   255
         Left            =   7320
         TabIndex        =   12
         Top             =   6840
         Width           =   2625
      End
      Begin VB.PictureBox picItem 
         BackColor       =   &H00000000&
         Height          =   540
         Left            =   120
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   10
         Top             =   240
         Width           =   540
      End
      Begin VB.CommandButton cmdNewAuction 
         Caption         =   "Nueva subasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   8520
         TabIndex        =   8
         Top             =   270
         Width           =   1400
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5340
         TabIndex        =   3
         Top             =   270
         Width           =   1400
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6930
         TabIndex        =   4
         Top             =   270
         Width           =   1400
      End
      Begin VB.TextBox txtOfertar 
         Height          =   285
         Left            =   1500
         TabIndex        =   5
         Top             =   6810
         Width           =   1065
      End
      Begin VB.CommandButton cmdOfertar 
         Caption         =   "Ofertar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   6
         Top             =   6810
         Width           =   1125
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3750
         TabIndex        =   2
         Top             =   270
         Width           =   1400
      End
      Begin VB.TextBox txtBuscar 
         Height          =   315
         Left            =   780
         TabIndex        =   1
         Top             =   340
         Width           =   2745
      End
      Begin VB.CommandButton cmdBuyOut 
         Caption         =   "Comprar (Directo)"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4050
         TabIndex        =   7
         Top             =   6810
         Width           =   1755
      End
      Begin VB.CommandButton cmdCancelAuction 
         Caption         =   "$528"
         Height          =   285
         Left            =   4050
         TabIndex        =   11
         Top             =   6810
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblCantidad 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Monedas de oro"
         Height          =   225
         Left            =   150
         TabIndex        =   9
         Top             =   6840
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmSubastas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sI As Byte
Sub RefreshList()
    With Me.lstItems
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        .ColumnHeaders.Add , , "Item"
        .ColumnHeaders.Add , , "Cantidad"
        .ColumnHeaders.Add , , "Precio (Actual)"
        .ColumnHeaders.Add , , "Precio (Final)"
        .ColumnHeaders.Add , , "Subastador"
        .ColumnHeaders.Add , , "Tiempo restante"
        .ColumnHeaders.Add , , "Comprador"
        
        .View = lvwReport
        .GridLines = True
        
        Dim item As ListItem, i As Byte
        
        For i = 1 To 100
            If Not lstSubastas(i).active = False Then
                Set item = .ListItems.Add(, , lstSubastas(i).OBJIndex & "  ," & i)
                item.SubItems(1) = CStr(lstSubastas(i).cant)
                item.SubItems(2) = CStr(lstSubastas(i).actOfert)
                item.SubItems(3) = CStr(lstSubastas(i).fnlOfert)
                item.SubItems(4) = CStr(lstSubastas(i).nckVndedor)
                item.SubItems(5) = CStr(lstSubastas(i).hsDura) & "h" & IIf(lstSubastas(i).mnDura <> 0, " " & CStr(lstSubastas(i).mnDura) & "m", "")
                item.SubItems(6) = CStr(IIf(lstSubastas(i).nckCmprdor = "", "Ninguno", lstSubastas(i).nckCmprdor))
            End If
        Next i
    End With
        
End Sub

Public Sub LimpiarSubastas()
On Error Resume Next
    Dim i As Byte
    For i = 1 To 100
        With lstSubastas(i)
            .active = False
            .actOfert = 0
            .cant = 0
            .fnlOfert = 0
            .hsDura = 0
            .mnDura = 0
            .nckCmprdor = ""
            .nckVndedor = ""
            .OBJIndex = ""
        End With
    Next i
End Sub

Private Sub cmdActualizar_Click()
    Call WriteSRequest
End Sub

Private Sub cmdBuscar_Click()
    Dim i As Byte
    For i = 1 To lstItems.ListItems.count
        If UCase$(mid$(lstItems.SelectedItem.Text, 1, Len(txtBuscar.Text))) = UCase$(txtBuscar.Text) Then
            lstItems.ListItems.item(i).Selected = True
        End If
    Next i
End Sub

Private Sub cmdBuyOut_Click()
    If sI > 0 Then
        Call WriteSComprar(sI)
        Call WriteSRequest
    End If
End Sub

Private Sub cmdNewAuction_Click()
    frmSubastaNueva.Show , Me
End Sub

Private Sub cmdOfertar_Click()
    If sI > 0 Then
        If val(txtOfertar.Text) > lstSubastas(sI).actOfert Then
            Call WriteSOfrecer(sI, val(txtOfertar.Text))
            Call WriteSRequest
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
    Unload frmSubastaNueva
End Sub


Private Sub lstItems_Click()
    If lstItems.ListItems.count <= 0 Then Exit Sub
    picItem.Cls
    sI = val(ReadField(2, lstItems.SelectedItem.Text, Asc(",")))
    Call TileEngine.Draw_Grh_Hdc(picItem.hdc, CInt(lstSubastas(sI).grhindex), 0, 0)
    cmdBuyOut.Enabled = IIf(lstSubastas(sI).fnlOfert = 0, False, True)
    txtOfertar.Text = lstSubastas(sI).actOfert + 100
End Sub
