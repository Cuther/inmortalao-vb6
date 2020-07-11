VERSION 5.00
Begin VB.Form frmComerciar 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   3960
      Index           =   0
      Left            =   780
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   2580
      Width           =   2460
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   3960
      Index           =   1
      Left            =   3720
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   2580
      Width           =   2490
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   900
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   1620
      Width           =   480
   End
   Begin VB.TextBox Cantidad 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   0
      Text            =   "1"
      Top             =   6960
      Width           =   510
   End
   Begin VB.Timer tmrNumber 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   30
      Top             =   30
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   7
      Top             =   1530
      Width           =   2985
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Index           =   1
      Left            =   5100
      TabIndex        =   6
      Top             =   1890
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Index           =   2
      Left            =   5520
      TabIndex        =   5
      Top             =   1530
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   3
      Left            =   1560
      TabIndex        =   4
      Top             =   1830
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   0
      Left            =   585
      Tag             =   "1"
      Top             =   6855
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   1
      Left            =   4200
      Tag             =   "1"
      Top             =   6855
      Width           =   2175
   End
   Begin VB.Image Command2 
      Height          =   345
      Left            =   6480
      Tag             =   "1"
      Top             =   180
      Width           =   345
   End
   Begin VB.Image cmdMasMenos 
      Height          =   420
      Index           =   1
      Left            =   3840
      Tag             =   "1"
      Top             =   6870
      Width           =   195
   End
   Begin VB.Image cmdMasMenos 
      Height          =   420
      Index           =   0
      Left            =   2940
      Tag             =   "1"
      Top             =   6870
      Width           =   195
   End
End
Attribute VB_Name = "frmComerciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez



Option Explicit

Public LastIndex1 As Integer
Public LastIndex2 As Integer
Public LasActionBuy As Boolean
Private lIndex As Byte

Private m_Number As Integer
Private m_Increment As Integer
Private m_Interval As Integer

Private Sub cantidad_Change()
    If val(Cantidad.Text) < 1 Then
        Cantidad.Text = 1
    End If
    
    If val(Cantidad.Text) > MAX_INVENTORY_OBJS Then
        Cantidad.Text = 1
    End If
    
    If lIndex = 0 Then
        If List1(0).ListIndex <> -1 Then
            'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
            Label1(1).Caption = CalculateSellPrice(NPCInventory(List1(0).ListIndex + 1).Valor, val(Cantidad.Text)) 'No mostramos numeros reales
        End If
    Else
        If List1(1).ListIndex <> -1 Then
            Label1(1).Caption = CalculateBuyPrice(Inventario.Valor(List1(1).ListIndex + 1), val(Cantidad.Text)) 'No mostramos numeros reales
        End If
    End If
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Audio.PlayWave(SND_CLICK)
Command2.Picture = LoadInterface("salir-down")
Command2.Tag = "1"
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Command2.Tag = "0" Then
    Command2.Picture = LoadInterface("salir-over")
    Command2.Tag = "1"
End If

End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
Call WriteCommerceEnd
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Button = vbLeftButton) Then
    Call Auto_Drag(Me.hwnd)
Else
    Call WriteCommerceEnd
End If

End Sub

Private Sub Form_Load()
Me.Picture = LoadInterface("comercio")
m_Number = 1
m_Interval = 30
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Image1(0).Tag = "1" Then
    Image1(0).Picture = Nothing
    Image1(0).Tag = "0"
End If

If Image1(1).Tag = "1" Then
    Image1(1).Picture = Nothing
    Image1(1).Tag = "0"
End If

If cmdMasMenos(0).Tag = "1" Then
    cmdMasMenos(0).Picture = Nothing
    cmdMasMenos(0).Tag = "0"
End If

If cmdMasMenos(1).Tag = "1" Then
    cmdMasMenos(1).Picture = Nothing
    cmdMasMenos(1).Tag = "0"
End If

If Command2.Tag = "1" Then
    Command2.Picture = Nothing
    Command2.Tag = "0"
End If

End Sub


''
' Calculates the selling price of an item (The price that a merchant will sell you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.

Private Function CalculateSellPrice(ByRef objValue As Single, ByVal objAmount As Long) As Long
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 19/08/2008
'Last modify by: Franco Zeoli (Noich)
'*************************************************
    On Error GoTo error
    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateSellPrice = CCur(objValue * 1000000) / 1000000 * objAmount + 0.5
    
    Exit Function
error:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.Number
End Function
''
' Calculates the buying price of an item (The price that a merchant will buy you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.
Private Function CalculateBuyPrice(ByRef objValue As Single, ByVal objAmount As Long) As Long
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 19/08/2008
'Last modify by: Franco Zeoli (Noich)
'*************************************************
    On Error GoTo error
    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateBuyPrice = Fix(CCur(objValue * 1000000) / 1000000 * objAmount)
    
    Exit Function
error:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.Number
End Function

Private Sub Image1_Click(Index As Integer)

Call Audio.PlayWave(SND_CLICK)


If List1(Index).List(List1(Index).ListIndex) = "(Nada)" Or _
   List1(Index).ListIndex < 0 Then Exit Sub

If Not IsNumeric(Cantidad.Text) Or Cantidad.Text = 0 Then Exit Sub

Select Case Index
    Case 0
        frmComerciar.List1(0).SetFocus
        LastIndex1 = List1(0).ListIndex
        LasActionBuy = True
        If UserGLD >= CalculateSellPrice(NPCInventory(List1(0).ListIndex + 1).Valor, val(Cantidad.Text)) Then
            Call WriteCommerceBuy(List1(0).ListIndex + 1, Cantidad.Text)
        Else
            AddtoRichTextBox frmMain.RecChat, "No tenés suficiente oro.", 2, 51, 223, 1, 1
            Exit Sub
        End If
   
   Case 1
        frmComerciar.List1(1).SetFocus
        LastIndex2 = List1(1).ListIndex
        LasActionBuy = False
        Call WriteCommerceSell(List1(1).ListIndex + 1, Cantidad.Text)
        
        
End Select

List1(0).Clear
List1(1).Clear


End Sub


Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index = 0 Then
    Image1(Index).Picture = LoadInterface("comprar-down")
    Image1(Index).Tag = "1"
ElseIf Index = 1 Then
    Image1(Index).Picture = LoadInterface("vender-down")
    Image1(Index).Tag = "1"
End If

End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index = 0 Then
    If Image1(Index).Tag = "0" Then
        Image1(Index).Picture = LoadInterface("comprar-over")
        Image1(Index).Tag = "1"
    End If
ElseIf Index = 1 Then
    If Image1(Index).Tag = "0" Then
        Image1(Index).Picture = LoadInterface("vender-over")
        Image1(Index).Tag = "1"
    End If
End If

End Sub

Private Sub List1_Click(Index As Integer)

lIndex = Index

Select Case Index
    Case 0
        
        Label1(0).Caption = NPCInventory(List1(0).ListIndex + 1).name
        Label1(1).Caption = CalculateSellPrice(NPCInventory(List1(0).ListIndex + 1).Valor, val(Cantidad.Text)) 'No mostramos numeros reales
        Label1(2).Caption = NPCInventory(List1(0).ListIndex + 1).Amount
        
        If Label1(2).Caption <> 0 Then
            Select Case NPCInventory(List1(0).ListIndex + 1).OBJType
                Case eObjType.otWeapon
                    Label1(3).Caption = "Golpe:" & NPCInventory(List1(0).ListIndex + 1).MinHit & "\" & NPCInventory(List1(0).ListIndex + 1).MaxHit
                    Label1(3).Visible = True
                Case eObjType.otArmadura
                    Label1(3).Caption = "Defensa:" & NPCInventory(List1(0).ListIndex + 1).Def
                    Label1(3).Visible = True
                Case Else
                    Label1(3).Visible = False
                    
            End Select
            
            Picture1.Cls
            If NPCInventory(List1(0).ListIndex + 1).grhindex Then
                Call TileEngine.Draw_Grh_Hdc(Picture1.hdc, NPCInventory(List1(0).ListIndex + 1).grhindex, 0, 0)
            Else
                Picture1.Picture = Nothing
            End If
        End If
    
    Case 1
        Label1(0).Caption = Inventario.ItemName(List1(1).ListIndex + 1)
        Label1(1).Caption = CalculateBuyPrice(Inventario.Valor(List1(1).ListIndex + 1), val(Cantidad.Text)) 'No mostramos numeros reales
        Label1(2).Caption = Inventario.Amount(List1(1).ListIndex + 1)
        
        If Label1(2).Caption <> 0 Then
        
            Select Case Inventario.OBJType(List1(1).ListIndex + 1)
                Case eObjType.otWeapon
                    Label1(3).Caption = "Golpe:" & Inventario.MinHit(List1(1).ListIndex + 1) & "\" & Inventario.MaxHit(List1(1).ListIndex + 1)
                    Label1(3).Visible = True
                Case eObjType.otArmadura
                    Label1(3).Caption = "Defensa:" & Inventario.Def(List1(1).ListIndex + 1)
                    Label1(3).Visible = True
                Case Else
                    Label1(3).Visible = False
            End Select
            
            Picture1.Cls
            If Inventario.grhindex(List1(1).ListIndex + 1) Then
                Call TileEngine.Draw_Grh_Hdc(Picture1.hdc, Inventario.grhindex(List1(1).ListIndex + 1), 0, 0)
            Else
                Picture1.Picture = Nothing
            End If
        End If
        
End Select

If Label1(2).Caption = 0 Then ' 27/08/2006 - GS > No mostrar imagen ni nada, cuando no ahi nada que mostrar.
    Label1(3).Visible = False
    Picture1.Visible = False
Else
    Picture1.Visible = True
    Picture1.refresh
End If

End Sub

Private Sub cmdMasMenos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0
        cmdMasMenos(Index).Picture = LoadInterface("menos-down")
        cmdMasMenos(Index).Tag = "1"
        Cantidad.Text = str((val(Cantidad.Text) - 1))
        m_Increment = -1
    Case 1
        cmdMasMenos(Index).Picture = LoadInterface("mas-down")
        cmdMasMenos(Index).Tag = "1"
        m_Increment = 1
End Select

tmrNumber.Interval = 30
tmrNumber.Enabled = True

End Sub

Private Sub cmdMasMenos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index
    Case 0
        If cmdMasMenos(Index).Tag = "0" Then
            cmdMasMenos(Index).Picture = LoadInterface("menos-over")
            cmdMasMenos(Index).Tag = "1"
        End If
    Case 1
        If cmdMasMenos(Index).Tag = "0" Then
            cmdMasMenos(Index).Picture = LoadInterface("mas-over")
            cmdMasMenos(Index).Tag = "1"
        End If
End Select

End Sub

Private Sub cmdMasMenos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
tmrNumber.Enabled = False
End Sub

Private Sub tmrNumber_Timer()

Const MIN_NUMBER = 1
Const MAX_NUMBER = 10000

    m_Number = m_Number + m_Increment
    If m_Number < MIN_NUMBER Then
        m_Number = MIN_NUMBER
    ElseIf m_Number > MAX_NUMBER Then
        m_Number = MAX_NUMBER
    End If

    Cantidad.Text = Format$(m_Number)
    
    If m_Interval > 1 Then
        m_Interval = m_Interval - 1
        tmrNumber.Interval = m_Interval
    End If

End Sub

