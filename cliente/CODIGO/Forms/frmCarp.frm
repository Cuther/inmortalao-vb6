VERSION 5.00
Begin VB.Form frmCarp 
   Caption         =   "Carpintero"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4650
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   270
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2580
      Width           =   1710
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Construir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2580
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2580
      Width           =   1710
   End
   Begin VB.ListBox lstArmas 
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4080
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   1140
      TabIndex        =   0
      Text            =   "1"
      Top             =   2160
      Width           =   3165
   End
   Begin VB.Label lblCantidad 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   4
      Top             =   2190
      Width           =   855
   End
End
Attribute VB_Name = "frmCarp"
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

Private Sub Command3_Click()
    On Error Resume Next

    Call WriteCraftCarpenter(ObjCarpintero(lstArmas.ListIndex + 1), Val(txtCantidad.Text))

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
