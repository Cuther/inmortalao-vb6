VERSION 5.00
Begin VB.Form frmHerrero 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Herrero"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5355
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstArmaduras 
      Height          =   2205
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Armas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ar&maduras"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.ListBox lstArmas 
      Height          =   2205
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   4815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Construir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "1"
      Top             =   3060
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   7
      Top             =   3075
      Width           =   855
   End
End
Attribute VB_Name = "frmHerrero"
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

Private Sub Command1_Click()
    lstArmaduras.Visible = False
    lstArmas.Visible = True
End Sub

Private Sub Command2_Click()
    lstArmaduras.Visible = True
    lstArmas.Visible = False
End Sub

Private Sub Command3_Click()
On Error Resume Next

    If lstArmas.Visible Then
        Call WriteCraftBlacksmith(ArmasHerrero(lstArmas.ListIndex + 1), val(txtCantidad.Text))
    Else
        Call WriteCraftBlacksmith(ArmadurasHerrero(lstArmaduras.ListIndex + 1), val(txtCantidad.Text))
    End If

    Unload Me
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
End Sub

Private Sub Form_Load()
    lstArmaduras.Visible = True
    lstArmas.Visible = False
End Sub
