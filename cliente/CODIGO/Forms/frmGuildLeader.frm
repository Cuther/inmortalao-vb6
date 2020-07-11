VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración del Clan"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   ClipControls    =   0   'False
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Borrar Clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   16
      Top             =   5160
      Width           =   2775
   End
   Begin VB.CommandButton Command8 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   6000
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Editar URL de la web del clan"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Editar Codex o Descripcion"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   3930
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      Caption         =   "Clanes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   2895
      Begin VB.ListBox guildslist 
         Height          =   1425
         ItemData        =   "frmGuildLeader.frx":03F6
         Left            =   120
         List            =   "frmGuildLeader.frx":03F8
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":03FA
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1800
         Width           =   2655
      End
   End
   Begin VB.Frame txtnews 
      Caption         =   "Notícias"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":054C
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox txtguildnews 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Miembros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton Command2 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":069E
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ListBox members 
         Height          =   1425
         ItemData        =   "frmGuildLeader.frx":07F0
         Left            =   120
         List            =   "frmGuildLeader.frx":07F2
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solicitudes de ingreso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   2895
      Begin VB.CommandButton Command1 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":07F4
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1560
         Width           =   2655
      End
      Begin VB.ListBox solicitudes 
         Height          =   1230
         ItemData        =   "frmGuildLeader.frx":0946
         Left            =   120
         List            =   "frmGuildLeader.frx":0948
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Miembros 
         Caption         =   "El clan cuenta con x miembros"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmGuildLeader"
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

Private Const MAX_NEWS_LENGTH As Integer = 512


Private Sub Command1_Click()
    If solicitudes.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembershipRequests
    Call WriteGuildMemberInfo(solicitudes.List(solicitudes.ListIndex))

    'Unload Me
End Sub

Private Sub Command2_Click()
    If members.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembers
    Call WriteGuildMemberInfo(members.List(members.ListIndex))

    'Unload Me
End Sub

Private Sub Command3_Click()
    Dim k As String

    k = Replace(txtguildnews, vbCrLf, "º")
    
    Call WriteGuildUpdateNews(k)
End Sub

Private Sub Command4_Click()
    frmGuildBrief.EsLeader = True
    Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))

    'Unload Me
End Sub

Private Sub Command5_Click()
    Call frmGuildDetails.Show(vbModal, frmGuildLeader)
    
    'Unload Me
End Sub

Private Sub Command6_Click()
Call frmGuildURL.Show(vbModeless, frmGuildLeader)
'Unload Me
End Sub


Private Sub Command7_Click()
    Dim inputm As String
    inputm = InputBox("Atencion: La accion que quiere hacer es peligrosa, ya que una vez hecho no hay vuelta atras. Si de en verdad quiere eliminar el clan escribe" & Chr(34) & "BORRAR CLAN" & Chr(34) & " en mayusculas y sin comillas.", "ADVERTENCIA")
    
    If inputm = "BORRAR CLAN" Then
        Call WriteDelClan
        Unload Me
    End If
    
End Sub

Private Sub Command8_Click()
    Unload Me
End Sub

Private Sub txtguildnews_Change()
    If Len(txtguildnews.Text) > MAX_NEWS_LENGTH Then _
        txtguildnews.Text = Left$(txtguildnews.Text, MAX_NEWS_LENGTH)
End Sub
