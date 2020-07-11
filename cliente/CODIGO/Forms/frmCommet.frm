VERSION 5.00
Begin VB.Form frmCommet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oferta de paz o alianza"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
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
   ScaleHeight     =   2820
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   120
      MouseIcon       =   "frmCommet.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   495
      Left            =   2400
      MouseIcon       =   "frmCommet.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmCommet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_PROPOSAL_LENGTH As Integer = 520

Public nombre As String


Private Sub Command1_Click()

    Call WriteGuildRejectNewMember(nombre, Replace(Replace(Text1.Text, ",", " "), vbCrLf, " "))

    Dim i As Long
    For i = 0 To frmGuildLeader.solicitudes.ListCount - 1
        If frmGuildLeader.solicitudes.List(i) = nombre Then
            frmGuildLeader.solicitudes.RemoveItem i
            Exit For
        End If
    Next i
    
    Me.Hide
    Unload frmCharInfo

    Unload Me

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Detalle de rechazo de membresía"
    Me.Text1.MaxLength = 50
End Sub

Private Sub Text1_Change()
    If Len(Text1.Text) > MAX_PROPOSAL_LENGTH Then _
        Text1.Text = Left$(Text1.Text, MAX_PROPOSAL_LENGTH)
End Sub
