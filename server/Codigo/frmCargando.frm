VERSION 5.00
Begin VB.Form frmCargando 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Inmortal AO"
   ClientHeight    =   2445
   ClientLeft      =   1410
   ClientTop       =   3000
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   5415
   End
   Begin VB.PictureBox cargar 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   5355
      TabIndex        =   4
      Top             =   1080
      Width           =   5415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Inmortal AO"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando, por favor espere..."
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   3
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   2
      Left            =   90
      TabIndex        =   0
      Top             =   1560
      Width           =   5370
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    End
End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

End Sub
