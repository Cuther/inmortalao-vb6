VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Compresor/Descompresor                 InmortalAO"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1440
      List            =   "Form1.frx":0013
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Descomprimir"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Comprimir"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   4
      Top             =   1080
      Width           =   75
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command5_Click()


Dim a As String
Dim b As Byte
Dim c As Byte


b = Combo2.ListIndex



If b = 0 Then
a = graph_path
c = Graphics
ElseIf b = 1 Then
a = init_path
c = Scripts
ElseIf b = 2 Then
a = inter_path
c = Interface
ElseIf b = 3 Then
a = map_path
c = Maps
ElseIf b = 4 Then
a = wav_path
c = Wav
End If

        Form1.Show
        Form1.Label1.Caption = "Comprimiendo " & Combo2.List(Combo2.ListIndex)
        Compress_Files c, a, resource_path
        Form1.Label1.Caption = "TERMINADO!"
        ProgressBar1.Value = 0

End Sub

Private Sub Command6_Click()

Dim a As String
Dim b As Byte
Dim c As Byte

b = Combo2.ListIndex

If b = 0 Then
c = Graphics
ElseIf b = 1 Then
c = Scripts
ElseIf b = 2 Then
c = Interface
ElseIf b = 3 Then
c = Maps
ElseIf b = 4 Then
c = Wav
End If

        Form1.Show
        Form1.Label1.Caption = "Descomprimiendo " & Combo2.List(Combo2.ListIndex)
        Extract_All_Files c, resource_path
        Form1.Label1.Caption = "TERMINADO!"
        ProgressBar1.Value = 0

End Sub

Private Sub Form_Load()
resource_path = App.Path & "\Resources\"
init_path = App.Path & "\Descomprimido\Init\"
graph_path = App.Path & "\Descomprimido\graficos\"
inter_path = App.Path & "\Descomprimido\Interface\"
map_path = App.Path & "\Descomprimido\mapas\"
wav_path = App.Path & "\Descomprimido\Wav\"
End Sub

