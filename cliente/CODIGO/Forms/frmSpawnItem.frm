VERSION 5.00
Begin VB.Form frmSpawnItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crear Item"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   3180
      ItemData        =   "frmSpawnItem.frx":0000
      Left            =   2880
      List            =   "frmSpawnItem.frx":0002
      TabIndex        =   6
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   240
      ScaleHeight     =   2400
      ScaleWidth      =   2280
      TabIndex        =   2
      Top             =   1200
      Width           =   2280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      MouseIcon       =   "frmSpawnItem.frx":0004
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3840
      Width           =   2685
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4560
      MouseIcon       =   "frmSpawnItem.frx":0156
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   4320
      Width           =   810
   End
   Begin VB.Frame Frame1 
      Caption         =   "Imagen item"
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "RTA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   720
      TabIndex        =   7
      Top             =   600
      Width           =   75
   End
End
Attribute VB_Name = "frmSpawnItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

   


Private Sub Command1_Click()
Call WriteCreateItem(List1.ListIndex + 1, 1)
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
List1.Clear

                      Dim i As Long, N As Long
                            For i = 1 To UBound(objs)
                                If InStr(1, Tilde(objs(i).name), Tilde(Text2.Text)) Then
                                    List1.AddItem (i & " " & objs(i).name & ".")
                                    N = N + 1
                                End If
                            Next
                            
                            If N = 0 Then
                               Label1.Caption = "No hubo resultados de la busqueda: " & Text2.Text
                            Else
                               Label1.Caption = "Hubo " & N & " resultados de la busqueda: " & Text2.Text
                            End If
                        
End Sub


Private Sub List1_Click()
Picture1.Cls
Call TileEngine.Draw_Grh_Hdc(Picture1.hdc, objs(List1.ListIndex + 1).Grh, 0, 0)
End Sub
