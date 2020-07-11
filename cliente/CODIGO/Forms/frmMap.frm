VERSION 5.00
Begin VB.Form frmMap 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape shpMap 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   9480
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   135
   End
   Begin VB.Image imgMap 
      Appearance      =   0  'Flat
      Height          =   8655
      Left            =   180
      Top             =   180
      Width           =   11655
   End
   Begin VB.Label lblMMAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Posición cursor:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   1440
      TabIndex        =   1
      Top             =   8520
      Width           =   1545
   End
   Begin VB.Label lbIMAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Posición:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   1440
      TabIndex        =   0
      Top             =   8280
      Width           =   855
   End
   Begin VB.Line ln2 
      BorderColor     =   &H00694843&
      X1              =   787
      X2              =   787
      Y1              =   589
      Y2              =   12
   End
   Begin VB.Line ln1 
      BorderColor     =   &H00694843&
      X1              =   786
      X2              =   786
      Y1              =   589
      Y2              =   12
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xx As Integer, yy As Integer

Private Sub Form_Load()
    Me.Picture = LoadInterface("mundo")
    
    Me.Top = frmMain.Top
    Me.Left = frmMain.Left
    
    Aplicar_Transparencia Me.hwnd, 237

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        Unload Me
        frmMain.SetFocus
    End If
End Sub

Private Sub imgMap_DblClick()

    #If ClienteGM = 1 Then
            Call WriteWarpChar("YO", MapTable(xx, yy), 50, 50)
           
    #End If
     SetMapPoint
End Sub


Private Sub imgMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    X = X / 15 - 12
    Y = Y / 15 - 12
    xx = X / 26 + 1
    yy = Y / 25 + 1
                
    If xx <= 0 Or yy <= 0 Or xx > 30 Or yy > 23 Then Exit Sub
    lblMMAP.Caption = "Posición cursor: " & MapNames(MapTable(xx, yy)) & " (" & MapTable(xx, yy) & ")"



End Sub

Private Sub imgMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Unload Me
        frmMain.SetFocus
    End If
End Sub
Public Sub SetMapPoint()
    Dim i As Long, ii As Long
    For i = 1 To 30
        For ii = 1 To 23
            If MapTable(i, ii) = UserMap Then
                shpMap.Top = (ii - 1) * 25 + 12 + 9
                shpMap.Left = (i - 1) * 26 + 12 + 9
                
                lbIMAP.Caption = "Posición: " & MapNames(MapTable(i, ii)) & " (" & MapTable(i, ii) & ")"
                Exit For
            End If
        Next ii
    Next i
    
End Sub
