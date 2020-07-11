VERSION 5.00
Begin VB.Form frmMap 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   442
   ShowInTaskbar   =   0   'False
   Begin VB.Shape shpMap 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   135
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
      Left            =   120
      TabIndex        =   1
      Top             =   7560
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
      Left            =   120
      TabIndex        =   0
      Top             =   7320
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
   Begin VB.Image imgMap 
      Appearance      =   0  'Flat
      Height          =   8970
      Left            =   0
      Top             =   0
      Width           =   6630
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xx As Single, yy As Single

Private Sub Form_Load()

    Me.Picture = LoadInterface("imagenmundo")
    
    'Me.Top = frmMain.Top
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
End Sub


Private Sub imgMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    X = X * 0.066666666666667
    Y = Y * 0.066666666666667
    xx = Int(X / 26 + 1)
    yy = Int(Y / 26 + 1)
                
    If xx <= 0 Or yy <= 0 Or xx > 17 Or yy > 23 Then Exit Sub
    lblMMAP.Caption = "Posición cursor: " & MapNameTable(xx, yy) & " (" & MapTable(xx, yy) & ")"
    
End Sub

Private Sub imgMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Unload Me
        frmMain.SetFocus
    End If
End Sub
Public Sub SetMapPoint()
    

    Dim X As Long, Y As Long
    Dim gotPos As Boolean
    shpMap.Visible = False
    
    For Y = 1 To 23
        For X = 1 To 17
            If MapTable(X, Y) = UserMap Then
                shpMap.Left = (X - 1) * 26 + 26 / 2 - shpMap.width / 2
                shpMap.Top = (Y - 1) * 26 + 26 / 2 - shpMap.height / 2
                shpMap.Visible = True
                lbIMAP.Caption = "Posición: " & MapNameTable(X, Y) & " (" & MapTable(X, Y) & ")"
                Exit Sub
            End If
        Next X
    Next Y
        
        
End Sub
