VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ButtonHolder As Image
Attribute ButtonHolder.VB_VarHelpID = -1

Private RollOverButtonImage As Picture
Private ClickButtonImage As Picture

Private SoundClick As Boolean

Private FormActual As Form

Private OverButton As Boolean

Public Sub Initialize(ByRef imgButton As Image, ByRef imgRollOverPath As String, _
            ByRef imgClickPath As String, ByRef Formulario As Form, Optional ByVal MakeClickSound As Boolean = True)
    Set ButtonHolder = imgButton
    Set FormActual = Formulario
    
    Set RollOverButtonImage = LoadInterface(imgRollOverPath)
    Set ClickButtonImage = LoadInterface(imgClickPath)
    SoundClick = MakeClickSound
    
End Sub

Private Sub ButtonHolder_Click()
    If SoundClick Then
        Call Audio.PlayWave(SND_CLICK)
    End If
End Sub

Private Sub ButtonHolder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonHolder.Picture = ClickButtonImage
End Sub

Private Sub ButtonHolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not OverButton Then

        If Not FormActual.LastPressed Is Nothing Then _
            FormActual.LastPressed.ToggleToNormal
        
        ButtonHolder.Picture = RollOverButtonImage
        OverButton = True
        
        Set FormActual.LastPressed = Me
    End If
End Sub

Public Sub ToggleToNormal()
    If Not OverButton Then Exit Sub
    
    ButtonHolder.Picture = Nothing
    OverButton = False
End Sub

Private Sub ButtonHolder_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OverButton = False
    ButtonHolder_MouseMove Button, Shift, X, Y
End Sub


