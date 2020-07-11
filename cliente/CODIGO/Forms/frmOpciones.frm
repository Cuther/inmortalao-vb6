VERSION 5.00
Begin VB.Form frmOpciones 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6885
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
   ScaleHeight     =   7065
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Salir"
      Height          =   360
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   6555
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sonido"
      Height          =   3495
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   3255
      Begin VB.CheckBox chkMidi 
         Caption         =   "Música"
         Height          =   255
         Left            =   150
         TabIndex        =   25
         Top             =   850
         Width           =   2985
      End
      Begin VB.CheckBox chkAudio 
         Caption         =   "Audio"
         Height          =   285
         Left            =   150
         TabIndex        =   24
         Top             =   530
         Width           =   2955
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Repetir música"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   23
         Top             =   1150
         Width           =   2985
      End
      Begin VB.HScrollBar scrVolume 
         Height          =   315
         LargeChange     =   15
         Left            =   150
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   22
         Top             =   1800
         Width           =   2895
      End
      Begin VB.HScrollBar scrAmbient 
         Enabled         =   0   'False
         Height          =   315
         LargeChange     =   15
         Left            =   150
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   21
         Top             =   2400
         Width           =   2895
      End
      Begin VB.HScrollBar scrMidi 
         Height          =   315
         LargeChange     =   15
         Left            =   150
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   20
         Top             =   3000
         Width           =   2895
      End
      Begin VB.CheckBox chkInvertir 
         Caption         =   "Invertir canales de audio (L/R)"
         Height          =   255
         Left            =   150
         TabIndex        =   19
         Top             =   240
         Width           =   2985
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de audio"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   28
         Top             =   1560
         Width           =   2835
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de sonidos ambientales"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   27
         Top             =   2160
         Width           =   2865
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de música"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   26
         Top             =   2760
         Width           =   2865
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información"
      Height          =   1575
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   4440
      Width           =   3255
      Begin VB.CommandButton cmdAyuda 
         Caption         =   "¿Necesitás ayuda?"
         Height          =   345
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CommandButton cmdWeb 
         Caption         =   "http://&inmortalao.com.ar"
         Height          =   345
         Index           =   0
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   300
         Width           =   2895
      End
      Begin VB.CommandButton cmdWeb 
         Caption         =   "Control de cuentas"
         Height          =   345
         Index           =   1
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   690
         Width           =   2895
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Apariencia y performance"
      Height          =   3105
      Left            =   3480
      TabIndex        =   9
      Top             =   120
      Width           =   3285
      Begin VB.CheckBox chkop 
         Caption         =   "Ver Nombre del Mapa"
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   11
         Top             =   240
         Width           =   2715
      End
      Begin VB.ListBox lstSkin 
         Enabled         =   0   'False
         Height          =   1860
         ItemData        =   "frmOpciones.frx":0000
         Left            =   180
         List            =   "frmOpciones.frx":0007
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Skins Instalados"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   13
         Top             =   600
         Width           =   2925
      End
      Begin VB.Label lblSkinData 
         BackStyle       =   0  'Transparent
         Caption         =   "Autor: Marius"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   2760
         Width           =   2895
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "General"
      Height          =   3705
      Left            =   3480
      TabIndex        =   3
      Top             =   3240
      Width           =   3285
      Begin VB.CheckBox chkop 
         Caption         =   "Habilitar mensajes globales"
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   180
         TabIndex        =   7
         Top             =   270
         Width           =   2715
      End
      Begin VB.ListBox lstIgnore 
         Enabled         =   0   'False
         Height          =   2010
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "Temporalmente Deshabilitado"
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Deshabilitar cursores gráficos"
         Height          =   285
         Index           =   5
         Left            =   180
         TabIndex        =   5
         Top             =   600
         Width           =   2715
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Habilitar chat faccionario"
         Height          =   285
         Index           =   6
         Left            =   180
         TabIndex        =   4
         Top             =   960
         Width           =   2715
      End
      Begin VB.Label Label4 
         Caption         =   "Lista de ignorados"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   1320
         Width           =   2925
      End
   End
   Begin VB.CommandButton cmdControles 
      Caption         =   "Configurar controles"
      Height          =   360
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6120
      Width           =   3255
   End
   Begin VB.Frame frmIdioma 
      Caption         =   "Lenguaje"
      Height          =   705
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.ComboBox cmbLanguage 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOpciones.frx":0016
         Left            =   180
         List            =   "frmOpciones.frx":0020
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private loading As Boolean
Private change As Boolean
Private Sub chkAudio_Click()
    If chkAudio.value = vbUnchecked Then
        Sound = 0
        scrVolume.Enabled = False
    Else
        Sound = 1
        scrVolume.Enabled = True
        scrVolume.value = Audio.SoundVolume
    End If
    
    Audio.SoundRefresh
    change = True
End Sub

Private Sub chkInvertir_Click()
    InvertCanal = IIf(chkInvertir.value = vbChecked, 1, 0)
End Sub

Private Sub chkMidi_Click()
    If Music = IIf(chkMidi.value = vbChecked, 1, 0) Then Exit Sub
    If chkMidi.value = vbUnchecked Then
        Music = 0
        scrMidi.Enabled = False
    ElseIf Not Music Then
        Music = 1
        scrMidi.Enabled = True
        scrMidi.value = Audio.MusicVolume
    End If
    Audio.MusicRefresh
    change = True
End Sub

Private Sub chkop_Click(Index As Integer)
    Dim Opcion As Byte
    Opcion = IIf(chkop(Index).value = vbChecked, 1, 0)
    Select Case Index
        Case 0
            RepitMusic = Opcion
            
        Case 5
            Cursores = Opcion
            msgbox_ok "Para que los cambios en esta opción sean reflejados, deberá reiniciar el cliente."
            
        Case 6
            ChatFaccionario = Opcion
        
        Case 9
            ChatGlobal = Opcion
            
    End Select
End Sub

Private Sub cmdAyuda_Click()
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    Call ShellExecute(0, "Open", "http://inmortalao.com.ar/wiki/", "", App.Path, 0)
End Sub

Private Sub cmdCerrar_Click()
    If change Then
        If MsgBox("¿Desea guardar los cambios hechos?", vbYesNo) = vbYes Then
            SaveConfig
        End If
    End If
    
    Unload Me
    frmMain.SetFocus
End Sub


Private Sub cmdControles_Click()
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    Call frmCustomKeys.Show(vbModal, Me)
End Sub

Private Sub cmdWeb_Click(Index As Integer)
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    Call ShellExecute(0, "Open", "http://inmortalao.com.ar/" & IIf(Index = 1, "panel-de-usuario.php", ""), "", App.Path, 0)
End Sub

Private Sub Form_Load()
    loading = True      'Prevent sounds when setting check's values
    
    scrMidi.min = -5000
    scrMidi.max = 0
    
    scrVolume.min = 0
    scrVolume.max = 100

    If VolumeSound > 100 Then VolumeSound = 100
    scrVolume.value = VolumeSound
    If VolumeMusic > 0 Then VolumeMusic = 0
    scrMidi.value = VolumeMusic
    
    chkAudio.value = IIf(Sound = 1, vbChecked, vbUnchecked)
    chkMidi.value = IIf(Music = 1, vbChecked, vbUnchecked)
    chkInvertir.value = IIf(InvertCanal = 1, vbChecked, vbUnchecked)
    chkop(0).value = IIf(RepitMusic = 1, vbChecked, vbUnchecked)
    
    cmbLanguage.ListIndex = 0
    cmbLanguage.Enabled = False
    
    loading = False
    
    change = False
End Sub

Private Sub scrMidi_Change()
    Audio.MusicVolume = scrMidi.value
    change = True
End Sub

Private Sub scrVolume_Change()
    Audio.SoundVolume = scrVolume.value
    change = True
End Sub
Private Sub scrMidi_Scroll()
    Audio.MusicVolume = scrMidi.value
    change = True
End Sub

Private Sub scrVolume_Scroll()
    Audio.SoundVolume = scrVolume.value
    change = True
End Sub
