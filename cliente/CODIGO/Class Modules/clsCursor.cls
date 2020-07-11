VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GCL_HCURSOR = -12

Public Enum CursorType
    E_NORMAL = 0
    E_ATTACK
    E_ARROW
    E_CAST
    E_WAIT
    E_SHOOT
    E_SHIP
End Enum

Private Const NUM_CURSORS = 6

#If Cursor = 1 Then
    Private hndlList(0 To NUM_CURSORS) As IPictureDisp
#Else
    Private hndlList(0 To NUM_CURSORS) As Long
#End If

Public Sub Parse_Form(ByRef aFrm As Form, Optional ByVal cType As CursorType = E_NORMAL)

On Error Resume Next

Dim aControl As Control
Dim lngHandle As Long
If CursoresStandar = 1 Then
    Select Case cType
        Case E_NORMAL
            lngHandle = vbDefault
        Case E_ATTACK
            lngHandle = vbCrosshair
        Case E_ARROW
            lngHandle = vbCrosshair
        Case E_CAST
            lngHandle = vbCrosshair
        Case E_WAIT
            lngHandle = vbHourglass
        Case E_SHOOT
            lngHandle = vbCrosshair
        Case E_SHIP
            lngHandle = vbCrosshair
    End Select
    
    aFrm.MouseIcon = Nothing
    aFrm.MousePointer = lngHandle

Else
    For Each aControl In aFrm.Controls
        aControl.MousePointer = vbCustom
        #If Cursor = 0 Then
            SetClassLong (aControl.hwnd), GCL_HCURSOR, hndlList(cType)
        #Else
            Set aControl.MouseIcon = hndlList(cType)
        #End If
    Next
    
    aFrm.MousePointer = vbCustom
    #If Cursor = 0 Then
        SetClassLong (aFrm.hwnd), GCL_HCURSOR, hndlList(cType)
    #Else
        Set aFrm.MouseIcon = hndlList(cType)
    #End If
End If

End Sub

Public Sub Init()
#If Cursor = 0 Then
    hndlList(E_NORMAL) = LoadCursorFromFile(resource_path & "main.cur")
    hndlList(E_ATTACK) = LoadCursorFromFile(resource_path & "attack.cur")
    hndlList(E_ARROW) = LoadCursorFromFile(resource_path & "arrow.cur")
    hndlList(E_CAST) = LoadCursorFromFile(resource_path & "cast.cur")
    hndlList(E_WAIT) = LoadCursorFromFile(resource_path & "wait.cur")
    hndlList(E_SHOOT) = LoadCursorFromFile(resource_path & "shoot.cur")
    hndlList(E_SHIP) = LoadCursorFromFile(resource_path & "ship.cur")
#Else
    Set hndlList(E_NORMAL) = LoadResData("MAIN", "CUSTOM")
    Set hndlList(E_ATTACK) = LoadResPicture("ATTACK", "CUSTOM")
    Set hndlList(E_ARROW) = LoadResPicture("ARROW", "CUSTOM")
    Set hndlList(E_CAST) = LoadResPicture("CAST", "CUSTOM")
    Set hndlList(E_WAIT) = LoadResPicture("WAIT", vbResCursor)
    Set hndlList(E_SHOOT) = LoadResPicture("SHOOT", vbResCursor)
    Set hndlList(E_SHIP) = LoadResPicture("SHIP", vbResCursor)
#End If
End Sub

