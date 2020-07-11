Attribute VB_Name = "AoDefenderDebugger"
Private Declare Function IsDebuggerPresent Lib "kernel32" () As Long

Public Function AoDefDebugger() As Boolean
    If IsDebuggerPresent Then
        AoDefDebugger = True
        Exit Function
    End If
    AoDefDebugger = False
End Function

Public Sub AoDefAntiDebugger()
    MsgBox "Se ha detectado un intento de Debuggear, su cliente será cerrado.!", vbCritical, "Atencion!"
End Sub

