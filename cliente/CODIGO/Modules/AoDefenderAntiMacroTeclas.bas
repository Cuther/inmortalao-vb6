Attribute VB_Name = "AoDefenderAntiMacroTeclas"
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal nCode As Long) As Integer
Public Function AoDefMacrer(ByVal KeyCode As Integer) As Boolean
If Not GetAsyncKeyState(KeyCode) < 0 Then
AoDefMacrer = True
Else
AoDefMacrer = False
End If
End Function

