Attribute VB_Name = "AntiExternos"
Option Explicit
 
 
 
 
Declare Function EnumWindows Lib "user32" ( _
                 ByVal wndenmprc As Long, _
                 ByVal lParam As Long) As Long
 
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
                 ByVal hwnd As Long, _
                 ByVal lpString As String, _
                 ByVal cch As Long) As Long
 
 
Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                 ByVal hwnd As Long, _
                 ByVal wMsg As Long, _
                 ByVal wParam As Long, _
                 lParam As Any) As Long
 
 
 
Const WM_SYSCOMMAND = &H112
Const SC_CLOSE = &HF060&
 

Public Sub Externos()
    Call EnumWindows(AddressOf recorrerVentanas, 0)
End Sub
 

Public Function recorrerVentanas(ByVal hwnd As Long, ByVal param As Long) As Long
 
Dim buffer As String * 256
Dim tWindows As String
Dim Size_buffer As Long
 
 
    Size_buffer = GetWindowText(hwnd, buffer, Len(buffer))
   
    tWindows = Left$(buffer, Size_buffer)
   
    'Mod Marius los agregamos aca así lo hace mas rapido y gasta menos recursos.
    If InStr(UCase(tWindows), "CHEAT") <> 0 Or _
    InStr(UCase(tWindows), "MACRO") <> 0 Or _
    InStr(UCase(tWindows), "ENGINE") <> 0 Or _
    InStr(UCase(tWindows), "SPEED") <> 0 Or _
    UCase(tWindows) = "WPE PRO" Or _
    UCase(tWindows) = "TINYTASK" Or _
    UCase(tWindows) = "KINGDOMS OF CAMELOT EN FACEBOOK - MOZILLA FIREFOX" Then
    'wpe pro es el editor de paquetes con el que pegan rapido
    'el kindom es un auto potas que tambien usaban los mismo del wpe
        SendMessage hwnd, WM_SYSCOMMAND, SC_CLOSE, ByVal 0&
    End If
 
   
    recorrerVentanas = 1
End Function

