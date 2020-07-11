Attribute VB_Name = "AoDefenderEncryptServer"
Option Explicit

Private Function ConvToHex(x As Integer) As String
    If x > 9 Then
        ConvToHex = Chr(x + 55)
    Else
        ConvToHex = CStr(x)
    End If
End Function

' función que codifica el dato
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Function AoDefServEncrypt(DataValue As Variant) As Variant
    
    Dim x As Long
    Dim Temp As String
    Dim TempNum As Integer
    Dim TempChar As String
    Dim TempChar2 As String
    
    For x = 1 To Len(DataValue)
        TempChar2 = mid(DataValue, x, 1)
        TempNum = Int(Asc(TempChar2) / 16)
        
        If ((TempNum * 16) < Asc(TempChar2)) Then
               
            TempChar = ConvToHex(Asc(TempChar2) - (TempNum * 16))
            Temp = Temp & ConvToHex(TempNum) & TempChar
        Else
            Temp = Temp & ConvToHex(TempNum) & "0"
        
        End If
    Next x
    
    
    AoDefServEncrypt = Temp
End Function
Private Function ConvToInt(x As String) As Integer
    
    Dim x1 As String
    Dim x2 As String
    Dim Temp As Integer
    
    x1 = mid(x, 1, 1)
    x2 = mid(x, 2, 1)
    
    If IsNumeric(x1) Then
        Temp = 16 * Int(x1)
    Else
        Temp = (Asc(x1) - 55) * 16
    End If
    
    If IsNumeric(x2) Then
        Temp = Temp + Int(x2)
    Else
        Temp = Temp + (Asc(x2) - 55)
    End If
    
    ' retorno
    ConvToInt = Temp
    
End Function

' función que decodifica el dato
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function AoDefServDecrypt(DataValue As Variant) As Variant
    
    Dim x As Long
    Dim Temp As String
    Dim HexByte As String
    
    For x = 1 To Len(DataValue) Step 2
        
        HexByte = mid(DataValue, x, 2)
        Temp = Temp & Chr(ConvToInt(HexByte))
        
    Next x
    ' retorno
    AoDefServDecrypt = Temp
    
End Function



