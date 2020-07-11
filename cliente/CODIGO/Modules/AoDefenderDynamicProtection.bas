Attribute VB_Name = "AoDefenderDynamicProtection"
Option Explicit
Public AoDefResult As Long
Public Function AoDefProtectDynamic() As Long
    'Initialize randomizer
   AoDefResult = AoDefResult + 1
   If AoDefResult = 254 Then
   AoDefResult = 15
   End If
   AoDefProtectDynamic = AoDefResult
End Function

