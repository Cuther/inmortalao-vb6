Attribute VB_Name = "AoDefenderDynamicProtection"
Option Explicit
Private Const AoDefPocMin As Integer = 1
Private Const AoDefPocMax As Long = 1000000
Public AoDefResult As Long
Public Function AoDefProtectDynamic() As Long
    'Initialize randomizer
   AoDefResult = AoDefResult + 1
   If AoDefResult = 999999 Then
   AoDefResult = 1
   End If
   AoDefProtectDynamic = AoDefResult
End Function

