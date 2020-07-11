Attribute VB_Name = "AoDefenderMultiClient"
Option Explicit

Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByRef lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Const ERROR_ALREADY_EXISTS = 183&

Private mutexHID As Long

Private Function CreateNamedMutex(ByRef mutexName As String) As Boolean
    Dim sa As SECURITY_ATTRIBUTES
    
    With sa
        .bInheritHandle = 0
        .lpSecurityDescriptor = 0
        .nLength = LenB(sa)
    End With
    
    mutexHID = CreateMutex(sa, False, "Global\" & mutexName)
    
    CreateNamedMutex = Not (Err.LastDllError = ERROR_ALREADY_EXISTS) 'check if the mutex already existed
End Function


Public Function AoDefMultiClient() As Boolean
' UniqueNameThatActuallyCouldBeAnything
   
    If CreateNamedMutex("AoDefenderAntiMultiClient") Then

        AoDefMultiClient = False
    Else

        AoDefMultiClient = True
    End If
End Function
Public Sub AoDefMultiClientOff()

    Call ReleaseMutex(mutexHID)
    Call CloseHandle(mutexHID)
End Sub

Public Sub AoDefMultiClientOn()
    MsgBox "No se puede ejecutar mas de un cliente a la vez.", vbCritical, "Atencion!"
End Sub


