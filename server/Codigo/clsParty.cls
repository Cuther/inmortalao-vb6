VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private p_members(1 To Grupo_MAXMEMBERS) As Integer
Private p_Fundador As Integer
Private p_CantMiembros As Integer
Private g_DPL As Integer 'Division per levels

Public Sub Class_Initialize()
    p_CantMiembros = 0
    g_DPL = 0
End Sub
Public Sub UpdateLevels()
    Dim i As Long
    Dim U As Integer
    
    g_DPL = 0
    For i = 1 To Grupo_MAXMEMBERS
        U = p_members(i)
        If U > 0 Then
            If sAbs(CLng(UserList(p_members(1)).Stats.ELV) - CLng(UserList(U).Stats.ELV)) <= 5 Then
                g_DPL = g_DPL + 1
            End If
        End If
    Next i
End Sub
Public Sub ObtenerExito(ByVal ExpGanada As Long, ByVal mapa As Integer, x As Integer, Y As Integer)
Dim i As Integer
Dim UI As Integer
Dim expRep As Long

expRep = Round(ExpGanada / g_DPL)
For i = 1 To Grupo_MAXMEMBERS
    UI = p_members(i)
    If UI > 0 Then
        If mapa = UserList(UI).Pos.map And UserList(UI).flags.Muerto = 0 Then
        If Distance(UserList(UI).Pos.x, UserList(UI).Pos.Y, x, Y) <= Grupo_MAXDISTANCIA Then
        If sAbs(UserList(p_members(1)).Stats.ELV - UserList(UI).Stats.ELV) <= MAXGrupoDELTALEVEL Then
                    UserList(UI).Stats.Exp = UserList(UI).Stats.Exp + expRep
                   ' If UserList(UI).Stats.Exp > MAXEXP Then UserList(UI).Stats.Exp = MAXEXP
                    Call CheckUserLevel(UI)
                    Call WriteUpdateUserStats(UI)
                    WriteMsg UI, 21, expRep
        End If
        End If
        End If
    End If
Next i

End Sub

Public Sub MandarMensajeAConsola(ByVal texto As String, ByVal Sender As String)
'feo feo, muy feo acceder a senddata desde aca, pero BUEEEEEEEEEEE...
Dim i As Integer

For i = 1 To Grupo_MAXMEMBERS
    If p_members(i) > 0 Then
        Call WriteConsoleMsg(1, p_members(i), " [" & Sender & "] " & texto, FontTypeNames.FONTTYPE_GRUPO)
    End If
Next i

End Sub

Public Function EsGrupoLeader(ByVal UserIndex As Integer) As Boolean
    EsGrupoLeader = (UserIndex = p_Fundador)
End Function

Public Function NuevoMiembro(ByVal UserIndex As Integer) As Boolean

Dim i As Integer
i = 1
While i <= Grupo_MAXMEMBERS And p_members(i) > 0
    i = i + 1
Wend

If i <= Grupo_MAXMEMBERS Then
    p_members(i) = UserIndex
    NuevoMiembro = True
    p_CantMiembros = p_CantMiembros + 1
    UpdateLevels
Else
    NuevoMiembro = False
End If

End Function

Public Function SaleMiembro(ByVal UserIndex As Integer) As Boolean
Dim i As Integer
Dim j As Integer
i = 1
SaleMiembro = False
While i <= Grupo_MAXMEMBERS And p_members(i) <> UserIndex
    i = i + 1
Wend

If i = 1 Then
    'sale el founder, el grupo se disuelve
    SaleMiembro = True
    Call MandarMensajeAConsola("El lider del grupo lo abandona y este se disuelve.", "Servidor")
    For j = Grupo_MAXMEMBERS To 1 Step -1
        If p_members(j) > 0 Then
            
            Call WriteConsoleMsg(1, p_members(j), " Abandonas el grupo liderado por " & UserList(p_members(1)).Name, FontTypeNames.FONTTYPE_GRUPO)
            
            Call MandarMensajeAConsola(UserList(p_members(j)).Name & " abandona el grupo.", "Servidor")
            
            LimpiarFlags j
            WriteGrupo j
            FlushBuffer j
            
            p_CantMiembros = p_CantMiembros - 1
            UpdateLevels
            p_members(j) = 0
        End If
    Next j
Else
    If i <= Grupo_MAXMEMBERS Then
        Call MandarMensajeAConsola(UserList(p_members(i)).Name & " abandona el grupo.", "Servidor")
        p_CantMiembros = p_CantMiembros - 1
        LimpiarFlags i
        UpdateLevels
        p_members(i) = 0
        CompactMemberList
    End If
End If

End Function

Public Function HacerLeader(ByVal UserIndex As Integer) As Boolean
Dim i As Integer
Dim OldLeader As Integer
Dim oldExp As Double
Dim UserIndexIndex As Integer

UserIndexIndex = 0
HacerLeader = True

For i = 1 To Grupo_MAXMEMBERS
    If p_members(i) > 0 Then
        If p_members(i) = UserIndex Then
            UserIndexIndex = i
        End If
    End If
Next i

If Not HacerLeader Then Exit Function

If UserIndexIndex = 0 Then
    'catastrofe! esto no deberia pasar nunca! pero como es AO.... :p
    Call LogError("INCONSISTENCIA DE PARTIES")
    Call SendData(SendTarget.ToADMINS, 0, PrepareMessageConsoleMsg(1, " Inconsistencia de parties en HACERLEADER (UII = 0), AVISE A UN PROGRAMADOR ESTO ES UNA CATASTROFE!!!!", FontTypeNames.FONTTYPE_GUILD))
    HacerLeader = False
    Exit Function
End If

OldLeader = p_members(1)

p_members(1) = p_members(UserIndexIndex)
p_members(UserIndexIndex) = OldLeader

p_Fundador = p_members(1)

End Function


Public Sub ObtenerMiembrosOnline(ByRef MemberList As String)

Dim i As Integer
MemberList = "Nombre(Exp): "
For i = 1 To Grupo_MAXMEMBERS
    If p_members(i) > 0 Then
        MemberList = MemberList & " - " & UserList(p_members(i)).Name
    End If
Next i

MemberList = MemberList & "."

End Sub

Public Function MiembroLista(ByVal i As Byte) As String
    MiembroLista = UserList(p_members(i)).Name
End Function


Public Function PuedeEntrar(ByVal UserIndex As Integer, ByRef razon As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 09/29/07
'Last Modification By: Lucas Tavolaro Ortiz (Tavo)
' - 09/29/07 There is no level prohibition
'***************************************************
'DEFINE LAS REGLAS DEL JUEGO PARA DEJAR ENTRAR A MIEMBROS
Dim esImperio As Boolean
Dim esNeutral As Boolean
Dim esRepubli As Boolean
Dim MyLevel As Integer
Dim i As Integer
Dim rv As Boolean
Dim UI As Integer

If UserList(UserIndex).faccion.Renegado = 1 Then _
    Call WriteConsoleMsg(1, UserIndex, "Los renegados no pueden integrar en grupos.", FONTTYPE_BROWNB)

rv = True
esImperio = (UserList(UserIndex).faccion.ArmadaReal = 1 Or UserList(UserIndex).faccion.Ciudadano = 1)
esNeutral = (UserList(UserIndex).faccion.FuerzasCaos = 1 Or UserList(UserIndex).faccion.Renegado)
esRepubli = (UserList(UserIndex).faccion.Republicano = 1 Or UserList(UserIndex).faccion.Milicia = 1)
MyLevel = UserList(UserIndex).Stats.ELV

rv = Distancia(UserList(p_members(1)).Pos, UserList(UserIndex).Pos) <= MAXDISTANCIAINGRESOGrupo
If rv Then
    rv = (p_members(Grupo_MAXMEMBERS) = 0)
    If rv Then
        For i = 1 To Grupo_MAXMEMBERS
            UI = p_members(i)
            'pongo los casos que evitarian que pueda entrar
            'aspirante armada en Grupo crimi
            If UI > 0 Then
                If esImperio And (esRene(UI) Or esCaos(UI) Or esMili(UI) Or esRepu(UI)) Then
                    razon = "Los miembros del imperio no entran a un grupo con bandos diferentes."
                    rv = False
                End If
                
                'aspirante caos en Grupo ciuda
                If esNeutral And (esArmada(UI) Or esCiuda(UI) Or esMili(UI) Or esRepu(UI)) Then
                    razon = "Los miembros neutrales no entran a un grupo con bandos diferentes."
                    rv = False
                End If
                
                'aspirante caos en Grupo ciuda
                If esRepubli And (esArmada(UI) Or esCiuda(UI) Or esRene(UI) Or esCaos(UI)) Then
                    razon = "Los miembros republicanos no entran a un grupo con bandos diferentes."
                    rv = False
                End If
            
                If Not rv Then Exit For 'violate una programacion estructurada
            End If
        Next i
    Else
        razon = "La mayor cantidad de miembros es " & Grupo_MAXMEMBERS
    End If
Else
    razon = "Te encuentras muy lejos del fundador."
End If

PuedeEntrar = rv

End Function
Private Sub CompactMemberList()
Dim i As Integer
Dim freeIndex As Integer
i = 1
While i <= Grupo_MAXMEMBERS
    If p_members(i) = 0 And freeIndex = 0 Then
        freeIndex = i
    ElseIf p_members(i) > 0 And freeIndex > 0 Then
        p_members(freeIndex) = p_members(i)
        p_members(i) = 0
        
        'muevo el de la pos i a freeindex
        i = freeIndex
        freeIndex = 0
    End If
    i = i + 1
Wend

End Sub

Public Function CantMiembros() As Integer
CantMiembros = p_CantMiembros
End Function
