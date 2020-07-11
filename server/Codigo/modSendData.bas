Attribute VB_Name = "modSendData"

Option Explicit

Public Enum SendTarget
    ToAll = 1
    ToMap
    ToPCArea
    ToAllButIndex
    ToMapButIndex
    ToGM
    ToNPCArea
    ToGuildMembers
    ToADMINS
    ToPCAreaButIndex
    ToDiosesYclan
    ToClanArea
    ToDeadArea
    ToGrupoArea
    
    ToReal
    ToCaos
    Tomili
    
    ToImpeMap
    ToRepuMap
End Enum

Public Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndData As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus) - Rewrite of original
'Last Modify Date: 01/08/2007
'Last modified by: (liquid)
'**************************************************************
On Error GoTo hayerror
    Dim loopC As Long
    Dim map As Integer
    Dim tempIndex As Integer

    
    Select Case sndRoute
        Case SendTarget.ToPCArea
            Call SendToUserArea(sndIndex, sndData)
        
            Exit Sub
        
        Case SendTarget.ToADMINS
            For loopC = 1 To LastUser
                If UserList(loopC).ConnID <> -1 Then
                    If EsDIOS(loopC) Then
                        Call EnviarDatosASlot(loopC, sndData)
                   End If
                End If
            Next loopC
            DoEvents
            Exit Sub
        
        Case SendTarget.ToGM
            For loopC = 1 To LastUser
                If UserList(loopC).ConnID <> -1 Then
                    If EsCONSE(loopC) Then
                        Call EnviarDatosASlot(loopC, sndData)
                   End If
                End If
            Next loopC
            DoEvents
            Exit Sub
        
        Case SendTarget.ToAll
            For loopC = 1 To LastUser
                If UserList(loopC).ConnID <> -1 Then
                    If UserList(loopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call EnviarDatosASlot(loopC, sndData)
                    End If
                End If
            Next loopC
            DoEvents
            Exit Sub
        
        Case SendTarget.ToAllButIndex
            For loopC = 1 To LastUser
                If (UserList(loopC).ConnID <> -1) And (loopC <> sndIndex) Then
                    If UserList(loopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call EnviarDatosASlot(loopC, sndData)
                    End If
                End If
            Next loopC
            DoEvents
            Exit Sub
        
        Case SendTarget.ToMap
            Call SendToMap(sndIndex, sndData)
            Exit Sub
          
        Case SendTarget.ToMapButIndex
            Call SendToMapButIndex(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToGuildMembers
            loopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            While loopC > 0
                If (UserList(loopC).ConnID <> -1) Then
                    Call EnviarDatosASlot(loopC, sndData)
                End If
                loopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend
            DoEvents
            Exit Sub
        
        Case SendTarget.ToDeadArea
            Call SendToDeadUserArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToPCAreaButIndex
            Call SendToUserAreaButindex(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToClanArea
            Call SendToUserGuildArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToGrupoArea
            Call SendToUserGrupoArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToNPCArea
            Call SendToNpcArea(sndIndex, sndData)
            Exit Sub
        
            
        Case SendTarget.ToReal
            For loopC = 1 To LastUser
                If (UserList(loopC).ConnID <> -1) Then
                    If UserList(loopC).flags.UserLogged And (UserList(loopC).faccion.ArmadaReal = 1 Or EsCONSE(loopC)) Then  'Esta logeado como usuario?
                        Call EnviarDatosASlot(loopC, sndData)
                    End If
                End If
            Next loopC
            DoEvents
            Exit Sub
            
        Case SendTarget.Tomili
            For loopC = 1 To LastUser
                If (UserList(loopC).ConnID <> -1) Then
                    If UserList(loopC).flags.UserLogged And (UserList(loopC).faccion.Milicia = 1 Or EsCONSE(loopC)) Then  'Esta logeado como usuario?
                        Call EnviarDatosASlot(loopC, sndData)
                    End If
                End If
            Next loopC
            DoEvents
            Exit Sub
 
        Case SendTarget.ToCaos
            For loopC = 1 To LastUser
                If (UserList(loopC).ConnID <> -1) Then
                    If UserList(loopC).flags.UserLogged And (UserList(loopC).faccion.FuerzasCaos = 1 Or EsCONSE(loopC)) Then    'Esta logeado como usuario?
                        Call EnviarDatosASlot(loopC, sndData)
                    End If
                End If
            Next loopC
            DoEvents
            Exit Sub
        
        'Add Marius =)
        Case SendTarget.ToImpeMap
            If Not MapaValido(sndIndex) Then Exit Sub
        
            For loopC = 1 To ConnGroups(sndIndex).CountEntrys
                tempIndex = ConnGroups(sndIndex).UserEntrys(loopC)
                
                If UserList(tempIndex).ConnIDValida And (esCiuda(tempIndex) Or esArmada(tempIndex)) Then
                    Call EnviarDatosASlot(tempIndex, sndData)
                End If
            Next loopC
            DoEvents
            Exit Sub
        
        Case SendTarget.ToRepuMap
            If Not MapaValido(sndIndex) Then Exit Sub
        
            For loopC = 1 To ConnGroups(sndIndex).CountEntrys
                tempIndex = ConnGroups(sndIndex).UserEntrys(loopC)
                
                If UserList(tempIndex).ConnIDValida And (esRepu(tempIndex) Or esMili(tempIndex)) Then
                    Call EnviarDatosASlot(tempIndex, sndData)
                End If
            Next loopC
            DoEvents
            Exit Sub
        '\Add
 
    End Select
    
    
    
    Exit Sub
    
hayerror:
    LogError ("Error en Senddata: " & err.Number & " desc: " & err.description & " SendTarget:" & sndRoute & " Index:" & sndIndex & " Data:" & sndData)
End Sub

Private Sub SendToUserArea(ByVal UserIndex As Integer, ByVal sdData As String)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim loopC As Long
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = UserList(UserIndex).Pos.map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(map) Then Exit Sub
    
    For loopC = 1 To ConnGroups(map).CountEntrys
        If loopC <= ConnGroups(map).CountEntrys Then
            tempIndex = ConnGroups(map).UserEntrys(loopC)
            
            If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
            
                    If UserList(tempIndex).ConnIDValida Then
                        Call EnviarDatosASlot(tempIndex, sdData)
                    End If

                End If
            End If
        Else
            Exit For
        End If
    Next loopC
    DoEvents
End Sub

Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal sdData As String)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim loopC As Long
    Dim TempInt As Integer
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = UserList(UserIndex).Pos.map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY

    If Not MapaValido(map) Then Exit Sub
    
    For loopC = 1 To ConnGroups(map).CountEntrys
        If loopC <= ConnGroups(map).CountEntrys Then
            tempIndex = ConnGroups(map).UserEntrys(loopC)
                
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
            If TempInt Then  'Esta en el area?
                TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
                If TempInt Then
                    If tempIndex <> UserIndex Then
                        If UserList(tempIndex).ConnIDValida Then
                            Call EnviarDatosASlot(tempIndex, sdData)
                        End If
                    End If
                End If
            End If
        Else
            Exit For
        End If
    Next loopC
    DoEvents
End Sub

Private Sub SendToDeadUserArea(ByVal UserIndex As Integer, ByVal sdData As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim loopC As Long
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = UserList(UserIndex).Pos.map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(map) Then Exit Sub
    
    For loopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(loopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                'Dead and admins read
                If UserList(tempIndex).ConnIDValida = True And (UserList(tempIndex).flags.Muerto = 1 Or (EsDIOS(tempIndex)) <> 0) Then
                    Call EnviarDatosASlot(tempIndex, sdData)
                End If
            End If
        End If
    Next loopC
    DoEvents
End Sub

Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal sdData As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim loopC As Long
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = UserList(UserIndex).Pos.map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(map) Then Exit Sub
    
    If UserList(UserIndex).GuildIndex = 0 Then Exit Sub
    
    For loopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(loopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida And (UserList(tempIndex).GuildIndex = UserList(UserIndex).GuildIndex Or ((EsDIOS(tempIndex)) And (UserList(tempIndex).flags.Privilegios And PlayerType.VIP) = 0)) Then
                    Call EnviarDatosASlot(tempIndex, sdData)
                End If
            End If
        End If
    Next loopC
    DoEvents
End Sub

Private Sub SendToUserGrupoArea(ByVal UserIndex As Integer, ByVal sdData As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim loopC As Long
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = UserList(UserIndex).Pos.map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(map) Then Exit Sub
    
    If UserList(UserIndex).GrupoIndex = 0 Then Exit Sub
    
    For loopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(loopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida And UserList(tempIndex).GrupoIndex = UserList(UserIndex).GrupoIndex Then
                    Call EnviarDatosASlot(tempIndex, sdData)
                End If
            End If
        End If
    Next loopC
    DoEvents
End Sub

Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal sdData As String)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim loopC As Long
    Dim TempInt As Integer
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = Npclist(NpcIndex).Pos.map
    AreaX = Npclist(NpcIndex).AreasInfo.AreaPerteneceX
    AreaY = Npclist(NpcIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(map) Then Exit Sub
    
    For loopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(loopC)
        
        TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
        If TempInt Then  'Esta en el area?
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
                'Add Marius Agregamos el if para ver si asi logramos cagar el bud de ver dobles
                'Solo test, si no funciona hay que sacar!
                If UserList(tempIndex).Pos.map = map Then
                    
                    If TempInt Then
                        If UserList(tempIndex).ConnIDValida Then
                            Call EnviarDatosASlot(tempIndex, sdData)
                        End If
                    End If
            
                End If
                'Add
        End If
    Next loopC
    DoEvents
End Sub

Public Sub SendToAreaByPos(ByVal map As Integer, ByVal AreaX As Integer, ByVal AreaY As Integer, ByVal sdData As String)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim loopC As Long
    Dim TempInt As Integer
    Dim tempIndex As Integer
    
    AreaX = 2 ^ (AreaX \ 9)
    AreaY = 2 ^ (AreaY \ 9)
    
    If Not MapaValido(map) Then Exit Sub

    For loopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(loopC)
            
        TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
        If TempInt Then  'Esta en el area?
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
            If TempInt Then
                If UserList(tempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(tempIndex, sdData)
                End If
            End If
        End If
    Next loopC
    DoEvents
End Sub

Public Sub SendToMap(ByVal map As Integer, ByVal sdData As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 5/24/2007
'
'**************************************************************
    Dim loopC As Long
    Dim tempIndex As Integer
    
    If Not MapaValido(map) Then Exit Sub

    For loopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(loopC)
        
        If UserList(tempIndex).ConnIDValida Then
            Call EnviarDatosASlot(tempIndex, sdData)
        End If
    Next loopC
    DoEvents
End Sub

Public Sub SendToMapButIndex(ByVal UserIndex As Integer, ByVal sdData As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 5/24/2007
'
'**************************************************************
    Dim loopC As Long
    Dim map As Integer
    Dim tempIndex As Integer
    
    map = UserList(UserIndex).Pos.map
    
    If Not MapaValido(map) Then Exit Sub

    For loopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(loopC)
        
        If tempIndex <> UserIndex And UserList(tempIndex).ConnIDValida Then
            Call EnviarDatosASlot(tempIndex, sdData)
        End If
    Next loopC
    DoEvents
End Sub
