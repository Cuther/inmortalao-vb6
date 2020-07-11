Attribute VB_Name = "Trabajo"


Option Explicit

Private Const ENERGIA_TRABAJO_HERRERO As Byte = 2
Private Const ENERGIA_TRABAJO_NOHERRERO As Byte = 6
Public Const CARROMINERO As Integer = 880
'2


Public Sub DoPermanecerOculto(ByVal Userindex As Integer)
'********************************************************
'Autor: Nacho (Integer)
'Last Modif: 28/01/2007
'Chequea si ya debe mostrarse
'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
'********************************************************

UserList(Userindex).Counters.TiempoOculto = UserList(Userindex).Counters.TiempoOculto - 1
If UserList(Userindex).Counters.TiempoOculto <= 0 Then
    
    UserList(Userindex).Counters.TiempoOculto = IntervaloOculto
    If UserList(Userindex).Clase = eClass.Cazador And UserList(Userindex).Stats.UserSkills(eSkill.Ocultarse) > 90 Then
        If UserList(Userindex).Invent.ArmourEqpObjIndex = 648 Or UserList(Userindex).Invent.ArmourEqpObjIndex = 360 Then
            Exit Sub
        End If
    End If
    UserList(Userindex).Counters.TiempoOculto = 0
    UserList(Userindex).flags.Oculto = 0
    If UserList(Userindex).flags.Invisible = 0 Then
        Call WriteConsoleMsg(1, Userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(UserList(Userindex).Char.CharIndex, False))
    End If
End If

Exit Sub

Errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")


End Sub

Public Sub DoOcultarse(ByVal Userindex As Integer)
'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
'Modifique la fórmula y ahora anda bien.
On Error GoTo Errhandler

Dim Suerte As Double
Dim res As Integer
Dim Skill As Integer

Skill = UserList(Userindex).Stats.UserSkills(eSkill.Ocultarse)

Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100

res = RandomNumber(1, 100)

If res <= Suerte Then

    UserList(Userindex).flags.Oculto = 1
    UserList(Userindex).Counters.TiempoOculto = IntervaloOculto
  
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(UserList(Userindex).Char.CharIndex, True))

    Call WriteConsoleMsg(2, Userindex, "¡Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
    Call SubirSkill(Userindex, Ocultarse)
Else
    '[CDT 17-02-2004]
    If Not UserList(Userindex).flags.UltimoMensaje = 4 Then
        Call WriteConsoleMsg(2, Userindex, "¡No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
        UserList(Userindex).flags.UltimoMensaje = 4
    End If
    '[/CDT]
End If

UserList(Userindex).Counters.Ocultando = UserList(Userindex).Counters.Ocultando + 1

Exit Sub

Errhandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub


Public Sub DoNavega(ByVal Userindex As Integer, ByRef Barco As ObjData, ByVal Slot As Integer)

Dim ModNave As Long

If UserList(Userindex).flags.Montando = 1 Then Exit Sub

If UserList(Userindex).flags.Invisible = 1 Or UserList(Userindex).flags.Oculto = 1 Then
    UserList(Userindex).flags.Oculto = 0
    UserList(Userindex).flags.Invisible = 0
    UserList(Userindex).Counters.TiempoOculto = 0
    UserList(Userindex).Counters.Invisibilidad = 0
    Call WriteConsoleMsg(1, Userindex, "Vuelves a ser visible.", FontTypeNames.FONTTYPE_BROWNI)
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(UserList(Userindex).Char.CharIndex, False))
End If
        
ModNave = ModNavegacion(UserList(Userindex).Clase)

If UserList(Userindex).Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
    Call WriteConsoleMsg(1, Userindex, "No tenes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(1, Userindex, "Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

UserList(Userindex).Invent.BarcoObjIndex = UserList(Userindex).Invent.Object(Slot).ObjIndex
UserList(Userindex).Invent.BarcoSlot = Slot

If UserList(Userindex).flags.Navegando = 0 Then
    
    UserList(Userindex).Char.Head = 0
    
    If UserList(Userindex).flags.Muerto = 0 Then
        UserList(Userindex).Char.body = 84
    Else
        UserList(Userindex).Char.body = 87
    End If
    
    UserList(Userindex).flags.Navegando = 1
    
Else
    
    UserList(Userindex).flags.Navegando = 0
    
    If UserList(Userindex).flags.Muerto = 0 Then
        UserList(Userindex).Char.Head = UserList(Userindex).OrigChar.Head
        
        If UserList(Userindex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(Userindex).Char.body = ObjData(UserList(Userindex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(Userindex)
        End If
        
        If UserList(Userindex).Invent.EscudoEqpObjIndex > 0 Then _
            UserList(Userindex).Char.ShieldAnim = ObjData(UserList(Userindex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then _
            UserList(Userindex).Char.WeaponAnim = ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(Userindex).Invent.CascoEqpObjIndex > 0 Then _
            UserList(Userindex).Char.CascoAnim = ObjData(UserList(Userindex).Invent.CascoEqpObjIndex).CascoAnim
    Else
        UserList(Userindex).Char.body = iCuerpoMuerto
        UserList(Userindex).Char.Head = iCabezaMuerto
    End If
End If

Call ChangeUserChar(Userindex, UserList(Userindex).Char.body, UserList(Userindex).Char.Head, UserList(Userindex).Char.heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
Call WriteNavigateToggle(Userindex)

End Sub

Public Function DoEquita(ByVal Userindex As Integer, ByVal Obj As Integer, ByVal Slot As Integer)
Dim ModEqui As Long
Dim SK As Long
Dim MK As Long

If UserList(Userindex).flags.Navegando = 1 Then Exit Function

If UserList(Userindex).flags.Invisible = 1 Or UserList(Userindex).flags.Oculto = 1 Then
    UserList(Userindex).flags.Oculto = 0
    UserList(Userindex).flags.Invisible = 0
    UserList(Userindex).Counters.TiempoOculto = 0
    UserList(Userindex).Counters.Invisibilidad = 0
    Call WriteConsoleMsg(1, Userindex, "Vuelves a ser visible.", FontTypeNames.FONTTYPE_BROWNI)
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(UserList(Userindex).Char.CharIndex, False))
End If



SK = UserList(Userindex).Stats.UserSkills(Equitacion)
MK = ObjData(Obj).MinSkill

If Not ClasePuedeUsarItem(Userindex, Obj) Or Not CheckRazaUsaRopa(Userindex, Obj) Then
    Exit Function
End If



If UserList(Userindex).Clase = Paladin Then
    If Obj = 1342 Or Obj = 1343 Then
        If Not SK > 23 Then
            Call WriteConsoleMsg(2, Userindex, "Para usar esta montura necesitas 24 puntos en equitación.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    ElseIf Obj = 1344 Then
        If Not SK > 37 Then
            Call WriteConsoleMsg(2, Userindex, "Para usar esta montura necesitas 38 puntos en equitación.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    ElseIf Obj = 1346 Or Obj = 1348 Then
        If Not SK > 64 Then
            Call WriteConsoleMsg(2, Userindex, "Para usar esta montura necesitas 65 puntos en equitación.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    ElseIf Obj = 1349 Then
        If Not SK > 69 Then
            Call WriteConsoleMsg(2, Userindex, "Para usar esta montura necesitas 70 puntos en equitación.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    Else

        Call WriteConsoleMsg(2, Userindex, "No puedes usar esta montura.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
ElseIf UserList(Userindex).Clase = Nigromante Then
    If Not SK > MK - 1 Then
        Call WriteConsoleMsg(2, Userindex, "Para usar esta montura necesitas " & MK & " puntos en equitación.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
Else
    If Obj = 1342 Or Obj = 1343 Or Obj = 1344 Or Obj = 1349 Or Obj = 1348 Then
        If Not SK > 41 Then
            Call WriteConsoleMsg(2, Userindex, "Para usar esta montura necesitas 42 puntos en equitación.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    ElseIf Obj = 1345 Or Obj = 1346 Or Obj = 1347 Then
        If Not SK > 69 Then
            Call WriteConsoleMsg(2, Userindex, "Para usar esta montura necesitas 70 puntos en equitación.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    Else
        Call WriteConsoleMsg(2, Userindex, "No puedes usar esta montura.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
End If


If UserList(Userindex).Stats.UserSkills(Equitacion) < ObjData(Obj).MinSkill Then
    Call WriteConsoleMsg(2, Userindex, "Para usar esta montura necesitas " & ObjData(Obj).MinSkill & " puntos en equitación.", FontTypeNames.FONTTYPE_INFO)
    Exit Function
End If

If UserList(Userindex).flags.Montando = 0 Then
    If MapData(UserList(Userindex).Pos.map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).Trigger >= 20 Then
        Exit Function
    End If
End If

UserList(Userindex).Invent.MonturaObjIndex = UserList(Userindex).Invent.Object(Slot).ObjIndex
UserList(Userindex).Invent.MonturaSlot = Slot
 
If UserList(Userindex).flags.Montando = 0 Then
    UserList(Userindex).Char.Head = 0
    If UserList(Userindex).flags.Muerto = 0 Then
        UserList(Userindex).Char.body = ObjData(Obj).Ropaje
    Else
        UserList(Userindex).Char.body = iCuerpoMuerto
        UserList(Userindex).Char.Head = iCabezaMuerto
    End If
    UserList(Userindex).Char.Head = UserList(Userindex).OrigChar.Head
    UserList(Userindex).flags.Montando = 1
    UserList(Userindex).Invent.Object(Slot).Equipped = 1
    Call UpdateUserInv(False, Userindex, UserList(Userindex).Invent.MonturaSlot)
Else
    UserList(Userindex).flags.Montando = 0
    If UserList(Userindex).flags.Muerto = 0 Then
        UserList(Userindex).Char.Head = UserList(Userindex).OrigChar.Head
        If UserList(Userindex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(Userindex).Char.body = ObjData(UserList(Userindex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(Userindex)
        End If
    Else
        UserList(Userindex).Char.body = iCuerpoMuerto
        UserList(Userindex).Char.Head = iCabezaMuerto
        UserList(Userindex).Char.ShieldAnim = NingunEscudo
        UserList(Userindex).Char.WeaponAnim = NingunArma
        UserList(Userindex).Char.CascoAnim = NingunCasco
    End If
    UserList(Userindex).Invent.Object(UserList(Userindex).Invent.MonturaSlot).Equipped = 0
    Call UpdateUserInv(False, Userindex, UserList(Userindex).Invent.MonturaSlot)
        
    UserList(Userindex).Invent.MonturaObjIndex = 0
    UserList(Userindex).Invent.MonturaSlot = 0
End If
 
Call ChangeUserChar(Userindex, UserList(Userindex).Char.body, UserList(Userindex).Char.Head, UserList(Userindex).Char.heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
Call WriteEquitateToggle(Userindex)
 
End Function
 
Public Function TieneObjetos(ByVal ItemIndex As Integer, ByVal Cant As Long, ByVal Userindex As Integer) As Boolean
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(Userindex).Invent.Object(i).ObjIndex = ItemIndex Then
        Total = Total + UserList(Userindex).Invent.Object(i).Amount
    End If
Next i

If Cant <= Total Then
    TieneObjetos = True
    Exit Function
End If
        
End Function
Function Have_Obj_To_Slot(ByVal ItemIndex As Integer, ByVal Slot As Byte, ByVal Userindex As Integer) As Long
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long
For i = 1 To Slot
    If UserList(Userindex).Invent.Object(i).ObjIndex = ItemIndex Then
        Have_Obj_To_Slot = Have_Obj_To_Slot + UserList(Userindex).Invent.Object(i).Amount
    End If
Next i
        
End Function
Function Have_Obj_Slot(ByVal ItemIndex As Integer, ByVal Userindex As Integer) As Integer
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(Userindex).Invent.Object(i).ObjIndex = ItemIndex Then
        Have_Obj_Slot = i
    End If
Next i
        
End Function
Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal Cant As Integer, ByVal Userindex As Integer) As Boolean
'Call LogTarea("Sub QuitarObjetos")

Dim i As Integer
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(Userindex).Invent.Object(i).ObjIndex = ItemIndex Then
        
        Call Desequipar(Userindex, i)
        
        UserList(Userindex).Invent.Object(i).Amount = UserList(Userindex).Invent.Object(i).Amount - Cant
        If (UserList(Userindex).Invent.Object(i).Amount <= 0) Then
            Cant = Abs(UserList(Userindex).Invent.Object(i).Amount)
            UserList(Userindex).Invent.Object(i).Amount = 0
            UserList(Userindex).Invent.Object(i).ObjIndex = 0
        Else
            Cant = 0
        End If
        
        Call UpdateUserInv(False, Userindex, i)
        
        If (Cant = 0) Then
            QuitarObjetos = True
            Exit Function
        End If
    End If
Next i

End Function

Sub HerreroQuitarMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer)
    If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH, Userindex)
    If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP, Userindex)
    If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO, Userindex)
End Sub

Sub carpinteroQuitarMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer)
    If ObjData(ItemIndex).Madera > 0 Then
        Call QuitarObjetos(Leña, ObjData(ItemIndex).Madera, Userindex)
    End If
End Sub


Sub druidaQuitarMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer)
    If ObjData(ItemIndex).raies > 0 Then
        Call QuitarObjetos(Raiz, ObjData(ItemIndex).raies, Userindex)
    End If
End Sub

Sub SastreQuitarMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer)
    If ObjData(ItemIndex).PielLobo Then _
        Call QuitarObjetos(PielLobo, ObjData(ItemIndex).PielLobo, Userindex)
        
    If ObjData(ItemIndex).PielOso Then _
        Call QuitarObjetos(PielOso, ObjData(ItemIndex).PielOso, Userindex)
    
    If ObjData(ItemIndex).PielLoboInvernal > 0 Then _
        Call QuitarObjetos(PielLoboInvernal, ObjData(ItemIndex).PielLoboInvernal, Userindex)
End Sub

Function CarpinteroTieneMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer) As Boolean
    
    If ObjData(ItemIndex).Madera > 0 Then
        If Not TieneObjetos(Leña, CLng(ObjData(ItemIndex).Madera) * CLng(Cant), Userindex) Then
            Call WriteConsoleMsg(1, Userindex, "No tenes suficientes madera.", FontTypeNames.FONTTYPE_INFO)
            CarpinteroTieneMateriales = False
            Exit Function
    End If
End If
    
    CarpinteroTieneMateriales = True

End Function

Function druidaTieneMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer) As Boolean
    
    If ObjData(ItemIndex).raies > 0 Then
        If Not TieneObjetos(Raiz, CLng(ObjData(ItemIndex).raies) * CLng(Cant), Userindex) Then
            Call WriteConsoleMsg(1, Userindex, "No tenes suficientes raices.", FontTypeNames.FONTTYPE_INFO)
            druidaTieneMateriales = False
            Exit Function
    End If
End If
    
    druidaTieneMateriales = True

End Function



Function SastreTieneMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer) As Boolean
    If ObjData(ItemIndex).PielLobo Then
        If TieneObjetos(PielLobo, CLng(ObjData(ItemIndex).PielLobo) * CLng(Cant), Userindex) = False Then
            Call WriteConsoleMsg(1, Userindex, "No tenes suficientes pieles.", FontTypeNames.FONTTYPE_INFO)
            SastreTieneMateriales = False
            Exit Function
        End If
    End If

    If ObjData(ItemIndex).PielOso Then
        If TieneObjetos(PielOso, CLng(ObjData(ItemIndex).PielOso) * CLng(Cant), Userindex) = False Then
            Call WriteConsoleMsg(1, Userindex, "No tenes suficientes pieles.", FontTypeNames.FONTTYPE_INFO)
            SastreTieneMateriales = False
            Exit Function
        End If
    End If
    
    If ObjData(ItemIndex).PielLoboInvernal Then
        If TieneObjetos(PielLoboInvernal, CLng(ObjData(ItemIndex).PielLoboInvernal) * CLng(Cant), Userindex) = False Then
            Call WriteConsoleMsg(1, Userindex, "No tenes suficientes pieles.", FontTypeNames.FONTTYPE_INFO)
            SastreTieneMateriales = False
            Exit Function
        End If
    End If
    
    SastreTieneMateriales = True

End Function

 
Function HerreroTieneMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer) As Boolean
    If ObjData(ItemIndex).LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH, Userindex) Then
                    Call WriteConsoleMsg(1, Userindex, "No tenes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingP > 0 Then
            If Not TieneObjetos(LingotePlata, ObjData(ItemIndex).LingP, Userindex) Then
                    Call WriteConsoleMsg(1, Userindex, "No tenes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingO > 0 Then
            If Not TieneObjetos(LingoteOro, ObjData(ItemIndex).LingO, Userindex) Then
                    Call WriteConsoleMsg(1, Userindex, "No tenes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    HerreroTieneMateriales = True
End Function

Public Function PuedeConstruir(ByVal Userindex As Integer, ByVal ItemIndex As Integer) As Boolean
PuedeConstruir = HerreroTieneMateriales(Userindex, ItemIndex) And UserList(Userindex).Stats.UserSkills(eSkill.Herreria) >= _
 ObjData(ItemIndex).SkHerreria
End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ArmasHerrero)
    If ArmasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
For i = 1 To UBound(ArmadurasHerrero)
    If ArmadurasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
PuedeConstruirHerreria = False
End Function


Public Sub HerreroConstruirItem(ByVal Userindex As Integer, ByVal ItemIndex As Integer)

If PuedeConstruir(Userindex, ItemIndex) And PuedeConstruirHerreria(ItemIndex) Then
    
    'Sacamos energía
    If UserList(Userindex).Clase = eClass.Herrero Then
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(Userindex).Stats.MinSTA >= ENERGIA_TRABAJO_HERRERO Then
            UserList(Userindex).Stats.MinSTA = UserList(Userindex).Stats.MinSTA - ENERGIA_TRABAJO_HERRERO
            Call WriteUpdateSta(Userindex)
        Else
            Call WriteConsoleMsg(1, Userindex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(Userindex).Stats.MinSTA >= ENERGIA_TRABAJO_NOHERRERO Then
            UserList(Userindex).Stats.MinSTA = UserList(Userindex).Stats.MinSTA - ENERGIA_TRABAJO_NOHERRERO
            Call WriteUpdateSta(Userindex)
        Else
            Call WriteConsoleMsg(1, Userindex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    Call HerreroQuitarMateriales(Userindex, ItemIndex)
    ' AGREGAR FX
    If ObjData(ItemIndex).OBJType = eOBJType.otWeapon Then
        Call WriteConsoleMsg(1, Userindex, "Has construido el arma!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otESCUDO Then
        Call WriteConsoleMsg(1, Userindex, "Has construido el escudo!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otCASCO Then
        Call WriteConsoleMsg(1, Userindex, "Has construido el casco!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otArmadura Then
        Call WriteConsoleMsg(1, Userindex, "Has construido la armadura!.", FontTypeNames.FONTTYPE_INFO)
    End If
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(Userindex, MiObj) Then
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    End If
    
    Call SubirSkill(Userindex, Herreria)
    Call UpdateUserInv(True, Userindex, 0)
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(MARTILLOHERRERO, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
    UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1
End If
End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjCarpintero)
    If ObjCarpintero(i) = ItemIndex Then
        PuedeConstruirCarpintero = True
        Exit Function
    End If
Next i
PuedeConstruirCarpintero = False

End Function


Public Function PuedeConstruirDruida(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjDruida)
    If ObjDruida(i) = ItemIndex Then
        PuedeConstruirDruida = True
        Exit Function
    End If
Next i
PuedeConstruirDruida = False

End Function


Public Function PuedeConstruirSastre(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjSastre)
    If ObjSastre(i) = ItemIndex Then
        PuedeConstruirSastre = True
        Exit Function
    End If
Next i
PuedeConstruirSastre = False

End Function


Public Sub CarpinteroConstruirItem(ByVal Userindex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer)

If CarpinteroTieneMateriales(Userindex, ItemIndex, Cant) And _
   UserList(Userindex).Stats.UserSkills(eSkill.Carpinteria) >= _
   ObjData(ItemIndex).SkCarpinteria And _
   PuedeConstruirCarpintero(ItemIndex) And _
   UserList(Userindex).Invent.AnilloEqpObjIndex = SERRUCHO_CARPINTERO Then
   
    'Sacamos energía
    If UserList(Userindex).Clase = eClass.Carpintero Then
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(Userindex).Stats.MinSTA >= ENERGIA_TRABAJO_HERRERO Then
            UserList(Userindex).Stats.MinSTA = UserList(Userindex).Stats.MinSTA - ENERGIA_TRABAJO_HERRERO
            Call WriteUpdateSta(Userindex)
        Else
            Call WriteConsoleMsg(1, Userindex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(Userindex).Stats.MinSTA >= ENERGIA_TRABAJO_NOHERRERO Then
            UserList(Userindex).Stats.MinSTA = UserList(Userindex).Stats.MinSTA - ENERGIA_TRABAJO_NOHERRERO
            Call WriteUpdateSta(Userindex)
        Else
            Call WriteConsoleMsg(1, Userindex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    Call carpinteroQuitarMateriales(Userindex, ItemIndex, Cant)
    Call WriteConsoleMsg(1, Userindex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
    
    Dim MiObj As Obj
    MiObj.Amount = Cant
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(Userindex, MiObj) Then
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    End If

    Call SubirSkill(Userindex, Carpinteria)
    Call UpdateUserInv(True, Userindex, 0)
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))


    UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

End If
End Sub


Public Sub druidaConstruirItem(ByVal Userindex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer)

If druidaTieneMateriales(Userindex, ItemIndex, Cant) And _
   UserList(Userindex).Stats.UserSkills(eSkill.alquimia) >= _
   ObjData(ItemIndex).SkPociones And _
   PuedeConstruirDruida(ItemIndex) And _
   UserList(Userindex).Invent.AnilloEqpObjIndex = OLLA Then
   
    'Sacamos energía
    If UserList(Userindex).Clase = eClass.Druida Then
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(Userindex).Stats.MinSTA >= ENERGIA_TRABAJO_HERRERO Then
            UserList(Userindex).Stats.MinSTA = UserList(Userindex).Stats.MinSTA - ENERGIA_TRABAJO_HERRERO
            Call WriteUpdateSta(Userindex)
        Else
            Call WriteConsoleMsg(1, Userindex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(Userindex).Stats.MinSTA >= ENERGIA_TRABAJO_NOHERRERO Then
            UserList(Userindex).Stats.MinSTA = UserList(Userindex).Stats.MinSTA - ENERGIA_TRABAJO_NOHERRERO
            Call WriteUpdateSta(Userindex)
        Else
            Call WriteConsoleMsg(1, Userindex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    Call druidaQuitarMateriales(Userindex, ItemIndex, Cant)
    Call WriteConsoleMsg(1, Userindex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
    
    Dim MiObj As Obj
    MiObj.Amount = Cant
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(Userindex, MiObj) Then
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    End If

    Call SubirSkill(Userindex, alquimia)
    Call UpdateUserInv(True, Userindex, 0)
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))


    UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

End If
End Sub



Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer
    Select Case Lingote
        Case iMinerales.HierroCrudo
            MineralesParaLingote = 14 * ModTrabajo
        Case iMinerales.PlataCruda
            MineralesParaLingote = 20 * ModTrabajo
        Case iMinerales.OroCrudo
            MineralesParaLingote = 35 * ModTrabajo
        Case Else
            MineralesParaLingote = 10000
    End Select
End Function


Public Sub DoLingotes(ByVal Userindex As Integer)

On Error GoTo hayerror

    Dim Slot As Integer
    Dim obji As Integer

    Slot = UserList(Userindex).flags.TargetObjInvSlot
    obji = UserList(Userindex).Invent.Object(Slot).ObjIndex
    
    If UserList(Userindex).Invent.Object(Slot).Amount < MineralesParaLingote(obji) Or _
        ObjData(obji).OBJType <> eOBJType.otMinerales Then
            Call WriteConsoleMsg(1, Userindex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
    End If
    
    UserList(Userindex).Invent.Object(Slot).Amount = UserList(Userindex).Invent.Object(Slot).Amount - MineralesParaLingote(obji)
    If UserList(Userindex).Invent.Object(Slot).Amount < 1 Then
        UserList(Userindex).Invent.Object(Slot).Amount = 0
        UserList(Userindex).Invent.Object(Slot).ObjIndex = 0
    End If
    
    Dim nPos As WorldPos
    Dim MiObj As Obj
    MiObj.Amount = 1 * ModTrabajo
    MiObj.ObjIndex = ObjData(UserList(Userindex).flags.TargetObjInvIndex).LingoteIndex
    If Not MeterItemEnInventario(Userindex, MiObj) Then
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    End If
    Call UpdateUserInv(False, Userindex, Slot)
    Call WriteConsoleMsg(1, Userindex, "¡Has obtenido un lingote!", FontTypeNames.FONTTYPE_INFO)

    UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

Exit Sub

hayerror:

LogError ("Error en dolingotes: " & err.Number & "desc: " & err.description)

End Sub

Function ModNavegacion(ByVal Clase As eClass) As Single

Select Case Clase
    Case eClass.Mercenario
        ModNavegacion = 1
    Case eClass.Pescador
        ModNavegacion = 1.2
    Case Else
        ModNavegacion = 2.3
End Select

End Function


Function ModFundicion(ByVal Clase As eClass) As Single

Select Case Clase
    Case eClass.Minero
        ModFundicion = 1
    Case eClass.Herrero
        ModFundicion = 1.2
    Case Else
        ModFundicion = 3
End Select

End Function

Function ModCarpinteria(ByVal Clase As eClass) As Integer

Select Case Clase
    Case eClass.Carpintero
        ModCarpinteria = 1
    Case Else
        ModCarpinteria = 3
End Select

End Function


Function Modalquimia(ByVal Clase As eClass) As Integer

Select Case Clase
    Case eClass.Druida
        Modalquimia = 1
    Case Else
        Modalquimia = 3
End Select

End Function


Function ModSastreria(ByVal Clase As eClass) As Integer

Select Case Clase
    Case eClass.Sastre
        ModSastreria = 1
    Case Else
        ModSastreria = 3
End Select

End Function


Function ModHerreriA(ByVal Clase As eClass) As Single
Select Case Clase
    Case eClass.Herrero
        ModHerreriA = 1
    Case eClass.Minero
        ModHerreriA = 1.2
    Case Else
        ModHerreriA = 4
End Select

End Function

Function ModDomar(ByVal Clase As eClass) As Integer
    Select Case Clase
        Case eClass.Druida
            ModDomar = 6
        Case eClass.Cazador
            ModDomar = 6
        Case eClass.Clerigo
            ModDomar = 7
        Case Else
            ModDomar = 10
    End Select
End Function

Function FreeMascotaIndex(ByVal Userindex As Integer) As Integer
'***************************************************
'Author: Unknown
'Last Modification: 02/03/09
'02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
'***************************************************
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(Userindex).MascotasType(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function
        End If
    Next j
End Function

Sub DoDomar(ByVal Userindex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Nacho (Integer)
'Last Modification: 02/03/2009
'12/15/2008: ZaMa - Limits the number of the same type of pet to 2.
'02/03/2009: ZaMa - Las criaturas domadas en zona segura, esperan afuera (desaparecen).
'***************************************************

Dim puntosDomar As Integer
Dim puntosRequeridos As Integer
Dim CanStay As Boolean
Dim petType As Integer
Dim NroPets As Integer


If Npclist(NpcIndex).MaestroUser = Userindex Then
    Call WriteConsoleMsg(1, Userindex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(Userindex).NroMascotas < MAXMASCOTAS Then
    
    If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
        Call WriteConsoleMsg(1, Userindex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not PuedeDomarMascota(Userindex, NpcIndex) Then
        Call WriteConsoleMsg(1, Userindex, "No puedes domar mas de dos criaturas del mismo tipo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    puntosDomar = CInt(UserList(Userindex).Stats.UserAtributos(eAtributos.Carisma)) * CInt(UserList(Userindex).Stats.UserSkills(eSkill.Domar))
    puntosRequeridos = Npclist(NpcIndex).flags.Domable
    
    If puntosRequeridos <= puntosDomar And RandomNumber(1, 5) = 1 Then
        Dim index As Integer
        UserList(Userindex).NroMascotas = UserList(Userindex).NroMascotas + 1
        index = FreeMascotaIndex(Userindex)
        UserList(Userindex).MascotasIndex(index) = NpcIndex
        UserList(Userindex).MascotasType(index) = Npclist(NpcIndex).Numero
        
        Npclist(NpcIndex).MaestroUser = Userindex
        
        Call FollowAmo(NpcIndex)
        Call ReSpawnNpc(Npclist(NpcIndex))
        
        Call WriteConsoleMsg(1, Userindex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)
        
        ' Es zona segura?
        CanStay = (MapInfo(UserList(Userindex).Pos.map).Pk = True)
        
        If Not CanStay Then
            petType = Npclist(NpcIndex).Numero
            NroPets = UserList(Userindex).NroMascotas
            
            Call QuitarNPC(NpcIndex)
            
            UserList(Userindex).MascotasType(index) = petType
            UserList(Userindex).NroMascotas = NroPets
            
            Call WriteConsoleMsg(1, Userindex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
        End If

    Else
        If Not UserList(Userindex).flags.UltimoMensaje = 5 Then
            Call WriteConsoleMsg(1, Userindex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
            UserList(Userindex).flags.UltimoMensaje = 5
        End If
    End If
    
    'Entreno domar. Es un 30% más dificil si no sos druida.
    If UserList(Userindex).Clase = eClass.Druida Or (RandomNumber(1, 3) < 3) Then
        Call SubirSkill(Userindex, Domar)
    End If
Else
    Call WriteConsoleMsg(1, Userindex, "No puedes controlar más criaturas.", FontTypeNames.FONTTYPE_INFO)
End If
End Sub

''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal Userindex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: ZaMa
'This function checks how many NPCs of the same type have
'been tamed by the user.
'Returns True if that amount is less than two.
'***************************************************
    Dim i As Long
    Dim numMascotas As Long
    
    For i = 1 To MAXMASCOTAS
        If UserList(Userindex).MascotasType(i) = Npclist(NpcIndex).Numero Then
            numMascotas = numMascotas + 1
        End If
    Next i
    
    If numMascotas <= 1 Then PuedeDomarMascota = True
    
End Function

Sub DoAdminInvisible(ByVal Userindex As Integer)
    
    If UserList(Userindex).flags.AdminInvisible = 0 Then
        UserList(Userindex).flags.AdminInvisible = 1
        UserList(Userindex).flags.Invisible = 1
        UserList(Userindex).flags.Oculto = 1
        UserList(Userindex).flags.OldBody = UserList(Userindex).Char.body
        UserList(Userindex).flags.OldHead = UserList(Userindex).Char.Head
        UserList(Userindex).Char.body = 0
        UserList(Userindex).Char.Head = 0
    Else
        UserList(Userindex).flags.AdminInvisible = 0
        UserList(Userindex).flags.Invisible = 0
        UserList(Userindex).flags.Oculto = 0
        UserList(Userindex).Counters.TiempoOculto = 0
        UserList(Userindex).Char.body = UserList(Userindex).flags.OldBody
        UserList(Userindex).Char.Head = UserList(Userindex).flags.OldHead
    End If
    
    'vuelve a ser visible por la fuerza
    Call ChangeUserChar(Userindex, UserList(Userindex).Char.body, UserList(Userindex).Char.Head, UserList(Userindex).Char.heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(UserList(Userindex).Char.CharIndex, False))
End Sub

Sub TratarDeHacerFogata(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj
Dim posMadera As WorldPos

If Not LegalPos(map, X, Y) Then Exit Sub

With posMadera
    .map = map
    .X = X
    .Y = Y
End With

If MapData(map, X, Y).ObjInfo.ObjIndex <> 58 Then
    Call WriteConsoleMsg(1, Userindex, "Necesitas clickear sobre Leña para hacer ramitas", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Distancia(posMadera, UserList(Userindex).Pos) > 2 Then
    Call WriteConsoleMsg(1, Userindex, "Estás demasiado lejos para prender la fogata.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(Userindex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(1, Userindex, "No puedes hacer fogatas estando muerto.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If MapData(map, X, Y).ObjInfo.Amount < 3 Then
    Call WriteConsoleMsg(1, Userindex, "Necesitas por lo menos tres troncos para hacer una fogata.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If


If UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) < 6 Then
    Suerte = 3
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 6 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 34 Then
    Suerte = 2
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 35 Then
    Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    Obj.ObjIndex = FOGATA_APAG
    Obj.Amount = MapData(map, X, Y).ObjInfo.Amount \ 3
    
    Call WriteConsoleMsg(1, Userindex, "Has hecho " & Obj.Amount & " fogatas.", FontTypeNames.FONTTYPE_INFO)
    
    Call MakeObj(Obj, map, X, Y)
    
    'Seteamos la fogata como el nuevo TargetObj del user
    UserList(Userindex).flags.TargetObj = FOGATA_APAG
Else
    '[CDT 17-02-2004]
    If Not UserList(Userindex).flags.UltimoMensaje = 10 Then
        Call WriteConsoleMsg(1, Userindex, "No has podido hacer la fogata.", FontTypeNames.FONTTYPE_INFO)
        UserList(Userindex).flags.UltimoMensaje = 10
    End If
    '[/CDT]
End If

Call SubirSkill(Userindex, Supervivencia)


End Sub

Public Sub DoPescar(ByVal Userindex As Integer, Optional ByVal red As Boolean)
On Error GoTo Errhandler

Dim Suerte As Integer
Dim res As Integer

If UserList(Userindex).Clase = eClass.Pescador Then
    Call QuitarSta(Userindex, EsfuerzoPescarPescador)
Else
    Call QuitarSta(Userindex, EsfuerzoPescarGeneral)
End If

Dim Skill As Integer
Skill = UserList(Userindex).Stats.UserSkills(eSkill.Pesca)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res <= 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    Dim Pez As Integer
    
    MiObj.Amount = ModTrabajo + IIf(red, RandomNumber(7, 15) * ModTrabajo, 0)
    If HayAgua(UserList(Userindex).Pos.map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y) Then
        If UCase$(Left$(Tilde(MapInfo(UserList(Userindex).Pos.map).Name), 14)) = "OCEANO ABIERTO" Then
            If UserList(Userindex).Clase = eClass.Pescador Then
                Pez = 900
            Else
                Pez = 545
            End If
        Else
            Pez = 546
        End If
    Else
        Pez = 139
    End If
    
    MiObj.ObjIndex = Pez
    
    If Not MeterItemEnInventario(Userindex, MiObj) Then
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    End If
    
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_PESCAR, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
Else
    '[CDT 17-02-2004]
    If Not UserList(Userindex).flags.UltimoMensaje = 6 Then
      Call WriteConsoleMsg(2, Userindex, "¡No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
      UserList(Userindex).flags.UltimoMensaje = 6
    End If
    '[/CDT]
End If

Call SubirSkill(Userindex, Pesca)

UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

Exit Sub

Errhandler:
    Call LogError("Error en DoPescar")
End Sub


''
' Try to steal an item / gold to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 24/07/028
'Last Modification By: Marco Vanotti (MarKoxX)
' - 24/07/08 Now it calls to WriteUpdateGold(VictimaIndex and LadrOnIndex) when the thief stoles gold. (MarKoxX)
'*************************************************

If Not MapInfo(UserList(VictimaIndex).Pos.map).Pk Then Exit Sub

If UserList(LadrOnIndex).flags.Seguro Then
    Call WriteConsoleMsg(1, LadrOnIndex, "Debes quitar el seguro para robar", FontTypeNames.FONTTYPE_FIGHT)
    Exit Sub
End If

If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

If UserList(VictimaIndex).Faccion.FuerzasCaos = 1 And UserList(LadrOnIndex).Faccion.FuerzasCaos = 1 Then
    Call WriteConsoleMsg(1, LadrOnIndex, "No puedes robar a otros miembros de las fuerzas del caos", FontTypeNames.FONTTYPE_FIGHT)
    Exit Sub
End If


Call QuitarSta(LadrOnIndex, 15)

Dim GuantesHurto As Boolean

If Not PuedeRobar(LadrOnIndex, VictimaIndex) Then Exit Sub

If UserList(VictimaIndex).flags.Privilegios And (PlayerType.User Or PlayerType.VIP) Then
    Dim Suerte As Integer
    Dim res As Integer
    
    If UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 10 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= -1 Then
                        Suerte = 35
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 20 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 11 Then
                        Suerte = 30
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 30 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 21 Then
                        Suerte = 28
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 40 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 31 Then
                        Suerte = 24
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 50 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 41 Then
                        Suerte = 22
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 60 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 51 Then
                        Suerte = 20
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 70 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 61 Then
                        Suerte = 18
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 80 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 71 Then
                        Suerte = 15
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 90 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 81 Then
                        Suerte = 10
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) < 100 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 91 Then
                        Suerte = 7
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) = 100 Then
                        Suerte = 5
    End If
    res = RandomNumber(1, Suerte)
        
    If res < 3 Then 'Exito robo
       
        If (RandomNumber(1, 50) < 25) And (UserList(LadrOnIndex).Clase = eClass.Ladron) Then
            If TieneObjetosRobables(VictimaIndex) Then
                Call RobarObjeto(LadrOnIndex, VictimaIndex)
            Else
                Call WriteConsoleMsg(2, LadrOnIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else 'Roba oro
            If UserList(VictimaIndex).Stats.GLD > 0 Then
                Dim N As Integer
                
                If UserList(LadrOnIndex).Clase = eClass.Ladron Then
                    N = RandomNumber(100, 1000)
                Else
                    N = RandomNumber(1, 100)
                End If
                If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                
                UserList(LadrOnIndex).Stats.GLD = UserList(LadrOnIndex).Stats.GLD + N
                If UserList(LadrOnIndex).Stats.GLD > MAXORO Then _
                    UserList(LadrOnIndex).Stats.GLD = MAXORO
                
                Call WriteConsoleMsg(2, LadrOnIndex, "Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name, FontTypeNames.FONTTYPE_INFO)
                Call WriteUpdateGold(LadrOnIndex) 'Le actualizamos la billetera al ladron
                
                Call WriteUpdateGold(VictimaIndex) 'Le actualizamos la billetera a la victima
                Call FlushBuffer(VictimaIndex)
            Else
                Call WriteConsoleMsg(2, LadrOnIndex, UserList(VictimaIndex).Name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    Else
        Call WriteConsoleMsg(2, LadrOnIndex, "¡No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(1, VictimaIndex, "¡" & UserList(LadrOnIndex).Name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
        Call FlushBuffer(VictimaIndex)
    End If
    
    Call SubirSkill(LadrOnIndex, Robar)
End If


End Sub

''
' Check if one item is stealable
'
' @param VictimaIndex Specifies reference to victim
' @param Slot Specifies reference to victim's inventory slot
' @return If the item is stealable
Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
' Agregué los barcos
' Esta funcion determina qué objetos son robables.

Dim OI As Integer

OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

ObjEsRobable = _
ObjData(OI).OBJType <> eOBJType.otLlaves And _
UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And _
ObjData(OI).Real = 0 And _
ObjData(OI).Caos = 0 And _
ObjData(OI).OBJType <> eOBJType.otMonturas And _
ObjData(OI).OBJType <> eOBJType.otBarcos

End Function

''
' Try to steal an item to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen
Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'Call LogTarea("Sub RobarObjeto")
Dim flag As Boolean
Dim i As Integer
flag = False

If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
    i = 1
    Do While Not flag And i <= MAX_INVENTORY_SLOTS
        'Hay objeto en este slot?
        If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
           If ObjEsRobable(VictimaIndex, i) Then
                 If RandomNumber(1, 10) < 4 Then flag = True
           End If
        End If
        If Not flag Then i = i + 1
    Loop
Else
    i = 20
    Do While Not flag And i > 0
      'Hay objeto en este slot?
      If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
         If ObjEsRobable(VictimaIndex, i) Then
               If RandomNumber(1, 10) < 4 Then flag = True
         End If
      End If
      If Not flag Then i = i - 1
    Loop
End If

If flag Then
    Dim MiObj As Obj
    Dim num As Byte
    'Cantidad al azar
    num = RandomNumber(1, 5)
                
    If num > UserList(VictimaIndex).Invent.Object(i).Amount Then
         num = UserList(VictimaIndex).Invent.Object(i).Amount
    End If
                
    MiObj.Amount = num
    MiObj.ObjIndex = UserList(VictimaIndex).Invent.Object(i).ObjIndex
    
    UserList(VictimaIndex).Invent.Object(i).Amount = UserList(VictimaIndex).Invent.Object(i).Amount - num
                
    If UserList(VictimaIndex).Invent.Object(i).Amount <= 0 Then
          Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)
    End If
            
    Call UpdateUserInv(False, VictimaIndex, CByte(i))
                
    If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)
    End If
    
    Call WriteConsoleMsg(1, LadrOnIndex, "Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)
Else
    Call WriteConsoleMsg(1, LadrOnIndex, "No has logrado robar ningún objeto.", FontTypeNames.FONTTYPE_INFO)
End If

'If exiting, cancel de quien es robado
Call CancelExit(VictimaIndex)

End Sub

Public Sub DoApuñalar(ByVal Userindex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal Daño As Integer)
'***************************************************
'Autor: Nacho (Integer) & Unknown (orginal version)
'Last Modification: 04/17/08 - (NicoNZ)
'Simplifique la cuenta que hacia para sacar la suerte
'y arregle la cuenta que hacia para sacar el daño
'***************************************************
Dim Suerte As Integer
Dim Skill As Integer

Skill = UserList(Userindex).Stats.UserSkills(eSkill.Apuñalar)

Select Case UserList(Userindex).Clase
    Case eClass.Asesino
        Suerte = Int(((0.00003 * Skill - 0.002) * Skill + 0.098) * Skill + 4.25)
    
    Case eClass.Clerigo, eClass.Paladin
        Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
    
    Case eClass.Bardo
        Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
    
    Case Else
        Suerte = Int(0.0361 * Skill + 4.39)
End Select


If RandomNumber(0, 100) < Suerte Then
    If VictimUserIndex <> 0 Then
        If UserList(Userindex).Clase = eClass.Asesino Then
            Daño = Round(Daño * 1.4, 0)
        Else
            Daño = Round(Daño * 1.5, 0)
        End If
        
        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - Daño
        Call WriteConsoleMsg(2, Userindex, "Has apuñalado a " & UserList(VictimUserIndex).Name & " por " & Daño, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(2, VictimUserIndex, "Te ha apuñalado " & UserList(Userindex).Name & " por " & Daño, FontTypeNames.FONTTYPE_FIGHT)
        
        Call FlushBuffer(VictimUserIndex)
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - Int(Daño * 2)
        Call WriteConsoleMsg(2, Userindex, "Has apuñalado la criatura por " & Int(Daño * 2), FontTypeNames.FONTTYPE_FIGHT)
        Call SubirSkill(Userindex, Apuñalar)
        '[Alejo]
        Call CalcularDarExp(Userindex, VictimNpcIndex, Daño * 2)
    End If
Else
    Call WriteConsoleMsg(2, Userindex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)
End If

End Sub


Public Sub QuitarSta(ByVal Userindex As Integer, ByVal Cantidad As Integer)
    UserList(Userindex).Stats.MinSTA = UserList(Userindex).Stats.MinSTA - Cantidad
    If UserList(Userindex).Stats.MinSTA < 0 Then UserList(Userindex).Stats.MinSTA = 0
    Call WriteUpdateSta(Userindex)
End Sub

Public Sub DoTalar(ByVal Userindex As Integer)
On Error GoTo Errhandler

    Dim Suerte As Integer
    Dim res As Integer
    
    If UserList(Userindex).Clase = eClass.Leñador Then
        Call QuitarSta(Userindex, EsfuerzoTalarLeñador)
    Else
        Call QuitarSta(Userindex, EsfuerzoTalarGeneral)
    End If
    
    Dim Skill As Integer
    Skill = UserList(Userindex).Stats.UserSkills(eSkill.Talar)
    Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
    res = RandomNumber(1, Suerte)
    
    If res <= 6 Then
        Dim nPos As WorldPos
        Dim MiObj As Obj
        
        If UserList(Userindex).Clase = eClass.Leñador Then
            MiObj.Amount = RandomNumber(1, 6) * ModTrabajo
        Else
            MiObj.Amount = RandomNumber(1, 2) * ModTrabajo
        End If
        
        MiObj.ObjIndex = Leña
        
        
        If Not MeterItemEnInventario(Userindex, MiObj) Then
            Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
        End If
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_TALAR, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
    Else
        '[CDT 17-02-2004]
        If Not UserList(Userindex).flags.UltimoMensaje = 8 Then
            Call WriteConsoleMsg(1, Userindex, "¡No has obtenido leña!", FontTypeNames.FONTTYPE_INFO)
            UserList(Userindex).flags.UltimoMensaje = 8
        End If
        '[/CDT]
    End If
    
    Call SubirSkill(Userindex, Talar)
    
    
    UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1
    
    Exit Sub

Errhandler:
    Call LogError("Error en DoTalar")

End Sub
Public Sub DoMineria(ByVal Userindex As Integer)
On Error GoTo Errhandler

Dim Suerte As Integer
Dim res As Integer
Dim metal As Integer

If UserList(Userindex).Clase = eClass.Minero Then
    Call QuitarSta(Userindex, EsfuerzoExcavarMinero)
Else
    Call QuitarSta(Userindex, EsfuerzoExcavarGeneral)
End If

Dim Skill As Integer
Skill = UserList(Userindex).Stats.UserSkills(eSkill.Mineria)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res <= 5 Then
    Dim MiObj As Obj
    Dim nPos As WorldPos
    
    If UserList(Userindex).flags.TargetObj = 0 Then Exit Sub
    
    MiObj.ObjIndex = ObjData(UserList(Userindex).flags.TargetObj).MineralIndex
    
    If UserList(Userindex).Clase = eClass.Minero Then
        MiObj.Amount = RandomNumber(1, 6) * ModTrabajo
    Else
        MiObj.Amount = RandomNumber(1, 2) * ModTrabajo
    End If
    
    If Not MeterItemEnInventario(Userindex, MiObj) Then _
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
        
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_MINERO, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
End If

Call SubirSkill(Userindex, Mineria)


UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

Exit Sub

Errhandler:
    Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoMeditar(ByVal Userindex As Integer)

UserList(Userindex).Counters.IdleCount = 0

Dim Suerte As Integer
Dim res As Integer
Dim Cant As Integer

Dim TActual As Long
TActual = GetTickCount() And &H7FFFFFFF
    
If UserList(Userindex).Stats.MinMAN >= UserList(Userindex).Stats.MaxMAN Then Exit Sub

If UserList(Userindex).Stats.UserSkills(eSkill.Meditar) <= 10 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= -1 Then
                    Suerte = 35
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) <= 20 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= 11 Then
                    Suerte = 30
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) <= 30 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= 21 Then
                    Suerte = 28
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) <= 40 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= 31 Then
                    Suerte = 24
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) <= 50 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= 41 Then
                    Suerte = 22
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) <= 60 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= 51 Then
                    Suerte = 20
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) <= 70 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= 61 Then
                    Suerte = 18
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) <= 80 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= 71 Then
                    Suerte = 15
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) <= 90 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= 81 Then
                    Suerte = 12
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) < 100 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= 91 Then
                    Suerte = 8
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) = 100 Then
                    Suerte = 7
End If

'Mannakia
If UserList(Userindex).Invent.MagicIndex <> 0 Then
    If ObjData(UserList(Userindex).Invent.MagicIndex).EfectoMagico = eMagicType.AceleraMana Then
        Suerte = Suerte - Porcentaje(Suerte, 30)  ' nose algo razonable
    End If
End If
'Mannakia

res = RandomNumber(1, Suerte)

If res = 1 Then
    
    Cant = Porcentaje(UserList(Userindex).Stats.MaxMAN, PorcentajeRecuperoMana)
    
    If Cant <= 0 Then Cant = 1
    UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MinMAN + Cant
    If UserList(Userindex).Stats.MinMAN > UserList(Userindex).Stats.MaxMAN Then _
        UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MaxMAN
    
    Call WriteUpdateMana(Userindex)
    Call SubirSkill(Userindex, Meditar)
End If

End Sub
Public Function DoTrabajar(ByVal Userindex As Integer)
    If Not IntervaloPermiteTrabajar(Userindex, True) Then Exit Function
    
    With UserList(Userindex)
        If .Stats.MinSTA < 2 Then
            Call WriteConsoleMsg(2, Userindex, "Dejas de trabajar.", FontTypeNames.FONTTYPE_BROWNI)
            .flags.Trabajando = False
            Exit Function
        End If
        
        If .flags.Lingoteando Then
            Call DoLingotes(Userindex)
        ElseIf .Invent.AnilloEqpSlot <> 0 Then
            Select Case .Invent.AnilloEqpObjIndex
                Case RED_PESCA
                    Call DoPescar(Userindex, True)
                    
                Case CAÑA_PESCA
                    Call DoPescar(Userindex)
                    
                Case PIQUETE_MINERO
                    Call DoMineria(Userindex)
                  
                Case HACHA_LEÑADOR
                    Call DoTalar(Userindex)
                    
                Case TIJERAS
                    Call DoBotanica(Userindex)
            End Select
        End If
    End With
End Function




Public Sub DoBotanica(ByVal Userindex As Integer)
On Error GoTo Errhandler

Dim Suerte As Integer
Dim res As Integer

If UserList(Userindex).Clase = eClass.Druida Then
    Call QuitarSta(Userindex, EsfuerzoBotanicaDruida)
Else
    Call QuitarSta(Userindex, EsfuerzoBotanicaGeneral)
End If

Dim Skill As Integer
Skill = UserList(Userindex).Stats.UserSkills(eSkill.botanica)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res <= 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    If UserList(Userindex).Clase = eClass.Druida Then
        MiObj.Amount = RandomNumber(1, 6) * ModTrabajo
    Else
        MiObj.Amount = RandomNumber(1, 2) * ModTrabajo
    End If
    
    MiObj.ObjIndex = Raiz
    
    
    If Not MeterItemEnInventario(Userindex, MiObj) Then
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    End If

Else
    '[CDT 17-02-2004]
    If Not UserList(Userindex).flags.UltimoMensaje = 8 Then
        Call WriteConsoleMsg(1, Userindex, "¡No has obtenido raices!", FontTypeNames.FONTTYPE_INFO)
        UserList(Userindex).flags.UltimoMensaje = 8
    End If
    '[/CDT]
End If

Call SubirSkill(Userindex, botanica)


UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

Exit Sub

Errhandler:
    Call LogError("Error en DoTalar")

End Sub






Public Sub SastreConstruirItem(ByVal Userindex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer)

If SastreTieneMateriales(Userindex, ItemIndex, Cant) And _
   UserList(Userindex).Stats.UserSkills(eSkill.Sastreria) >= _
   ObjData(ItemIndex).SkSastreria And _
   PuedeConstruirSastre(ItemIndex) And _
   UserList(Userindex).Invent.AnilloEqpObjIndex = COSTURERO Then
   
    'Sacamos energía
    If UserList(Userindex).Clase = eClass.Sastre Then
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(Userindex).Stats.MinSTA >= ENERGIA_TRABAJO_HERRERO Then
            UserList(Userindex).Stats.MinSTA = UserList(Userindex).Stats.MinSTA - ENERGIA_TRABAJO_HERRERO
            Call WriteUpdateSta(Userindex)
        Else
            Call WriteConsoleMsg(1, Userindex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(Userindex).Stats.MinSTA >= ENERGIA_TRABAJO_NOHERRERO Then
            UserList(Userindex).Stats.MinSTA = UserList(Userindex).Stats.MinSTA - ENERGIA_TRABAJO_NOHERRERO
            Call WriteUpdateSta(Userindex)
        Else
            Call WriteConsoleMsg(1, Userindex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    Call SastreQuitarMateriales(Userindex, ItemIndex, Cant)
    Call WriteConsoleMsg(1, Userindex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
    
    Dim MiObj As Obj
    MiObj.Amount = Cant
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(Userindex, MiObj) Then
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    End If

    Call SubirSkill(Userindex, Sastreria)
    Call UpdateUserInv(True, Userindex, 0)
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

    UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

End If
End Sub
