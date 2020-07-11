Attribute VB_Name = "Trabajo"


Option Explicit

Private Const ENERGIA_TRABAJO_HERRERO As Byte = 2
Private Const ENERGIA_TRABAJO_NOHERRERO As Byte = 6
Public Const CARROMINERO As Integer = 880
'2


Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)
'********************************************************
'Autor: Nacho (Integer)
'Last Modif: 28/01/2007
'Chequea si ya debe mostrarse
'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
'********************************************************

UserList(UserIndex).Counters.TiempoOculto = UserList(UserIndex).Counters.TiempoOculto - 1
If UserList(UserIndex).Counters.TiempoOculto <= 0 Then
    
    UserList(UserIndex).Counters.TiempoOculto = IntervaloOculto
    If UserList(UserIndex).Clase = eClass.Cazador And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) > 90 Then
        If UserList(UserIndex).Invent.ArmourEqpObjIndex = 648 Or UserList(UserIndex).Invent.ArmourEqpObjIndex = 360 Then
            Exit Sub
        End If
    End If
    UserList(UserIndex).Counters.TiempoOculto = 0
    UserList(UserIndex).flags.Oculto = 0
    If UserList(UserIndex).flags.Invisible = 0 Then
        Call WriteConsoleMsg(1, UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
    End If
End If

Exit Sub

Errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")


End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)
'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
'Modifique la fórmula y ahora anda bien.
On Error GoTo Errhandler

Dim Suerte As Double
Dim res As Integer
Dim Skill As Integer

Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse)

Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100

res = RandomNumber(1, 100)

If res <= Suerte Then

    UserList(UserIndex).flags.Oculto = 1
    UserList(UserIndex).Counters.TiempoOculto = IntervaloOculto
  
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, True))

    Call WriteConsoleMsg(2, UserIndex, "¡Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
    Call SubirSkill(UserIndex, Ocultarse)
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 4 Then
        Call WriteConsoleMsg(2, UserIndex, "¡No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 4
    End If
    '[/CDT]
End If

UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando + 1

Exit Sub

Errhandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub


Public Sub DoNavega(ByVal UserIndex As Integer, ByRef Barco As ObjData, ByVal Slot As Integer)

Dim ModNave As Long

If UserList(UserIndex).flags.Montando = 1 Then Exit Sub

'Add Marius Captura la Bandera
If UserList(UserIndex).Pos.map = Bandera_mapa Then Exit Sub
'\Add

If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then
    UserList(UserIndex).flags.Oculto = 0
    UserList(UserIndex).flags.Invisible = 0
    UserList(UserIndex).Counters.TiempoOculto = 0
    UserList(UserIndex).Counters.Invisibilidad = 0
    Call WriteConsoleMsg(1, UserIndex, "Vuelves a ser visible.", FontTypeNames.FONTTYPE_BROWNI)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
End If
        
ModNave = ModNavegacion(UserList(UserIndex).Clase)

If UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
    Call WriteConsoleMsg(1, UserIndex, "No tenes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(1, UserIndex, "Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
UserList(UserIndex).Invent.BarcoSlot = Slot

If UserList(UserIndex).flags.Navegando = 0 Then
    
    UserList(UserIndex).Char.Head = 0
    
    If UserList(UserIndex).flags.Muerto = 0 Then
        UserList(UserIndex).Char.body = 84
    Else
        UserList(UserIndex).Char.body = 87
    End If
    
    UserList(UserIndex).flags.Navegando = 1
    
Else
    
    UserList(UserIndex).flags.Navegando = 0
    
    If UserList(UserIndex).flags.Muerto = 0 Then
        UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
        
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(UserIndex).Char.body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(UserIndex)
        End If
        
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
    Else
        UserList(UserIndex).Char.body = iCuerpoMuerto
        UserList(UserIndex).Char.Head = iCabezaMuerto
    End If
End If

Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
Call WriteNavigateToggle(UserIndex)

End Sub

Public Function DoEquita(ByVal UserIndex As Integer, ByVal obj As Integer, ByVal Slot As Integer)
Dim ModEqui As Long
Dim SK As Long
Dim MK As Long

With UserList(UserIndex)
    If .flags.Navegando = 1 Then Exit Function
    
    If .flags.Invisible = 1 Or .flags.Oculto = 1 Then
        .flags.Oculto = 0
        .flags.Invisible = 0
        .Counters.TiempoOculto = 0
        .Counters.Invisibilidad = 0
        Call WriteConsoleMsg(1, UserIndex, "Vuelves a ser visible.", FontTypeNames.FONTTYPE_BROWNI)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
    End If
    
    
    
    SK = .Stats.UserSkills(Equitacion)
    MK = ObjData(obj).MinSkill
    
    If obj <> 1351 Then 'Si no es el unicornio obviamos esta validacion
        If (Not ClasePuedeUsarItem(UserIndex, obj)) Then
            Call WriteConsoleMsg(2, UserIndex, "Tu clase no puede usar esta montura.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        If (Not CheckRazaUsaRopa(UserIndex, obj)) Then
            Call WriteConsoleMsg(2, UserIndex, "Tu raza no puede usar esta montura.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    End If
    
    If .Stats.UserSkills(Equitacion) < ObjData(obj).MinSkill Then
        Call WriteConsoleMsg(2, UserIndex, "Para usar esta montura necesitas " & ObjData(obj).MinSkill & " puntos en equitación.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    
    If .flags.Montando = 0 Then
        If MapData(.Pos.map, .Pos.x, .Pos.Y).Trigger >= 20 Then
            Exit Function
        End If
        
        'Add Marius No montar en doungeons ni en Captura la bandera
        If MapInfo(.Pos.map).Zona = "DUNGEON" Or .Pos.map = Bandera_mapa Then
            Exit Function
        End If
    End If
    
    .Invent.MonturaObjIndex = .Invent.Object(Slot).ObjIndex
    .Invent.MonturaSlot = Slot
     
    If .flags.Montando = 0 Then
        .Char.Head = 0
        If .flags.Muerto = 0 Then
            .Char.body = ObjData(obj).Ropaje
        Else
            .Char.body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
        End If
        .Char.Head = .OrigChar.Head
        .flags.Montando = 1
        .Invent.Object(Slot).Equipped = 1
        Call UpdateUserInv(False, UserIndex, .Invent.MonturaSlot)
    Else
        .flags.Montando = 0
        If .flags.Muerto = 0 Then
            .Char.Head = .OrigChar.Head
            If .Invent.ArmourEqpObjIndex > 0 Then
                .Char.body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
            Else
                Call DarCuerpoDesnudo(UserIndex)
            End If
        Else
            .Char.body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
        End If
        .Invent.Object(.Invent.MonturaSlot).Equipped = 0
        Call UpdateUserInv(False, UserIndex, .Invent.MonturaSlot)
            
        .Invent.MonturaObjIndex = 0
        .Invent.MonturaSlot = 0
    End If
     
    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    Call WriteEquitateToggle(UserIndex)
End With

End Function
 
Public Function TieneObjetos(ByVal ItemIndex As Integer, ByVal Cant As Long, ByVal UserIndex As Integer) As Boolean
Dim i As Integer
Dim Total As Long

    For i = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
            Total = Total + UserList(UserIndex).Invent.Object(i).Amount
        End If
    Next i
    
    If Cant <= Total Then
        TieneObjetos = True
        Exit Function
End If
        
End Function
Function Have_Obj_To_Slot(ByVal ItemIndex As Integer, ByVal Slot As Byte, ByVal UserIndex As Integer) As Long
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long

    For i = 1 To Slot
        If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
            Have_Obj_To_Slot = Have_Obj_To_Slot + UserList(UserIndex).Invent.Object(i).Amount
        End If
    Next i
        
End Function
Function Have_Obj_Slot(ByVal ItemIndex As Integer, ByVal UserIndex As Integer) As Integer
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long

    For i = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
            Have_Obj_Slot = i
        End If
    Next i
        
End Function
Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal Cant As Integer, ByVal UserIndex As Integer) As Boolean
'Call LogTarea("Sub QuitarObjetos")

Dim i As Integer
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
        
        Call Desequipar(UserIndex, i)
        
        UserList(UserIndex).Invent.Object(i).Amount = UserList(UserIndex).Invent.Object(i).Amount - Cant
        If (UserList(UserIndex).Invent.Object(i).Amount <= 0) Then
            Cant = Abs(UserList(UserIndex).Invent.Object(i).Amount)
            UserList(UserIndex).Invent.Object(i).Amount = 0
            UserList(UserIndex).Invent.Object(i).ObjIndex = 0
        Else
            Cant = 0
        End If
        
        Call UpdateUserInv(False, UserIndex, i)
        
        If (Cant = 0) Then
            QuitarObjetos = True
            Exit Function
        End If
    End If
Next i

End Function

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, Cant As Integer)
    If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH * CLng(Cant), UserIndex)
    If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP * CLng(Cant), UserIndex)
    If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO * CLng(Cant), UserIndex)
End Sub

Sub carpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer)
    If ObjData(ItemIndex).Madera > 0 Then
        Call QuitarObjetos(Leña, ObjData(ItemIndex).Madera * CLng(Cant), UserIndex)
    End If
End Sub


Sub druidaQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer)
    If ObjData(ItemIndex).raies > 0 Then
        Call QuitarObjetos(Raiz, ObjData(ItemIndex).raies * CLng(Cant), UserIndex)
    End If
End Sub

Sub SastreQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer)
    If ObjData(ItemIndex).PielLobo Then _
        Call QuitarObjetos(PielLobo, ObjData(ItemIndex).PielLobo * Cant, UserIndex)
        
    If ObjData(ItemIndex).PielOso Then _
        Call QuitarObjetos(PielOso, ObjData(ItemIndex).PielOso * CLng(Cant), UserIndex)
    
    If ObjData(ItemIndex).PielLoboInvernal > 0 Then _
        Call QuitarObjetos(PielLoboInvernal, ObjData(ItemIndex).PielLoboInvernal * CLng(Cant), UserIndex)
End Sub

Function CarpinteroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer) As Boolean
    
    If ObjData(ItemIndex).Madera > 0 Then
        If Not TieneObjetos(Leña, CLng(ObjData(ItemIndex).Madera) * CLng(Cant), UserIndex) Then
            Call WriteConsoleMsg(1, UserIndex, "No tenes suficientes madera.", FontTypeNames.FONTTYPE_INFO)
            CarpinteroTieneMateriales = False
            Exit Function
        End If
    End If
    
    CarpinteroTieneMateriales = True

End Function

Function druidaTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer) As Boolean
    
    If ObjData(ItemIndex).raies > 0 Then
        If Not TieneObjetos(Raiz, CLng(ObjData(ItemIndex).raies) * CLng(Cant), UserIndex) Then
            Call WriteConsoleMsg(1, UserIndex, "No tenes suficientes raices.", FontTypeNames.FONTTYPE_INFO)
            druidaTieneMateriales = False
            Exit Function
    End If
End If
    
    druidaTieneMateriales = True

End Function



Function SastreTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer) As Boolean
    If ObjData(ItemIndex).PielLobo Then
        If TieneObjetos(PielLobo, CLng(ObjData(ItemIndex).PielLobo) * CLng(Cant), UserIndex) = False Then
            Call WriteConsoleMsg(1, UserIndex, "No tenes suficientes pieles.", FontTypeNames.FONTTYPE_INFO)
            SastreTieneMateriales = False
            Exit Function
        End If
    End If

    If ObjData(ItemIndex).PielOso Then
        If TieneObjetos(PielOso, CLng(ObjData(ItemIndex).PielOso) * CLng(Cant), UserIndex) = False Then
            Call WriteConsoleMsg(1, UserIndex, "No tenes suficientes pieles.", FontTypeNames.FONTTYPE_INFO)
            SastreTieneMateriales = False
            Exit Function
        End If
    End If
    
    If ObjData(ItemIndex).PielLoboInvernal Then
        If TieneObjetos(PielLoboInvernal, CLng(ObjData(ItemIndex).PielLoboInvernal) * CLng(Cant), UserIndex) = False Then
            Call WriteConsoleMsg(1, UserIndex, "No tenes suficientes pieles.", FontTypeNames.FONTTYPE_INFO)
            SastreTieneMateriales = False
            Exit Function
        End If
    End If
    
    SastreTieneMateriales = True

End Function

 
Function HerreroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer) As Boolean
    If ObjData(ItemIndex).LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH * CLng(Cant), UserIndex) Then
                    Call WriteConsoleMsg(1, UserIndex, "No tenes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingP > 0 Then
            If Not TieneObjetos(LingotePlata, ObjData(ItemIndex).LingP * CLng(Cant), UserIndex) Then
                    Call WriteConsoleMsg(1, UserIndex, "No tenes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingO > 0 Then
            If Not TieneObjetos(LingoteOro, ObjData(ItemIndex).LingO * CLng(Cant), UserIndex) Then
                    Call WriteConsoleMsg(1, UserIndex, "No tenes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    HerreroTieneMateriales = True
End Function

Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer) As Boolean
    PuedeConstruir = HerreroTieneMateriales(UserIndex, ItemIndex, Cant) And _
                     UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) >= _
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


Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer)

Dim OtroUserIndex As Integer

If PuedeConstruir(UserIndex, ItemIndex, Cant) And PuedeConstruirHerreria(ItemIndex) Then
    
    If UserList(UserIndex).flags.Comerciando Then
        OtroUserIndex = UserList(UserIndex).ComUsu.DestUsu
           
        If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
            Call WriteConsoleMsg(1, UserIndex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
            Call WriteConsoleMsg(1, OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
           
            Call LimpiarComercioSeguro(UserIndex)
            Call Protocol.FlushBuffer(OtroUserIndex)
        End If
    End If
    
    'Sacamos energía
    If UserList(UserIndex).Clase = eClass.Herrero Then
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(UserIndex).Stats.MinSTA >= ENERGIA_TRABAJO_HERRERO Then
            UserList(UserIndex).Stats.MinSTA = UserList(UserIndex).Stats.MinSTA - ENERGIA_TRABAJO_HERRERO
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteConsoleMsg(1, UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(UserIndex).Stats.MinSTA >= ENERGIA_TRABAJO_NOHERRERO Then
            UserList(UserIndex).Stats.MinSTA = UserList(UserIndex).Stats.MinSTA - ENERGIA_TRABAJO_NOHERRERO
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteConsoleMsg(1, UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    Call HerreroQuitarMateriales(UserIndex, ItemIndex, Cant)
    ' AGREGAR FX
    If ObjData(ItemIndex).OBJType = eOBJType.otWeapon Then
        Call WriteConsoleMsg(1, UserIndex, "Has construido el arma!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otESCUDO Then
        Call WriteConsoleMsg(1, UserIndex, "Has construido el escudo!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otCASCO Then
        Call WriteConsoleMsg(1, UserIndex, "Has construido el casco!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otArmadura Then
        Call WriteConsoleMsg(1, UserIndex, "Has construido la armadura!.", FontTypeNames.FONTTYPE_INFO)
    End If
    Dim MiObj As obj
    MiObj.Amount = Cant
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    Call SubirSkill(UserIndex, Herreria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MARTILLOHERRERO, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
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


Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer)

Dim OtroUserIndex As Integer

If CarpinteroTieneMateriales(UserIndex, ItemIndex, Cant) And _
   UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) >= _
   ObjData(ItemIndex).SkCarpinteria And _
   PuedeConstruirCarpintero(ItemIndex) And _
   UserList(UserIndex).Invent.AnilloEqpObjIndex = SERRUCHO_CARPINTERO Then
    
    If UserList(UserIndex).flags.Comerciando Then
        OtroUserIndex = UserList(UserIndex).ComUsu.DestUsu
           
        If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
            Call WriteConsoleMsg(1, UserIndex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
            Call WriteConsoleMsg(1, OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
           
            Call LimpiarComercioSeguro(UserIndex)
            Call Protocol.FlushBuffer(OtroUserIndex)
        End If
    End If
    
    'Sacamos energía
    If UserList(UserIndex).Clase = eClass.Carpintero Then
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(UserIndex).Stats.MinSTA >= ENERGIA_TRABAJO_HERRERO Then
            UserList(UserIndex).Stats.MinSTA = UserList(UserIndex).Stats.MinSTA - ENERGIA_TRABAJO_HERRERO
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteConsoleMsg(1, UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(UserIndex).Stats.MinSTA >= ENERGIA_TRABAJO_NOHERRERO Then
            UserList(UserIndex).Stats.MinSTA = UserList(UserIndex).Stats.MinSTA - ENERGIA_TRABAJO_NOHERRERO
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteConsoleMsg(1, UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    Call carpinteroQuitarMateriales(UserIndex, ItemIndex, Cant)
    Call WriteConsoleMsg(1, UserIndex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
    
    Dim MiObj As obj
    MiObj.Amount = Cant
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If

    Call SubirSkill(UserIndex, Carpinteria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))


    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

End If
End Sub


Public Sub druidaConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer)

If druidaTieneMateriales(UserIndex, ItemIndex, Cant) And _
   UserList(UserIndex).Stats.UserSkills(eSkill.alquimia) >= _
   ObjData(ItemIndex).SkPociones And _
   PuedeConstruirDruida(ItemIndex) And _
   UserList(UserIndex).Invent.AnilloEqpObjIndex = OLLA Then
   
    'Sacamos energía
    If UserList(UserIndex).Clase = eClass.Druida Then
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(UserIndex).Stats.MinSTA >= ENERGIA_TRABAJO_HERRERO Then
            UserList(UserIndex).Stats.MinSTA = UserList(UserIndex).Stats.MinSTA - ENERGIA_TRABAJO_HERRERO
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteConsoleMsg(1, UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(UserIndex).Stats.MinSTA >= ENERGIA_TRABAJO_NOHERRERO Then
            UserList(UserIndex).Stats.MinSTA = UserList(UserIndex).Stats.MinSTA - ENERGIA_TRABAJO_NOHERRERO
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteConsoleMsg(1, UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    Call druidaQuitarMateriales(UserIndex, ItemIndex, Cant)
    Call WriteConsoleMsg(1, UserIndex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
    
    Dim MiObj As obj
    MiObj.Amount = Cant
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If

    Call SubirSkill(UserIndex, alquimia)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))


    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

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


Public Sub DoLingotes(ByVal UserIndex As Integer)

On Error GoTo hayerror

    Dim Slot As Integer
    Dim obji As Integer
    Dim OtroUserIndex As Integer

    Slot = UserList(UserIndex).flags.TargetObjInvSlot
    obji = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    
    If UserList(UserIndex).flags.Comerciando Then
        OtroUserIndex = UserList(UserIndex).ComUsu.DestUsu
           
        If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
            Call WriteConsoleMsg(1, UserIndex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
            Call WriteConsoleMsg(1, OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
           
            Call LimpiarComercioSeguro(UserIndex)
            Call Protocol.FlushBuffer(OtroUserIndex)
        End If
    End If
    
    If UserList(UserIndex).Invent.Object(Slot).Amount < MineralesParaLingote(obji) Or _
        ObjData(obji).OBJType <> eOBJType.otMinerales Then
            Call WriteConsoleMsg(1, UserIndex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
    End If
    
    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - MineralesParaLingote(obji)
    If UserList(UserIndex).Invent.Object(Slot).Amount < 1 Then
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
    End If
    
    Dim nPos As WorldPos
    Dim MiObj As obj
    MiObj.Amount = 1 * ModTrabajo
    MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).LingoteIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    Call UpdateUserInv(False, UserIndex, Slot)
    Call WriteConsoleMsg(1, UserIndex, "¡Has obtenido un lingote!", FontTypeNames.FONTTYPE_INFO)
    
    'Sonido Lingoteamos
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(162, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))

    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

Exit Sub

hayerror:

LogError ("Error en dolingotes: " & err.Number & " desc: " & err.description & " Name: " & UserList(UserIndex).Name)

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

Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer
'***************************************************
'Author: Unknown
'Last Modification: 02/03/09
'02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
'***************************************************
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasType(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function
        End If
    Next j
End Function

Sub DoDomar(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
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


If Npclist(NpcIndex).MaestroUser = UserIndex Then
    Call WriteConsoleMsg(1, UserIndex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).NroMascotas < MAXMASCOTAS Then
    
    If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
        Call WriteConsoleMsg(1, UserIndex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not PuedeDomarMascota(UserIndex, NpcIndex) Then
        Call WriteConsoleMsg(1, UserIndex, "No puedes domar mas de dos criaturas del mismo tipo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    puntosDomar = CInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma)) * CInt(UserList(UserIndex).Stats.UserSkills(eSkill.Domar))
    puntosRequeridos = Npclist(NpcIndex).flags.Domable
    
    If puntosRequeridos <= puntosDomar And RandomNumber(1, 5) = 1 Then
        Dim index As Integer
        UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas + 1
        index = FreeMascotaIndex(UserIndex)
        UserList(UserIndex).MascotasIndex(index) = NpcIndex
        UserList(UserIndex).MascotasType(index) = Npclist(NpcIndex).Numero
        
        Npclist(NpcIndex).MaestroUser = UserIndex
        
        Call FollowAmo(NpcIndex)
        Call ReSpawnNpc(Npclist(NpcIndex))
        
        Call WriteConsoleMsg(1, UserIndex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)
        
        ' Es zona segura?
        CanStay = (MapInfo(UserList(UserIndex).Pos.map).Pk = True)
        
        If Not CanStay Then
            petType = Npclist(NpcIndex).Numero
            NroPets = UserList(UserIndex).NroMascotas
            
            Call QuitarNPC(NpcIndex)
            
            UserList(UserIndex).MascotasType(index) = petType
            UserList(UserIndex).NroMascotas = NroPets
            
            Call WriteConsoleMsg(1, UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
        End If

    Else
        If Not UserList(UserIndex).flags.UltimoMensaje = 5 Then
            Call WriteConsoleMsg(1, UserIndex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 5
        End If
    End If
    
    'Entreno domar. Es un 30% más dificil si no sos druida.
    If UserList(UserIndex).Clase = eClass.Druida Or (RandomNumber(1, 3) < 3) Then
        Call SubirSkill(UserIndex, Domar)
    End If
Else
    Call WriteConsoleMsg(1, UserIndex, "No puedes controlar más criaturas.", FontTypeNames.FONTTYPE_INFO)
End If
End Sub

''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: ZaMa
'This function checks how many NPCs of the same type have
'been tamed by the user.
'Returns True if that amount is less than two.
'***************************************************
    Dim i As Long
    Dim numMascotas As Long
    
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasType(i) = Npclist(NpcIndex).Numero Then
            numMascotas = numMascotas + 1
        End If
    Next i
    
    If numMascotas <= 1 Then PuedeDomarMascota = True
    
End Function

Sub DoAdminInvisible(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).flags.AdminInvisible = 0 Then
        UserList(UserIndex).flags.AdminInvisible = 1
        UserList(UserIndex).flags.Invisible = 1
        UserList(UserIndex).flags.Oculto = 1
        UserList(UserIndex).flags.OldBody = UserList(UserIndex).Char.body
        UserList(UserIndex).flags.OldHead = UserList(UserIndex).Char.Head
        UserList(UserIndex).Char.body = 0
        UserList(UserIndex).Char.Head = 0
        
        'Add Marius
        UserList(UserIndex).showName = False
    Else
        UserList(UserIndex).flags.AdminInvisible = 0
        UserList(UserIndex).flags.Invisible = 0
        UserList(UserIndex).flags.Oculto = 0
        UserList(UserIndex).Counters.TiempoOculto = 0
        UserList(UserIndex).Char.body = UserList(UserIndex).flags.OldBody
        UserList(UserIndex).Char.Head = UserList(UserIndex).flags.OldHead
        
        'Add Marius
        UserList(UserIndex).showName = True
    End If
            
    'Add Marius
    Call RefreshCharStatus(UserIndex)
    '\Add
    
    'vuelve a ser visible por la fuerza
    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
End Sub

Sub TratarDeHacerFogata(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

Dim Suerte As Byte
Dim exito As Byte
Dim obj As obj
Dim posMadera As WorldPos

If Not LegalPos(map, x, Y) Then Exit Sub

With posMadera
    .map = map
    .x = x
    .Y = Y
End With

If MapData(map, x, Y).ObjInfo.ObjIndex <> 58 Then
    Call WriteConsoleMsg(1, UserIndex, "Necesitas clickear sobre Leña para hacer ramitas", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Distancia(posMadera, UserList(UserIndex).Pos) > 2 Then
    Call WriteConsoleMsg(1, UserIndex, "Estás demasiado lejos para prender la fogata.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(1, UserIndex, "No puedes hacer fogatas estando muerto.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If MapData(map, x, Y).ObjInfo.Amount < 3 Then
    Call WriteConsoleMsg(1, UserIndex, "Necesitas por lo menos tres troncos para hacer una fogata.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If


If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) < 6 Then
    Suerte = 3
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 34 Then
    Suerte = 2
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 35 Then
    Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    obj.ObjIndex = FOGATA_APAG
    obj.Amount = MapData(map, x, Y).ObjInfo.Amount \ 3
    
    Call WriteConsoleMsg(1, UserIndex, "Has hecho " & obj.Amount & " fogatas.", FontTypeNames.FONTTYPE_INFO)
    
    Call MakeObj(obj, map, x, Y)
    
    'Seteamos la fogata como el nuevo TargetObj del user
    UserList(UserIndex).flags.TargetObj = FOGATA_APAG
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 10 Then
        Call WriteConsoleMsg(1, UserIndex, "No has podido hacer la fogata.", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 10
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Supervivencia)


End Sub

Public Sub DoPescar(ByVal UserIndex As Integer, Optional ByVal red As Boolean)
On Error GoTo Errhandler

Dim Suerte As Integer
Dim res As Integer

    If UserList(UserIndex).Clase = eClass.Pescador Then
        Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
    Else
        Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
    End If
    
    Dim Skill As Integer
    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Pesca)
    Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
    res = RandomNumber(1, Suerte)
    
    If res <= 6 Then
        Dim nPos As WorldPos
        Dim MiObj As obj
        Dim Pez As Integer
        
        MiObj.Amount = ModTrabajo + IIf(red, RandomNumber(7, 15) * ModTrabajo, 0)
        If HayAgua(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y) Then
            If UCase$(Left$(Tilde(MapInfo(UserList(UserIndex).Pos.map).Name), 14)) = "OCEANO ABIERTO" Then
                If UserList(UserIndex).Clase = eClass.Pescador Then
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
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
    Else
        '[CDT 17-02-2004]
        If Not UserList(UserIndex).flags.UltimoMensaje = 6 Then
          Call WriteConsoleMsg(2, UserIndex, "¡No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
          UserList(UserIndex).flags.UltimoMensaje = 6
        End If
        '[/CDT]
    End If
    
    Call SubirSkill(UserIndex, Pesca)
    
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

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

If UserList(VictimaIndex).faccion.FuerzasCaos = 1 And UserList(LadrOnIndex).faccion.FuerzasCaos = 1 Then
    Call WriteConsoleMsg(1, LadrOnIndex, "No puedes robar a otros miembros de las fuerzas del caos", FontTypeNames.FONTTYPE_FIGHT)
    Exit Sub
End If


Call QuitarSta(LadrOnIndex, 15)

Dim GuantesHurto As Boolean
Dim OtroUserIndex As Integer


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
       
        If UserList(VictimaIndex).flags.Comerciando Then
            OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(1, VictimaIndex, "¡¡Comercio cancelado, te están robando!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(1, OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
               
                Call LimpiarComercioSeguro(VictimaIndex)
                Call Protocol.FlushBuffer(OtroUserIndex)
            End If
        End If
        
                
                
        If (RandomNumber(1, 50) < 25) And (UserList(LadrOnIndex).Clase = eClass.Ladron) Then
            If TieneObjetosRobables(VictimaIndex) Then
                If UserList(VictimaIndex).flags.Comerciando Then
                    OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
    
                    If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                        Call WriteConsoleMsg(1, VictimaIndex, "¡¡Comercio cancelado, te están robando!!", FontTypeNames.FONTTYPE_TALK)
                        Call WriteConsoleMsg(1, OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                       
                        Call LimpiarComercioSeguro(VictimaIndex)
                        Call Protocol.FlushBuffer(OtroUserIndex)
                    End If
                End If
            
            
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
    ObjData(OI).Milicia = 0 And _
    ObjData(OI).Shop = 0 And ObjData(OI).EfectoMagico = eMagicType.Sacrificio And _
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
    Dim MiObj As obj
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

Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal Daño As Integer)
'***************************************************
'Autor: Nacho (Integer) & Unknown (orginal version)
'Last Modification: 04/17/08 - (NicoNZ)
'Simplifique la cuenta que hacia para sacar la suerte
'y arregle la cuenta que hacia para sacar el daño
'***************************************************
Dim Suerte As Integer
Dim Skill As Integer

Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar)

Select Case UserList(UserIndex).Clase
    Case eClass.Asesino
        Suerte = Int(((0.00003 * Skill - 0.002) * Skill + 0.098) * Skill + 4.25)
        '24 (con 100 skill)
    Case eClass.Clerigo, eClass.Paladin
        Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
        '14 (con 100 skill)
    Case eClass.Bardo
        Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
        '12 (con 100 skill)
    Case Else
        Suerte = Int(0.0361 * Skill + 4.39)
        '7 casi 4 (con 100 skill)
End Select

If RandomNumber(0, 100) <= Suerte Then
    If VictimUserIndex <> 0 Then
        If UserList(UserIndex).Clase = eClass.Asesino Then
            Daño = Round(Daño * 1.4, 0)
        Else
            Daño = Round(Daño * 1.5, 0)
        End If
        
        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - Daño
        Call WriteConsoleMsg(2, UserIndex, "Has apuñalado a " & UserList(VictimUserIndex).Name & " por " & Daño, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(2, VictimUserIndex, "Te ha apuñalado " & UserList(UserIndex).Name & " por " & Daño, FontTypeNames.FONTTYPE_FIGHT)
        
        Call FlushBuffer(VictimUserIndex)
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - Int(Daño * 2)
        Call WriteConsoleMsg(2, UserIndex, "Has apuñalado la criatura por " & Int(Daño * 2), FontTypeNames.FONTTYPE_FIGHT)
        Call SubirSkill(UserIndex, Apuñalar)
        '[Alejo]
        Call CalcularDarExp(UserIndex, VictimNpcIndex, Daño * 2)
    End If
Else
    Call WriteConsoleMsg(2, UserIndex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)
End If

End Sub


Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
    UserList(UserIndex).Stats.MinSTA = UserList(UserIndex).Stats.MinSTA - Cantidad
    If UserList(UserIndex).Stats.MinSTA < 0 Then UserList(UserIndex).Stats.MinSTA = 0
    Call WriteUpdateSta(UserIndex)
End Sub

Public Sub DoTalar(ByVal UserIndex As Integer)
On Error GoTo Errhandler

    Dim Suerte As Integer
    Dim res As Integer
    
    If UserList(UserIndex).Clase = eClass.Leñador Then
        Call QuitarSta(UserIndex, EsfuerzoTalarLeñador)
    Else
        Call QuitarSta(UserIndex, EsfuerzoTalarGeneral)
    End If
    
    Dim Skill As Integer
    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Talar)
    Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
    res = RandomNumber(1, Suerte)
    
    If res <= 6 Then
        Dim nPos As WorldPos
        Dim MiObj As obj
        
        If UserList(UserIndex).Clase = eClass.Leñador Then
            MiObj.Amount = RandomNumber(1, 6) * ModTrabajo
        Else
            MiObj.Amount = RandomNumber(1, 2) * ModTrabajo
        End If
        
        MiObj.ObjIndex = Leña
        
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
    Else
        '[CDT 17-02-2004]
        If Not UserList(UserIndex).flags.UltimoMensaje = 8 Then
            Call WriteConsoleMsg(1, UserIndex, "¡No has obtenido leña!", FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 8
        End If
        '[/CDT]
    End If
    
    Call SubirSkill(UserIndex, Talar)
    
    
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
    
    Exit Sub

Errhandler:
    Call LogError("Error en DoTalar")

End Sub
Public Sub DoMineria(ByVal UserIndex As Integer)
On Error GoTo Errhandler

Dim Suerte As Integer
Dim res As Integer
Dim metal As Integer

If UserList(UserIndex).Clase = eClass.Minero Then
    Call QuitarSta(UserIndex, EsfuerzoExcavarMinero)
Else
    Call QuitarSta(UserIndex, EsfuerzoExcavarGeneral)
End If

Dim Skill As Integer
Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Mineria)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res <= 5 Then
    Dim MiObj As obj
    Dim nPos As WorldPos
    
    If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
    
    MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObj).MineralIndex
    
    If UserList(UserIndex).Clase = eClass.Minero Then
        MiObj.Amount = RandomNumber(1, 6) * ModTrabajo
    Else
        MiObj.Amount = RandomNumber(1, 2) * ModTrabajo
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then _
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_MINERO, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
End If

Call SubirSkill(UserIndex, Mineria)


UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

Exit Sub

Errhandler:
    Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)

UserList(UserIndex).Counters.IdleCount = 0

Dim Suerte As Integer
Dim res As Integer
Dim Cant As Integer

Dim TActual As Long
TActual = GetTickCount() And &H7FFFFFFF
    
If UserList(UserIndex).Stats.MinMAN >= UserList(UserIndex).Stats.MaxMAN Then Exit Sub

If UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 81 Then
                    Suerte = 12
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) < 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 91 Then
                    Suerte = 8
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) = 100 Then
                    Suerte = 7
End If

'Mannakia
If UserList(UserIndex).Invent.MagicIndex <> 0 Then
    If ObjData(UserList(UserIndex).Invent.MagicIndex).EfectoMagico = eMagicType.AceleraMana Then
        Suerte = Suerte - Porcentaje(Suerte, 30)  ' nose algo razonable
    End If
End If
'Mannakia

res = RandomNumber(1, Suerte)

If res = 1 Then
    
    Cant = Porcentaje(UserList(UserIndex).Stats.MaxMAN, PorcentajeRecuperoMana)
    
    If Cant <= 0 Then Cant = 1
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + Cant
    If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then _
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
    
    Call WriteUpdateMana(UserIndex)
    Call SubirSkill(UserIndex, Meditar)
End If

End Sub
Public Function DoTrabajar(ByVal UserIndex As Integer)
    Dim OtroUserIndex As Integer
    
    If UserList(UserIndex).flags.Comerciando Then
        OtroUserIndex = UserList(UserIndex).ComUsu.DestUsu
           
        If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
            Call WriteConsoleMsg(1, UserIndex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
            Call WriteConsoleMsg(1, OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
           
            Call LimpiarComercioSeguro(UserIndex)
            Call Protocol.FlushBuffer(OtroUserIndex)
        End If
    End If
    
    If Not IntervaloPermiteTrabajar(UserIndex, True) Then Exit Function
    
    With UserList(UserIndex)
        If .Stats.MinSTA < 2 Then
            Call WriteConsoleMsg(2, UserIndex, "Dejas de trabajar.", FontTypeNames.FONTTYPE_BROWNI)
            .flags.Trabajando = False
            Exit Function
        End If
        
        If .flags.Lingoteando Then
            Call DoLingotes(UserIndex)
        ElseIf .Invent.AnilloEqpSlot <> 0 Then
            Select Case .Invent.AnilloEqpObjIndex
                Case RED_PESCA
                    Call DoPescar(UserIndex, True)
                    
                Case CAÑA_PESCA
                    Call DoPescar(UserIndex)
                    
                Case PIQUETE_MINERO
                    Call DoMineria(UserIndex)
                  
                Case HACHA_LEÑADOR
                    Call DoTalar(UserIndex)
                    
                Case TIJERAS
                    Call DoBotanica(UserIndex)
            End Select
        End If
    End With
End Function




Public Sub DoBotanica(ByVal UserIndex As Integer)
On Error GoTo Errhandler

Dim Suerte As Integer
Dim res As Integer

    If UserList(UserIndex).Clase = eClass.Druida Then
        Call QuitarSta(UserIndex, EsfuerzoBotanicaDruida)
    Else
        Call QuitarSta(UserIndex, EsfuerzoBotanicaGeneral)
    End If
    
    Dim Skill As Integer
    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.botanica)
    Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
    res = RandomNumber(1, Suerte)
    
    If res <= 6 Then
        Dim nPos As WorldPos
        Dim MiObj As obj
        
        If UserList(UserIndex).Clase = eClass.Druida Then
            MiObj.Amount = RandomNumber(1, 6) * ModTrabajo
        Else
            MiObj.Amount = RandomNumber(1, 2) * ModTrabajo
        End If
        
        MiObj.ObjIndex = Raiz
        
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
    
    Else
        '[CDT 17-02-2004]
        If Not UserList(UserIndex).flags.UltimoMensaje = 8 Then
            Call WriteConsoleMsg(1, UserIndex, "¡No has obtenido raices!", FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 8
        End If
        '[/CDT]
    End If
    
    Call SubirSkill(UserIndex, botanica)
    
    
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

Exit Sub

Errhandler:
    Call LogError("Error en DoTalar")

End Sub






Public Sub SastreConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer)

If SastreTieneMateriales(UserIndex, ItemIndex, Cant) And _
   UserList(UserIndex).Stats.UserSkills(eSkill.Sastreria) >= _
   ObjData(ItemIndex).SkSastreria And _
   PuedeConstruirSastre(ItemIndex) And _
   UserList(UserIndex).Invent.AnilloEqpObjIndex = COSTURERO Then
   
    'Sacamos energía
    If UserList(UserIndex).Clase = eClass.Sastre Then
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(UserIndex).Stats.MinSTA >= ENERGIA_TRABAJO_HERRERO Then
            UserList(UserIndex).Stats.MinSTA = UserList(UserIndex).Stats.MinSTA - ENERGIA_TRABAJO_HERRERO
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteConsoleMsg(1, UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(UserIndex).Stats.MinSTA >= ENERGIA_TRABAJO_NOHERRERO Then
            UserList(UserIndex).Stats.MinSTA = UserList(UserIndex).Stats.MinSTA - ENERGIA_TRABAJO_NOHERRERO
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteConsoleMsg(1, UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    Call SastreQuitarMateriales(UserIndex, ItemIndex, Cant)
    Call WriteConsoleMsg(1, UserIndex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
    
    Dim MiObj As obj
    MiObj.Amount = Cant
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If

    Call SubirSkill(UserIndex, Sastreria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))

    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

End If
End Sub
