Attribute VB_Name = "modUpdate"

Public Directory As String
Dim bDone As Boolean, dError As Boolean, f As Integer
Public dMain As Boolean
Public dNum As Integer

Private Type CBChar
    ch(4096) As Byte
End Type
Private Type UNZIPUSERFUNCTION
    UNZIPPrntFunction As Long
    UNZIPSndFunction As Long
    UNZIPReplaceFunction  As Long
    UNZIPPassword As Long
    UNZIPMessage  As Long
    UNZIPService  As Long
    TotalSizeComp As Long
    TotalSize As Long
    CompFactor As Long
    NumFiles As Long
    Comment As Integer
End Type
Private Type UNZIPOPTIONS
    ExtractOnlyNewer  As Long
    SpaceToUnderScore As Long
    PromptToOverwrite As Long
    fQuiet As Long
    ncflag As Long
    ntflag As Long
    nvflag As Long
    nUflag As Long
    nzflag As Long
    ndflag As Long
    noflag As Long
    naflag As Long
    nZIflag As Long
    C_flag As Long
    FPrivilege As Long
    Zip As String
    extractdir As String
End Type
Private Type ZIPnames
    s(0 To 99) As String
End Type
Public Declare Function Wiz_SingleEntryUnzip Lib "unzip32.dll" (ByVal ifnc As Long, ByRef ifnv As ZIPnames, ByVal xfnc As Long, ByRef xfnv As ZIPnames, dcll As UNZIPOPTIONS, Userf As UNZIPUSERFUNCTION) As Long
Public Actualizando As Boolean

Public Sub Unzip(Zip As String, extractdir As String)
On Error GoTo err_Unzip

Dim Resultado As Long
Dim intContadorFicheros As Integer

Dim FuncionesUnZip As UNZIPUSERFUNCTION
Dim OpcionesUnZip As UNZIPOPTIONS

Dim NombresFicherosZip As ZIPnames, NombresFicheros2Zip As ZIPnames

NombresFicherosZip.s(0) = vbNullChar
NombresFicheros2Zip.s(0) = vbNullChar
FuncionesUnZip.UNZIPMessage = 0&
FuncionesUnZip.UNZIPPassword = 0&
FuncionesUnZip.UNZIPPrntFunction = DevolverDireccionMemoria(AddressOf UNFuncionParaProcesarMensajes)
FuncionesUnZip.UNZIPReplaceFunction = DevolverDireccionMemoria(AddressOf UNFuncionReplaceOptions)
FuncionesUnZip.UNZIPService = 0&
FuncionesUnZip.UNZIPSndFunction = 0&
OpcionesUnZip.ndflag = 1 'Carpetas incluidas >> [Bug Fixed]
OpcionesUnZip.C_flag = 1
OpcionesUnZip.fQuiet = 2
OpcionesUnZip.noflag = 1
OpcionesUnZip.Zip = Zip
OpcionesUnZip.extractdir = extractdir

Resultado = Wiz_SingleEntryUnzip(0, NombresFicherosZip, 0, NombresFicheros2Zip, OpcionesUnZip, FuncionesUnZip)

Exit Sub
err_Unzip:
    frmMain.lblEstado.Caption = "Error al descomprimir, pruebe que el archivo sea un original .ZIP y no .RAR"
    MsgBox "Unzip: " + Err.Description, vbExclamation
    Err.Clear
End Sub

Private Function UNFuncionParaProcesarMensajes(ByRef fname As CBChar, ByVal X As Long) As Long
On Error GoTo err_UNFuncionParaProcesarMensajes

    UNFuncionParaProcesarMensajes = 0

Exit Function
err_UNFuncionParaProcesarMensajes:
    MsgBox "UNFuncionParaProcesarMensajes: " + Err.Description, vbExclamation
    Err.Clear
End Function

Private Function UNFuncionReplaceOptions(ByRef p As CBChar, ByVal L As Long, ByRef m As CBChar, ByRef Name As CBChar) As Integer
On Error GoTo err_UNFuncionReplaceOptions

    UNFuncionParaProcesarPassword = 0

Exit Function
err_UNFuncionReplaceOptions:
    MsgBox "UNFuncionParaProcesarPassword: " + Err.Description, vbExclamation
    Err.Clear
End Function
Public Function DevolverDireccionMemoria(Direccion As Long) As Long
On Error GoTo err_DevolverDireccionMemoria

    DevolverDireccionMemoria = Direccion

Exit Function
err_DevolverDireccionMemoria:
    MsgBox "DevolverDireccionMemoria: " + Err.Description, vbExclamation
    Err.Clear
End Function

Sub Analizar()
    Dim i As Integer, iX As Integer, tX As Integer, DifX As Integer
    frmMain.lblEstado.Caption = "Buscando actualizaciones"
    DoEvents

    iX = CInt(frmMain.mInet.OpenURL("http://inmortalao.com.ar/parches/listado.txt")) 'Host
    tX = LeerInt(App.Path & "\INIT\Update.ini")
    DifX = iX - tX

    If Not (DifX = 0) Then
        frmMain.cmdPlay.Picture = LoadInterface("iniciargris")
        frmMain.cmdPlay.Enabled = False
    
        For i = 1 To DifX
            dNum = i + tX
            Directory = App.Path & "\INIT\Parche" & dNum & ".zip"
            dError = False

            frmMain.lblEstado.Caption = "Descargando Parche " & dNum
            DoEvents

            With frmMain.mInet
                .AccessType = icUseDefault
                .URL = "http://inmortalao.com.ar/parches/Parche" & dNum & ".zip"
                .Execute , "GET"
                
                dMain = False
                Do While Not dMain
                    DoEvents
                Loop
            End With

            If dError Then
                MsgBox "Hubo un error al descargar el parche numero " & dNum & "."
                Exit Sub
            Else
                frmMain.lblEstado.Caption = "Extrayendo Parche " & dNum & "/" & DifX
                DoEvents
            End If

            Unzip Directory, App.Path & "\"
            Kill Directory

            DoEvents
        Next i
    End If

    Call GuardarInt(App.Path & "\INIT\Update.ini", iX)

    frmMain.cmdPlay.Enabled = True
    Actualizando = False
End Sub
Private Function LeerInt(ByVal Ruta As String) As Integer
    f = FreeFile
    Open Ruta For Input As f
    LeerInt = Input$(LOF(f), #f)
    Close #f
End Function

Private Sub GuardarInt(ByVal Ruta As String, ByVal data As Integer)
    f = FreeFile
    Open Ruta For Output As f
    Print #f, data
    Close #f
End Sub

