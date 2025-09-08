VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Software"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Procedure     : clsMycommerce
' DateTime      : 29/02/2016 20:34
' Author        : clayton
' Revisado em   : 29/02/2016
' Objetivo      : Ficara nessa classes apenas atributos/metodos relacionados ao mycommerce.exe/Visual Software
'---------------------------------------------------------------------------------------
'
Option Explicit


Private mvarNomeEmpresa As String    'local copy
Private mvarTelefoneEmpresa As String    'local copy
Private mvarCidadeEmpresa As String    'local copy
Private mvarVersaoSoftware As String    'local copy
Private mvarNomeSoftware As String    'local copy
'
Private mvarTitleBarDefaultColor1 As Long
Private mvarTitleBarDefaultColor2 As Long
Private mvarTitleBarFonteDefaultColor As Long
'
Private mvarTitleBarColor1 As Long
Private mvarTitleBarColor2 As Long
Private mvarTitleBarFonteColor As Long
'
Private mvarisTitleBarCustom As Boolean

Public ReconectDB As Boolean

Public Property Get TitleBarDefaultColor1() As Long
    TitleBarDefaultColor1 = mvarTitleBarDefaultColor1
End Property

Public Property Get TitleBarColor1() As Long
    TitleBarColor1 = mvarTitleBarColor1
End Property


Public Property Get isTitleBarCustom() As Boolean
    isTitleBarCustom = mvarisTitleBarCustom
End Property

Public Property Get NomeSoftware() As String
    NomeSoftware = mvarNomeSoftware
End Property
Public Property Let NomeSoftware(ByVal vData As String)
    mvarNomeSoftware = vData
End Property


Public Property Get VersaoSoftware() As String
    VersaoSoftware = mvarVersaoSoftware
End Property
Public Property Let VersaoSoftware(ByVal vData As String)
    mvarVersaoSoftware = vData
End Property

Public Property Let CidadeEmpresa(ByVal vData As String)
    mvarCidadeEmpresa = vData
End Property


Public Property Get CidadeEmpresa() As String
    CidadeEmpresa = mvarCidadeEmpresa
End Property

Private Property Let TelefoneEmpresa(ByVal vData As String)
    mvarTelefoneEmpresa = vData
End Property


Public Property Get TelefoneEmpresa() As String
    TelefoneEmpresa = mvarTelefoneEmpresa
End Property

Private Property Let NomeEmpresa(ByVal vData As String)
    mvarNomeEmpresa = vData
End Property


Public Property Get NomeEmpresa() As String
    NomeEmpresa = mvarNomeEmpresa
End Property



Private Sub Class_Initialize()
    Dim vvar As Variant
    On Error Resume Next

    NomeEmpresa = "VISUAL SOFTWARE"
    TelefoneEmpresa = "(46) 3532-2005"
    VersaoSoftware = "1"
    NomeSoftware = "Mycommerce"


    On Error GoTo 0

    'Call clsSoftware.SetConexao(CN)                                                 '-// para rastrear a queda de conexao

End Sub

Function CriarPastaSuporteRemoto()
    Dim PastaSistema As String
    Dim PastaNova As String

    PastaSistema = App.Path
    PastaNova = PastaSistema & "\Suporte Remoto"

    If Dir(PastaNova, vbDirectory) = "" Then MkDir (PastaNova)

    '' COPIA OS ARQUIVOS PARA A PASTA NOVA E DEPOIS REMOVE
    If Dir(PastaNova & "\Remoto.exe", vbArchive) = "" Then
        FileCopy PastaSistema & "\Remoto.exe", PastaNova & "\Remoto.exe"
        '#COPIA OS ARQUIVOS
        FileCopy PastaSistema & "\MSRC4Plugin.dsm", PastaNova & "\MSRC4Plugin.dsm"
        FileCopy PastaSistema & "\rc4.key", PastaNova & "\rc4.key"
        FileCopy PastaSistema & "\SCHook.dll", PastaNova & "\SCHook.dll"
        FileCopy PastaSistema & "\vnchooks.dll", PastaNova & "\vnchooks.dll"
        FileCopy PastaSistema & "\icon1.ico", PastaNova & "\icon1.ico"
        FileCopy PastaSistema & "\icon2.ico", PastaNova & "\icon2.ico"
        FileCopy PastaSistema & "\background.bmp", PastaNova & "\background.bmp"
        FileCopy PastaSistema & "\logo.bmp", PastaNova & "\logo.bmp"
        FileCopy PastaSistema & "\helpdesk.ini", PastaNova & "\helpdesk.txt"
        '#FIM DA COPIA
    End If

    '#APAGA O ARQUIVO QUE DA CONFLITO NA CHAMADA DIRETA DA RAIZ DO SISTEMA
    If Dir(PastaSistema & "\helpdesk.txt", vbArchive) <> "" Then Kill (PastaSistema & "\helpdesk.txt")
    '#FIM

    Dim objShell As Object
    Dim objShortCut As Object
    Dim strUserProfile As String

    'TAREFA 27327
    '    On Error Resume Next
    '    Set objShell = CreateObject("Wscript.Shell")
    '    strUserProfile = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
    '    Set objShortCut = objShell.CreateShortcut(strUserProfile & "\Desktop\Suporte Remoto.lnk")
    '    objShortCut.TargetPath = PastaNova & "\Remoto.exe"
    '    objShortCut.Description = "Suporte Remoto Visual Software"
    '    objShortCut.IconLocation = PastaNova & "\Icon1.ico"
    '    objShortCut.Save
    '    Set objShell = Nothing
    '    Set objShortCut = Nothing
    '    On Error GoTo 0

    ''ATALHO DO TEAMVIEWER
    If Dir(App.Path & "\Suporte Remoto\TV\TeamViewer.exe", vbArchive) <> "" Then
        On Error Resume Next
        Set objShell = CreateObject("Wscript.Shell")
        strUserProfile = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
        Set objShortCut = objShell.CreateShortcut(strUserProfile & "\Desktop\TeamViewer Visual Software.lnk")
        objShortCut.TargetPath = App.Path & "\Suporte Remoto\TV\TeamViewer.exe"
        objShortCut.Description = "Suporte Remoto Visual Software"
        'objShortCut.IconLocation = PastaNova & "\Icon1.ico"
        objShortCut.save
        Set objShell = Nothing
        Set objShortCut = Nothing
        On Error GoTo 0
    End If

    If Dir(App.Path & "\SuporteWeb.exe", vbArchive) <> "" Then
        On Error Resume Next
        Set objShell = CreateObject("Wscript.Shell")
        strUserProfile = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
        Set objShortCut = objShell.CreateShortcut(strUserProfile & "\Desktop\Suporte Web Visual Software.lnk")
        objShortCut.TargetPath = App.Path & "\SuporteWeb.exe"
        objShortCut.Description = "Suporte Web Visual Software"
        'objShortCut.IconLocation = PastaNova & "\Icon1.ico"
        objShortCut.save
        Set objShell = Nothing
        Set objShortCut = Nothing
        On Error GoTo 0
    End If
    ''


End Function






Private Sub Class_Terminate()
'-//SETA DE VOLTA AS CORES PADROES
'Call MyF.GetColorTitleBarDeafult(mvarTitleBarDefaultColor1, mvarTitleBarDefaultColor2, mvarTitleBarFonteDefaultColor)
'WriteINI App.path & "\" & clsVar.ArquivoConfiguracao, "Cores", GetComputer(), mvarTitleBarDefaultColor1 & "," & mvarTitleBarDefaultColor2 & "," & mvarTitleBarFonteDefaultColor
End Sub




