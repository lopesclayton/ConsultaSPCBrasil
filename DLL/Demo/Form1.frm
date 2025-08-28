VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   2520
      Width           =   735
   End
   Begin MSWinsockLib.Winsock SocketSPC 
      Left            =   5280
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   240
      Width           =   9615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   2055
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3240
      Width           =   6735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   7200
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   2640
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim Servidor As String
 Dim Porta As String
 Dim Conectou As String
    Dim a As New Consulta_SPCBrasil.SPC_Brasil_Consulta


Private Sub Command1_Click()
    
    Servidor = "hxh.spc.org.br"
    Porta = "3348"

    Conectou = ""
    
    SocketSPC.Close
    SocketSPC.Connect Servidor, Porta

    Do While Conectou = ""
        DoEvents
    Loop

    DoEvents

MsgBox Len(Text1)

a = "SPC10001420                000126826020122012229918943000856181307724112    03 0000000                 000000000000000000000000000000000000000000000000000  0000000000000000000000000000000000000000000A        0000000000000000000000000000000                    N0"

    SocketSPC.SendData (a & vbCrLf)


End Sub

Private Sub Command2_Click()


    With a

        .Codigo_Transacao = "SPC100"
        .Versao = "014"
        .Meio_de_Acesso = "20"
        .Codigo_Estacao = String(16, "X")
        .Codigo_Operador = "1268260"                       '"1260597" ' "1180891"
        .Senha_Operador = "20122012"
        .Tipo_Documento = "2"
        .CPF_CNPJ = "29918943000856"
        .RG = "                "
        .Codigo_Produto = "03"
        .Tipo_Resposta = "A"
        .Informacaoes_Obito = "N"
        .Informacoes_RG = "N"


        .CNPJ_Empresa = "04717891000152"
        .IDLiberacao = "MY2012A10SO000197246E"


    End With

    a.Gravar_Resposta_Arquivo = "c:\resp.txt"

    strEnviar = a.Monta_String_Envio                       '("hxh.spc.org.br", "3338")    '3338  3348


    Servidor = "hxh.spc.org.br"
    Porta = "3348"

    Conectou = ""

    SocketSPC.Close
    SocketSPC.Connect Servidor, Porta

    Do While Conectou = ""
        DoEvents
    Loop

    SocketSPC.SendData (strEnviar)

    a.Trata_Respostas


    'DoEvents
    'Text1 = ""
    'Open "C:\spc.txt" For Input As #1
    'Do While Not EOF(1)
    'Line Input #1, linha
    '    Text1 = Text1 & linha
    'Loop
    'Close #1

End Sub

Private Sub SocketSPC_Connect()
10   Conectou = "S"
End Sub

Private Sub SocketSPC_DataArrival(ByVal bytesTotal As Long)
    Dim Buffer As String
    Dim Recebidos As Long
    Dim Completo As Boolean

    'On Error GoTo SocketSPC_DataArrival_Error


    Buffer = ""

    '20  lblStatus.Caption = "Recebendo os Dados..."
    '30  Pb1.Value = 0
    Debug.Print bytesTotal

    Do While Recebidos < bytesTotal
        SocketSPC.GetData Buffer, , 1024
        respostasocket = respostasocket & Buffer
        Recebidos = Recebidos + Len(Buffer)
    Loop


    Text2 = Text2 & respostasocket

    If IsEmpty(respostasocket) Then
        Completo = True
    ElseIf Asc(Right(respostasocket, 1)) = 0 Then
        Completo = True
    End If

    If Completo = True Then
        MsgBox "terminou realmente"
        a.StrResposta = Text2
    End If
    




'MsgBox bytesTotals


On Error GoTo 0
Exit Sub

SocketSPC_DataArrival_Error:
'Def
FrameIMG.Visible = False
MsgBox "Erro: " & Err.Number & vbCrLf & _
        "Linha: " & Erl & vbCrLf & _
        "(" & Err.Description & ")" & vbCrLf & _
        "in procedure SocketSPC_DataArrival do Formulário FrmConsultaSPC_Sophus", vbCritical, "Erro"

End Sub

Private Sub SocketSPC_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "erro"
10  Conectou = "E"
20  msg_SocketSPC_Erro = Description
End Sub

Private Sub SocketSPC_SendComplete()
DoEvents
End Sub

Private Sub Text3_Change()
On Error Resume Next
Label1.Caption = Mid(Text2, Text3, Text4 - Text3)
On Error GoTo 0
End Sub

Private Sub Text4_Change()
On Error Resume Next
Label1.Caption = Mid(Text2, Text3, Text4 - Text3)
On Error GoTo 0

End Sub
