VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SPC_Brasil_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Header_Resposta"
Attribute VB_Ext_KEY = "Member1" ,"Produtos"
Attribute VB_Ext_KEY = "Member2" ,"REG_005"
Attribute VB_Ext_KEY = "Member3" ,"REG_006"
Attribute VB_Ext_KEY = "Member4" ,"REG_007"
Attribute VB_Ext_KEY = "Member5" ,"REG_008"
Attribute VB_Ext_KEY = "Member6" ,"REG_009"
Attribute VB_Ext_KEY = "Member7" ,"REG_010"
Attribute VB_Ext_KEY = "Member8" ,"REG_011"
Attribute VB_Ext_KEY = "Member9" ,"REG_020"
Attribute VB_Ext_KEY = "Member10" ,"REG_025"
Attribute VB_Ext_KEY = "Member11" ,"REG_026"
Attribute VB_Ext_KEY = "Member12" ,"REG_027"
Attribute VB_Ext_KEY = "Member13" ,"REG_028"
Attribute VB_Ext_KEY = "Member14" ,"REG_030"
Attribute VB_Ext_KEY = "Member15" ,"REG_040"
Attribute VB_Ext_KEY = "Member16" ,"REG_050"
Attribute VB_Ext_KEY = "Member17" ,"REG_060"
Attribute VB_Ext_KEY = "Member18" ,"REG_070"
Attribute VB_Ext_KEY = "Member19" ,"REG_080"
Attribute VB_Ext_KEY = "Member20" ,"REG_081"
Attribute VB_Ext_KEY = "Member21" ,"REG_082"
Attribute VB_Ext_KEY = "Member22" ,"REG_084"
Attribute VB_Ext_KEY = "Member23" ,"REG_085"
Attribute VB_Ext_KEY = "Member24" ,"REG_086"
Attribute VB_Ext_KEY = "Member25" ,"REG_087"
Attribute VB_Ext_KEY = "Member26" ,"REG_088"
Attribute VB_Ext_KEY = "Member27" ,"REG_089"
Attribute VB_Ext_KEY = "Member28" ,"REG_090"
Attribute VB_Ext_KEY = "Member29" ,"REG_092"
Attribute VB_Ext_KEY = "Member30" ,"REG_095"
Attribute VB_Ext_KEY = "Member31" ,"REG_096"
Attribute VB_Ext_KEY = "Member32" ,"REG_097"
Attribute VB_Ext_KEY = "Member33" ,"REG_CAPITALSOCIAL"
Option Explicit

Dim objDom     As Object
Dim strTipoConsumidor As String                                      'F FISICA / J JURIDICA

Public Enum eTipoCheque
    CCF = 1
    Cheque_Lojista = 2
    Contra_Ordem = 3
    Contra_Ordem_Documento_Diferente = 4
    Contra_Ordem_Agencia_Diferente = 5
    Contra_ordem_Agencia_Conta_Diferente = 6
End Enum

Public Enum eTipoRespostaSpc
    Host = 1
    xml = 2
End Enum


Public Enum eTipoConsumidorSpcBrasil
    Fisica = 1
    Juridica = 2
End Enum

Public Enum eAmbienteSPCBrasil
    Producao = 1
    Homologacao = 2
End Enum


Dim Cod_Servico As String

'local variable(s) to hold property value(s)
Private mvarCodigo_Transacao As String                               'local copy
Private mvarVersao As String                                         'local copy
Private mvarCodigo_Operador As String                                'local copy
Private mvarSenha_Operador As String                                 'local copy
Private mvarTipo_Documento As String                                 'local copy
Private mvarCPF_CNPJ As String                                       'local copy
Private mvarRG As String                                             'local copy
Private mvarCodigo_Produto As String                                 'local copy
'local variable(s) to hold property value(s)
Private mvarTipo_Resposta As String                                  'local copy
Private mvarInformacaoes_Obito As String                             'local copy
Private mvarInformacoes_RG As String                                 'local copy
Private mvarHeader_Resposta As Header_Resposta
Private mvarREG_005 As REG_005
Private mvarREG_006 As REG_006
Private mvarREG_007 As REG_007
Private mvarREG_008 As REG_008
Private mvarREG_009 As REG_009
Private mvarREG_011 As REG_011
Private mvarREG_010 As REG_010
Private mvarREG_020 As REG_020
Private mvarREG_025 As REG_025
Private mvarREG_026 As REG_026
Private mvarREG_027 As REG_027
Private mvarREG_028 As REG_028
Private mvarREG_030 As REG_030
Private mvarREG_040 As REG_040
Private mvarREG_050 As REG_050
Private mvarREG_060 As REG_060
Private mvarREG_070 As REG_070
Private mvarREG_080 As REG_080
Private mvarREG_081 As REG_081
Private mvarREG_082 As REG_082
Private mvarREG_084 As REG_084
Private mvarREG_085 As REG_085
Private mvarREG_086 As REG_086
Private mvarREG_087 As REG_087
Private mvarREG_088 As REG_088
Private mvarREG_089 As REG_089
Private mvarREG_090 As REG_090
Private mvarREG_092 As REG_092
Private mvarREG_095 As REG_095
Private mvarREG_096 As REG_096
Private mvarREG_097 As REG_097
'local variable(s) to hold property value(s)
Private mvarMeio_de_Acesso As String                                 'local copy
'local variable(s) to hold property value(s)
Private mvarCodigo_Estacao As String                                 'local copy
'local variable(s) to hold property value(s)
Private mvarIDLiberacao As String                                    'local copy
'local variable(s) to hold property value(s)
Private mvarCNPJ_Empresa As String                                   'local copy
'local variable(s) to hold property value(s)
Private mvarMsg_Erro As String                                       'local copy
'local variable(s) to hold property value(s)
Private mvarErro As Boolean                                          'local copy
'local variable(s) to hold property value(s)
Private mvarStrResposta As String                                    'local copy
'local variable(s) to hold property value(s)
Private mvarUsarArquivo As String                                    'local copy
'local variable(s) to hold property value(s)
Private mvarGravar_Resposta_Arquivo As String                        'local copy
'local variable(s) to hold property value(s)
Private mvarGravar_String_Envio As String                            'local copy
'local variable(s) to hold property value(s)
Private mvarbytesRecebidos As String                                 'local copy
Private mvarProdutos As Produtos
Private mvarREG_CAPITALSOCIAL As REG_CAPITALSOCIAL
Private mvarREG_HISTORICOPAGAMENTO As REG_HISTORICOPGTO


Public Property Get REG_HISTORICOPAGAMENTO() As REG_HISTORICOPGTO
    If mvarREG_HISTORICOPAGAMENTO Is Nothing Then
        Set mvarREG_HISTORICOPAGAMENTO = New REG_HISTORICOPGTO
    End If


    Set REG_HISTORICOPAGAMENTO = mvarREG_HISTORICOPAGAMENTO
End Property


Public Property Set REG_HISTORICOPAGAMENTO(vData As REG_HISTORICOPGTO)
    Set mvarREG_HISTORICOPAGAMENTO = vData
End Property



Public Property Get REG_CAPITALSOCIAL() As REG_CAPITALSOCIAL
    Set REG_CAPITALSOCIAL = mvarREG_CAPITALSOCIAL
End Property


Public Property Set REG_CAPITALSOCIAL(vData As REG_CAPITALSOCIAL)
    Set mvarREG_CAPITALSOCIAL = vData
End Property



Public Property Get Produtos() As Produtos
    If mvarProdutos Is Nothing Then
        Set mvarProdutos = New Produtos
    End If


    Set Produtos = mvarProdutos
End Property


Public Property Set Produtos(vData As Produtos)
    Set mvarProdutos = vData
End Property

Public Property Let bytesRecebidos(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bytesRecebidos = 5
    mvarbytesRecebidos = vData
End Property


Public Property Get bytesRecebidos() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bytesRecebidos
    bytesRecebidos = mvarbytesRecebidos
End Property



Public Property Let Gravar_String_Envio(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Gravar_String_Envio = 5
    mvarGravar_String_Envio = vData
End Property


Public Property Get Gravar_String_Envio() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Gravar_String_Envio
    Gravar_String_Envio = mvarGravar_String_Envio
End Property



Public Property Let Gravar_Resposta_Arquivo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Gravar_Resposta_Arquivo = 5
    mvarGravar_Resposta_Arquivo = vData
End Property


Public Property Get Gravar_Resposta_Arquivo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Gravar_Resposta_Arquivo
    Gravar_Resposta_Arquivo = mvarGravar_Resposta_Arquivo
End Property



Public Property Let UsarArquivo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UsarArquivo = 5
    mvarUsarArquivo = vData
End Property


Public Property Get UsarArquivo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UsarArquivo
    UsarArquivo = mvarUsarArquivo
End Property





Private Sub Class_Initialize()


'create the mHeader_Resposta object when the SPC_Brasil_Consulta class is created
    Set mvarHeader_Resposta = New Header_Resposta
    'create the mREG_007 object when the SPC_Brasil_Consulta class is created
    Set mvarREG_007 = New REG_007
    'create the mREG_008 object when the SPC_Brasil_Consulta class is created
    Set mvarREG_008 = New REG_008
    'create the mREG_009 object when the SPC_Brasil_Consulta class is created
    Set mvarREG_009 = New REG_009
    'create the mREG_011 object when the SPC_Brasil_Consulta class is created
    Set mvarREG_011 = New REG_011
    'create the mREG_080 object when the SPC_Brasil_Consulta class is created
    Set mvarREG_080 = New REG_080
    'create the mREG_087 object when the SPC_Brasil_Consulta class is created
    Set mvarREG_087 = New REG_087
    'create the mREG_095 object when the SPC_Brasil_Consulta class is created
    Set mvarREG_095 = New REG_095

    'Set mySocket = Nothing
    'Set mySocket = New MSWinsockLib.Winsock

    'create the mREG_CAPITALSOCIAL object when the SPC_Brasil_Consulta class is created
    Set mvarREG_CAPITALSOCIAL = New REG_CAPITALSOCIAL
End Sub

Private Sub Class_Terminate()
    Set mvarREG_HISTORICOPAGAMENTO = Nothing
    Set mvarREG_CAPITALSOCIAL = Nothing
    Set mvarProdutos = Nothing
    Set mvarREG_097 = Nothing
    Set mvarREG_096 = Nothing
    Set mvarREG_095 = Nothing
    Set mvarREG_092 = Nothing
    Set mvarREG_090 = Nothing
    Set mvarREG_089 = Nothing
    Set mvarREG_088 = Nothing
    Set mvarREG_087 = Nothing
    Set mvarREG_086 = Nothing
    Set mvarREG_085 = Nothing
    Set mvarREG_084 = Nothing
    Set mvarREG_082 = Nothing
    Set mvarREG_081 = Nothing
    Set mvarREG_080 = Nothing
    Set mvarREG_070 = Nothing
    Set mvarREG_060 = Nothing
    Set mvarREG_050 = Nothing
    Set mvarREG_040 = Nothing
    Set mvarREG_030 = Nothing
    Set mvarREG_028 = Nothing
    Set mvarREG_027 = Nothing
    Set mvarREG_026 = Nothing
    Set mvarREG_025 = Nothing
    Set mvarREG_020 = Nothing
    Set mvarREG_010 = Nothing
    Set mvarREG_011 = Nothing
    Set mvarREG_009 = Nothing
    Set mvarREG_008 = Nothing
    Set mvarREG_007 = Nothing
    Set mvarREG_006 = Nothing
    Set mvarREG_005 = Nothing
    Set mvarHeader_Resposta = Nothing


End Sub

Public Property Let CNPJ_Empresa(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CNPJ_Empresa = 5
    mvarCNPJ_Empresa = vData
End Property

Public Property Get CNPJ_Empresa() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CNPJ_Empresa
    CNPJ_Empresa = mvarCNPJ_Empresa
End Property

Public Property Let Codigo_Estacao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Codigo_Estacao = 5
    mvarCodigo_Estacao = vData
    mvarCodigo_Estacao = PadL(mvarCodigo_Estacao, 16)
End Property

Public Property Get Codigo_Estacao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Codigo_Estacao
    Codigo_Estacao = mvarCodigo_Estacao
End Property

Public Property Let Codigo_Operador(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Codigo_Operador = 5
    mvarCodigo_Operador = vData
    mvarCodigo_Operador = StrZero(mvarCodigo_Operador, 10)
End Property

Public Property Get Codigo_Operador() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Codigo_Operador
    Codigo_Operador = mvarCodigo_Operador
End Property

Public Property Let Codigo_Produto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Codigo_Produto = 5
    mvarCodigo_Produto = vData
    mvarCodigo_Produto = StrZero(mvarCodigo_Produto, 2)
End Property

Public Property Get Codigo_Produto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Codigo_Produto
    Codigo_Produto = mvarCodigo_Produto
End Property

Public Property Let Codigo_Transacao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Codigo_Transacao = 5
    If vData = "" Then vData = "SPC100"
    mvarCodigo_Transacao = PadL(vData, 6)
End Property

Public Property Get Codigo_Transacao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Codigo_Transacao
    Codigo_Transacao = mvarCodigo_Transacao
End Property

Function Monta_String_Envio() As String

    Dim Texto  As String

    On Error GoTo Consulta_Error


    Msg_Erro = ""
    StrResposta = ""

    ''--//VALIDACAO DO ID
    If mvarIDLiberacao <> "MY2012A10SO000197246E" Then               'CodigoPadrao
        If (mvarIDLiberacao) <> (CodigoLiberacao(mvarCNPJ_Empresa)) Then
            MsgBox "Código de  Liberação Inválido", vbCritical, "Erro"
            Monta_String_Envio = "ERRO"
            Exit Function
        End If
    End If
    ''--// FIM

    Texto = mvarCodigo_Transacao                                     ' 001-006  TAMANHO 06  DEFAULT SPC100
    Texto = Texto & mvarVersao                                       ' 007-009  TAMANHO 03 DEFAULT  014
    Texto = Texto & mvarMeio_de_Acesso                               ' 010-011  TAMANHO 02
    Texto = Texto & mvarCodigo_Estacao                               ' 012-027  TAMANHO 16
    Texto = Texto & mvarCodigo_Operador                              ' 028-037  TAMANHO 10
    Texto = Texto & mvarSenha_Operador                               ' 038-045  TAMANHO 08
    Texto = Texto & mvarTipo_Documento                               ' 046-046  TAMANHO 01
    Texto = Texto & mvarCPF_CNPJ                                     ' 047-060  TAMANHO 14
    Texto = Texto & mvarRG                                           ' 061-076  TAMANHO 16
    Texto = Texto & mvarCodigo_Produto                               ' 077-078  TAMANHO 02

    '' CHEQUES
    Texto = Texto & " "                                              ' 079-079  TAMANHO 01   TIPO DE ENTRADA DE DADOS DO CHEQUE
    Texto = Texto & "000"                                            ' 080-082  TAMANHO 03   CODIGO DO BANCO
    Texto = Texto & "0000"                                           ' 083-086  TAMANHO 04  CODIGO DA AGENCIA
    Texto = Texto & Space$(15)                                       ' 087-101  TAMANHO 15  NUMERO DA CONTA
    Texto = Texto & Space$(2)                                        ' 102-103  TAMANHO 02  DIGITO DA CONTA CORRENTE
    Texto = Texto & "000000"                                         ' 104-109  TAMANHO 06  CHEQUE INICIAL
    Texto = Texto & "0"                                              ' 110-110  TAMANHO 01  DIGITO DO CHEQUE INICIAL
    Texto = Texto & "00000000"                                       ' 111-118  TAMANHO 08  CMC7-1
    Texto = Texto & "0000000000"                                     ' 119-128  TAMANHO 10  CMC7-2
    Texto = Texto & "000000000000"                                   ' 129-140  TAMANHO 12  CMC7-3
    Texto = Texto & "00"                                             ' 141-142  TAMANHO 02  QUANTIDADE DE CHEQUES
    Texto = Texto & "000000000000"                                   ' 143-154  TAMANHO 12  VALOR DE CHEQUES
    Texto = Texto & Space$(2)                                        ' 155-156  TAMANHO 02  TIPO CREDITO
    Texto = Texto & "000"                                            ' 157-159  TAMANHO 03  INTERVALO ENTRE PARCELAS
    Texto = Texto & "00000000"                                       ' 160-167  TAMANHO 08  DATA VENCIMENTO PRIMEIRA PARCELA
    Texto = Texto & "000"                                            ' 168-170  TAMANHO 03  QUANTIDADE PARCELAS
    Texto = Texto & "000000000000"                                   ' 171-182  TAMANHO 12  VALOR TOTAL COMPRA
    ''FIM CHEQUES

    Texto = Texto & "00000000"                                       ' 183-190  TAMANHO 08 CEP ORIGEM DA CONSULTA
    Texto = Texto & "000000"                                         ' 191-196  TAMANHO 06 DATA ABERTURA CONTA CORRENTE
    Texto = Texto & "000"                                            ' 197-199  TAMANHO 03 PRACA DE COMPENSACAO DO CHEQUE

    Texto = Texto & mvarTipo_Resposta                                ' 200-200  TAMANHO 01
    Texto = Texto & Space$(8)                                        ' 201-208  TAMANHO 08  NOVA SENHA
    Texto = Texto & "00000000"                                       ' 209-216  TAMANHO 08  CANCELAMENTO DE GARANTIA DE CHEQUES
    Texto = Texto & "00"                                             ' 217-218  TAMANHO 02  MOTIVO DO CANCELAMENTO DE GARANTIA DE CHEQUES
    Texto = Texto & "00000000"                                       ' 219-226  TAMANHO 08  ENTIDADE URA
    Texto = Texto & "0000000000"                                     ' 227-236  TAMANHO 10  CODIGO AUTENTICACAO URA
    Texto = Texto & "000"                                            ' 237-239  TAMANHO 03  SEQ URA
    Texto = Texto & Space$(12)                                       ' 240-251  TAMANHO 12  RESERVADO BRANCO

    'Texto = Texto & Space$(1)                              ' 252-252  TAMANHO 01  INFORMACOES COMPLEMENTARES

    Texto = Texto & Space$(8)                                        ' 253-260  TAMANHO 08  CEP DO CONSUMIDOR
    Texto = Texto & "N"                                              ' 261-261  TAMANHO 01  CALCULAR SCORE DE CREDITO
    Texto = Texto & "0"                                              ' 261-261  TAMANHO 01  CALCULAR SCORE DE CREDITO

    'Texto = Texto & " "                                    ' 262-262  TAMANHO 01
    'Texto = Texto & mvarInformacaoes_Obito                 ' 263-263  TAMANHO 01
    'Texto = Texto & mvarInformacoes_RG                     ' 264-264  TAMANHO 01
    'Texto = Texto & Space$(20)                             ' 265-284  TAMANHO 20
    'Texto = Texto & Space$(8)                              ' 285-292  TAMANHO 08
    'Texto = Texto & Space$(2)                              ' 293-294  TAMANHO 02
    'Texto = Texto & Space$(1)                              ' 295-295  TAMANHO 01
    'Texto = Texto & Chr(13)                                '          Byte Finalizador



    Erro = False

    If Gravar_String_Envio <> "" Then
        Open Gravar_String_Envio For Output As #1
        Print #1, Texto & vbCrLf
        Close #1
    End If

    Monta_String_Envio = Texto & vbCrLf


    On Error GoTo 0
    Exit Function

Consulta_Error:
    Monta_String_Envio = "ERRO"
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Consulta of Módulo de classe SPC_Brasil_Consulta"
End Function

Public Property Let CPF_CNPJ(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CPF_CNPJ = 5
    mvarCPF_CNPJ = vData
    mvarCPF_CNPJ = Trim(RemoveNaoNumericos(mvarCPF_CNPJ))
    mvarCPF_CNPJ = StrZero(mvarCPF_CNPJ, 14)

End Property

Public Property Get CPF_CNPJ() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CPF_CNPJ
    CPF_CNPJ = mvarCPF_CNPJ
End Property

Public Property Let Erro(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Erro = 5
    mvarErro = vData
End Property

Public Property Get Erro() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Erro
    Erro = mvarErro
End Property

Private Function Formata_Data(ByVal Data As String) As String
    If Data <> "" Then
        Data = Left(Data, 2) & "/" & Mid(Data, 3, 2) & "/" & Right(Data, 4)
    End If

    Formata_Data = Data
End Function

Private Function Formata_Hora(ByVal Hora As String) As String
    If Hora <> "" Then
        Hora = Left(Hora, 2) & ":" & Mid(Hora, 3, 2) & ":" & Right(Hora, 2)
    End If

    Formata_Hora = Hora
End Function

Private Function Formata_Valor(ByVal Valor As String) As String
    If Valor <> "" Then
        Valor = Format(Valor / 100, "0.00")
    End If
    Formata_Valor = Format(Valor, "#,##0.00")
End Function

Private Function Formata_CEP(ByVal CEP As String) As String
    If CEP <> "" Then
        CEP = Replace(CEP, "-", "")
        CEP = Format(CEP, "00000-000")
    End If

    Formata_CEP = CEP

End Function

Private Function Formata_Fone(ByVal Texto As String) As String
    Dim tam    As Integer
    Texto = RemoveNaoNumericos(Texto)
    tam = Len(Texto)

    Dim NonoDig As Boolean

    Select Case Len(Texto)                                           'Left(Texto, 2)

    Case 11, 9                                                       'Not "11", "21", "91", "92", "93", "94", "95", "96", "97", "98", "99"
        NonoDig = True

    Case Else
        NonoDig = False
    End Select

    If tam = 0 Then
        Formata_Fone = ""
        Exit Function
    ElseIf tam > 10 Then
        If NonoDig = False Then
            Texto = Left(Texto, 10)
            tam = 10
        End If
    End If

    If NonoDig = True And (tam = 9 Or tam >= 11) Then
        Texto = Format(Texto, "(00)00000-0000")
    Else
        Texto = String(10 - tam, "0") & Texto
        Texto = Format(Texto, "(00)0000-0000")
    End If

    Formata_Fone = Texto
    Debug.Print Formata_Fone
End Function


Public Property Get Header_Resposta() As Header_Resposta
    Set Header_Resposta = mvarHeader_Resposta
End Property

Public Property Set Header_Resposta(vData As Header_Resposta)
    Set mvarHeader_Resposta = vData
End Property

Public Property Let IDLiberacao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IDLiberacao = 5
    mvarIDLiberacao = vData
End Property

Public Property Get IDLiberacao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IDLiberacao
    IDLiberacao = mvarIDLiberacao
End Property

Public Property Let Informacaoes_Obito(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Informacaoes_Obito = 5
    mvarInformacaoes_Obito = vData
End Property

Public Property Get Informacaoes_Obito() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Informacaoes_Obito
    Informacaoes_Obito = mvarInformacaoes_Obito
End Property

Public Property Let Informacoes_RG(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Informacoes_RG = 5
    mvarInformacoes_RG = vData
    mvarInformacoes_RG = "N"                                         ' PADRAO PQ SO FUNCIONA NO CEARA ESSA OPCAO
End Property

Public Property Get Informacoes_RG() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Informacoes_RG
    Informacoes_RG = mvarInformacoes_RG
End Property

Public Property Let Meio_de_Acesso(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Meio_de_Acesso = 5
    mvarMeio_de_Acesso = vData
    If mvarMeio_de_Acesso = "" Then mvarMeio_de_Acesso = "20"
    mvarMeio_de_Acesso = StrZero(mvarMeio_de_Acesso, 2)
End Property

Public Property Get Meio_de_Acesso() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Meio_de_Acesso
    Meio_de_Acesso = mvarMeio_de_Acesso
End Property

Sub Monta_Header(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
    Dim Caminho As String

    If TipoResposta = Host Then
        With mvarHeader_Resposta
            .Tipo_Documento = Trim(Mid(Texto, 46, 1))
            .CPF_CNPJ = Trim(Mid(Texto, 47, 14))
            If .Tipo_Documento = "1" Then .CPF_CNPJ = Right(.CPF_CNPJ, 11)
            .RG = Trim(Mid(Texto, 61, 16))
            .Nome = Trim(Mid(Texto, 77, 60))
            .Data_Nascimento_Fundacao = Formata_Data(Trim(Mid(Texto, 137, 8)))
            .Endereco = Trim(Mid(Texto, 145, 60))
            .Complemento = Trim(Mid(Texto, 205, 30))
            .Bairro = Trim(Mid(Texto, 235, 30))
            .Cidade = Trim(Mid(Texto, 265, 30))
            .UF = Trim(Mid(Texto, 295, 2))
            .CEP = Formata_CEP(Trim(Mid(Texto, 297, 8)))
            .Telefone = Trim(Mid(Texto, 305, 8))
            .Data_Ultima_Atu_Endereco = Formata_Data(Trim(Mid(Texto, 313, 8)))
            .Total_SPC = Trim(Mid(Texto, 321, 4))
            .Total_Cheque_CCF = Trim(Mid(Texto, 325, 4))
            .Total_Cheque_Lojista = Trim(Mid(Texto, 329, 4))
            .Total_Cheque_ContraOrdenados = Trim(Mid(Texto, 333, 4))
            .Total_Alertas = Trim(Mid(Texto, 337, 4))
            .Qtde_Dias_Consulta_Realizada = Trim(Mid(Texto, 341, 3))
            .Total_Consultas_Realizadas = Trim(Mid(Texto, 344, 4))
            .Total_Protesto = Trim(Mid(Texto, 348, 4))
            .Total_Acoes = Trim(Mid(Texto, 352, 4))
            .Total_Ocorrencias_Credito_Concedidos = Trim(Mid(Texto, 356, 4))
            .Total_Cheques_Contra_Ordenados_DIF_Consultado = Trim(Mid(Texto, 360, 4))
            .Total_Cheques_ContraOrdenados_AG_Diferente = Trim(Mid(Texto, 364, 4))
            .Total_Cheques_ContraOrdenados_AGConta_Diferente = Trim(Mid(Texto, 368, 4))
            .Qtde_Dias_Contumacia = Trim(Mid(Texto, 372, 3))
            .Total_Cheques_Sustados_Linha21 = Trim(Mid(Texto, 375, 4))
            .Numero_Protocolo = Trim(Mid(Texto, 408, 12))
            .Data_Resposta = Formata_Data(Trim(Mid(Texto, 420, 8)))
            .Sucesso_Transacao = Trim(Mid(Texto, 434, 1))
            .Horas_Resposta = Formata_Hora(Trim(Mid(Texto, 428, 6)))
            .Indicativo_Consulta_Parcial = Trim(Mid(Texto, 445, 1))
        End With


        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 455, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoResposta = xml Then

        With mvarHeader_Resposta
            Caminho = ""
            .Data_Resposta = Formata_Data_XML(PegaValorCampoXML(Caminho, 0, "", "data"))
            .Horas_Resposta = Mid(PegaValorCampoXML(Caminho, 0, "", "data"), 12, 8)
            .Numero_Protocolo = PegaValorCampoXML(Caminho, 0, "protocolo", "numero")
            .Numero_Protocolo = .Numero_Protocolo & PegaValorCampoXML(Caminho, 0, "protocolo", "digito")
            '
            If strTipoConsumidor = "F" Then
                '--//INFORMACOES ESPECIFICAS PARA PESSOA FISICA
                Caminho = "consumidor/consumidor-pessoa-fisica"
                .CPF_CNPJ = PegaValorCampoXML(Caminho, 0, "cpf", "numero")
                .RG = PegaValorCampoXML(Caminho, 0, "", "numero-rg")
                .Nome = PegaValorCampoXML(Caminho, 0, "", "nome")
                .Data_Nascimento_Fundacao = Formata_Data_XML(PegaValorCampoXML(Caminho, 0, "", "data-nascimento"))
                .Telefone = PegaValorCampoXML(Caminho, 0, "telefone-residencial", "numero-ddd")
                .Telefone = .Telefone & PegaValorCampoXML(Caminho, 0, "telefone-residencial", "numero")
                .Telefone = Formata_Fone(.Telefone)
                '--// FIM DAS INFORMACOES
            ElseIf strTipoConsumidor = "J" Then
                '--//INFORMACOES ESPECIFICAS PARA PESSOA JURIDICA
                Caminho = "consumidor/consumidor-pessoa-juridica"
                .CPF_CNPJ = PegaValorCampoXML(Caminho, 0, "cnpj", "numero")
                .RG = ""                                             'PegaValorCampoXML(Caminho, 0, "", "inscricao-estadual")
                .Nome = PegaValorCampoXML(Caminho, 0, "", "razao-social")
                .Data_Nascimento_Fundacao = Formata_Data_XML(PegaValorCampoXML(Caminho, 0, "", "data-fundacao"))
                '--// FIM DAS INFORMACOES
            End If

            '--//INFORMACOES COMUNS ENTRE PESSOA FISICA E JURIDICA
            .Tipo_Documento = IIf(Len(.CPF_CNPJ) = 11, "CPF", "CNPJ")
            .Endereco = PegaValorCampoXML(Caminho, 0, "endereco", "logradouro")
            .Endereco = .Endereco & "," & PegaValorCampoXML(Caminho, 0, "endereco", "numero")
            .Complemento = PegaValorCampoXML(Caminho, 0, "endereco", "complemento")
            .Bairro = PegaValorCampoXML(Caminho, 0, "endereco", "bairro")
            .Cidade = PegaValorCampoXML(Caminho, 0, "endereco/cidade", "nome")
            .UF = PegaValorCampoXML(Caminho, 0, "endereco/cidade/estado", "sigla-uf")
            .CEP = Formata_CEP(PegaValorCampoXML(Caminho, 0, "endereco", "cep"))
            '--// FIM DAS INFORMACOES

            '.Data_Ultima_Atu_Endereco = PegaValorCampoXML( Caminho, "", "numero-rg")

            '--// RESUMO DAS INFORMAÇÕES
            .Total_SPC = PegaValorCampoXML("spc", 0, "resumo", "quantidade-total")
            .Total_SPC = Format(.Total_SPC, "0000")

            '
            .Total_Cheque_CCF = PegaValorCampoXML("ccf", 0, "resumo", "quantidade-total")
            .Total_Cheque_CCF = Format(.Total_Cheque_CCF, "0000")
            '
            .Total_Cheque_Lojista = PegaValorCampoXML("cheque-lojista", 0, "resumo", "quantidade-total")
            .Total_Cheque_Lojista = Format(.Total_Cheque_Lojista, "0000")
            '
            .Total_Cheque_ContraOrdenados = PegaValorCampoXML("contra-ordem", 0, "resumo", "quantidade-total")
            .Total_Cheque_ContraOrdenados = Format(.Total_Cheque_ContraOrdenados, "0000")
            '
            .Total_Alertas = PegaValorCampoXML("alerta-documento", 0, "resumo", "quantidade-total")
            .Total_Alertas = Format(.Total_Alertas, "0000")
            '
            .Total_InfoPoderJudiciario = PegaValorCampoXML("informacao-poder-judiciario", 0, "resumo", "quantidade-total")
            .Total_InfoPoderJudiciario = Format(.Total_InfoPoderJudiciario, "0000")
            '
            .Qtde_Dias_Consulta_Realizada = PegaValorCampoXML("consulta-realizada", 0, "", "quantidade-dias-consultados")
            .Qtde_Dias_Consulta_Realizada = Format(.Qtde_Dias_Consulta_Realizada, "0000")
            '
            .Total_Consultas_Realizadas = PegaValorCampoXML("consulta-realizada", 0, "resumo", "quantidade-total")
            .Total_Consultas_Realizadas = Format(.Total_Consultas_Realizadas, "0000")
            '
            .Total_Protesto = PegaValorCampoXML("protesto", 0, "resumo", "quantidade-total")
            .Total_Protesto = Format(.Total_Protesto, "0000")
            '
            .Total_Acoes = PegaValorCampoXML("acao", 0, "resumo", "quantidade-total")
            .Total_Acoes = Format(.Total_Acoes, "0000")
            '
            .Total_Ocorrencias_Credito_Concedidos = PegaValorCampoXML("credito-concedido", 0, "resumo", "quantidade-total")
            .Total_Ocorrencias_Credito_Concedidos = Format(.Total_Ocorrencias_Credito_Concedidos, "0000")
            '
            .Total_Cheques_Contra_Ordenados_DIF_Consultado = PegaValorCampoXML("contra-ordem-documento-diferente", 0, "resumo", "quantidade-total")
            .Total_Cheques_Contra_Ordenados_DIF_Consultado = Format(.Total_Cheques_Contra_Ordenados_DIF_Consultado, "0000")
            '
            .Total_Cheques_ContraOrdenados_AG_Diferente = PegaValorCampoXML("contra-ordem-agencia-diferente", 0, "resumo", "quantidade-total")
            .Total_Cheques_ContraOrdenados_AG_Diferente = Format(.Total_Cheques_ContraOrdenados_AG_Diferente, "0000")
            '
            .Total_Cheques_ContraOrdenados_AGConta_Diferente = PegaValorCampoXML("contra-ordem-agencia-conta-diferente", 0, "resumo", "quantidade-total")
            .Total_Cheques_ContraOrdenados_AGConta_Diferente = Format(.Total_Cheques_ContraOrdenados_AGConta_Diferente, "0000")
            '
            On Error Resume Next
            .Qtde_Dias_Contumacia = DateDiff("d", Formata_Data_XML(PegaValorCampoXML("contumacia", 0, "", "data-inicial-periodo-consultado")), Formata_Data_XML(PegaValorCampoXML("contumacia", 0, "", "data-final-periodo-consultado")))
            On Error GoTo 0
            .Total_Cheques_Sustados_Linha21 = PegaValorCampoXML("contumacia", 0, "resumo", "quantidade-total")
            .Total_Cheques_Sustados_Linha21 = Format(.Total_Cheques_Sustados_Linha21, "0000")
            .Sucesso_Transacao = 1

            '
            .Total_RegistroConsulta = PegaValorCampoXML("registro-consulta", 0, "resumo", "quantidade-total")
            .Total_RegistroConsulta = Format(.Total_RegistroConsulta, "0000")
            '
            .Total_HistoricoPagamento = PegaValorCampoXML("historico-pagamento", 0, "resumo", "quantidade-total")
            .Total_HistoricoPagamento = Format(.Total_HistoricoPagamento, "0000")
            '
            .Total_ControleSocietario = PegaValorCampoXML("socio", 0, "resumo", "quantidade-total")
            .Total_ControleSocietario = Format(.Total_ControleSocietario, "0000")
            '
            .Total_PendenciaSerasa = PegaValorCampoXML("pendencia-financeira", 0, "resumo", "quantidade-total")
            .Total_PendenciaSerasa = Format(.Total_PendenciaSerasa, "0000")

            '
            .Total_ParticipacaoFalencia = PegaValorCampoXML("participacao-falencia", 0, "resumo", "quantidade-total")
            .Total_ParticipacaoFalencia = Format(.Total_ParticipacaoFalencia, "0000")






        End With

    End If

End Sub

Sub Monta_Reg_005(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)

'' -- Resposta de Bases Externas Indisponíveis - máximo de 10 ocorrências

    If TipoResposta = Host Then
        With REG_005.Add
            .Codigo_Base = Trim(Mid(Texto, 3, 5))
            .Descricao_Base = Trim(Mid(Texto, 8, 50))
            .Mensagem = Trim(Mid(Texto, 58, 80))
        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 138, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoResposta = xml Then

        With REG_005.Add
            .Codigo_Base = ""                                        'XML NAO RETORNA ESSES DADOS
            .Descricao_Base = ""                                     'XML NAO RETORNA ESSES DADOS
            .Mensagem = ""                                           'XML NAO RETORNA ESSES DADOS
        End With

    End If

End Sub

Sub Monta_Reg_006(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
    Dim Caminho As String
    Dim I      As Integer
    '' -- Resposta de Grafia PJ - máximo de 10 ocorrências

    If TipoResposta = Host Then
        With REG_006.Add
            .Razao_Social = Trim(Mid(Texto, 3, 100))
        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 108, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoResposta = xml Then
        Caminho = "grafia-pj/detalhe-grafia-pj "
        For I = 0 To objDom.selectNodes("resultado/" & Caminho).length - 1
            With REG_006.Add
                .Razao_Social = PegaValorCampoXML(Caminho, I, "", "razao-social")
            End With
        Next I
    End If
End Sub

Sub Monta_Reg_007(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
'' -- Resposta de informação de óbito

    If TipoResposta = Host Then
        With REG_007
            .Data_Obito = Formata_Data(Trim(Mid(Texto, 3, 8)))
            .Cartorio = Trim(Mid(Texto, 11, 80))
            .Cidade = Trim(Mid(Texto, 91, 50))
            .UF = Trim(Mid(Texto, 141, 2))
        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 143, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoResposta = xml Then
        With REG_007
            .Data_Obito = ""                                         ' NÃO ACHEI NENHUM CPF COM ESSE RETORNO
            .Cartorio = ""                                           ' NÃO ACHEI NENHUM CPF COM ESSE RETORNO
            .Cidade = ""                                             ' NÃO ACHEI NENHUM CPF COM ESSE RETORNO
            .UF = ""                                                 ' NÃO ACHEI NENHUM CPF COM ESSE RETORNO
        End With
    End If
End Sub

Sub Monta_Reg_008(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
'' -- Resposta de Score de crédito

    Dim Caminho As String

    If TipoResposta = Host Then
        With REG_008
            .Horizonte = Trim(Mid(Texto, 3, 2))
            .Valor = Val(Trim(Mid(Texto, 5, 3)))
            .Probabilidade = Formata_Valor(Trim(Mid(Texto, 8, 6)))
            .Classe_Risco = Trim(Mid(Texto, 14, 40))
            .Menssagem_Curta = Trim(Mid(Texto, 54, 40))
            .Menssagem_Interpretativa_Score = Trim(Mid(Texto, 94, 450))
            .Messagem_Interpretativa_Probabilidade = Trim(Mid(Texto, 544, 150))
        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 694, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoResposta = xml Then
        Caminho = "spc-score-12-meses/detalhe-spc-score-12-meses"
        With REG_008
            .Horizonte = PegaValorCampoXML(Caminho, 0, "", "horizonte")
            .Valor = PegaValorCampoXML(Caminho, 0, "", "score")
            .Probabilidade = PegaValorCampoXML(Caminho, 0, "", "probabilidade")
            .Classe_Risco = PegaValorCampoXML(Caminho, 0, "", "classe")
            .Menssagem_Curta = ""
            .Menssagem_Interpretativa_Score = PegaValorCampoXML(Caminho, 0, "", "mesagem-interpretativa-score")
            .Messagem_Interpretativa_Probabilidade = ""
        End With

    End If

End Sub

Sub Monta_Reg_009(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
'' --SCORE DE CREDITO

    If TipoResposta = Host Then
        With REG_009
            .Menssagem_Score = Trim(Mid(Texto, 3, 100))
        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 103, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoResposta = xml Then
        With REG_009
            .Menssagem_Score = ""
        End With
    End If
End Sub

Sub Monta_Reg_010(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
    Dim Caminho As String
    Dim I      As Integer

    If TipoResposta = Host Then
        With REG_010.Add
            .Data_Inclusao = Formata_Data(Trim(Mid(Texto, 3, 8)))
            .Data_Vencimento = Formata_Data(Trim(Mid(Texto, 11, 8)))
            .Comprador_Fiador = Trim(Mid(Texto, 19, 1))
            .Numero_Contrato = Trim(Mid(Texto, 20, 30))
            .Valor_debito = Formata_Valor(Trim(Mid(Texto, 50, 15)))
            .Nome_Associado = Trim(Mid(Texto, 65, 60))
            .Cidade = Trim(Mid(Texto, 125, 30))
            .UF = Trim(Mid(Texto, 155, 2))
            .Nome_Consumidor = Trim(Mid(Texto, 157, 60))
            .Entidade_Origem = Trim(Mid(Texto, 217, 30))

        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 247, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoResposta = xml Then
        Caminho = "spc/detalhe-spc"
        For I = 0 To objDom.selectNodes("resultado/" & Caminho).length - 1
            With REG_010.Add
                .Data_Inclusao = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-inclusao"))
                .Data_Vencimento = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-vencimento"))
                .Comprador_Fiador = PegaValorCampoXML(Caminho, I, "", "comprador-fiador-avalista")
                .Numero_Contrato = PegaValorCampoXML(Caminho, I, "", "contrato")
                .Valor_debito = PegaValorCampoXML(Caminho, I, "", "valor")
                .Nome_Associado = PegaValorCampoXML(Caminho, I, "", "nome-associado")
                .Entidade_Origem = PegaValorCampoXML(Caminho, I, "", "nome-entidade")
                .Nome_Consumidor = ""
                .Cidade = PegaValorCampoXML(Caminho, I, "cidade-associado", "nome")
                .UF = PegaValorCampoXML(Caminho, I, "cidade-associado/estado", "sigla-uf")


            End With
        Next I
    End If

End Sub

Sub Monta_Reg_020(ByVal Texto As String, ByVal TipoCheque As eTipoCheque, ByVal TipoResposta As eTipoRespostaSpc)
    Dim Caminho As String
    Dim I      As Integer
    Dim strTipoCheque As String

    '' RESPOSTA DOS CHEQUES

    '--// CCF NO XML
    '--// CHEQUE LOJISTA NO XML
    '--// CONTRA ORDEM NO XML

    If TipoResposta = Host Then
        With REG_020.Add
            .Tipo_Ocorrencia = Trim(Mid(Texto, 3, 1))
            .Data_Inclusao = Formata_Data(Trim(Mid(Texto, 4, 8)))
            .Hora_Inclusao = Formata_Hora(Trim(Mid(Texto, 12, 6)))
            .Data_Ocorrencia = Formata_Data(Trim(Mid(Texto, 18, 8)))
            .Data_Emissao = Formata_Data(Trim(Mid(Texto, 26, 8)))
            .Origem = Trim(Mid(Texto, 34, 60))
            .Codigo_Banco = Trim(Mid(Texto, 94, 3))
            .Nome_Banco = Trim(Mid(Texto, 97, 30))
            .Codigo_Agencia = Trim(Mid(Texto, 127, 4))
            .Conta_Corrente = Trim(Mid(Texto, 131, 12))
            .Numero_Cheque = Trim(Mid(Texto, 143, 6))
            .Numero_Cheque_Final = Trim(Mid(Texto, 149, 6))
            .Qtde_Ocorrencias = Trim(Mid(Texto, 155, 4))
            .Valor_Documento = Formata_Valor(Trim(Mid(Texto, 159, 15)))
            .Codigo_Alinea = Trim(Mid(Texto, 174, 2))
            .Motivo = Trim(Mid(Texto, 176, 20))
            .Nome_Consumidor = Trim(Mid(Texto, 196, 60))
            .Entidade_Origem = Trim(Mid(Texto, 256, 30))
        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 286, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoResposta = xml Then

        If TipoCheque = CCF Then
            strTipoCheque = "CCF"
            Caminho = "ccf/detalhe-ccf"
        ElseIf TipoCheque = Cheque_Lojista Then
            strTipoCheque = "CHQUE LOJISTA"
            Caminho = "cheque-lojista/detalhe-cheque-lojista"
        ElseIf TipoCheque = Contra_Ordem Then
            strTipoCheque = "CONTRA ORDEM"
            Caminho = "contra-ordem/detalhe-contra-ordem"
        ElseIf TipoCheque = Contra_Ordem_Documento_Diferente Then
            strTipoCheque = "CONTRA ORDEM DOCUMENTO DIFERENTE"
            Caminho = "contra-ordem-documento-diferente/detalhe-contra-ordem-documento-diferente"
        ElseIf TipoCheque = Contra_Ordem_Agencia_Diferente Then
            strTipoCheque = "CONTRA ORDEM AGENCIA DIFERENTE"
            Caminho = "contra-ordem-agencia-diferente/detalhe-contra-ordem-agencia-diferente"
        ElseIf TipoCheque = Contra_ordem_Agencia_Conta_Diferente Then
            strTipoCheque = "CONTRA ORDEM AGENCIA/CONTA DIFERENTE"
            Caminho = "contra-ordem-agencia-conta-diferente/detalhe-contra-ordem-agencia-conta-diferente"
        End If

        For I = 0 To objDom.selectNodes("resultado/" & Caminho).length - 1
            With REG_020.Add
                .Tipo_Ocorrencia = strTipoCheque
                If TipoCheque = Cheque_Lojista Then
                    .Data_Inclusao = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-inclusao"))
                    .Hora_Inclusao = Mid(PegaValorCampoXML(Caminho, I, "", "data-inclusao"), 12, 8)
                    .Data_Emissao = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-inclusao"))
                    .Data_Ocorrencia = .Data_Inclusao
                    .Origem = PegaValorCampoXML(Caminho, I, "", "nome-associado")
                    .Qtde_Ocorrencias = 1
                    .Codigo_Alinea = PegaValorCampoXML(Caminho, I, "alinea", "codigo")
                    .Motivo = PegaValorCampoXML(Caminho, I, "alinea", "descricao")
                    .Codigo_Banco = PegaValorCampoXML(Caminho, I, "cheque-inicial/dados-bancarios/banco", "codigo")
                    .Nome_Banco = PegaValorCampoXML(Caminho, I, "cheque-inicial/dados-bancarios/banco", "nome")
                    .Codigo_Agencia = PegaValorCampoXML(Caminho, I, "cheque-inicial/dados-bancarios", "numero-agencia")
                    .Conta_Corrente = PegaValorCampoXML(Caminho, I, "cheque-inicial/dados-bancarios", "numero-conta-corrente")
                    .Numero_Cheque = PegaValorCampoXML(Caminho, I, "cheque-inicial", "numero")
                    .Numero_Cheque_Final = ""                        'PegaValorCampoXML(Caminho, I, "cheque-final", "numero")
                    .Entidade_Origem = PegaValorCampoXML(Caminho, I, "", "nome-entidade")
                ElseIf TipoCheque = CCF Then
                    .Data_Inclusao = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-ultimo-cheque"))
                    .Hora_Inclusao = Mid(PegaValorCampoXML(Caminho, I, "", "data-ultimo-cheque"), 12, 8)
                    .Data_Emissao = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-ultimo-cheque"))
                    .Data_Ocorrencia = .Data_Inclusao
                    .Origem = PegaValorCampoXML(Caminho, I, "", "origem")
                    .Qtde_Ocorrencias = PegaValorCampoXML(Caminho, I, "", "quantidade")
                    .Codigo_Alinea = PegaValorCampoXML(Caminho, I, "motivo", "codigo")
                    .Motivo = PegaValorCampoXML(Caminho, I, "motivo", "descricao")
                    .Codigo_Banco = PegaValorCampoXML(Caminho, I, "ultimo-cheque/dados-bancarios/banco", "codigo")
                    .Nome_Banco = PegaValorCampoXML(Caminho, I, "ultimo-cheque/dados-bancarios/banco", "nome")
                    .Codigo_Agencia = PegaValorCampoXML(Caminho, I, "ultimo-cheque/dados-bancarios", "numero-agencia")
                ElseIf TipoCheque = Contra_Ordem Then
                    .Data_Inclusao = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-inclusao"))
                    .Hora_Inclusao = Mid(PegaValorCampoXML(Caminho, I, "", "data-inclusao"), 12, 8)
                    .Data_Emissao = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-inclusao"))
                    .Data_Ocorrencia = .Data_Inclusao
                    .Origem = PegaValorCampoXML(Caminho, I, "", "origem")
                    .Qtde_Ocorrencias = 1
                    .Codigo_Alinea = PegaValorCampoXML(Caminho, I, "motivo", "codigo")
                    .Motivo = PegaValorCampoXML(Caminho, I, "motivo", "descricao")
                    .Codigo_Banco = PegaValorCampoXML(Caminho, I, "cheque-inicial/dados-bancarios/banco", "codigo")
                    .Nome_Banco = PegaValorCampoXML(Caminho, I, "cheque-inicial/dados-bancarios/banco", "nome")
                    .Codigo_Agencia = PegaValorCampoXML(Caminho, I, "cheque-inicial/dados-bancarios", "numero-agencia")
                    .Conta_Corrente = PegaValorCampoXML(Caminho, I, "cheque-inicial/dados-bancarios", "numero-conta-corrente")
                    .Numero_Cheque = PegaValorCampoXML(Caminho, I, "cheque-inicial", "numero")
                    .Numero_Cheque_Final = PegaValorCampoXML(Caminho, I, "cheque-final", "numero")
                    .Entidade_Origem = .Origem
                ElseIf TipoCheque = Contra_Ordem_Documento_Diferente Or TipoCheque = Contra_Ordem_Agencia_Diferente Or TipoCheque = Contra_ordem_Agencia_Conta_Diferente Then
                    .Data_Inclusao = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-inclusao"))
                    .Hora_Inclusao = Mid(PegaValorCampoXML(Caminho, I, "", "data-inclusao"), 12, 8)
                    .Data_Emissao = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-inclusao"))
                    .Data_Ocorrencia = .Data_Inclusao
                    .Origem = PegaValorCampoXML(Caminho, I, "", "origem")
                    .Qtde_Ocorrencias = 1
                    .Codigo_Alinea = PegaValorCampoXML(Caminho, I, "motivo", "codigo")
                    .Motivo = PegaValorCampoXML(Caminho, I, "motivo", "descricao")
                    .Codigo_Banco = PegaValorCampoXML(Caminho, I, "cheque-inicial/dados-bancarios/banco", "codigo")
                    .Nome_Banco = PegaValorCampoXML(Caminho, I, "cheque-inicial/dados-bancarios/banco", "nome")
                    .Codigo_Agencia = PegaValorCampoXML(Caminho, I, "cheque-inicial/dados-bancarios", "numero-agencia")
                    .Conta_Corrente = PegaValorCampoXML(Caminho, I, "cheque-inicial/dados-bancarios", "numero-conta-corrente")
                    .Numero_Cheque = PegaValorCampoXML(Caminho, I, "cheque-inicial", "numero")
                    .Numero_Cheque_Final = PegaValorCampoXML(Caminho, I, "cheque-final", "numero")
                    .Entidade_Origem = PegaValorCampoXML(Caminho, I, "", "informante")
                End If

                .Valor_Documento = ""
                .Nome_Consumidor = ""
            End With
        Next I
    End If
End Sub

Sub Monta_Reg_025(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
    Dim Caminho As String
    Dim I      As Integer
    '' -- BANCO, AGENCIA, CONTA, COM DOCUMENTO DIFERENTE DO CONSULTADO

    If TipoResposta = Host Then
        With REG_025.Add
            .Codigo_Banco = Trim(Mid(Texto, 3, 3))
            .Codigo_Agencia = Trim(Mid(Texto, 6, 4))
            .Conta_Corrente = Trim(Mid(Texto, 10, 15))
            .Tipo_Documento = Trim(Mid(Texto, 25, 1))
            .CPF_CNPJ = Trim(Mid(Texto, 26, 14))
        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 40, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)
    ElseIf TipoResposta = xml Then
        Caminho = "banco-agencia-conta-documento-diferente/detalhe-banco-agencia-conta-documento-diferente"
        For I = 0 To objDom.selectNodes("resultado/" & Caminho).length - 1
            With REG_025.Add
                .Codigo_Banco = PegaValorCampoXML(Caminho, I, "dados-bancarios/banco", "codigo")
                .Codigo_Agencia = PegaValorCampoXML(Caminho, I, "dados-bancarios", "numero-agencia")
                .Conta_Corrente = PegaValorCampoXML(Caminho, I, "dados-bancarios", "numero-conta-corrente")
                .Tipo_Documento = ""
                .CPF_CNPJ = PegaValorCampoXML(Caminho, I, "", "documento")
            End With
        Next I
    End If
End Sub

Sub Monta_Reg_027(ByVal Texto As String)
'' -- INFORMACOA ONLINE INSTITUICAO

    With REG_027.Add
        .Menssagem = Trim(Mid(Texto, 3, 70))
    End With

    ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
    StrResposta = Mid(StrResposta, 73, Len(StrResposta))

    ''' PEGA O CODIGO DO PROXIMO SERVICO
    Cod_Servico = Left(StrResposta, 2)

End Sub

Sub Monta_Reg_028(ByVal Texto As String)
'' -- CHEQUE SUSTADO

    With REG_028.Add
        .Codigo_Banco = Trim(Mid(Texto, 3, 3))
        .Codigo_Agencia = Trim(Mid(Texto, 6, 4))
        .Conta_Corrente = Trim(Mid(Texto, 10, 15))
        .Numero_Cheque = Trim(Mid(Texto, 25, 6))
        .Codigo_Alinea = Trim(Mid(Texto, 31, 2))
        .Origem = Trim(Mid(Texto, 33, 60))
        .Data_Abertura_ContaCorrente = Formata_Data(Trim(Mid(Texto, 93, 8)))
    End With

    ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
    StrResposta = Mid(StrResposta, 101, Len(StrResposta))

    ''' PEGA O CODIGO DO PROXIMO SERVICO
    Cod_Servico = Left(StrResposta, 2)

End Sub

Sub Monta_Reg_030(ByVal Texto As String, ByVal TipoRestposta As eTipoRespostaSpc)
'' -- PROTESTOS

    Dim Caminho As String
    Dim I      As Integer

    If TipoRestposta = Host Then
        With REG_030.Add
            .Data_Inclusao = Formata_Data(Trim(Mid(Texto, 3, 8)))
            .Data_Protesto = Formata_Data(Trim(Mid(Texto, 11, 8)))
            .Cartorio = Trim(Mid(Texto, 19, 20))
            .Cidade = Trim(Mid(Texto, 39, 30))
            .UF = Trim(Mid(Texto, 69, 2))
            .Requerente = Trim(Mid(Texto, 71, 60))
            .Valor = Formata_Valor(Trim(Mid(Texto, 131, 15)))
            .Nome_Consumidor = Trim(Mid(Texto, 146, 60))
            .Entidade_Origem = Trim(Mid(Texto, 206, 30))
        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 236, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoRestposta = xml Then
        Caminho = "protesto/detalhe-protesto"
        For I = 0 To objDom.selectNodes("resultado/" & Caminho).length - 1
            With REG_030.Add
                .Data_Inclusao = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-inclusao"))
                .Data_Protesto = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-protesto"))
                .Requerente = PegaValorCampoXML(Caminho, I, "", "requerente-credor")
                .Valor = PegaValorCampoXML(Caminho, I, "", "valor")
                .Cartorio = PegaValorCampoXML(Caminho, I, "cartorio", "nome")
                .Cidade = PegaValorCampoXML(Caminho, I, "cartorio/cidade", "nome")
                .UF = PegaValorCampoXML(Caminho, I, "cartorio/cidade/estado", "sigla-uf")
                .Nome_Consumidor = ""
                .Entidade_Origem = ""
            End With
        Next I
    End If
End Sub

Sub Monta_Reg_040(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
''--// Resposta de Ações - máximo de 05 ocorrências

    Dim Caminho As String
    Dim I      As Integer


    If TipoResposta = Host Then
        With REG_040.Add
            .Data_Inclusao = Formata_Data(Trim(Mid(Texto, 3, 8)))
            .Data_Acao = Formata_Data(Trim(Mid(Texto, 11, 8)))
            .Tipo_Acao = Trim(Mid(Texto, 19, 20))
            .Vara = Trim(Mid(Texto, 39, 20))
            .Cidade = Trim(Mid(Texto, 59, 30))
            .UF = Trim(Mid(Texto, 89, 2))
            .Autor = Trim(Mid(Texto, 91, 60))
            .Valor = Formata_Valor(Trim(Mid(Texto, 151, 15)))
            .Nome_Consumidor = Trim(Mid(Texto, 166, 60))
            .Entidade_Origem = Trim(Mid(Texto, 266, 30))
        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 256, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoResposta = Host Then
        Caminho = "acao/detalhe-acao"
        For I = 0 To objDom.selectNodes("resultado/" & Caminho).length - 1
            With REG_040.Add
                .Data_Inclusao = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-inclusao"))
                .Data_Acao = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-acao"))
                .Tipo_Acao = PegaValorCampoXML(Caminho, I, "tipo-acao", "descricao")
                .Autor = PegaValorCampoXML(Caminho, I, "", "nome-autor")
                .Valor = PegaValorCampoXML(Caminho, I, "", "valor-acao")
                .Vara = PegaValorCampoXML(Caminho, I, "vara", "nome")
                .Cidade = PegaValorCampoXML(Caminho, I, "vara/comarca", "nome")
                .UF = PegaValorCampoXML(Caminho, I, "vara/comarca/estado", "sigla-uf")

                .Nome_Consumidor = ""
                .Entidade_Origem = ""
            End With
        Next I
    End If
End Sub

Sub Monta_Reg_050(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
''--// Resposta de Credito Concedito - máximo de 09 ocorrências

    Dim Caminho As String
    Dim I      As Integer


    If TipoResposta = Host Then
        With REG_050.Add
            .Mes_Ano = Trim(Mid(Texto, 3, 2)) & "/" & Trim(Mid(Texto, 5, 4))
            .Qdte = Trim(Mid(Texto, 9, 3))
            .Valor = Formata_Valor(Trim(Mid(Texto, 12, 15)))
            .Nome_Consumidor = Trim(Mid(Texto, 27, 60))
        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 87, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoResposta = xml Then
        Caminho = "credito-concedido/detalhe-credito-concedido "
        For I = 0 To objDom.selectNodes("resultado/" & Caminho).length - 1
            With REG_050.Add
                .Mes_Ano = PegaValorCampoXML(Caminho, I, "", "mes-vencimento") & "/" & PegaValorCampoXML(Caminho, I, "", "ano-vencimento")
                .Qdte = PegaValorCampoXML(Caminho, I, "", "quantidade-credito")
                .Valor = Replace(PegaValorCampoXML(Caminho, I, "", "valor"), ".", ",")
                .Valor = Format(.Valor, "0.00")
                .Nome_Consumidor = ""
            End With
        Next I
    End If
End Sub

Sub Monta_Reg_060(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
''--// Resposta de Alerta - máximo de 10 ocorrências
''--// alerta-documento  quando for resposta xml

    Dim Caminho As String
    Dim I      As Integer

    If TipoResposta = Host Then
        With REG_060.Add
            .Data_Inclusao = Formata_Data(Trim(Mid(Texto, 3, 8)))
            .Motivo = Trim(Mid(Texto, 11, 20))
            .Tipo_Documento = Trim(Mid(Texto, 31, 20))
            .Observacao = Trim(Mid(Texto, 51, 100))
            .Cidade = Trim(Mid(Texto, 151, 30))
            .UF = Trim(Mid(Texto, 181, 2))
            .Nome_Consumidor = Trim(Mid(Texto, 183, 60))
            .Entidade_Origem = Trim(Mid(Texto, 243, 30))
        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 273, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoResposta = xml Then
        Caminho = "alerta-documento/detalhe-alerta-documento"
        For I = 0 To objDom.selectNodes("resultado/" & Caminho).length - 1
            With REG_060.Add
                .Data_Inclusao = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-inclusao"))
                .Motivo = PegaValorCampoXML(Caminho, I, "", "motivo")
                .Tipo_Documento = PegaValorCampoXML(Caminho, I, "tipo-documento-alerta", "nome")
                .Observacao = PegaValorCampoXML(Caminho, I, "", "observacao")
                .Cidade = ""                                         ' nao tem no xml
                .UF = ""                                             ' nao tem no xml
                .Nome_Consumidor = ""                                ' nao tem no xml
                .Entidade_Origem = PegaValorCampoXML(Caminho, I, "", "entidade-origem")
            End With
        Next I
    End If

End Sub

Sub Monta_Reg_070(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
''--/// RESPOSTA DE CONSULTAS REALIZADAS

    Dim Caminho As String
    Dim I      As Integer

    If TipoResposta = Host Then
        With REG_070.Add
            .Data_Consulta = Formata_Data(Trim(Mid(Texto, 3, 8)))
            .Hora_Consulta = Formata_Hora(Trim(Mid(Texto, 11, 4)))
            .Nome_Associado = Trim(Mid(Texto, 15, 60))
            .Cidade = Trim(Mid(Texto, 75, 30))
            .UF = Trim(Mid(Texto, 105, 2))
            .Entidade_Origem = Trim(Mid(Texto, 107, 30))
        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 137, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoResposta = xml Then
        Caminho = "consulta-realizada/detalhe-consulta-realizada"
        For I = 0 To objDom.selectNodes("resultado/" & Caminho).length - 1
            With REG_070.Add
                .Data_Consulta = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-consulta"))
                .Hora_Consulta = Mid(PegaValorCampoXML(Caminho, I, "", "data-consulta"), 12, 8)
                .Nome_Associado = PegaValorCampoXML(Caminho, I, "", "nome-associado")
                .Cidade = PegaValorCampoXML(Caminho, I, "origem-associado", "nome")
                .UF = PegaValorCampoXML(Caminho, I, "origem-associado/estado", "sigla-uf")
                .Entidade_Origem = PegaValorCampoXML(Caminho, I, "", "nome-entidade-origem")
            End With
        Next I
    End If

End Sub

Sub Monta_Reg_080(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
    Dim Caminho As String

    ''--/// Resposta de Dados Telefônicos (produtos que consultam somente 1 telefone) - máximo de 1 ocorrência
    ''--// <telefone-vinculado-consumidor>  Quando for  XML

    If TipoResposta = Host Then
        With REG_080
            .Assinante = Trim(Mid(Texto, 3, 60))
            .Tipo_Logradouro = Trim(Mid(Texto, 63, 3))
            .Endereco = Trim(Mid(Texto, 66, 50))
            .Numero = Trim(Mid(Texto, 116, 5))
            .Cidade = Trim(Mid(Texto, 121, 30))
            .Bairro = Trim(Mid(Texto, 151, 20))
            .CEP = Formata_CEP(Trim(Mid(Texto, 171, 9)))
            .UF = Trim(Mid(Texto, 180, 2))
        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 182, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoResposta = xml Then
        Caminho = "telefone-vinculado-consumidor/detalhe-telefone-vinculado-consumidor/assinante"
        With REG_080
            .Assinante = ""                                          'PegaValorCampoXML(Caminho, "telefone", "numero-ddd")
            .Tipo_Logradouro = ""                                    'PegaValorCampoXML(Caminho, "endereco", "logradouro")
            .Endereco = PegaValorCampoXML(Caminho, 0, "endereco", "logradouro")
            .Numero = PegaValorCampoXML(Caminho, 0, "endereco", "numero")
            .Bairro = PegaValorCampoXML(Caminho, 0, "endereco", "bairro")
            .CEP = PegaValorCampoXML(Caminho, 0, "endereco", "cep")
            Caminho = "telefone-vinculado-consumidor/detalhe-telefone-vinculado-consumidor/assinante/endereco"
            .Cidade = PegaValorCampoXML(Caminho, 0, "cidade", "nome")
            Caminho = "telefone-vinculado-consumidor/detalhe-telefone-vinculado-consumidor/assinante/endereco/cidade"
            .UF = PegaValorCampoXML(Caminho, 0, "estado", "sigla-uf")
        End With
    End If

End Sub

Sub Monta_Reg_081(ByVal Texto As String)

''--/// Resposta de Consulta a Vários Telefones - máximo de 10 ocorrências

    With REG_081.Add
        .Tipo_Logradouro = Trim(Mid(Texto, 3, 1))
        .DDD = Trim(Mid(Texto, 4, 3))
        .Telefone = Trim(Mid(Texto, 7, 10))
        .Assinante = Trim(Mid(Texto, 17, 60))
        .Tipo_Logradouro = Trim(Mid(Texto, 73, 3))
        .Endereco = Trim(Mid(Texto, 80, 50))
        .Numero = Trim(Mid(Texto, 130, 5))
        .Cidade = Trim(Mid(Texto, 135, 30))
        .Bairro = Trim(Mid(Texto, 165, 20))
        .CEP = Formata_CEP(Trim(Mid(Texto, 185, 9)))
        .UF = Trim(Mid(Texto, 194, 2))
    End With

    ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
    StrResposta = Mid(StrResposta, 196, Len(StrResposta))

    ''' PEGA O CODIGO DO PROXIMO SERVICO
    Cod_Servico = Left(StrResposta, 2)


End Sub

Sub Monta_Reg_082(ByVal Texto As String)

''--/// Resposta de Telefones Vinculados ao Documento do Telefone Consultado - máximo de  10 ocorrências

    With REG_082.Add
        .Tipo_Logradouro = Trim(Mid(Texto, 3, 1))
        .DDD = Trim(Mid(Texto, 4, 3))
        .Telefone = Trim(Mid(Texto, 7, 10))
        .Assinante = Trim(Mid(Texto, 17, 60))
        .Tipo_Logradouro = Trim(Mid(Texto, 73, 3))
        .Endereco = Trim(Mid(Texto, 80, 50))
        .Numero = Trim(Mid(Texto, 130, 5))
        .Cidade = Trim(Mid(Texto, 135, 30))
        .Bairro = Trim(Mid(Texto, 165, 20))
        .CEP = Formata_CEP(Trim(Mid(Texto, 185, 9)))
        .UF = Trim(Mid(Texto, 194, 2))
    End With

    ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
    StrResposta = Mid(StrResposta, 196, Len(StrResposta))

    ''' PEGA O CODIGO DO PROXIMO SERVICO
    Cod_Servico = Left(StrResposta, 2)


End Sub

Sub Monta_Reg_084(ByVal Texto As String)

''--/// Resposta de Telefones Vinculados ao Documento Consultado - máximo de 10 ocorrências

    With REG_084.Add
        .Tipo_Telefone = Trim(Mid(Texto, 3, 1))
        .DDD = Trim(Mid(Texto, 4, 3))
        .Telefone = Trim(Mid(Texto, 7, 10))
        .Assinante = Trim(Mid(Texto, 17, 60))
        .Tipo_Logradouro = Trim(Mid(Texto, 77, 3))
        .Endereco = Trim(Mid(Texto, 80, 50))
        .Numero = Trim(Mid(Texto, 130, 5))
        .Cidade = Trim(Mid(Texto, 135, 30))
        .Bairro = Trim(Mid(Texto, 165, 20))
        .CEP = Formata_CEP(Trim(Mid(Texto, 185, 9)))
        .UF = Trim(Mid(Texto, 194, 2))
    End With

    ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
    StrResposta = Mid(StrResposta, 196, Len(StrResposta))

    ''' PEGA O CODIGO DO PROXIMO SERVICO
    Cod_Servico = Left(StrResposta, 2)


End Sub

Sub Monta_Reg_085(ByVal Texto As String)
    With REG_085.Add
        .DDD = Trim(Mid(Texto, 3, 3))
        .Telefone = Trim(Mid(Texto, 6, 10))
        .Data_PrimeiraConsulta = Formata_Data(Trim(Mid(Texto, 16, 8)))
        .Data_UltimaConsulta = Formata_Data(Trim(Mid(Texto, 24, 8)))
    End With

    ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
    StrResposta = Mid(StrResposta, 32, Len(StrResposta))

    ''' PEGA O CODIGO DO PROXIMO SERVICO
    Cod_Servico = Left(StrResposta, 2)



End Sub

Sub Monta_Reg_086(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
    Dim Caminho As String
    Dim I      As Integer

    If TipoResposta = Host Then
        With REG_086.Add
            .Endereco = Trim(Mid(Texto, 3, 50))
            .Complemento = Trim(Mid(Texto, 53, 30))
            .Bairro = Trim(Mid(Texto, 83, 30))
            .Cidade = Trim(Mid(Texto, 113, 30))
            .CEP = Formata_CEP(Right(Trim(Mid(Texto, 143, 9)), 8))
            .UF = Trim(Mid(Texto, 152, 2))
        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 154, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoResposta = xml Then
        Caminho = "ultimo-endereco-informado/detalhe-ultimo-endereco-informado"
        For I = 0 To objDom.selectNodes("resultado/" & Caminho).length - 1
            With REG_086.Add
                .Endereco = PegaValorCampoXML(Caminho, I, "endereco", "logradouro")
                .Complemento = PegaValorCampoXML(Caminho, I, "endereco", "complemento")
                .Bairro = PegaValorCampoXML(Caminho, I, "endereco", "bairro")
                .CEP = PegaValorCampoXML(Caminho, I, "endereco", "cep")
                .Cidade = PegaValorCampoXML(Caminho, I, "endereco/cidade", "nome")
                .UF = PegaValorCampoXML(Caminho, I, "endereco/cidade/estado", "sigla-uf")
            End With
        Next I
    End If

End Sub

Sub Monta_Reg_087(ByVal Texto As String)

''--/// RESPOSTA DE CONSULTAS REALIZADAS

    With REG_087
        .Endereco = Trim(Mid(Texto, 3, 85))
        .Cidade = Trim(Mid(Texto, 88, 30))
        .Bairro = Trim(Mid(Texto, 118, 30))
        .CEP = Formata_CEP(Trim(Mid(Texto, 148, 9)))
        .UF = Trim(Mid(Texto, 157, 2))
    End With

    ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
    StrResposta = Mid(StrResposta, 159, Len(StrResposta))

    ''' PEGA O CODIGO DO PROXIMO SERVICO
    Cod_Servico = Left(StrResposta, 2)


End Sub

Sub Monta_Reg_092(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
''--/// RESPOSTA DE CHEQUES SUSTADOS
    Dim Caminho As String
    Dim I      As Integer

    If TipoResposta = Host Then
        With REG_092.Add
            .Data_Incluso = Formata_Data(Trim(Mid(Texto, 3, 8)))
            .Data_Ocorrencia = Formata_Data(Trim(Mid(Texto, 11, 8)))
            .Origem = Trim(Mid(Texto, 19, 60))
            .Codigo_Banco = Trim(Mid(Texto, 79, 3))
            .Nome_Banco = Trim(Mid(Texto, 82, 30))
            .Codigo_Agencia = Trim(Mid(Texto, 112, 4))
            .Conta_Corrente = Trim(Mid(Texto, 116, 15))
            .Cheque_Inicial = Trim(Mid(Texto, 131, 6))
            .Cheque_Final = Trim(Mid(Texto, 137, 6))
            .Valor = Formata_Valor(Trim(Mid(Texto, 143, 15)))
            .Codigo_Alinea = Trim(Mid(Texto, 158, 2))
            .Entidade_Origem = Trim(Mid(Texto, 160, 30))
        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 190, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoResposta = xml Then
        Caminho = "contumacia/detalhe-contumacia"
        For I = 0 To objDom.selectNodes("resultado/" & Caminho).length - 1
            With REG_092.Add
                .Data_Incluso = ""
                .Data_Ocorrencia = ""
                .Origem = ""
                .Codigo_Banco = ""
                .Nome_Banco = ""
                .Codigo_Agencia = ""
                .Conta_Corrente = ""
                .Cheque_Inicial = ""
                .Cheque_Final = ""
                .Valor = ""
                .Codigo_Alinea = ""
                .Entidade_Origem = ""
            End With
        Next I
    End If
End Sub

Sub Monta_Reg_095(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
''--// RESPOSTA DA SINTESE CADASTRAL

    Dim Caminho As String
    Dim I      As Integer

    If TipoResposta = Host Then

        With mvarREG_095
            .Nome = Trim(Mid(Texto, 3, 60))
            .Nome_Fantasia = Trim(Mid(Texto, 63, 60))
            .RG = Trim(Mid(Texto, 123, 16))
            .Titulo_Eleitor = Trim(Mid(Texto, 139, 13))
            .Data_Nasc_Fundacao = Formata_Data(Trim(Mid(Texto, 152, 8)))
            .Situacao_CPF_CNPJ = Trim(Mid(Texto, 160, 1))
            .Nome_Mae = Trim(Mid(Texto, 161, 60))
            .Endereco = Trim(Mid(Texto, 221, 60))
            .Complemento = Trim(Mid(Texto, 281, 30))
            .Bairro = Trim(Mid(Texto, 311, 30))
            .Cidade = Trim(Mid(Texto, 341, 30))
            .UF = Trim(Mid(Texto, 371, 2))
            .CEP = Formata_CEP(Trim(Mid(Texto, 373, 8)))
            .Data_Ultima_Alteracao_Endereco = Formata_Data(Trim(Mid(Texto, 381, 8)))
            .DDD = Trim(Mid(Texto, 389, 4))
            .Telefone = Trim(Mid(Texto, 393, 8))
            .Situacao_CPF_CNPJ_RF = Trim(Mid(Texto, 401, 1))
            .Data_Situacao_RF = Trim(Mid(Texto, 402, 8))
            .Hora_Situacao_RF = Trim(Mid(Texto, 410, 6))
            .Origem_CPF = Trim(Mid(Texto, 416, 65))
            .Sexo = Trim(Mid(Texto, 481, 1))
            .Signo = Trim(Mid(Texto, 482, 11))
            .Idade = Trim(Mid(Texto, 493, 3))
            .IE = Trim(Mid(Texto, 496, 18))
            .UF_IE = Trim(Mid(Texto, 514, 2))
            .Situacao_IE = Trim(Mid(Texto, 516, 1))
            .Data_Situacao_IE = Formata_Data(Trim(Mid(Texto, 517, 8)))
            .Hora_Situacao_IE = Formata_Hora(Trim(Mid(Texto, 525, 6)))
            .Valor_Capital_Social = Formata_Valor(Trim(Mid(Texto, 727, 15)))
            .Razao_Antecessora = Trim(Mid(Texto, 742, 80))
        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 822, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoResposta = xml Then

        If strTipoConsumidor = "F" Then Caminho = "consumidor/consumidor-pessoa-fisica"
        If strTipoConsumidor = "J" Then Caminho = "consumidor/consumidor-pessoa-juridica"

        With mvarREG_095
            If strTipoConsumidor = "J" Then
                .Nome = PegaValorCampoXML(Caminho, I, "", "razao-social")
                .Nome_Fantasia = PegaValorCampoXML(Caminho, I, "", "nome-comercial")
                .Data_Nasc_Fundacao = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-fundacao"))
                .Situacao_CPF_CNPJ = PegaValorCampoXML(Caminho, I, "situacao-cnpj", "descricao-situacao")
                .Situacao_CPF_CNPJ_RF = .Situacao_CPF_CNPJ
                .Data_Situacao_RF = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "situacao-cnpj ", "data-situacao"))
                .Hora_Situacao_RF = Mid(PegaValorCampoXML(Caminho, I, "situacao-cnpj", "data-situacao"), 12, 8)
                .Descricao_Natureza_JUridica = PegaValorCampoXML(Caminho, I, "natureza-juridica", "descricao")
                .Codigo_Natureza_Juridica = PegaValorCampoXML(Caminho, I, "natureza-juridica", "codigo")
                .Descricao_CNAE = PegaValorCampoXML(Caminho, I, "atividade-economica-principal", "descricao")
                .CNAE = PegaValorCampoXML(Caminho, I, "atividade-economica-principal", "codigo")

                .IE = ""
                .UF_IE = ""
                .Situacao_IE = ""
                .Data_Situacao_IE = ""
                .Hora_Situacao_IE = ""
                .Valor_Capital_Social = PegaValorCampoXML("capital-social", 0, "detalhe-capital-social", "valor-capital-social")
                .Razao_Antecessora = ""
            Else
                .Nome = PegaValorCampoXML(Caminho, I, "", "nome")
                .Data_Nasc_Fundacao = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-nascimento"))
                .Idade = PegaValorCampoXML(Caminho, I, "", "idade")
                .Nome_Mae = PegaValorCampoXML(Caminho, I, "", "nome-mae")
                .RG = PegaValorCampoXML(Caminho, I, "", "numero-rg")
                .Titulo_Eleitor = PegaValorCampoXML(Caminho, I, "", "numero-titulo-eleitor")
                .Sexo = PegaValorCampoXML(Caminho, I, "", "sexo")
                .Signo = PegaValorCampoXML(Caminho, I, "", "signo")
                .Origem_CPF = PegaValorCampoXML(Caminho, I, "cpf", "regiao-origem")
                .Situacao_CPF_CNPJ = PegaValorCampoXML(Caminho, I, "situacao-cpf", "descricao-situacao")
                .Situacao_CPF_CNPJ_RF = .Situacao_CPF_CNPJ
                .Data_Situacao_RF = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "situacao-cpf", "data-situacao"))
                .Hora_Situacao_RF = Mid(PegaValorCampoXML(Caminho, I, "situacao-cpf", "data-situacao"), 12, 8)
            End If


            .Endereco = PegaValorCampoXML(Caminho, I, "endereco", "logradouro")
            .Endereco = .Endereco & "," & PegaValorCampoXML(Caminho, I, "endereco", "numero")
            .Complemento = PegaValorCampoXML(Caminho, I, "endereco", "complemento")
            .Bairro = PegaValorCampoXML(Caminho, I, "endereco", "bairro")
            .CEP = PegaValorCampoXML(Caminho, I, "endereco", "cep")
            .Cidade = PegaValorCampoXML(Caminho, I, "endereco/cidade", "nome")
            .UF = PegaValorCampoXML(Caminho, I, "endereco/cidade/estado", "sigla-uf")

            .Data_Ultima_Alteracao_Endereco = ""
            .DDD = ""
            .Telefone = ""


        End With
    End If

End Sub


Sub Monta_Reg_CapitalSocial(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
''--// RESPOSTA DA SINTESE CADASTRAL

    Dim Caminho As String
    Dim I      As Integer

    If TipoResposta = Host Then
    ElseIf TipoResposta = xml Then

        Caminho = "capital-social"

        On Error Resume Next
        With mvarREG_CAPITALSOCIAL
            .Nacionalidade = PegaValorCampoXML(Caminho, 0, "detalhe-capital-social", "nacionalidade")
            .Natureza = PegaValorCampoXML(Caminho, 0, "detalhe-capital-social", "natureza")
            .Origem = PegaValorCampoXML(Caminho, 0, "detalhe-capital-social", "origem")
            .DataAtualizacao = PegaValorCampoXML(Caminho, 0, "detalhe-capital-social", "data-ultima-atualizacao")
            .ValorCapitalAutorizado = PegaValorCampoXML(Caminho, 0, "detalhe-capital-social", "valor-capital-autorizado")
            .ValorCapitalRealizado = PegaValorCampoXML(Caminho, 0, "detalhe-capital-social", "valor-capital-realizado")
            .ValorCapitalSocial = PegaValorCampoXML(Caminho, 0, "detalhe-capital-social", "valor-capital-social")
        End With
        On Error GoTo 0
    End If

End Sub

Sub Monta_Reg_HistoricoPagamento(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
''--// RESPOSTA DA SINTESE CADASTRAL

    Dim Caminho As String
    Dim I      As Integer

    If TipoResposta = Host Then
    ElseIf TipoResposta = xml Then

        Caminho = "historico-pagamento"

        On Error Resume Next
        With mvarREG_HISTORICOPAGAMENTO
            ''FALTA IMPLEMENTAR
        End With
        On Error GoTo 0
    End If

End Sub



Public Property Let Msg_Erro(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Msg_Erro = 5
    mvarMsg_Erro = vData
End Property

Public Property Get Msg_Erro() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Msg_Erro
    Msg_Erro = mvarMsg_Erro
End Property

Public Property Get REG_005() As REG_005
    If mvarREG_005 Is Nothing Then
        Set mvarREG_005 = New REG_005
    End If


    Set REG_005 = mvarREG_005
End Property

Public Property Set REG_005(vData As REG_005)
    Set mvarREG_005 = vData
End Property

Public Property Get REG_006() As REG_006
    If mvarREG_006 Is Nothing Then
        Set mvarREG_006 = New REG_006
    End If


    Set REG_006 = mvarREG_006
End Property

Public Property Set REG_006(vData As REG_006)
    Set mvarREG_006 = vData
End Property

Public Property Get REG_007() As REG_007
    Set REG_007 = mvarREG_007
End Property

Public Property Set REG_007(vData As REG_007)
    Set mvarREG_007 = vData
End Property

Public Property Get REG_008() As REG_008
    Set REG_008 = mvarREG_008
End Property

Public Property Set REG_008(vData As REG_008)
    Set mvarREG_008 = vData
End Property

Public Property Get REG_009() As REG_009
    Set REG_009 = mvarREG_009
End Property

Public Property Set REG_009(vData As REG_009)
    Set mvarREG_009 = vData
End Property

Public Property Get REG_010() As REG_010
    If mvarREG_010 Is Nothing Then
        Set mvarREG_010 = New REG_010
    End If


    Set REG_010 = mvarREG_010
End Property

Public Property Set REG_010(vData As REG_010)
    Set mvarREG_010 = vData
End Property

Public Property Get REG_011() As REG_011
    Set REG_011 = mvarREG_011
End Property

Public Property Set REG_011(vData As REG_011)
    Set mvarREG_011 = vData
End Property

Public Property Get REG_020() As REG_020
    If mvarREG_020 Is Nothing Then
        Set mvarREG_020 = New REG_020
    End If


    Set REG_020 = mvarREG_020
End Property

Public Property Set REG_020(vData As REG_020)
    Set mvarREG_020 = vData
End Property

Public Property Get REG_025() As REG_025
    If mvarREG_025 Is Nothing Then
        Set mvarREG_025 = New REG_025
    End If


    Set REG_025 = mvarREG_025
End Property

Public Property Set REG_025(vData As REG_025)
    Set mvarREG_025 = vData
End Property

Public Property Get REG_026() As REG_026
    If mvarREG_026 Is Nothing Then
        Set mvarREG_026 = New REG_026
    End If


    Set REG_026 = mvarREG_026
End Property

Public Property Set REG_026(vData As REG_026)
    Set mvarREG_026 = vData
End Property

Public Property Get REG_027() As REG_027
    If mvarREG_027 Is Nothing Then
        Set mvarREG_027 = New REG_027
    End If


    Set REG_027 = mvarREG_027
End Property

Public Property Set REG_027(vData As REG_027)
    Set mvarREG_027 = vData
End Property

Public Property Get REG_028() As REG_028
    If mvarREG_028 Is Nothing Then
        Set mvarREG_028 = New REG_028
    End If


    Set REG_028 = mvarREG_028
End Property

Public Property Set REG_028(vData As REG_028)
    Set mvarREG_028 = vData
End Property

Public Property Get REG_030() As REG_030
    If mvarREG_030 Is Nothing Then
        Set mvarREG_030 = New REG_030
    End If


    Set REG_030 = mvarREG_030
End Property

Public Property Set REG_030(vData As REG_030)
    Set mvarREG_030 = vData
End Property

Public Property Get REG_040() As REG_040
    If mvarREG_040 Is Nothing Then
        Set mvarREG_040 = New REG_040
    End If


    Set REG_040 = mvarREG_040
End Property

Public Property Set REG_040(vData As REG_040)
    Set mvarREG_040 = vData
End Property

Public Property Get REG_050() As REG_050
    If mvarREG_050 Is Nothing Then
        Set mvarREG_050 = New REG_050
    End If


    Set REG_050 = mvarREG_050
End Property

Public Property Set REG_050(vData As REG_050)
    Set mvarREG_050 = vData
End Property

Public Property Get REG_060() As REG_060
    If mvarREG_060 Is Nothing Then
        Set mvarREG_060 = New REG_060
    End If


    Set REG_060 = mvarREG_060
End Property

Public Property Set REG_060(vData As REG_060)
    Set mvarREG_060 = vData
End Property

Public Property Get REG_070() As REG_070
    If mvarREG_070 Is Nothing Then
        Set mvarREG_070 = New REG_070
    End If


    Set REG_070 = mvarREG_070
End Property

Public Property Set REG_070(vData As REG_070)
    Set mvarREG_070 = vData
End Property

Public Property Get REG_080() As REG_080
    Set REG_080 = mvarREG_080
End Property

Public Property Set REG_080(vData As REG_080)
    Set mvarREG_080 = vData
End Property

Public Property Get REG_081() As REG_081
    If mvarREG_081 Is Nothing Then
        Set mvarREG_081 = New REG_081
    End If


    Set REG_081 = mvarREG_081
End Property

Public Property Set REG_081(vData As REG_081)
    Set mvarREG_081 = vData
End Property

Public Property Get REG_082() As REG_082
    If mvarREG_082 Is Nothing Then
        Set mvarREG_082 = New REG_082
    End If


    Set REG_082 = mvarREG_082
End Property

Public Property Set REG_082(vData As REG_082)
    Set mvarREG_082 = vData
End Property

Public Property Get REG_084() As REG_084
    If mvarREG_084 Is Nothing Then
        Set mvarREG_084 = New REG_084
    End If


    Set REG_084 = mvarREG_084
End Property

Public Property Set REG_084(vData As REG_084)
    Set mvarREG_084 = vData
End Property

Public Property Get REG_085() As REG_085
    If mvarREG_085 Is Nothing Then
        Set mvarREG_085 = New REG_085
    End If


    Set REG_085 = mvarREG_085
End Property

Public Property Set REG_085(vData As REG_085)
    Set mvarREG_085 = vData
End Property

Public Property Get REG_086() As REG_086
    If mvarREG_086 Is Nothing Then
        Set mvarREG_086 = New REG_086
    End If


    Set REG_086 = mvarREG_086
End Property

Public Property Set REG_086(vData As REG_086)
    Set mvarREG_086 = vData
End Property

Public Property Get REG_087() As REG_087
    Set REG_087 = mvarREG_087
End Property

Public Property Set REG_087(vData As REG_087)
    Set mvarREG_087 = vData
End Property

Public Property Get REG_088() As REG_088
    If mvarREG_088 Is Nothing Then
        Set mvarREG_088 = New REG_088
    End If


    Set REG_088 = mvarREG_088
End Property

Public Property Set REG_088(vData As REG_088)
    Set mvarREG_088 = vData
End Property

Public Property Get REG_089() As REG_089
    If mvarREG_089 Is Nothing Then
        Set mvarREG_089 = New REG_089
    End If


    Set REG_089 = mvarREG_089
End Property

Public Property Set REG_089(vData As REG_089)
    Set mvarREG_089 = vData
End Property

Public Property Get REG_090() As REG_090
    If mvarREG_090 Is Nothing Then
        Set mvarREG_090 = New REG_090
    End If


    Set REG_090 = mvarREG_090
End Property

Public Property Set REG_090(vData As REG_090)
    Set mvarREG_090 = vData
End Property

Public Property Get REG_092() As REG_092
    If mvarREG_092 Is Nothing Then
        Set mvarREG_092 = New REG_092
    End If


    Set REG_092 = mvarREG_092
End Property

Public Property Set REG_092(vData As REG_092)
    Set mvarREG_092 = vData
End Property

Public Property Get REG_095() As REG_095
    Set REG_095 = mvarREG_095
End Property

Public Property Set REG_095(vData As REG_095)
    Set mvarREG_095 = vData
End Property

Public Property Get REG_096() As REG_096
    If mvarREG_096 Is Nothing Then
        Set mvarREG_096 = New REG_096
    End If


    Set REG_096 = mvarREG_096
End Property

Public Property Set REG_096(vData As REG_096)
    Set mvarREG_096 = vData
End Property

Public Property Get REG_097() As REG_097
    If mvarREG_097 Is Nothing Then
        Set mvarREG_097 = New REG_097
    End If


    Set REG_097 = mvarREG_097
End Property

Public Property Set REG_097(vData As REG_097)
    Set mvarREG_097 = vData
End Property

Public Property Let RG(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RG = 5
    mvarRG = vData
    mvarRG = PadL(mvarRG, 16)
End Property

Public Property Get RG() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RG
    RG = mvarRG
End Property

Public Property Let Senha_Operador(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Senha_Operador = 5
    mvarSenha_Operador = vData
    mvarSenha_Operador = PadL(mvarSenha_Operador, 8)
End Property

Public Property Get Senha_Operador() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Senha_Operador
    Senha_Operador = mvarSenha_Operador
End Property



Public Property Let StrResposta(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StrResposta = 5
    mvarStrResposta = vData
End Property

Public Property Get StrResposta() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StrResposta
    StrResposta = mvarStrResposta
End Property

Public Property Let Tipo_Documento(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tipo_Documento = 5
    mvarTipo_Documento = vData
End Property

Public Property Get Tipo_Documento() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tipo_Documento
    Tipo_Documento = mvarTipo_Documento
End Property

Public Property Let Tipo_Resposta(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tipo_Resposta = 5
    mvarTipo_Resposta = vData
End Property

Public Property Get Tipo_Resposta() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tipo_Resposta
    Tipo_Resposta = mvarTipo_Resposta
End Property

Function Trata_Respostas()

''' REMOVE O CARETER 0 QUE INDICA FINALIZADOR
    StrResposta = Replace(StrResposta, Chr(0), "")

    ''' USANDO PARA TESTE
    If Dir(UsarArquivo, vbArchive) <> "" Then
        MsgBox "USANDO ARQUIVO*****"
        StrResposta = Pega_Resultado_Arquivo(UsarArquivo)
        Trata_Respostas
        Exit Function
    End If
    ''''FIM


    If Gravar_Resposta_Arquivo <> "" Then
        Open Gravar_Resposta_Arquivo For Output As #1
        Print #1, StrResposta
        Close #1
    End If

    If StrResposta = "TIMEOUT" Then
        Erro = True
        Msg_Erro = StrResposta
        Exit Function
    End If

    Monta_Header Mid(StrResposta, 1, 455), Host

    If Header_Resposta.Sucesso_Transacao = "S" Then
        Do While StrResposta <> ""
            If Cod_Servico = "05" Then Monta_Reg_005 Mid(StrResposta, 1, 137), Host
            If Cod_Servico = "06" Then Monta_Reg_006 Mid(StrResposta, 1, 107), Host
            If Cod_Servico = "07" Then Monta_Reg_007 Mid(StrResposta, 1, 142), Host
            If Cod_Servico = "08" Then Monta_Reg_008 Mid(StrResposta, 1, 693), Host
            If Cod_Servico = "09" Then Monta_Reg_009 Mid(StrResposta, 1, 102), Host
            If Cod_Servico = "10" Then Monta_Reg_010 Mid(StrResposta, 1, 246), Host
            If Cod_Servico = "20" Then Monta_Reg_020 Mid(StrResposta, 1, 285), 0, Host
            If Cod_Servico = "25" Then Monta_Reg_025 Mid(StrResposta, 1, 39), Host
            If Cod_Servico = "27" Then Monta_Reg_027 Mid(StrResposta, 1, 72)
            If Cod_Servico = "28" Then Monta_Reg_028 Mid(StrResposta, 1, 100)
            If Cod_Servico = "30" Then Monta_Reg_030 Mid(StrResposta, 1, 235), Host
            If Cod_Servico = "40" Then Monta_Reg_040 Mid(StrResposta, 1, 255), Host
            If Cod_Servico = "50" Then Monta_Reg_050 Mid(StrResposta, 1, 86), Host
            If Cod_Servico = "60" Then Monta_Reg_060 Mid(StrResposta, 1, 272), Host
            If Cod_Servico = "70" Then Monta_Reg_070 Mid(StrResposta, 1, 136), Host
            If Cod_Servico = "80" Then Monta_Reg_080 Mid(StrResposta, 1, 181), Host
            If Cod_Servico = "81" Then Monta_Reg_081 Mid(StrResposta, 1, 195)
            If Cod_Servico = "82" Then Monta_Reg_082 Mid(StrResposta, 1, 195)
            If Cod_Servico = "84" Then Monta_Reg_084 Mid(StrResposta, 1, 195)
            If Cod_Servico = "85" Then Monta_Reg_085 Mid(StrResposta, 1, 31)
            If Cod_Servico = "86" Then Monta_Reg_086 Mid(StrResposta, 1, 153), Host
            If Cod_Servico = "87" Then Monta_Reg_087 Mid(StrResposta, 1, 158)
            If Cod_Servico = "92" Then Monta_Reg_092 Mid(StrResposta, 1, 189), Host
            If Cod_Servico = "95" Then Monta_Reg_095 Mid(StrResposta, 1, 821), Host
            If Cod_Servico = "96" Then Monta_Reg_096 Mid(StrResposta, 1, 236), Host
            If Cod_Servico = "97" Then Monta_Reg_097 Mid(StrResposta, 1, 157), Host


            If Cod_Servico = "11" Or Cod_Servico = "26" Or Cod_Servico = "88" Or Cod_Servico = "89" Or Cod_Servico = "90" Then
                MsgBox "Atenção Esse CPF Retornou dados incompletos favor, passe esse CPF/CNPJ para o Desenvolvedor", vbCritical, "Erro"
                Erro = True
                Exit Do
            End If

            DoEvents
        Loop
    Else
        Erro = True
        Msg_Erro = Mid(StrResposta, 5, 100)
    End If

End Function

Public Property Let Versao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Versao = 5
    mvarVersao = vData
    If mvarVersao = "" Then mvarVersao = "014"
    mvarVersao = StrZero(mvarVersao, 3)
End Property

Public Property Get Versao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Versao
    Versao = mvarVersao
End Property

Sub Monta_Reg_096(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
'' -- Sócios Proprietarios

    Dim Caminho As String
    Dim I      As Integer

    If TipoResposta = Host Then
        With REG_096.Add
            .Tipo_Documento = Trim(Mid(Texto, 3, 1))
            .CPF_CNPJ = Trim(Mid(Texto, 4, 14))
            .Nome_Socio_Proprietario = Trim(Mid(Texto, 18, 60))
            .Percentual_Participacao = Formata_Valor(Trim(Mid(Texto, 78, 5)))
            .Data_Entrada = Formata_Data(Trim(Mid(Texto, 83, 8)))
            .Data_Saida = Formata_Data(Trim(Mid(Texto, 91, 8)))
            .Condicao = Trim(Mid(Texto, 99, 40))
            .Tipo_Documento = Trim(Mid(Texto, 139, 1))
            .Cargo_Direcao = Trim(Mid(Texto, 140, 40))
            .Cargo_Administracao = Trim(Mid(Texto, 180, 40))
            .Valor_Participacao = Formata_Valor(Trim(Mid(Texto, 220, 17)))
        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 237, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoResposta = xml Then
        Caminho = "socio/detalhe-socio "
        For I = 0 To objDom.selectNodes("resultado/" & Caminho).length - 1
            With REG_096.Add
                .CPF_CNPJ = PegaValorCampoXML(Caminho, I, "", "documento")
                .Tipo_Documento = IIf(Len(.CPF_CNPJ) = 11, "CPF", "CNPJ")
                .Nome_Socio_Proprietario = PegaValorCampoXML(Caminho, I, "", "nome")
                .Percentual_Participacao = PegaValorCampoXML(Caminho, I, "", "porcentual-participacao")
                .Data_Entrada = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-entrada"))
                .Data_Saida = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-saida"))
                .Condicao = PegaValorCampoXML(Caminho, I, "", "condicao")
                .Cargo_Direcao = PegaValorCampoXML(Caminho, I, "", "cargo-direcao")
                .Cargo_Administracao = PegaValorCampoXML(Caminho, I, "", "cargo-administracao")
                .Valor_Participacao = Format(Replace(PegaValorCampoXML(Caminho, I, "", "valor-participacao"), ".", ","), "0.00")
            End With
        Next I
    End If
End Sub

Sub Monta_Reg_097(ByVal Texto As String, ByVal TipoResposta As eTipoRespostaSpc)
'' -- Participacao societaria da Pessoa Fisica

    Dim Caminho As String
    Dim I      As Integer

    If TipoResposta = Host Then

        With REG_097.Add
            .CNPJ = Trim(Mid(Texto, 3, 14))
            .Razao_Social = Trim(Mid(Texto, 17, 60))
            .Nome_Comercial = Trim(Mid(Texto, 77, 60))
            .Percentual_Participacao = Formata_Valor(Trim(Mid(Texto, 137, 5)))
            .Data_Entrada = Formata_Data(Trim(Mid(Texto, 142, 8)))
            .Data_Saida = Formata_Data(Trim(Mid(Texto, 150, 8)))
        End With

        ''' -// RETIRA O TEXTO JA TRABALHO DA STRING DE RESPOSTA
        StrResposta = Mid(StrResposta, 158, Len(StrResposta))

        ''' PEGA O CODIGO DO PROXIMO SERVICO
        Cod_Servico = Left(StrResposta, 2)

    ElseIf TipoResposta = xml Then
        Caminho = "participacao-empresa/detalhe-participacao-empresa"
        For I = 0 To objDom.selectNodes("resultado/" & Caminho).length - 1
            With REG_097.Add
                .CNPJ = PegaValorCampoXML(Caminho, I, "", "documento")
                .Razao_Social = PegaValorCampoXML(Caminho, I, "", "nome")
                .Nome_Comercial = ""
                .Percentual_Participacao = PegaValorCampoXML(Caminho, I, "", "porcentual-participacao")
                .Data_Entrada = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-entrada"))
                .Data_Saida = Formata_Data_XML(PegaValorCampoXML(Caminho, I, "", "data-saida"))
            End With
        Next I
    End If

End Sub


Function Pega_Resultado_Arquivo(ByVal Arquivo As String) As String
''' USADO SOMENTE PARA TESTE
    Dim Texto  As String

    Close #1
    Open Arquivo For Input As #1
    Do While Not EOF(1)
        Line Input #1, Texto
        Pega_Resultado_Arquivo = Pega_Resultado_Arquivo & Texto
    Loop
    Close #1

    Pega_Resultado_Arquivo = Replace(Pega_Resultado_Arquivo, vbCrLf, "")

End Function


Public Function ConsultarViaWebService(ByVal Usuario As String, _
                                       ByVal Senha As String, _
                                       ByVal CodigoProduto As String, _
                                       ByVal TipoConsumidor As eTipoConsumidorSpcBrasil, _
                                       ByVal NumeroDocumento As String, _
                                       ByVal Insumos As Insumos, _
                                       ByVal Ambiente As eAmbienteSPCBrasil) As Boolean



    Dim sRequisicao As String
    Dim strUrl As String
    Dim strSoapAction As String
    Dim I      As Integer
    Dim strOutput As String

    'Usuario = "395793"
    'Senha = "12012016"
    'ACodigoProduto = "116"
    'ATipoConsumidor = "F"
    'ACPF = "00000000191"
    'ACPF = "00829161600"



    Erro = False
    Msg_Erro = ""

    If TipoConsumidor = Fisica Then
        strTipoConsumidor = "F"
    ElseIf TipoConsumidor = Juridica Then
        strTipoConsumidor = "J"
    End If


    If UsarArquivo <> "" Then
        If Dir(UsarArquivo, vbArchive) <> "" Then
            MsgBox "USANDO O ARQUIVO DE RESPOSTA"
            strOutput = Pega_Resultado_Arquivo(UsarArquivo)
            GoTo Resposta
        End If
    End If

    'PRODUCAO
    If Ambiente = eAmbienteSPCBrasil.Producao Then
        strUrl = "https://servicos.spc.org.br/spc/remoting/ws/consulta/consultaWebService"
    ElseIf Ambiente = eAmbienteSPCBrasil.Homologacao Then
        'HOMOLOGACAO
        strUrl = "https://treina.spc.org.br/spc/remoting/ws/consulta/consultaWebService"
    End If

    '' consultar
    strSoapAction = "http://webservice.consulta.spcjava.spcbrasil.org/consultaWebService/consultarRequest/consultar"
    sRequisicao = ""
    sRequisicao = sRequisicao & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" "
    sRequisicao = sRequisicao & "xmlns:web=""http://webservice.consulta.spcjava.spcbrasil.org/"">"
    sRequisicao = sRequisicao & "<soapenv:Header/><soapenv:Body><web:filtro>"
    sRequisicao = sRequisicao & "<codigo-produto>" & CodigoProduto & "</codigo-produto>"    ' codigo da consulta spc mix
    sRequisicao = sRequisicao & "<tipo-consumidor>" + strTipoConsumidor + "</tipo-consumidor>"
    sRequisicao = sRequisicao & "<documento-consumidor>" + Trim(NumeroDocumento) + "</documento-consumidor>"


    For I = 1 To Insumos.Count
        sRequisicao = sRequisicao & "<codigo-insumo-opcional>" & Insumos(I).Codigo & "</codigo-insumo-opcional>"    'HISTORICO DE PAGAMENTO
        'sRequisicao = sRequisicao & "<codigo-insumo-opcional>5193</codigo-insumo-opcional>"    'HISTORICO DE PAGAMENTO
        'sRequisicao = sRequisicao & "<codigo-insumo-opcional>77</codigo-insumo-opcional>" 'SCORE 03 MESES
        'sRequisicao = sRequisicao & "<codigo-insumo-opcional>78</codigo-insumo-opcional>"    'SCORE 12 MESES
    Next I
    sRequisicao = sRequisicao & "</web:filtro></soapenv:Body></soapenv:Envelope>"



    ' listar produtos
    'strSoapAction = "http://webservice.consulta.spcjava.spcbrasil.org/consultaWebService/consultarRequest/listarProdutos"
    'ARequisicao = ""
    'ARequisicao = ARequisicao & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" "
    'ARequisicao = ARequisicao & "xmlns:web=""http://webservice.consulta.spcjava.spcbrasil.org/"">"
    'ARequisicao = ARequisicao & "<soapenv:Header/><soapenv:Body>"
    'ARequisicao = ARequisicao & "</soapenv:Body></soapenv:Envelope>"



    strOutput = PostWebservice(Usuario, Senha, strUrl, strSoapAction, sRequisicao)
    strOutput = Replace(strOutput, "UTF-8", "ISO-8859-1")


Resposta:

    If Erro = False Then
        StrResposta = strOutput
        Trata_Respostas_WebService
    End If

    ConsultarViaWebService = True


End Function


Public Function ListarProdutosDisponiveis(ByVal Usuario As String, _
                                          ByVal Senha As String, _
                                          ByVal Ambiente As eAmbienteSPCBrasil) As Boolean



    Dim sRequisicao As String
    Dim strUrl As String
    Dim strSoapAction As String

    Dim strOutput As String


    Erro = False
    Msg_Erro = ""


    'PRODUCAO
    If Ambiente = eAmbienteSPCBrasil.Producao Then
        strUrl = "https://servicos.spc.org.br/spc/remoting/ws/consulta/consultaWebService"
    ElseIf Ambiente = eAmbienteSPCBrasil.Homologacao Then
        'HOMOLOGACAO
        strUrl = "https://treina.spc.org.br/spc/remoting/ws/consulta/consultaWebService"
    End If


    ' listar produtos
    strSoapAction = "http://webservice.consulta.spcjava.spcbrasil.org/consultaWebService/consultarRequest/listarProdutos"
    sRequisicao = ""
    sRequisicao = sRequisicao & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" "
    sRequisicao = sRequisicao & "xmlns:web=""http://webservice.consulta.spcjava.spcbrasil.org/"">"
    sRequisicao = sRequisicao & "<soapenv:Header/><soapenv:Body>"
    sRequisicao = sRequisicao & "</soapenv:Body></soapenv:Envelope>"


    strOutput = PostWebservice(Usuario, Senha, strUrl, strSoapAction, sRequisicao)
    strOutput = Replace(strOutput, "UTF-8", "ISO-8859-1")

    If Erro = False Then
        StrResposta = strOutput
        Trata_Respostas_ListaProdutos
    End If

    ListarProdutosDisponiveis = True


End Function

Private Function PostWebservice(ByVal User As String, ByVal Password As String, ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String) As String
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim strRet As String
    Dim Auth   As String
    'Dim intPos1          As Integer
    'Dim intPos2          As Integer

    On Error GoTo Err_PW

    ' Cria o Objeto DOMDocument and XMLHTTP
    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.XMLHTTP")

    ' Carrega o XML
    objDom.async = False
    objDom.LoadXML XmlBody

    ' Abre Webservice
    objXmlHttp.Open "POST", AsmxUrl, False

    ' Cria o Cabecalho
    'objXmlHttp.setTimeouts 120, 120, 120, 120
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", SoapActionUrl

    If (User & Password) <> "" Then
        Auth = Base64Encode(User & ":" & Password)
        objXmlHttp.setRequestHeader "Authorization", "Basic " & Auth
    End If

    ' Envia XML
    objXmlHttp.send objDom.xml


    ' Pega a Resposta do Webservice
    strRet = objXmlHttp.ResponseText

    ' Mata o Objeto
    Set objXmlHttp = Nothing

    ' Retorna o valor
    PostWebservice = strRet

    Exit Function
Err_PW:
    PostWebservice = "Error: " & Err.Number & " - " & Err.Description
    Erro = True
    Msg_Erro = PostWebservice

End Function

Private Function PegaValorCampoXML(ByVal Caminho As String, Item As Integer, ByVal SubCaminho As String, ByVal NomeAtributo As String, Optional ByVal ProcurarNode As Boolean) As String
    Dim Pai    As String
    Dim Node   As Object


    Pai = "/resultado/"
    Caminho = Pai & Caminho

    If Right(Caminho, 1) = "/" Then Caminho = Left(Caminho, Len(Caminho) - 1)

    On Error Resume Next
    Set Node = objDom.selectNodes(Caminho)
    If ProcurarNode = True Then                                      ' opcao pega o valor do Node e nao do atributo
        PegaValorCampoXML = Node.Item(0).selectSingleNode(NomeAtributo).Text
    ElseIf SubCaminho = "" Then
        PegaValorCampoXML = Node.Item(Item).Attributes.getNamedItem(NomeAtributo).Text
    Else
        PegaValorCampoXML = Node.Item(Item).selectNodes(SubCaminho).Item(0).Attributes.getNamedItem(NomeAtributo).Text
    End If
    Set Node = Nothing
    On Error GoTo 0

End Function

Private Sub Trata_Respostas_WebService()
    Dim Texto  As String

    Texto = StrResposta

    If InStr(Texto, "<S:Fault") > 0 Then
        ' HOUVE ERRO NA CONSULTA
        Texto = GetStringBetween(Texto, "<S:Fault ", "</S:Fault>")
        Erro = True
    Else
        Texto = GetStringBetween(Texto, "<ns2:resultado", "</ns2:resultado>")
        Erro = False
    End If

    If Texto = "" Then
        Texto = "<resultado> " & Texto & "</resultado>"
    Else
        Texto = "<resultado " & Texto & "</resultado>"
    End If


    Set objDom = CreateObject("MSXML2.DOMDocument")
    objDom.async = False

    If objDom.LoadXML(Texto) = False Then
        'MsgBox "Erro ao Carrega resposta do XML " & objDom.parseErro'r, vbCritical, "Erro"
        Erro = "Erro ao Carrega resposta do XML " & objDom.parseError
        Exit Sub
    Else

        If Erro = True Then
            Msg_Erro = PegaValorCampoXML("", 0, "", "faultcode", True) & " - " & PegaValorCampoXML("", 0, "", "faultstring", True)
            Msg_Erro = RemoveCaracterUTF8(Msg_Erro)
            Debug.Print Msg_Erro
        Else
            Monta_Reg_095 StrResposta, xml
            Monta_Reg_CapitalSocial StrResposta, xml                 '--//INFORMACOES SOBRE O CAPITAL SOCIAL DA EMPRESA
            Monta_Header StrResposta, xml                            '--// INFORMACOES DE CABECALHO / CONSUMIDOR
            Monta_Reg_005 StrResposta, xml                           '--// BASE EXTERNAS ( NO XML NAO TEM ESSAS OPCOES)
            Monta_Reg_006 StrResposta, xml                           '--// GRAFIA PJ
            Monta_Reg_007 StrResposta, xml                           '--// OBITOS
            Monta_Reg_008 StrResposta, xml                           '--// SCORE  ** NAO ACHEI NENHUM CPF/CNPJ QUE RETORNASSE
            Monta_Reg_009 StrResposta, xml                           '--// SCORE  ** NAO ACHEI NENHUM CPF/CNPJ QUE RETORNASSE
            Monta_Reg_010 StrResposta, xml                           '--// REGISTROS SPC
            Monta_Reg_020 StrResposta, Cheque_Lojista, xml           '--// CHEQUES LOJISTA
            Monta_Reg_020 StrResposta, CCF, xml                      '--// CHEQUES CCF
            Monta_Reg_020 StrResposta, Contra_Ordem, xml             '--// CHEQUES CONTRA ORDEM
            Monta_Reg_020 StrResposta, Contra_Ordem_Documento_Diferente, xml    '--// CHEQUES CONTRA ORDEM DOCUMENTO DIFERENTE (contra-ordem-documento-diferente)
            Monta_Reg_020 StrResposta, Contra_Ordem_Agencia_Diferente, xml    '--// CHEQUES CONTRA ORDEM AGENCIA DIFERENTE (contra-ordem-agencia-diferente)
            Monta_Reg_020 StrResposta, Contra_ordem_Agencia_Conta_Diferente, xml    '--// CHEQUES CONTRA ORDEM AGENCIA/CONTA DIFERENTE (contra-ordem-agencia-conta-diferente)
            Monta_Reg_025 StrResposta, xml                           '--// BANCO-AGENCIA-CONTA-DOCUMENTO-DIFERENTE
            Monta_Reg_030 StrResposta, xml                           '--// PROTESTO
            Monta_Reg_040 StrResposta, xml                           '--// ACOES
            Monta_Reg_050 StrResposta, xml                           '--// CREDITO CONCEDIDO
            Monta_Reg_060 StrResposta, xml                           '--// ALERTA DE DOCUMENTOS
            Monta_Reg_070 StrResposta, xml                           '--// CONSULTAS REALIZADAS
            Monta_Reg_080 StrResposta, xml                           '--// INFORMACOES DE TELEFONE NO NOME DO CLIENTE
            Monta_Reg_086 StrResposta, xml                           '--// ULTIMOS ENDERECOS INFORMADOS
            Monta_Reg_092 StrResposta, xml                           '--// CONTUMACIA  ** NAO ACHEI NENHUM CPF QUE RETORNASSE
            Monta_Reg_096 StrResposta, xml                           '--// SOCIOS
            Monta_Reg_097 StrResposta, xml                           '--// PARTICIPACAO SOCIETÁRIA
        End If

        '        Exit Sub

    End If

End Sub

Private Sub Trata_Respostas_ListaProdutos()
    Dim I      As Integer
    Dim X      As Integer
    Dim Texto  As String
    Dim vVar   As Variant
    Dim strInsumo As String
    Dim vInsumo As Variant
    Dim Codigo As String
    Dim Nome   As String
    Dim Tipo   As String
    Dim strTemp As String

    Texto = StrResposta

    If InStr(Texto, "<S:Fault") > 0 Then
        ' HOUVE ERRO NA CONSULTA
        Texto = GetStringBetween(Texto, "<S:Fault ", "</S:Fault>")
        Erro = True
    Else
        Texto = GetStringBetween(Texto, "<ns2:produtos", "</ns2:produtos>")
        Erro = False
    End If

    Texto = "<produtos " & Texto & "</produtos>"


    Set objDom = CreateObject("MSXML2.DOMDocument")
    objDom.async = False

    If objDom.LoadXML(Texto) = False Then
        'MsgBox "Erro ao Carrega resposta do XML " & objDom.parseErro'r, vbCritical, "Erro"
        Erro = "Erro ao Carrega resposta do XML " & objDom.parseError
        Exit Sub
    Else

        If Erro = True Then
            Msg_Erro = GetStringBetween(Texto, "<faultcode>", "</faultcode>")
            Msg_Erro = Msg_Erro & " - " & GetStringBetween(Texto, "<faultstring>", "</faultstring>")
            Msg_Erro = Replace(Msg_Erro, "<![CDATA[", "")
            Msg_Erro = Replace(Msg_Erro, "]]>", "")
            Msg_Erro = RemoveCaracterUTF8(Msg_Erro)
            Debug.Print Msg_Erro
        Else
            vVar = Split(Texto, "<produto>")
            For I = 1 To UBound(vVar)

                Codigo = GetStringBetween(vVar(I), "<codigo>", "</codigo>")
                Nome = GetStringBetween(vVar(I), "<nome>", "</nome>")
                Tipo = "A"                                           'F-FISICA, J-JURIDICA - A-AMBAS

                strTemp = GetStringBetween(vVar(I), "<insumo-opcional", "</produto>", 0, True)
                If strTemp <> "" Then
                    strInsumo = "<insumo-opcional" & strTemp
                    vInsumo = Split(strInsumo, "<insumo-opcional")

                    With Produtos.Add
                        .Codigo = Codigo
                        .Nome = Nome
                        .Tipo = Tipo

                        For X = 1 To UBound(vInsumo)
                            With .Insumos.Add
                                .Codigo = GetStringBetween(vInsumo(X), "codigo", "opcional")
                                .Nome = GetStringBetween(vInsumo(X), "nome", "codigo")

                            End With
                        Next X
                    End With
                End If
            Next I
        End If

    End If

End Sub

