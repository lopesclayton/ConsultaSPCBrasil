VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Header_Resposta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarCodigo_Transacao As String                     'local copy
Private mvarVersao As String                               'local copy
Private mvarMeio_de_Acesso As String                       'local copy
Private mvarCodigo_Estacao As String                       'local copy
Private mvarCodigo_Operador As String                      'local copy
Private mvarSenha_Operador As String                       'local copy
Private mvarTipo_Documento As String                       'local copy
Private mvarCPF_CNPJ As String                             'local copy
Private mvarRG As String                                   'local copy
Private mvarNome As String                                 'local copy
Private mvarData_Nascimento_Fundacao As String             'local copy
Private mvarEndereco As String                             'local copy
Private mvarComplemento As String                          'local copy
Private mvarBairro As String                               'local copy
Private mvarCidade As String                               'local copy
Private mvarUF As String                                   'local copy
Private mvarCEP As String                                  'local copy
Private mvarTelefone As String                             'local copy
Private mvarData_Ultima_Atu_Endereco As String             'local copy
Private mvarTotal_SPC As String                            'local copy
Private mvarTotal_Cheque_CCF As String                     'local copy
Private mvarTotal_Cheque_Lojista As String                 'local copy
Private mvarTotal_Cheque_ContraOrdenados As String         'local copy
Private mvarTotal_Alertas As String                        'local copy
Private mvarQtde_Dias_Consulta_Realizada As String         'local copy
Private mvarTotal_Consultas_Realizadas As String           'local copy
Private mvarTotal_Protesto As String                       'local copy
Private mvarTotal_Acoes As String                          'local copy
Private mvarTotal_Ocorrencias_Credito_Concedidos As String 'local copy
Private mvarTotal_Cheques_Contra_Ordenados_DIF_Consultado As String    'local copy
Private mvarTotal_Cheques_ContraOrdenados_AG_Diferente As String    'local copy
Private mvarTotal_Cheques_ContraOrdenados_AGConta_Diferente As String    'local copy
Private mvarQtde_Dias_Contumacia As String                 'local copy
Private mvarTotal_Cheques_Sustados_Linha21 As String       'local copy
Private mvarTotal_TalaoCheques_ContraOrdenados As String   'local copy
'local variable(s) to hold property value(s)
Private mvarTotal_TalaoCheques_ContraOrdenados_DOC_Diferente As String    'local copy
'local variable(s) to hold property value(s)
Private mvarTotal_TaloesCheques_ContraOrdenados_AG_Diferente As String    'local copy
'local variable(s) to hold property value(s)
Private mvarTotal_TaloesCheques_ContraOrdenados_AGCC_Diferente As String    'local copy
'local variable(s) to hold property value(s)
Private mvarTotal_Menssagem_Online As String               'local copy
Private mvarTotal_Cheques_Sustados_Online As String        'local copy
'local variable(s) to hold property value(s)
Private mvarInformacao_Obito As String                     'local copy
Private mvarTipo_Resposta As String                        'local copy
Private mvarNumero_Protocolo As String                     'local copy
Private mvarData_Resposta As String                        'local copy
Private mvarHoras_Resposta As String                       'local copy
Private mvarSucesso_Transacao As String                    'local copy
Private mvarCodigo_Entidade As String                      'local copy
Private mvarTotal_ContaCorrente_Doc_Diferente As String    'local copy
'local variable(s) to hold property value(s)
Private mvarIndicativo_Consulta_Parcial As String          'local copy
Private mvarTotal_RegistroConsulta As String
Private mvarTotal_InfoPoderJudiciario As String
Private mvarTotal_HistoricoPagamento As String
Private mvarTotal_ControleSocietario As String
Private mvarTotal_PendenciaSerasa As String
Private mvarTotal_ParticipacaoFalencia As String


Public Property Let Indicativo_Consulta_Parcial(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Indicativo_Consulta_Parcial = 5
    mvarIndicativo_Consulta_Parcial = vData
End Property


Public Property Get Indicativo_Consulta_Parcial() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Indicativo_Consulta_Parcial
    Indicativo_Consulta_Parcial = mvarIndicativo_Consulta_Parcial
End Property



Public Property Let Total_ContaCorrente_Doc_Diferente(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_ContaCorrente_Doc_Diferente = 5
    mvarTotal_ContaCorrente_Doc_Diferente = vData
End Property


Public Property Get Total_ContaCorrente_Doc_Diferente() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_ContaCorrente_Doc_Diferente
    Total_ContaCorrente_Doc_Diferente = mvarTotal_ContaCorrente_Doc_Diferente
End Property



Public Property Let Codigo_Entidade(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Codigo_Entidade = 5
    mvarCodigo_Entidade = vData
End Property


Public Property Get Codigo_Entidade() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Codigo_Entidade
    Codigo_Entidade = mvarCodigo_Entidade
End Property



Public Property Let Sucesso_Transacao(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Sucesso_Transacao = 5
    mvarSucesso_Transacao = vData
End Property


Public Property Get Sucesso_Transacao() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Sucesso_Transacao
    Sucesso_Transacao = mvarSucesso_Transacao
End Property



Public Property Let Horas_Resposta(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Horas_Resposta = 5
    mvarHoras_Resposta = vData
End Property


Public Property Get Horas_Resposta() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Horas_Resposta
    Horas_Resposta = mvarHoras_Resposta
End Property



Public Property Let Data_Resposta(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_Resposta = 5
    mvarData_Resposta = vData
End Property


Public Property Get Data_Resposta() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_Resposta
    Data_Resposta = mvarData_Resposta
End Property



Public Property Let Numero_Protocolo(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Numero_Protocolo = 5
    mvarNumero_Protocolo = vData
End Property


Public Property Get Numero_Protocolo() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Numero_Protocolo
    Numero_Protocolo = mvarNumero_Protocolo
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



Public Property Let Informacao_Obito(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Informacao_Obito = 5
    mvarInformacao_Obito = vData
End Property


Public Property Get Informacao_Obito() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Informacao_Obito
    Informacao_Obito = mvarInformacao_Obito
End Property



Public Property Let Total_Cheques_Sustados_Online(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Cheques_Sustados_Online = 5
    mvarTotal_Cheques_Sustados_Online = vData
End Property


Public Property Get Total_Cheques_Sustados_Online() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Cheques_Sustados_Online
    Total_Cheques_Sustados_Online = mvarTotal_Cheques_Sustados_Online
    If Total_Cheques_Sustados_Online = "" Then Total_Cheques_Sustados_Online = "0000"
End Property



Public Property Let Total_Menssagem_Online(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Menssagem_Online = 5
    mvarTotal_Menssagem_Online = vData
End Property


Public Property Get Total_Menssagem_Online() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Menssagem_Online
    Total_Menssagem_Online = mvarTotal_Menssagem_Online
End Property



Public Property Let Total_TaloesCheques_ContraOrdenados_AGCC_Diferente(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_TaloesCheques_ContraOrdenados_AGCC_Diferente = 5
    mvarTotal_TaloesCheques_ContraOrdenados_AGCC_Diferente = vData
End Property


Public Property Get Total_TaloesCheques_ContraOrdenados_AGCC_Diferente() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_TaloesCheques_ContraOrdenados_AGCC_Diferente
    Total_TaloesCheques_ContraOrdenados_AGCC_Diferente = mvarTotal_TaloesCheques_ContraOrdenados_AGCC_Diferente
End Property



Public Property Let Total_TaloesCheques_ContraOrdenados_AG_Diferente(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_TaloesCheques_ContraOrdenados_AG_Diferente = 5
    mvarTotal_TaloesCheques_ContraOrdenados_AG_Diferente = vData
End Property


Public Property Get Total_TaloesCheques_ContraOrdenados_AG_Diferente() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_TaloesCheques_ContraOrdenados_AG_Diferente
    Total_TaloesCheques_ContraOrdenados_AG_Diferente = mvarTotal_TaloesCheques_ContraOrdenados_AG_Diferente
End Property



Public Property Let Total_TalaoCheques_ContraOrdenados_DOC_Diferente(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_TalaoCheques_ContraOrdenados_DOC_Diferente = 5
    mvarTotal_TalaoCheques_ContraOrdenados_DOC_Diferente = vData
End Property


Public Property Get Total_TalaoCheques_ContraOrdenados_DOC_Diferente() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_TalaoCheques_ContraOrdenados_DOC_Diferente
    Total_TalaoCheques_ContraOrdenados_DOC_Diferente = mvarTotal_TalaoCheques_ContraOrdenados_DOC_Diferente
End Property



Public Property Let Total_TalaoCheques_ContraOrdenados(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_TalaoCheques_ContraOrdenados = 5
    mvarTotal_TalaoCheques_ContraOrdenados = vData
End Property


Public Property Get Total_TalaoCheques_ContraOrdenados() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_TalaoCheques_ContraOrdenados
    Total_TalaoCheques_ContraOrdenados = mvarTotal_TalaoCheques_ContraOrdenados
End Property



Public Property Let Total_Cheques_Sustados_Linha21(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Cheques_Sustados_Linha21 = 5
    mvarTotal_Cheques_Sustados_Linha21 = vData
End Property


Public Property Get Total_Cheques_Sustados_Linha21() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Cheques_Sustados_Linha21
    Total_Cheques_Sustados_Linha21 = mvarTotal_Cheques_Sustados_Linha21
End Property



Public Property Let Qtde_Dias_Contumacia(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Qtde_Dias_Contumacia = 5
    mvarQtde_Dias_Contumacia = vData
End Property


Public Property Get Qtde_Dias_Contumacia() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Qtde_Dias_Contumacia
    Qtde_Dias_Contumacia = mvarQtde_Dias_Contumacia
End Property



Public Property Let Total_Cheques_ContraOrdenados_AGConta_Diferente(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Cheques_ContraOrdenados_AGConta_Diferente = 5
    mvarTotal_Cheques_ContraOrdenados_AGConta_Diferente = vData
End Property


Public Property Get Total_Cheques_ContraOrdenados_AGConta_Diferente() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Cheques_ContraOrdenados_AGConta_Diferente
    Total_Cheques_ContraOrdenados_AGConta_Diferente = mvarTotal_Cheques_ContraOrdenados_AGConta_Diferente
End Property



Public Property Let Total_Cheques_ContraOrdenados_AG_Diferente(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Cheques_ContraOrdenados_AG_Diferente = 5
    mvarTotal_Cheques_ContraOrdenados_AG_Diferente = vData
End Property


Public Property Get Total_Cheques_ContraOrdenados_AG_Diferente() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Cheques_ContraOrdenados_AG_Diferente
    Total_Cheques_ContraOrdenados_AG_Diferente = mvarTotal_Cheques_ContraOrdenados_AG_Diferente
End Property



Public Property Let Total_Cheques_Contra_Ordenados_DIF_Consultado(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Cheques_Contra_Ordenados_DIF_Consultado = 5
    mvarTotal_Cheques_Contra_Ordenados_DIF_Consultado = vData
End Property


Public Property Get Total_Cheques_Contra_Ordenados_DIF_Consultado() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Cheques_Contra_Ordenados_DIF_Consultado
    Total_Cheques_Contra_Ordenados_DIF_Consultado = mvarTotal_Cheques_Contra_Ordenados_DIF_Consultado
End Property



Public Property Let Total_Ocorrencias_Credito_Concedidos(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Ocorrencias_Credito_Concedidos = 5
    mvarTotal_Ocorrencias_Credito_Concedidos = vData
End Property


Public Property Get Total_Ocorrencias_Credito_Concedidos() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Ocorrencias_Credito_Concedidos
    Total_Ocorrencias_Credito_Concedidos = mvarTotal_Ocorrencias_Credito_Concedidos
End Property



Public Property Let Total_Acoes(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Acoes = 5
    mvarTotal_Acoes = vData
End Property


Public Property Get Total_Acoes() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Acoes
    Total_Acoes = mvarTotal_Acoes
End Property



Public Property Let Total_Protesto(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Protesto = 5
    mvarTotal_Protesto = vData
End Property


Public Property Get Total_Protesto() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Protesto
    Total_Protesto = mvarTotal_Protesto
End Property



Public Property Let Total_Consultas_Realizadas(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Consultas_Realizadas = 5
    mvarTotal_Consultas_Realizadas = vData
End Property


Public Property Get Total_Consultas_Realizadas() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Consultas_Realizadas
    Total_Consultas_Realizadas = mvarTotal_Consultas_Realizadas
End Property



Public Property Let Qtde_Dias_Consulta_Realizada(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Qtde_Dias_Consulta_Realizada = 5
    mvarQtde_Dias_Consulta_Realizada = vData
End Property


Public Property Get Qtde_Dias_Consulta_Realizada() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Qtde_Dias_Consulta_Realizada
    Qtde_Dias_Consulta_Realizada = mvarQtde_Dias_Consulta_Realizada
End Property



Public Property Let Total_Alertas(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Alertas = 5
    mvarTotal_Alertas = vData
End Property


Public Property Get Total_Alertas() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Alertas
    Total_Alertas = mvarTotal_Alertas
End Property



Public Property Let Total_Cheque_ContraOrdenados(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Cheque_ContraOrdenados = 5
    mvarTotal_Cheque_ContraOrdenados = vData
End Property


Public Property Get Total_Cheque_ContraOrdenados() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Cheque_ContraOrdenados
    Total_Cheque_ContraOrdenados = mvarTotal_Cheque_ContraOrdenados
End Property



Public Property Let Total_Cheque_Lojista(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Cheque_Lojista = 5
    mvarTotal_Cheque_Lojista = vData
End Property


Public Property Get Total_Cheque_Lojista() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Cheque_Lojista
    Total_Cheque_Lojista = mvarTotal_Cheque_Lojista
End Property



Public Property Let Total_Cheque_CCF(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Cheque_CCF = 5
    mvarTotal_Cheque_CCF = vData
End Property


Public Property Get Total_Cheque_CCF() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Cheque_CCF
    Total_Cheque_CCF = mvarTotal_Cheque_CCF
End Property



Public Property Let Total_SPC(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_SPC = 5
    mvarTotal_SPC = vData
End Property


Public Property Get Total_SPC() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_SPC
    Total_SPC = mvarTotal_SPC
End Property



Public Property Let Data_Ultima_Atu_Endereco(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_Ultima_Atu_Endereco = 5
    mvarData_Ultima_Atu_Endereco = vData
End Property


Public Property Get Data_Ultima_Atu_Endereco() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_Ultima_Atu_Endereco
    Data_Ultima_Atu_Endereco = mvarData_Ultima_Atu_Endereco
End Property



Public Property Let Telefone(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Telefone = 5
    mvarTelefone = vData
End Property


Public Property Get Telefone() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Telefone
    Telefone = mvarTelefone
End Property



Public Property Let CEP(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Cep = 5
    mvarCEP = vData
End Property


Public Property Get CEP() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Cep
    CEP = mvarCEP
End Property



Public Property Let UF(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.UF = 5
    mvarUF = vData
End Property


Public Property Get UF() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.UF
    UF = mvarUF
End Property



Public Property Let Cidade(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Cidade = 5
    mvarCidade = vData
End Property


Public Property Get Cidade() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Cidade
    Cidade = mvarCidade
End Property



Public Property Let Bairro(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Bairro = 5
    mvarBairro = vData
End Property


Public Property Get Bairro() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Bairro
    Bairro = mvarBairro
End Property



Public Property Let Complemento(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Complemento = 5
    mvarComplemento = vData
End Property


Public Property Get Complemento() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Complemento
    Complemento = mvarComplemento
End Property



Public Property Let Endereco(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Endereco = 5
    mvarEndereco = vData
End Property


Public Property Get Endereco() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Endereco
    Endereco = mvarEndereco
End Property



Public Property Let Data_Nascimento_Fundacao(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_Nascimento_Fundacao = 5
    mvarData_Nascimento_Fundacao = vData
End Property


Public Property Get Data_Nascimento_Fundacao() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_Nascimento_Fundacao
    Data_Nascimento_Fundacao = mvarData_Nascimento_Fundacao
End Property



Public Property Let Nome(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Nome = 5
    mvarNome = vData
End Property


Public Property Get Nome() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Nome
    Nome = mvarNome
End Property



Public Property Let RG(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.RG = 5
    mvarRG = vData
End Property


Public Property Get RG() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.RG
    RG = mvarRG
End Property



Public Property Let CPF_CNPJ(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.CPF_CNPJ = 5
    mvarCPF_CNPJ = vData
End Property


Public Property Get CPF_CNPJ() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.CPF_CNPJ
    CPF_CNPJ = mvarCPF_CNPJ
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



Public Property Let Senha_Operador(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Senha_Operador = 5
    mvarSenha_Operador = vData
End Property


Public Property Get Senha_Operador() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Senha_Operador
    Senha_Operador = mvarSenha_Operador
End Property



Public Property Let Codigo_Operador(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Codigo_Operador = 5
    mvarCodigo_Operador = vData
End Property


Public Property Get Codigo_Operador() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Codigo_Operador
    Codigo_Operador = mvarCodigo_Operador
End Property



Public Property Let Codigo_Estacao(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Codigo_Estacao = 5
    mvarCodigo_Estacao = vData
End Property


Public Property Get Codigo_Estacao() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Codigo_Estacao
    Codigo_Estacao = mvarCodigo_Estacao
End Property



Public Property Let Meio_de_Acesso(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Meio_de_Acesso = 5
    mvarMeio_de_Acesso = vData
End Property


Public Property Get Meio_de_Acesso() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Meio_de_Acesso
    Meio_de_Acesso = mvarMeio_de_Acesso
End Property



Public Property Let Versao(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Versao = 5
    mvarVersao = vData
End Property


Public Property Get Versao() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Versao
    Versao = mvarVersao
End Property



Public Property Let Codigo_Transacao(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Codigo_Transacao = 5
    mvarCodigo_Transacao = vData
End Property


Public Property Get Codigo_Transacao() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Codigo_Transacao
    Codigo_Transacao = mvarCodigo_Transacao
End Property



Public Property Let Total_RegistroConsulta(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Alertas = 5
    mvarTotal_RegistroConsulta = vData
End Property


Public Property Get Total_RegistroConsulta() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Alertas
    Total_RegistroConsulta = mvarTotal_RegistroConsulta
End Property


Public Property Let Total_InfoPoderJudiciario(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Alertas = 5
    mvarTotal_InfoPoderJudiciario = vData
End Property


Public Property Get Total_InfoPoderJudiciario() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Alertas
    Total_InfoPoderJudiciario = mvarTotal_InfoPoderJudiciario
End Property




Public Property Let Total_HistoricoPagamento(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Alertas = 5
    mvarTotal_HistoricoPagamento = vData
End Property


Public Property Get Total_HistoricoPagamento() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Alertas
    Total_HistoricoPagamento = mvarTotal_HistoricoPagamento
End Property



Public Property Let Total_ControleSocietario(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Alertas = 5
    mvarTotal_ControleSocietario = vData
End Property


Public Property Get Total_ControleSocietario() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Alertas
    Total_ControleSocietario = mvarTotal_ControleSocietario
End Property




Public Property Let Total_PendenciaSerasa(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Alertas = 5
    mvarTotal_PendenciaSerasa = vData
End Property


Public Property Get Total_PendenciaSerasa() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Alertas
    Total_PendenciaSerasa = mvarTotal_PendenciaSerasa
End Property




Public Property Let Total_ParticipacaoFalencia(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Total_Alertas = 5
    mvarTotal_ParticipacaoFalencia = vData
End Property


Public Property Get Total_ParticipacaoFalencia() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Total_Alertas
    Total_ParticipacaoFalencia = mvarTotal_ParticipacaoFalencia
End Property


