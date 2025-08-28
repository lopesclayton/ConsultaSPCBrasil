VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_095"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit


'local variable(s) to hold property value(s)
Private mvarCodigo_Servico As String                       'local copy
Private mvarNome As String                                 'local copy
Private mvarNome_Fantasia As String                        'local copy
Private mvarRG As String                                   'local copy
Private mvarTitulo_Eleitor As String                       'local copy
Private mvarData_Nasc_Fundacao As String                   'local copy
Private mvarSituacao_CPF_CNPJ As String                    'local copy
Private mvarNome_Mae As String                             'local copy
Private mvarEndereco As String                             'local copy
Private mvarComplemento As String                          'local copy
Private mvarBairro As String                               'local copy
Private mvarCidade As String                               'local copy
Private mvarUF As String                                   'local copy
Private mvarCEP As String                                  'local copy
Private mvarData_Ultima_Alteracao_Endereco As String       'local copy
Private mvarDDD As String                                  'local copy
Private mvarTelefone As String                             'local copy
Private mvarSituacao_CPF_CNPJ_RF As String                 'local copy
Private mvarData_Situacao_RF As String                     'local copy
Private mvarHora_Situacao_RF As String                     'local copy
Private mvarOrigem_CPF As String                           'local copy
Private mvarSexo As String                                 'local copy
Private mvarSigno As String                                'local copy
Private mvarIdade As String                                'local copy
Private mvarIE As String                                   'local copy
Private mvarIdade_Consumidor As String                     'local copy
Private mvarUF_IE As String                                'local copy
Private mvarSituacao_IE As String                          'local copy
Private mvarData_Situacao_IE As String                     'local copy
Private mvarHora_Situacao_IE As String                     'local copy
Private mvarCodigo_Natureza_Juridica As String             'local copy
Private mvarDescricao_Natureza_JUridica As String          'local copy
Private mvarCNAE As String                                 'local copy
Private mvarDescricao_CNAE As String                       'local copy
Private mvarNIRC_NIRE As String                            'local copy
Private mvarValor_Capital_Social As String                 'local copy
Private mvarRazao_Antecessora As String                    'local copy
Public Property Let Razao_Antecessora(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Razao_Antecessora = 5
    mvarRazao_Antecessora = vData
End Property


Public Property Get Razao_Antecessora() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Razao_Antecessora
    Razao_Antecessora = mvarRazao_Antecessora
End Property



Public Property Let Valor_Capital_Social(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Valor_Capital_Social = 5
    mvarValor_Capital_Social = vData
End Property


Public Property Get Valor_Capital_Social() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Valor_Capital_Social
    Valor_Capital_Social = mvarValor_Capital_Social
End Property



Public Property Let NIRC_NIRE(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.NIRC_NIRE = 5
    mvarNIRC_NIRE = vData
End Property


Public Property Get NIRC_NIRE() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.NIRC_NIRE
    NIRC_NIRE = mvarNIRC_NIRE
End Property



Public Property Let Descricao_CNAE(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Descricao_CNAE = 5
    mvarDescricao_CNAE = vData
End Property


Public Property Get Descricao_CNAE() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Descricao_CNAE
    Descricao_CNAE = mvarDescricao_CNAE
End Property



Public Property Let CNAE(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.CNAE = 5
    mvarCNAE = vData
End Property


Public Property Get CNAE() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.CNAE
    CNAE = mvarCNAE
End Property



Public Property Let Descricao_Natureza_JUridica(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Descricao_Natureza_JUridica = 5
    mvarDescricao_Natureza_JUridica = vData
End Property


Public Property Get Descricao_Natureza_JUridica() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Descricao_Natureza_JUridica
    Descricao_Natureza_JUridica = mvarDescricao_Natureza_JUridica
End Property



Public Property Let Codigo_Natureza_Juridica(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Codigo_Natureza_Juridica = 5
    mvarCodigo_Natureza_Juridica = vData
End Property


Public Property Get Codigo_Natureza_Juridica() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Codigo_Natureza_Juridica
    Codigo_Natureza_Juridica = mvarCodigo_Natureza_Juridica
End Property



Public Property Let Hora_Situacao_IE(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Hora_Situacao_IE = 5
    mvarHora_Situacao_IE = vData
End Property


Public Property Get Hora_Situacao_IE() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Hora_Situacao_IE
    Hora_Situacao_IE = mvarHora_Situacao_IE
End Property



Public Property Let Data_Situacao_IE(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_Situacao_IE = 5
    mvarData_Situacao_IE = vData
End Property


Public Property Get Data_Situacao_IE() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_Situacao_IE
    Data_Situacao_IE = mvarData_Situacao_IE
End Property



Public Property Let Situacao_IE(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Situacao_IE = 5
    mvarSituacao_IE = vData
End Property


Public Property Get Situacao_IE() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Situacao_IE
    Situacao_IE = mvarSituacao_IE
End Property



Public Property Let UF_IE(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.UF_IE = 5
    mvarUF_IE = vData
End Property


Public Property Get UF_IE() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.UF_IE
    UF_IE = mvarUF_IE
End Property



Public Property Let Idade_Consumidor(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Idade_Consumidor = 5
    mvarIdade_Consumidor = vData
End Property


Public Property Get Idade_Consumidor() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Idade_Consumidor
    Idade_Consumidor = mvarIdade_Consumidor
End Property



Public Property Let IE(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.IE = 5
    mvarIE = vData
End Property


Public Property Get IE() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.IE
    IE = mvarIE
End Property



Public Property Let Idade(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Idade = 5
    mvarIdade = vData
End Property


Public Property Get Idade() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Idade
    Idade = mvarIdade
End Property



Public Property Let Signo(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Signo = 5
    mvarSigno = vData
End Property


Public Property Get Signo() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Signo
    Signo = mvarSigno
End Property



Public Property Let Sexo(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Sexo = 5
    mvarSexo = vData
End Property


Public Property Get Sexo() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Sexo
    Sexo = mvarSexo
End Property



Public Property Let Origem_CPF(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Origem_CPF = 5
    mvarOrigem_CPF = vData
End Property


Public Property Get Origem_CPF() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Origem_CPF
    Origem_CPF = mvarOrigem_CPF
End Property



Public Property Let Hora_Situacao_RF(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Hora_Situacao_RF = 5
    mvarHora_Situacao_RF = vData
End Property


Public Property Get Hora_Situacao_RF() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Hora_Situacao_RF
    Hora_Situacao_RF = mvarHora_Situacao_RF
End Property



Public Property Let Data_Situacao_RF(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_Situacao_RF = 5
    mvarData_Situacao_RF = vData
End Property


Public Property Get Data_Situacao_RF() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_Situacao_RF
    Data_Situacao_RF = mvarData_Situacao_RF
End Property



Public Property Let Situacao_CPF_CNPJ_RF(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Situacao_CPF_CNPJ_RF = 5
    mvarSituacao_CPF_CNPJ_RF = vData
End Property


Public Property Get Situacao_CPF_CNPJ_RF() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Situacao_CPF_CNPJ_RF
    Situacao_CPF_CNPJ_RF = mvarSituacao_CPF_CNPJ_RF
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



Public Property Let DDD(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.DDD = 5
    mvarDDD = vData
End Property


Public Property Get DDD() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.DDD
    DDD = mvarDDD
End Property



Public Property Let Data_Ultima_Alteracao_Endereco(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_Ultima_Alteracao_Endereco = 5
    mvarData_Ultima_Alteracao_Endereco = vData
End Property


Public Property Get Data_Ultima_Alteracao_Endereco() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_Ultima_Alteracao_Endereco
    Data_Ultima_Alteracao_Endereco = mvarData_Ultima_Alteracao_Endereco
End Property



Public Property Let CEP(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.CEP = 5
    mvarCEP = vData
End Property


Public Property Get CEP() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.CEP
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



Public Property Let Nome_Mae(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Nome_Mae = 5
    mvarNome_Mae = vData
End Property


Public Property Get Nome_Mae() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Nome_Mae
    Nome_Mae = mvarNome_Mae
End Property



Public Property Let Situacao_CPF_CNPJ(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Situacao_CPF_CNPJ = 5
    mvarSituacao_CPF_CNPJ = vData
End Property


Public Property Get Situacao_CPF_CNPJ() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Situacao_CPF_CNPJ
    Situacao_CPF_CNPJ = mvarSituacao_CPF_CNPJ
End Property



Public Property Let Data_Nasc_Fundacao(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_Nasc_Fundacao = 5
    mvarData_Nasc_Fundacao = vData
End Property


Public Property Get Data_Nasc_Fundacao() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_Nasc_Fundacao
    Data_Nasc_Fundacao = mvarData_Nasc_Fundacao
End Property



Public Property Let Titulo_Eleitor(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Titulo_Eleitor = 5
    mvarTitulo_Eleitor = vData
End Property


Public Property Get Titulo_Eleitor() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Titulo_Eleitor
    Titulo_Eleitor = mvarTitulo_Eleitor
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



Public Property Let Nome_Fantasia(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Nome_Fantasia = 5
    mvarNome_Fantasia = vData
End Property


Public Property Get Nome_Fantasia() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Nome_Fantasia
    Nome_Fantasia = mvarNome_Fantasia
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



Public Property Let Codigo_Servico(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Codigo_Servico = 5
    mvarCodigo_Servico = vData
End Property


Public Property Get Codigo_Servico() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Codigo_Servico
    Codigo_Servico = mvarCodigo_Servico
End Property



