VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_096x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCodigo_Servico As String                       'local copy
Private mvarTipo_Documento As String                       'local copy
Private mvarCPF_CNPJ As String                             'local copy
Private mvarNome_Socio_Proprietario As String              'local copy
Private mvarPercentual_Participacao As String              'local copy
Private mvarData_Entrada As String                         'local copy
Private mvarData_Saida As String                           'local copy
Private mvarCondicao As String                             'local copy
Private mvarTipo_Relacionamento As String                  'local copy
Private mvarCargo_Direcao As String                        'local copy
Private mvarCargo_Administracao As String                  'local copy
Private mvarValor_Participacao As String                   'local copy
Public Property Let Valor_Participacao(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Valor_Participacao = 5
    mvarValor_Participacao = vData
End Property


Public Property Get Valor_Participacao() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Valor_Participacao
    Valor_Participacao = mvarValor_Participacao
End Property



Public Property Let Cargo_Administracao(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Cargo_Administracao = 5
    mvarCargo_Administracao = vData
End Property


Public Property Get Cargo_Administracao() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Cargo_Administracao
    Cargo_Administracao = mvarCargo_Administracao
End Property



Public Property Let Cargo_Direcao(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Cargo_Direcao = 5
    mvarCargo_Direcao = vData
End Property


Public Property Get Cargo_Direcao() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Cargo_Direcao
    Cargo_Direcao = mvarCargo_Direcao
End Property



Public Property Let Tipo_Relacionamento(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Tipo_Relacionamento = 5
    mvarTipo_Relacionamento = vData
End Property


Public Property Get Tipo_Relacionamento() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Tipo_Relacionamento
    Tipo_Relacionamento = mvarTipo_Relacionamento
End Property



Public Property Let Condicao(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Condicao = 5
    mvarCondicao = vData
End Property


Public Property Get Condicao() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Condicao
    Condicao = mvarCondicao
End Property



Public Property Let Data_Saida(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_Saida = 5
    mvarData_Saida = vData
End Property


Public Property Get Data_Saida() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_Saida
    Data_Saida = mvarData_Saida
End Property



Public Property Let Data_Entrada(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_Entrada = 5
    mvarData_Entrada = vData
End Property


Public Property Get Data_Entrada() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_Entrada
    Data_Entrada = mvarData_Entrada
End Property



Public Property Let Percentual_Participacao(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Percentual_Participacao = 5
    mvarPercentual_Participacao = vData
End Property


Public Property Get Percentual_Participacao() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Percentual_Participacao
    Percentual_Participacao = mvarPercentual_Participacao
End Property



Public Property Let Nome_Socio_Proprietario(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Nome_Socio_Proprietario = 5
    mvarNome_Socio_Proprietario = vData
End Property


Public Property Get Nome_Socio_Proprietario() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Nome_Socio_Proprietario
    Nome_Socio_Proprietario = mvarNome_Socio_Proprietario
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



