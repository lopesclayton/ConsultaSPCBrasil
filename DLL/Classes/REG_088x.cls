VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_088x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarTipo_Documento As String                       'local copy
Private mvarCPF_CNPJ As String                             'local copy
Private mvarData_Ocorrencia As String                      'local copy
Private mvarCodigo_Banco As String                         'local copy
'local variable(s) to hold property value(s)
Private mvarNome_Banco As String                           'local copy
Private mvarCodigo_Agencia As String                       'local copy
Private mvarConta_Corrente As String                       'local copy
Private mvarNumero_Cheque_Inicial As String                'local copy
Private mvarNumero_Cheque_Final As String                  'local copy
Private mvarCodigo_Alinea As String                        'local copy
Private mvarValor As String                                'local copy
Private mvarOrigem As String                               'local copy
Private mvarEntidae_Origem As String                       'local copy
Public Property Let Entidae_Origem(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Entidae_Origem = 5
    mvarEntidae_Origem = vData
End Property


Public Property Get Entidae_Origem() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Entidae_Origem
    Entidae_Origem = mvarEntidae_Origem
End Property



Public Property Let Origem(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Origem = 5
    mvarOrigem = vData
End Property


Public Property Get Origem() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Origem
    Origem = mvarOrigem
End Property



Public Property Let Valor(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Valor = 5
    mvarValor = vData
End Property


Public Property Get Valor() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Valor
    Valor = mvarValor
End Property



Public Property Let Codigo_Alinea(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Codigo_Alinea = 5
    mvarCodigo_Alinea = vData
End Property


Public Property Get Codigo_Alinea() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Codigo_Alinea
    Codigo_Alinea = mvarCodigo_Alinea
End Property



Public Property Let Numero_Cheque_Final(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Numero_Cheque_Final = 5
    mvarNumero_Cheque_Final = vData
End Property


Public Property Get Numero_Cheque_Final() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Numero_Cheque_Final
    Numero_Cheque_Final = mvarNumero_Cheque_Final
End Property



Public Property Let Numero_Cheque_Inicial(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Numero_Cheque_Inicial = 5
    mvarNumero_Cheque_Inicial = vData
End Property


Public Property Get Numero_Cheque_Inicial() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Numero_Cheque_Inicial
    Numero_Cheque_Inicial = mvarNumero_Cheque_Inicial
End Property



Public Property Let Conta_Corrente(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Conta_Corrente = 5
    mvarConta_Corrente = vData
End Property


Public Property Get Conta_Corrente() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Conta_Corrente
    Conta_Corrente = mvarConta_Corrente
End Property



Public Property Let Codigo_Agencia(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Codigo_Agencia = 5
    mvarCodigo_Agencia = vData
End Property


Public Property Get Codigo_Agencia() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Codigo_Agencia
    Codigo_Agencia = mvarCodigo_Agencia
End Property



Public Property Let Nome_Banco(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Nome_Banco = 5
    mvarNome_Banco = vData
End Property


Public Property Get Nome_Banco() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Nome_Banco
    Nome_Banco = mvarNome_Banco
End Property



Public Property Let Codigo_Banco(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Codigo_Banco = 5
    mvarCodigo_Banco = vData
End Property


Public Property Get Codigo_Banco() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Codigo_Banco
    Codigo_Banco = mvarCodigo_Banco
End Property



Public Property Let Data_Ocorrencia(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_Ocorrencia = 5
    mvarData_Ocorrencia = vData
End Property


Public Property Get Data_Ocorrencia() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_Ocorrencia
    Data_Ocorrencia = mvarData_Ocorrencia
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



