VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_092x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCodigo_Servico As String                       'local copy
Private mvarData_Incluso As String                         'local copy
Private mvarData_Ocorrencia As String                      'local copy
Private mvarOrigem As String                               'local copy
Private mvarCodigo_Banco As String                         'local copy
Private mvarNome_Banco As String                           'local copy
Private mvarCodigo_Agencia As String                       'local copy
Private mvarConta_Corrente As String                       'local copy
Private mvarCheque_Inicial As String                       'local copy
Private mvarCheque_Final As String                         'local copy
Private mvarValor As String                                'local copy
Private mvarCodigo_Alinea As String                        'local copy
Private mvarEntidade_Origem As String                      'local copy
Public Property Let Entidade_Origem(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Entidade_Origem = 5
    mvarEntidade_Origem = vData
End Property


Public Property Get Entidade_Origem() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Entidade_Origem
    Entidade_Origem = mvarEntidade_Origem
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



Public Property Let Cheque_Final(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Cheque_Final = 5
    mvarCheque_Final = vData
End Property


Public Property Get Cheque_Final() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Cheque_Final
    Cheque_Final = mvarCheque_Final
End Property



Public Property Let Cheque_Inicial(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Cheque_Inicial = 5
    mvarCheque_Inicial = vData
End Property


Public Property Get Cheque_Inicial() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Cheque_Inicial
    Cheque_Inicial = mvarCheque_Inicial
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



Public Property Let Data_Incluso(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_Incluso = 5
    mvarData_Incluso = vData
End Property


Public Property Get Data_Incluso() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_Incluso
    Data_Incluso = mvarData_Incluso
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



