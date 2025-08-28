VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_010x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarData_Inclusao As String                        'local copy
Private mvarData_Vencimento As String                      'local copy
Private mvarComprador_Fiador As String                     'local copy
Private mvarNumero_Contrato As String                      'local copy
Private mvarValor_debito As String                         'local copy
Private mvarNome_Associado As String                       'local copy
Private mvarCidade As String                               'local copy
Private mvarUF As String                                   'local copy
Private mvarNome_Consumidor As String                      'local copy
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



Public Property Let Nome_Consumidor(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Nome_Consumidor = 5
    mvarNome_Consumidor = vData
End Property


Public Property Get Nome_Consumidor() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Nome_Consumidor
    Nome_Consumidor = mvarNome_Consumidor
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



Public Property Let Nome_Associado(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Nome_Associado = 5
    mvarNome_Associado = vData
End Property


Public Property Get Nome_Associado() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Nome_Associado
    Nome_Associado = mvarNome_Associado
End Property



Public Property Let Valor_debito(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Valor_debito = 5
    mvarValor_debito = vData
End Property


Public Property Get Valor_debito() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Valor_debito
    Valor_debito = mvarValor_debito
End Property



Public Property Let Numero_Contrato(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Numero_Contrato = 5
    mvarNumero_Contrato = vData
End Property


Public Property Get Numero_Contrato() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Numero_Contrato
    Numero_Contrato = mvarNumero_Contrato
End Property



Public Property Let Comprador_Fiador(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Comprador_Fiador = 5
    mvarComprador_Fiador = vData
End Property


Public Property Get Comprador_Fiador() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Comprador_Fiador
    Comprador_Fiador = mvarComprador_Fiador
End Property



Public Property Let Data_Vencimento(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_Vencimento = 5
    mvarData_Vencimento = vData
End Property


Public Property Get Data_Vencimento() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_Vencimento
    Data_Vencimento = mvarData_Vencimento
End Property



Public Property Let Data_Inclusao(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_Inclusao = 5
    mvarData_Inclusao = vData
End Property


Public Property Get Data_Inclusao() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_Inclusao
    Data_Inclusao = mvarData_Inclusao
End Property



