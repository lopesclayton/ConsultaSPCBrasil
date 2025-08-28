VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_030x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarData_Inclusao As String                        'local copy
'local variable(s) to hold property value(s)
Private mvarData_Protesto As String                        'local copy
'local variable(s) to hold property value(s)
Private mvarCartorio As String                             'local copy
'local variable(s) to hold property value(s)
Private mvarCidade As String                               'local copy
'local variable(s) to hold property value(s)
Private mvarUF As String                                   'local copy
'local variable(s) to hold property value(s)
Private mvarRequerente As String                           'local copy
'local variable(s) to hold property value(s)
Private mvarValor As String                                'local copy
'local variable(s) to hold property value(s)
Private mvarNome_Consumidor As String                      'local copy
'local variable(s) to hold property value(s)
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



Public Property Let Requerente(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Requerente = 5
    mvarRequerente = vData
End Property


Public Property Get Requerente() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Requerente
    Requerente = mvarRequerente
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



Public Property Let Cartorio(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Cartorio = 5
    mvarCartorio = vData
End Property


Public Property Get Cartorio() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Cartorio
    Cartorio = mvarCartorio
End Property



Public Property Let Data_Protesto(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_Protesto = 5
    mvarData_Protesto = vData
End Property


Public Property Get Data_Protesto() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_Protesto
    Data_Protesto = mvarData_Protesto
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



