VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_040x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarData_Inclusao As String                        'local copy
Private mvarData_Acao As String                            'local copy
Private mvarTipo_Acao As String                            'local copy
Private mvarVara As String                                 'local copy
Private mvarCidade As String                               'local copy
Private mvarUF As String                                   'local copy
Private mvarAutor As String                                'local copy
Private mvarValor As String                                'local copy
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



Public Property Let Autor(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Autor = 5
    mvarAutor = vData
End Property


Public Property Get Autor() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Autor
    Autor = mvarAutor
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



Public Property Let Vara(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Vara = 5
    mvarVara = vData
End Property


Public Property Get Vara() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Vara
    Vara = mvarVara
End Property



Public Property Let Tipo_Acao(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Tipo_Acao = 5
    mvarTipo_Acao = vData
End Property


Public Property Get Tipo_Acao() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Tipo_Acao
    Tipo_Acao = mvarTipo_Acao
End Property



Public Property Let Data_Acao(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_Acao = 5
    mvarData_Acao = vData
End Property


Public Property Get Data_Acao() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_Acao
    Data_Acao = mvarData_Acao
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



