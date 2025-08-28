VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_060x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarData_Inclusao As String                        'local copy
Private mvarMotivo As String                               'local copy
Private mvarTipo_Documento As String                       'local copy
Private mvarObservacao As String                           'local copy
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



Public Property Let Observacao(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Observacao = 5
    mvarObservacao = vData
End Property


Public Property Get Observacao() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Observacao
    Observacao = mvarObservacao
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



Public Property Let Motivo(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Motivo = 5
    mvarMotivo = vData
End Property


Public Property Get Motivo() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Motivo
    Motivo = mvarMotivo
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



