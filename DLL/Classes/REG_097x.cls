VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_097x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCodigo_Servico As String                       'local copy
'local variable(s) to hold property value(s)
Private mvarCNPJ As String                                 'local copy
Private mvarRazao_Social As String                         'local copy
Private mvarNome_Comercial As String                       'local copy
Private mvarPercentual_Participacao As String              'local copy
Private mvarData_Entrada As String                         'local copy
Private mvarData_Saida As String                           'local copy
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



Public Property Let Nome_Comercial(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Nome_Comercial = 5
    mvarNome_Comercial = vData
End Property


Public Property Get Nome_Comercial() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Nome_Comercial
    Nome_Comercial = mvarNome_Comercial
End Property



Public Property Let Razao_Social(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Razao_Social = 5
    mvarRazao_Social = vData
End Property


Public Property Get Razao_Social() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Razao_Social
    Razao_Social = mvarRazao_Social
End Property



Public Property Let CNPJ(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.CNPJ = 5
    mvarCNPJ = vData
End Property


Public Property Get CNPJ() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.CNPJ
    CNPJ = mvarCNPJ
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



