VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_028x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCodigo_Banco As String                         'local copy
Private mvarCodigo_Agencia As String                       'local copy
Private mvarConta_Corrente As String                       'local copy
'local variable(s) to hold property value(s)
Private mvarNumero_Cheque As String                        'local copy
'local variable(s) to hold property value(s)
Private mvarCodigo_Alinea As String                        'local copy
'local variable(s) to hold property value(s)
Private mvarOrigem As String                               'local copy
'local variable(s) to hold property value(s)
Private mvarData_Abertura_ContaCorrente As String          'local copy
Public Property Let Data_Abertura_ContaCorrente(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_Abertura_ContaCorrente = 5
    mvarData_Abertura_ContaCorrente = vData
End Property


Public Property Get Data_Abertura_ContaCorrente() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_Abertura_ContaCorrente
    Data_Abertura_ContaCorrente = mvarData_Abertura_ContaCorrente
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



Public Property Let Numero_Cheque(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Numero_Cheque = 5
    mvarNumero_Cheque = vData
End Property


Public Property Get Numero_Cheque() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Numero_Cheque
    Numero_Cheque = mvarNumero_Cheque
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



