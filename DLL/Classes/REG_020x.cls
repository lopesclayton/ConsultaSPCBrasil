VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_020x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarTipo_Ocorrencia As String                      'local copy
Private mvarData_Inclusao As String                        'local copy
Private mvarHora_Inclusao As String                        'local copy
Private mvarData_Ocorrencia As String                      'local copy
'local variable(s) to hold property value(s)
Private mvarData_Emissao As String                         'local copy
'local variable(s) to hold property value(s)
Private mvarOrigem As String                               'local copy
Private mvarCodigo_Banco As String                         'local copy
Private mvarNome_Banco As String                           'local copy
'local variable(s) to hold property value(s)
Private mvarCodigo_Agencia As String                       'local copy
Private mvarConta_Corrente As String                       'local copy
'local variable(s) to hold property value(s)
Private mvarNumero_Cheque As String                        'local copy
'local variable(s) to hold property value(s)
Private mvarNumero_Cheque_Final As String                  'local copy
'local variable(s) to hold property value(s)
Private mvarQtde_Ocorrencias As String                     'local copy
'local variable(s) to hold property value(s)
Private mvarValor_Documento As String                      'local copy
'local variable(s) to hold property value(s)
Private mvarCodigo_Alinea As String                        'local copy
Private mvarMotivo As String                               'local copy
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



Public Property Let Valor_Documento(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Valor_Documento = 5
    mvarValor_Documento = vData
End Property


Public Property Get Valor_Documento() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Valor_Documento
    Valor_Documento = mvarValor_Documento
End Property



Public Property Let Qtde_Ocorrencias(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Qtde_Ocorrencias = 5
    mvarQtde_Ocorrencias = vData
End Property


Public Property Get Qtde_Ocorrencias() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Qtde_Ocorrencias
    Qtde_Ocorrencias = mvarQtde_Ocorrencias
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



Public Property Let Data_Emissao(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_Emissao = 5
    mvarData_Emissao = vData
End Property


Public Property Get Data_Emissao() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_Emissao
    Data_Emissao = mvarData_Emissao
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



Public Property Let Hora_Inclusao(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Hora_Inclusao = 5
    mvarHora_Inclusao = vData
End Property


Public Property Get Hora_Inclusao() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Hora_Inclusao
    Hora_Inclusao = mvarHora_Inclusao
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



Public Property Let Tipo_Ocorrencia(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Tipo_Ocorrencia = 5
    mvarTipo_Ocorrencia = vData
End Property


Public Property Get Tipo_Ocorrencia() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Tipo_Ocorrencia
    Tipo_Ocorrencia = mvarTipo_Ocorrencia
End Property



