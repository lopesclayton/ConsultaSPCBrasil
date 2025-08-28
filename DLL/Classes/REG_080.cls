VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_080"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarAssinante As String                            'local copy
Private mvarTipo_Logradouro As String                      'local copy
Private mvarEndereco As String                             'local copy
Private mvarNumero As String                               'local copy
Private mvarCidade As String                               'local copy
Private mvarBairro As String                               'local copy
Private mvarCEP As String                                  'local copy
Private mvarUF As String                                   'local copy
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



Public Property Let CEP(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Cep = 5
    mvarCEP = vData
End Property


Public Property Get CEP() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Cep
    CEP = mvarCEP
End Property



Public Property Let Bairro(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Bairro = 5
    mvarBairro = vData
End Property


Public Property Get Bairro() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Bairro
    Bairro = mvarBairro
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



Public Property Let Numero(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.NUmero = 5
    mvarNumero = vData
End Property


Public Property Get Numero() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.NUmero
    Numero = mvarNumero
End Property



Public Property Let Endereco(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Endereco = 5
    mvarEndereco = vData
End Property


Public Property Get Endereco() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Endereco
    Endereco = mvarEndereco
End Property



Public Property Let Tipo_Logradouro(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Tipo_Logradouro = 5
    mvarTipo_Logradouro = vData
End Property


Public Property Get Tipo_Logradouro() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Tipo_Logradouro
    Tipo_Logradouro = mvarTipo_Logradouro
End Property



Public Property Let Assinante(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Assinante = 5
    mvarAssinante = vData
End Property


Public Property Get Assinante() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Assinante
    Assinante = mvarAssinante
End Property



