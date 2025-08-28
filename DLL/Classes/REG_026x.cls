VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_026x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCodigo_Banco As String                         'local copy
Private mvarNome_Banco As String                           'local copy
'local variable(s) to hold property value(s)
Private mvarCodigo_Agencia As String                       'local copy
'local variable(s) to hold property value(s)
Private mvarNome_Agencia As String                         'local copy
'local variable(s) to hold property value(s)
Private mvarEndereco As String                             'local copy
'local variable(s) to hold property value(s)
Private mvarBairro As String                               'local copy
'local variable(s) to hold property value(s)
Private mvarCidade As String                               'local copy
'local variable(s) to hold property value(s)
Private mvarUF As String                                   'local copy
'local variable(s) to hold property value(s)
Private mvarCEP As String                                  'local copy
'local variable(s) to hold property value(s)
Private mvarPraca As String                                'local copy
'local variable(s) to hold property value(s)
Private mvarMenssagem As String                            'local copy
Public Property Let Menssagem(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Menssagem = 5
    mvarMenssagem = vData
End Property


Public Property Get Menssagem() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Menssagem
    Menssagem = mvarMenssagem
End Property



Public Property Let Praca(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Praca = 5
    mvarPraca = vData
End Property


Public Property Get Praca() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Praca
    Praca = mvarPraca
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



Public Property Let Nome_Agencia(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Nome_Agencia = 5
    mvarNome_Agencia = vData
End Property


Public Property Get Nome_Agencia() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Nome_Agencia
    Nome_Agencia = mvarNome_Agencia
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



