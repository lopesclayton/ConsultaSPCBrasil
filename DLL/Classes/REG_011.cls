VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarNumero_RG As String                            'local copy
Private mvarUF_Emissao_RG As String                        'local copy
Private mvarNome As String                                 'local copy
Private mvarData_Nascimento As String                      'local copy
Private mvarNome_Mae As String                             'local copy
Private mvarNome_Pai As String                             'local copy
Private mvarCidade As String                               'local copy
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



Public Property Let Nome_Pai(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Nome_Pai = 5
    mvarNome_Pai = vData
End Property


Public Property Get Nome_Pai() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Nome_Pai
    Nome_Pai = mvarNome_Pai
End Property



Public Property Let Nome_Mae(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Nome_Mae = 5
    mvarNome_Mae = vData
End Property


Public Property Get Nome_Mae() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Nome_Mae
    Nome_Mae = mvarNome_Mae
End Property



Public Property Let Data_Nascimento(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_Nascimento = 5
    mvarData_Nascimento = vData
End Property


Public Property Get Data_Nascimento() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_Nascimento
    Data_Nascimento = mvarData_Nascimento
End Property



Public Property Let Nome(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Nome = 5
    mvarNome = vData
End Property


Public Property Get Nome() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Nome
    Nome = mvarNome
End Property



Public Property Let UF_Emissao_RG(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.UF_Emissao_RG = 5
    mvarUF_Emissao_RG = vData
End Property


Public Property Get UF_Emissao_RG() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.UF_Emissao_RG
    UF_Emissao_RG = mvarUF_Emissao_RG
End Property



Public Property Let Numero_RG(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Numero_RG = 5
    mvarNumero_RG = vData
End Property


Public Property Get Numero_RG() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Numero_RG
    Numero_RG = mvarNumero_RG
End Property



