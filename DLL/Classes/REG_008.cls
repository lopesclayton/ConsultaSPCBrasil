VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarHorizonte As String                            'local copy
Private mvarValor As String                                'local copy
Private mvarProbabilidade As String                        'local copy
Private mvarClasse_Risco As String                         'local copy
Private mvarMenssagem_Curta As String                      'local copy
Private mvarMenssagem_Interpretativa_Score As String       'local copy
Private mvarMessagem_Interpretativa_Probabilidade As String 'local copy
Public Property Let Messagem_Interpretativa_Probabilidade(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Messagem_Interpretativa_Probabilidade = 5
    mvarMessagem_Interpretativa_Probabilidade = vData
End Property


Public Property Get Messagem_Interpretativa_Probabilidade() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Messagem_Interpretativa_Probabilidade
    Messagem_Interpretativa_Probabilidade = mvarMessagem_Interpretativa_Probabilidade
End Property



Public Property Let Menssagem_Interpretativa_Score(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Menssagem_Interpretativa_Score = 5
    mvarMenssagem_Interpretativa_Score = vData
End Property


Public Property Get Menssagem_Interpretativa_Score() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Menssagem_Interpretativa_Score
    Menssagem_Interpretativa_Score = mvarMenssagem_Interpretativa_Score
End Property



Public Property Let Menssagem_Curta(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Menssagem_Curta = 5
    mvarMenssagem_Curta = vData
End Property


Public Property Get Menssagem_Curta() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Menssagem_Curta
    Menssagem_Curta = mvarMenssagem_Curta
End Property



Public Property Let Classe_Risco(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Classe_Risco = 5
    mvarClasse_Risco = vData
End Property


Public Property Get Classe_Risco() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Classe_Risco
    Classe_Risco = mvarClasse_Risco
End Property



Public Property Let Probabilidade(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Probabilidade = 5
    mvarProbabilidade = vData
End Property


Public Property Get Probabilidade() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Probabilidade
    Probabilidade = mvarProbabilidade
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



Public Property Let Horizonte(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Horizonte = 5
    mvarHorizonte = vData
End Property


Public Property Get Horizonte() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Horizonte
    Horizonte = mvarHorizonte
End Property



