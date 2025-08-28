VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_COMPROM_RENDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Option Explicit

'local variable(s) to hold property value(s)
Private mvarClassificacao As String 'local copy
Private mvarFaixa As String 'local copy
Private mvarPercentual As String 'local copy
Private mvarFlag_Endividamento As String 'local copy
Public Property Let Flag_Endividamento(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Flag_Endividamento = 5
    mvarFlag_Endividamento = vData
End Property


Public Property Get Flag_Endividamento() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Flag_Endividamento
    Flag_Endividamento = mvarFlag_Endividamento
End Property



Public Property Let Percentual(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Percentual = 5
    mvarPercentual = vData
End Property


Public Property Get Percentual() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Percentual
    Percentual = mvarPercentual
End Property



Public Property Let Faixa(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Faixa = 5
    mvarFaixa = vData
End Property


Public Property Get Faixa() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Faixa
    Faixa = mvarFaixa
End Property



Public Property Let Classificacao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.classificacao = 5
    mvarClassificacao = vData
End Property


Public Property Get Classificacao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.classificacao
    Classificacao = mvarClassificacao
End Property



