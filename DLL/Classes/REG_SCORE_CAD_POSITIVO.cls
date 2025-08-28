VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_SCORE_POSITIVO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarRisco As String 'local copy
Private mvarMensagem As String 'local copy
Private mvarProbabilidadeInadimplencia As String 'local copy
Private mvarScore As String 'local copy
Private mvarScoreMaximo As String 'local copy
Public Property Let ScoreMaximo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ScoreMaximo = 5
    mvarScoreMaximo = vData
End Property


Public Property Get ScoreMaximo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ScoreMaximo
    ScoreMaximo = mvarScoreMaximo
End Property



Public Property Let Score(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Score = 5
    mvarScore = vData
End Property


Public Property Get Score() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Score
    Score = mvarScore
End Property



Public Property Let ProbabilidadeInadimplencia(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProbabilidadeInadimplencia = 5
    mvarProbabilidadeInadimplencia = vData
End Property


Public Property Get ProbabilidadeInadimplencia() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProbabilidadeInadimplencia
    ProbabilidadeInadimplencia = mvarProbabilidadeInadimplencia
End Property



Public Property Let Mensagem(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Mensagem = 5
    mvarMensagem = vData
End Property


Public Property Get Mensagem() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Mensagem
    Mensagem = mvarMensagem
End Property



Public Property Let Risco(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Risco = 5
    mvarRisco = vData
End Property


Public Property Get Risco() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Risco
    Risco = mvarRisco
End Property



