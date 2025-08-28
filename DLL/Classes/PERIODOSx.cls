VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PERIODOSx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mvarPagamento_Comprovado As String 'local copy
Private mvarDescricao As String 'local copy
Private mvarPercentual As String 'local copy
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



Public Property Let Descricao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Descricao = 5
    mvarDescricao = vData
End Property


Public Property Get Descricao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Descricao
    Descricao = mvarDescricao
End Property



Public Property Let Pagamento_Comprovado(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Pagamento_Comprovado = 5
    mvarPagamento_Comprovado = vData
End Property


Public Property Get Pagamento_Comprovado() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Pagamento_Comprovado
    Pagamento_Comprovado = mvarPagamento_Comprovado
End Property



