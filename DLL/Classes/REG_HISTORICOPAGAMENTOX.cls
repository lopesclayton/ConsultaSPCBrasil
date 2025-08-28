VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_HISTORICOPGTOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit


'local variable(s) to hold property value(s)
Private mvarPeriodo As String 'local copy
Private mvarQuantidade As String 'local copy
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



Public Property Let Quantidade(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quantidade = 5
    mvarQuantidade = vData
End Property


Public Property Get Quantidade() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Quantidade
    Quantidade = mvarQuantidade
End Property



Public Property Let Periodo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Periodo = 5
    mvarPeriodo = vData
End Property


Public Property Get Periodo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Periodo
    Periodo = mvarPeriodo
End Property



