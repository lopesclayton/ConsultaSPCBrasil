VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_INDICE_GASTOSx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarNome As String 'local copy
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



