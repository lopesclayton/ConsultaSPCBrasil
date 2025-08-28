VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_LIMITECREDSUGx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarLimite_Sugerido As String 'local copy
Public Property Let Limite_Sugerido(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Limite_Sugerido = 5
    mvarLimite_Sugerido = vData
End Property


Public Property Get Limite_Sugerido() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Limite_Sugerido
    Limite_Sugerido = mvarLimite_Sugerido
End Property



