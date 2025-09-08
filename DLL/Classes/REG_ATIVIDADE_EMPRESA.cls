VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_ATIV_EMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarAtividadePrincipal As String 'local copy
Private mvarCodigoPrincipal As String 'local copy
Private mvarAtividadeSecundaria As String 'local copy
Private mvarCodigoSecundaria As String 'local copy
Public Property Let CodigoSecundaria(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CodigoSecundaria = 5
    mvarCodigoSecundaria = vData
End Property


Public Property Get CodigoSecundaria() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodigoSecundaria
    CodigoSecundaria = mvarCodigoSecundaria
End Property



Public Property Let AtividadeSecundaria(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AtividadeSecundaria = 5
    mvarAtividadeSecundaria = vData
End Property


Public Property Get AtividadeSecundaria() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AtividadeSecundaria
    AtividadeSecundaria = mvarAtividadeSecundaria
End Property



Public Property Let CodigoPrincipal(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CodigoPrincipal = 5
    mvarCodigoPrincipal = vData
End Property


Public Property Get CodigoPrincipal() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodigoPrincipal
    CodigoPrincipal = mvarCodigoPrincipal
End Property



Public Property Let AtividadePrincipal(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AtividadePrincipal = 5
    mvarAtividadePrincipal = vData
End Property


Public Property Get AtividadePrincipal() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AtividadePrincipal
    AtividadePrincipal = mvarAtividadePrincipal
End Property



