VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_005X"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCodigo_Base As String                          'local copy
Private mvarDescricao_Base As String                       'local copy
Private mvarMensagem As String                             'local copy
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



Public Property Let Descricao_Base(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Descricao_Base = 5
    mvarDescricao_Base = vData
End Property


Public Property Get Descricao_Base() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Descricao_Base
    Descricao_Base = mvarDescricao_Base
End Property



Public Property Let Codigo_Base(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Codigo_Base = 5
    mvarCodigo_Base = vData
End Property


Public Property Get Codigo_Base() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Codigo_Base
    Codigo_Base = mvarCodigo_Base
End Property



