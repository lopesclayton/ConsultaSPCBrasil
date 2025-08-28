VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProdutosX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Insumos"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCodigo As Long 'local copy
Private mvarNome As String 'local copy
Private mvarTipo As String 'local copy
Private mvarInsumos As Insumos


Public Property Get Insumos() As Insumos
    If mvarInsumos Is Nothing Then
        Set mvarInsumos = New Insumos
    End If


    Set Insumos = mvarInsumos
End Property


Public Property Set Insumos(vData As Insumos)
    Set mvarInsumos = vData
End Property
Private Sub Class_Terminate()
    Set mvarInsumos = Nothing
End Sub




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



Public Property Let Codigo(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Codigo = 5
    mvarCodigo = vData
End Property


Public Property Get Codigo() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Codigo
    Codigo = mvarCodigo
End Property



Public Property Let Tipo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tipo = 6
    mvarTipo = vData
End Property


Public Property Get Tipo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tipo
    Nome = mvarTipo
End Property

