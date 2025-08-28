VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_INDICE_PONTUALx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"PERIODOS"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarNome As String 'local copy
Private mvarPERIODOS As PERIODOS


Public Property Get PERIODOS() As PERIODOS
    If mvarPERIODOS Is Nothing Then
        Set mvarPERIODOS = New PERIODOS
    End If


    Set PERIODOS = mvarPERIODOS
End Property


Public Property Set PERIODOS(vData As PERIODOS)
    Set mvarPERIODOS = vData
End Property
Private Sub Class_Terminate()
    Set mvarPERIODOS = Nothing
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



