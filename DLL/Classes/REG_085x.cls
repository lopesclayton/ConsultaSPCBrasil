VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_085x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarDDD As String                                  'local copy
Private mvarTelefone As String                             'local copy
Private mvarData_PrimeiraConsulta As String                'local copy
Private mvarData_UltimaConsulta As String                  'local copy
Public Property Let Data_UltimaConsulta(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_UltimaConsulta = 5
    mvarData_UltimaConsulta = vData
End Property


Public Property Get Data_UltimaConsulta() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_UltimaConsulta
    Data_UltimaConsulta = mvarData_UltimaConsulta
End Property



Public Property Let Data_PrimeiraConsulta(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Data_PrimeiraConsulta = 5
    mvarData_PrimeiraConsulta = vData
End Property


Public Property Get Data_PrimeiraConsulta() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Data_PrimeiraConsulta
    Data_PrimeiraConsulta = mvarData_PrimeiraConsulta
End Property



Public Property Let Telefone(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Telefone = 5
    mvarTelefone = vData
End Property


Public Property Get Telefone() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Telefone
    Telefone = mvarTelefone
End Property



Public Property Let DDD(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.DDD = 5
    mvarDDD = vData
End Property


Public Property Get DDD() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.DDD
    DDD = mvarDDD
End Property



