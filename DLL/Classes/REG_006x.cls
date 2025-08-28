VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_006x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit


'local variable(s) to hold property value(s)
Private mvarRazao_Social As String                         'local copy
Public Property Let Razao_Social(ByVal vData As String)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.Razao_Social = 5
    mvarRazao_Social = vData
End Property


Public Property Get Razao_Social() As String
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.Razao_Social
    Razao_Social = mvarRazao_Social
End Property



