VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG_CAPITALSOCIAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'local variable(s) to hold property value(s)
Private mvarNacionalidade As String 'local copy
'local variable(s) to hold property value(s)
Private mvarNatureza As String 'local copy
'local variable(s) to hold property value(s)
Private mvarOrigem As String 'local copy
'local variable(s) to hold property value(s)
Private mvarDataAtualizacao As String 'local copy
'local variable(s) to hold property value(s)
Private mvarValorCapitalAutorizado As String 'local copy
'local variable(s) to hold property value(s)
Private mvarValorCapitalRealizado As String 'local copy
'local variable(s) to hold property value(s)
Private mvarValorCapitalSocial As String 'local copy
Public Property Let ValorCapitalSocial(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ValorCapitalSocial = 5
    mvarValorCapitalSocial = Format(vData, "#,##0.00")
End Property


Public Property Get ValorCapitalSocial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ValorCapitalSocial
    ValorCapitalSocial = mvarValorCapitalSocial
End Property



Public Property Let ValorCapitalRealizado(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ValorCapitalRealizado = 5
    mvarValorCapitalRealizado = Format(vData, "#,##0.00")
End Property


Public Property Get ValorCapitalRealizado() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ValorCapitalRealizado
    ValorCapitalRealizado = mvarValorCapitalRealizado
End Property



Public Property Let ValorCapitalAutorizado(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ValorCapitalAutorizado = 5
    mvarValorCapitalAutorizado = Format(vData, "#,##0.00")
End Property


Public Property Get ValorCapitalAutorizado() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ValorCapitalAutorizado
    ValorCapitalAutorizado = mvarValorCapitalAutorizado
End Property



Public Property Let DataAtualizacao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataAtualizacao = 5
    mvarDataAtualizacao = Format(Left(vData, 10), "DD/MM/YYYY")
End Property


Public Property Get DataAtualizacao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DataAtualizacao
    DataAtualizacao = mvarDataAtualizacao
End Property



Public Property Let Origem(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Origem = 5
    mvarOrigem = vData
End Property


Public Property Get Origem() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Origem
    Origem = mvarOrigem
End Property



Public Property Let Natureza(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Natureza = 5
    mvarNatureza = vData
End Property


Public Property Get Natureza() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Natureza
    Natureza = mvarNatureza
End Property



Public Property Let Nacionalidade(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Nacionalidade = 5
    mvarNacionalidade = vData
End Property


Public Property Get Nacionalidade() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Nacionalidade
    Nacionalidade = mvarNacionalidade
End Property



