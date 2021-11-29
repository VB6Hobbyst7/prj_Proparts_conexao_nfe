<<<<<<< HEAD
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProcessamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ClassType
        pk As String
        Chave As String
        valor As String
End Type
Private This As ClassType

Private Const queryName As String = "azsCreateQueryInserttblProcessamento"
Private Const queryScript As String = "INSERT INTO tblProcessamento ( pk, chave, valor ) SELECT [pk] AS strpk, [chave] AS strchave, [valor] AS strvalor;"

Property Get pk() As String: pk = This.pk: End Property
Property Let pk(pValue As String): This.pk = pValue: End Property
Property Get Chave() As String: Chave = This.Chave: End Property
Property Let Chave(pValue As String): This.Chave = pValue: End Property
Property Get valor() As String: valor = This.valor: End Property
Property Let valor(pValue As String): This.valor = pValue: End Property

Public Sub cadastrar()
Dim db As DAO.Database: Set db = CurrentDb

    qryDeleteExists queryName
    qryCreate queryName, queryScript
    
    Dim qdf As DAO.QueryDef: Set qdf = db.QueryDefs(queryName)
    
    With qdf
            .Parameters("pk") = pk
            .Parameters("chave") = Chave
            .Parameters("valor") = valor
            .Execute
    End With
    
    qryDeleteExists queryName
    
    db.Close: Set db = Nothing

End Sub

=======
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProcessamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ClassType
        pk As String
        Chave As String
        valor As String
End Type
Private This As ClassType



Private Const queryName As String = "azsCreateQueryInserttblProcessamento"
Private Const queryScript As String = "INSERT INTO tblProcessamento ( pk, chave, valor ) SELECT [pk] AS strpk, [chave] AS strchave, [valor] AS strvalor;"

Property Get pk() As String: pk = This.pk: End Property
Property Let pk(pValue As String): This.pk = pValue: End Property
Property Get Chave() As String: Chave = This.Chave: End Property
Property Let Chave(pValue As String): This.Chave = pValue: End Property
Property Get valor() As String: valor = This.valor: End Property
Property Let valor(pValue As String): This.valor = pValue: End Property

Public Sub cadastrar()
Dim db As DAO.Database: Set db = CurrentDb

    qryDeleteExists queryName
    qryCreate queryName, queryScript
    
    Dim qdf As DAO.QueryDef: Set qdf = db.QueryDefs(queryName)
    
    With qdf
            .Parameters("pk") = pk
            .Parameters("chave") = Chave
            .Parameters("valor") = valor
            .Execute
    End With
    
    qryDeleteExists queryName
    
    db.Close: Set db = Nothing

End Sub

>>>>>>> f4084cb29d769387d25e7b837853d2119e0da429
