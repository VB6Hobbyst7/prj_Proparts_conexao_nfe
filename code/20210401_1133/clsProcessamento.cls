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


Property Get pk() As String: pk = This.pk: End Property
Property Let pk(pValue As String): This.pk = pValue: End Property
Property Get Chave() As String: Chave = This.Chave: End Property
Property Let Chave(pValue As String): This.Chave = pValue: End Property
Property Get valor() As String: valor = This.valor: End Property
Property Let valor(pValue As String): This.valor = pValue: End Property


Sub cadastrar()
Dim db As dao.Database: Set db = CurrentDb
Dim qdf As dao.QueryDef: Set qdf = db.QueryDefs("INSERT_tblProcessamento")

With qdf
        .Parameters("pk") = pk
        .Parameters("chave") = Chave
        .Parameters("valor") = valor
        .Execute
End With
End Sub

