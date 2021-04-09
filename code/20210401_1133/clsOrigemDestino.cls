VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOrigemDestino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type ClassType
	Origem As String
	Destino As String
	Tipo As String
	Tag As String
End Type
Private This As ClassType


Property Get Origem() As String:Origem = This.Origem:End Property
Property Let Origem(pValue As String):This.Origem = pValue:End Property
Property Get Destino() As String:Destino = This.Destino:End Property
Property Let Destino(pValue As String):This.Destino = pValue:End Property
Property Get Tipo() As String:Tipo = This.Tipo:End Property
Property Let Tipo(pValue As String):This.Tipo = pValue:End Property
Property Get Tag() As String:Tag = This.Tag:End Property
Property Let Tag(pValue As String):This.Tag = pValue:End Property


sub cadastrar()
Dim db As DAO.Database: Set db = CurrentDb
Dim qdf As DAO.QueryDef: Set qdf = db.QueryDefs("INSERT_tblOrigemDestino")


With qdf
	.Parameters("Origem") = Origem
	.Parameters("Destino") = Destino
	.Parameters("Tipo") = Tipo
	.Parameters("Tag") = Tag
	.Execute
End With
End Sub

