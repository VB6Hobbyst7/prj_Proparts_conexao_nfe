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


Private This As ClassType


Private Type ClassType
	ID As Long
	Origem As String
	Destino As String
	Tipo As String
End Type


Property Get ID() As Long:ID = This.ID:End Property
Property Let ID(pValue As Long):This.ID = pValue:End Property
Property Get Origem() As String:Origem = This.Origem:End Property
Property Let Origem(pValue As String):This.Origem = pValue:End Property
Property Get Destino() As String:Destino = This.Destino:End Property
Property Let Destino(pValue As String):This.Destino = pValue:End Property
Property Get Tipo() As String:Tipo = This.Tipo:End Property
Property Let Tipo(pValue As String):This.Tipo = pValue:End Property


sub cadastrar(obj as clsOrigemDestino)
Dim db As DAO.Database: Set db = CurrentDb
Dim qdf As DAO.QueryDef: Set qdf = db.QueryDefs("INSERT_tblOrigemDestino")


With qdf
	.Parameters("strID") = obj.ID
	.Parameters("strOrigem") = obj.Origem
	.Parameters("strDestino") = obj.Destino
	.Parameters("strTipo") = obj.Tipo
	.Execute
End With
End Sub

