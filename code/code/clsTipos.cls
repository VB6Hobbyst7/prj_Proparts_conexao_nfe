VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private This As ClassType


Private Type ClassType
	ID As Long
	codMod As Long
	Descricao As String
End Type


Property Get ID() As Long:ID = This.ID:End Property
Property Let ID(pValue As Long):This.ID = pValue:End Property
Property Get codMod() As Long:codMod = This.codMod:End Property
Property Let codMod(pValue As Long):This.codMod = pValue:End Property
Property Get Descricao() As String:Descricao = This.Descricao:End Property
Property Let Descricao(pValue As String):This.Descricao = pValue:End Property


sub cadastrar(obj as clsTipos)
Dim db As DAO.Database: Set db = CurrentDb
Dim qdf As DAO.QueryDef: Set qdf = db.QueryDefs("INSERT_tblTipos")


With qdf
	.Parameters("strID") = obj.ID
	.Parameters("strcodMod") = obj.codMod
	.Parameters("strDescricao") = obj.Descricao
	.Execute
End With
End Sub

