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


Private Type ClassType
	codMod As Long
	Descricao As String
End Type
Private This As ClassType


Property Get codMod() As Long:codMod = This.codMod:End Property
Property Let codMod(pValue As Long):This.codMod = pValue:End Property
Property Get Descricao() As String:Descricao = This.Descricao:End Property
Property Let Descricao(pValue As String):This.Descricao = pValue:End Property


sub cadastrar()
Dim db As DAO.Database: Set db = CurrentDb
Dim qdf As DAO.QueryDef: Set qdf = db.QueryDefs("INSERT_tblTipos")


With qdf
	.Parameters("codMod") = codMod
	.Parameters("Descricao") = Descricao
	.Execute
End With
End Sub

