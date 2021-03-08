VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private This As ClassType


Private Type ClassType
	ID As Long
	TipoDeParametro As String
	ValorDoParametro As String
End Type


Property Get ID() As Long:ID = This.ID:End Property
Property Let ID(pValue As Long):This.ID = pValue:End Property
Property Get TipoDeParametro() As String:TipoDeParametro = This.TipoDeParametro:End Property
Property Let TipoDeParametro(pValue As String):This.TipoDeParametro = pValue:End Property
Property Get ValorDoParametro() As String:ValorDoParametro = This.ValorDoParametro:End Property
Property Let ValorDoParametro(pValue As String):This.ValorDoParametro = pValue:End Property


sub cadastrar(obj as clsParametros)
Dim db As DAO.Database: Set db = CurrentDb
Dim qdf As DAO.QueryDef: Set qdf = db.QueryDefs("INSERT_tblParametros")


With qdf
	.Parameters("strID") = obj.ID
	.Parameters("strTipoDeParametro") = obj.TipoDeParametro
	.Parameters("strValorDoParametro") = obj.ValorDoParametro
	.Execute
End With
End Sub

