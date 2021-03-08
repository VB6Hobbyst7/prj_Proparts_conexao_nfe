VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDadosConexaoNFeCTe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private This As ClassType


Private Type ClassType
	ID As Long
	ID_Empresa As Long
	ID_Tipo As Long
	codMod As Long
	codIntegrado As Long
	dhEmi As String
	CNPJ_emit As String
	Razao_emit As String
	CNPJ_Rem As String
	CPNJ_Dest As String
	CaminhoDoArquivo As String
	Chave As String
	Comando As String
End Type


Property Get ID() As Long:ID = This.ID:End Property
Property Let ID(pValue As Long):This.ID = pValue:End Property
Property Get ID_Empresa() As Long:ID_Empresa = This.ID_Empresa:End Property
Property Let ID_Empresa(pValue As Long):This.ID_Empresa = pValue:End Property
Property Get ID_Tipo() As Long:ID_Tipo = This.ID_Tipo:End Property
Property Let ID_Tipo(pValue As Long):This.ID_Tipo = pValue:End Property
Property Get codMod() As Long:codMod = This.codMod:End Property
Property Let codMod(pValue As Long):This.codMod = pValue:End Property
Property Get codIntegrado() As Long:codIntegrado = This.codIntegrado:End Property
Property Let codIntegrado(pValue As Long):This.codIntegrado = pValue:End Property
Property Get dhEmi() As String:dhEmi = This.dhEmi:End Property
Property Let dhEmi(pValue As String):This.dhEmi = pValue:End Property
Property Get CNPJ_emit() As String:CNPJ_emit = This.CNPJ_emit:End Property
Property Let CNPJ_emit(pValue As String):This.CNPJ_emit = pValue:End Property
Property Get Razao_emit() As String:Razao_emit = This.Razao_emit:End Property
Property Let Razao_emit(pValue As String):This.Razao_emit = pValue:End Property
Property Get CNPJ_Rem() As String:CNPJ_Rem = This.CNPJ_Rem:End Property
Property Let CNPJ_Rem(pValue As String):This.CNPJ_Rem = pValue:End Property
Property Get CPNJ_Dest() As String:CPNJ_Dest = This.CPNJ_Dest:End Property
Property Let CPNJ_Dest(pValue As String):This.CPNJ_Dest = pValue:End Property
Property Get CaminhoDoArquivo() As String:CaminhoDoArquivo = This.CaminhoDoArquivo:End Property
Property Let CaminhoDoArquivo(pValue As String):This.CaminhoDoArquivo = pValue:End Property
Property Get Chave() As String:Chave = This.Chave:End Property
Property Let Chave(pValue As String):This.Chave = pValue:End Property
Property Get Comando() As String:Comando = This.Comando:End Property
Property Let Comando(pValue As String):This.Comando = pValue:End Property


sub cadastrar(obj as clsDadosConexaoNFeCTe)
Dim db As DAO.Database: Set db = CurrentDb
Dim qdf As DAO.QueryDef: Set qdf = db.QueryDefs("INSERT_tblDadosConexaoNFeCTe")


With qdf
	.Parameters("strID") = obj.ID
	.Parameters("strID_Empresa") = obj.ID_Empresa
	.Parameters("strID_Tipo") = obj.ID_Tipo
	.Parameters("strcodMod") = obj.codMod
	.Parameters("strcodIntegrado") = obj.codIntegrado
	.Parameters("strdhEmi") = obj.dhEmi
	.Parameters("strCNPJ_emit") = obj.CNPJ_emit
	.Parameters("strRazao_emit") = obj.Razao_emit
	.Parameters("strCNPJ_Rem") = obj.CNPJ_Rem
	.Parameters("strCPNJ_Dest") = obj.CPNJ_Dest
	.Parameters("strCaminhoDoArquivo") = obj.CaminhoDoArquivo
	.Parameters("strChave") = obj.Chave
	.Parameters("strComando") = obj.Comando
	.Execute
End With
End Sub

