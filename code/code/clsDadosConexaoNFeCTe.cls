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


Private Type ClassType
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
Private This As ClassType


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


sub cadastrar()
Dim db As DAO.Database: Set db = CurrentDb
Dim qdf As DAO.QueryDef: Set qdf = db.QueryDefs("INSERT_tblDadosConexaoNFeCTe")


With qdf
	.Parameters("ID_Empresa") = ID_Empresa
	.Parameters("ID_Tipo") = ID_Tipo
	.Parameters("codMod") = codMod
	.Parameters("codIntegrado") = codIntegrado
	.Parameters("dhEmi") = dhEmi
	.Parameters("CNPJ_emit") = CNPJ_emit
	.Parameters("Razao_emit") = Razao_emit
	.Parameters("CNPJ_Rem") = CNPJ_Rem
	.Parameters("CPNJ_Dest") = CPNJ_Dest
	.Parameters("CaminhoDoArquivo") = CaminhoDoArquivo
	.Parameters("Chave") = Chave
	.Parameters("Comando") = Comando
	.Execute
End With
End Sub

