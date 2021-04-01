VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConexaoNFeCTeDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private pCol As New Collection
Private Const qryParametro As String = "SELECT tblParametros.ValorDoParametro FROM tblParametros WHERE (((tblParametros.TipoDeParametro) = 'strParametro'))"
Private Const qryProcessamento As String = "select nomecampo, valor from tblProcessamento where pk = 'strChave'"


Private Type ClassType
        ID_Empresa As String
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
        codUsuarioErp As String
        nomeUsuarioErp As String
        strFileName As String
        Comando As String
        codTipoEvento As String
        CFOP As String
        emit_UF As String
        rem_UF As String
        ChvAcesso As String
        
        CaminhoDeProcessados As String
End Type
Private This As ClassType


Property Get ID_Empresa() As String: ID_Empresa = This.ID_Empresa: End Property
Property Let ID_Empresa(pValue As String): This.ID_Empresa = pValue: End Property
Property Get ID_Tipo() As Long: ID_Tipo = This.ID_Tipo: End Property
Property Let ID_Tipo(pValue As Long): This.ID_Tipo = pValue: End Property
Property Get codMod() As Long: codMod = This.codMod: End Property
Property Let codMod(pValue As Long): This.codMod = pValue: End Property
Property Get codIntegrado() As Long: codIntegrado = This.codIntegrado: End Property
Property Let codIntegrado(pValue As Long): This.codIntegrado = pValue: End Property
Property Get dhEmi() As String: dhEmi = This.dhEmi: End Property
Property Let dhEmi(pValue As String): This.dhEmi = pValue: End Property
Property Get CNPJ_emit() As String: CNPJ_emit = This.CNPJ_emit: End Property
Property Let CNPJ_emit(pValue As String): This.CNPJ_emit = pValue: End Property
Property Get Razao_emit() As String: Razao_emit = This.Razao_emit: End Property
Property Let Razao_emit(pValue As String): This.Razao_emit = pValue: End Property
Property Get CNPJ_Rem() As String: CNPJ_Rem = This.CNPJ_Rem: End Property
Property Let CNPJ_Rem(pValue As String): This.CNPJ_Rem = pValue: End Property
Property Get CPNJ_Dest() As String: CPNJ_Dest = This.CPNJ_Dest: End Property
Property Let CPNJ_Dest(pValue As String): This.CPNJ_Dest = pValue: End Property

Property Get CaminhoDoArquivo() As String: CaminhoDoArquivo = This.CaminhoDoArquivo: End Property
Property Let CaminhoDoArquivo(pValue As String): This.CaminhoDoArquivo = pValue: End Property

Property Get Chave() As String: Chave = This.Chave: End Property
Property Let Chave(pValue As String): This.Chave = pValue: End Property

Property Get codUsuarioErp() As String: codUsuarioErp = This.codUsuarioErp: End Property
Property Let codUsuarioErp(pValue As String): This.codUsuarioErp = pValue: End Property
Property Get nomeUsuarioErp() As String: nomeUsuarioErp = This.nomeUsuarioErp: End Property
Property Let nomeUsuarioErp(pValue As String): This.nomeUsuarioErp = pValue: End Property

Property Get strFileName() As String: strFileName = This.strFileName: End Property
Property Let strFileName(pValue As String): This.strFileName = pValue: End Property

Property Get Comando() As String: Comando = This.Comando: End Property
Property Let Comando(pValue As String): This.Comando = pValue: End Property
Property Get codTipoEvento() As String: codTipoEvento = This.codTipoEvento: End Property
Property Let codTipoEvento(pValue As String): This.codTipoEvento = pValue: End Property
Property Get CFOP() As String: CFOP = This.CFOP: End Property
Property Let CFOP(pValue As String): This.CFOP = pValue: End Property
Property Get emit_UF() As String: emit_UF = This.emit_UF: End Property
Property Let emit_UF(pValue As String): This.emit_UF = pValue: End Property
Property Get rem_UF() As String: rem_UF = This.rem_UF: End Property
Property Let rem_UF(pValue As String): This.rem_UF = pValue: End Property

Property Get ChvAcesso() As String: ChvAcesso = This.ChvAcesso: End Property
Property Let ChvAcesso(pValue As String): This.ChvAcesso = pValue: End Property

Property Get CaminhoDeProcessados() As String: CaminhoDeProcessados = This.CaminhoDeProcessados: End Property
Property Let CaminhoDeProcessados(pValue As String): This.CaminhoDeProcessados = pValue: End Property



Sub carregar_dados(strChave As String)
Dim db As DAO.Database: Set db = CurrentDb
Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(Replace(qryProcessamento, "strChave", strChave))

'' ########################################################
'' RELACIONAMENTO DE DADOS COLETADOS COM A TABELA DESTINO
'' ########################################################

    Do While Not rst.EOF
        Select Case rst.Fields("nomecampo").Value
        
            Case "codMod": codMod = rst.Fields("valor").Value
            Case "dhEmi": dhEmi = rst.Fields("valor").Value
            Case "CNPJ_emit": CNPJ_emit = rst.Fields("valor").Value
            Case "Razao_emit": Razao_emit = rst.Fields("valor").Value
            Case "CNPJ_Rem": CNPJ_Rem = rst.Fields("valor").Value
            Case "CPNJ_Dest": CPNJ_Dest = rst.Fields("valor").Value
            Case "Chave": Chave = rst.Fields("valor").Value
            
            Case "CFOP": CFOP = rst.Fields("valor").Value
            Case "emit_UF": emit_UF = rst.Fields("valor").Value
            Case "rem_UF": rem_UF = rst.Fields("valor").Value
            Case "ChvAcesso": ChvAcesso = rst.Fields("valor").Value
            
            Case "CaminhoDoArquivo": CaminhoDoArquivo = rst.Fields("valor").Value
                        
        End Select
        
        rst.MoveNext
    Loop

db.Close

Set db = Nothing

End Sub

Sub cadastrar()
Dim db As DAO.Database: Set db = CurrentDb
Dim qdf As DAO.QueryDef: Set qdf = db.QueryDefs("INSERT_tblDadosConexaoNFeCTe")
Dim t As Variant

    With qdf
            
            '' qryUpdateID_EMPRESA
            .Parameters("ID_Empresa") = DLookup("[ID_Empresa]", "tblEmpresa", "[CNPJ_Empresa]='" & STRPontos(CNPJ_emit) & "'")

            .Parameters("ID_Tipo") = 0
            
            .Parameters("codMod") = codMod
            .Parameters("codIntegrado") = codIntegrado
            .Parameters("dhEmi") = dhEmi
            .Parameters("CNPJ_emit") = CNPJ_emit
            .Parameters("Razao_emit") = Razao_emit
            .Parameters("CNPJ_Rem") = CNPJ_Rem
            .Parameters("CPNJ_Dest") = CPNJ_Dest
            
            .Parameters("CFOP") = CFOP
            .Parameters("emit_UF") = emit_UF
            .Parameters("rem_UF") = rem_UF
            .Parameters("ChvAcesso") = ChvAcesso
            
            .Parameters("CaminhoDoArquivo") = CaminhoDoArquivo
            .Parameters("Chave") = Chave
            .Parameters("Comando") = Comando
            .Parameters("TipoEvento") = codTipoEvento
            .Parameters("registroValido") = 0
            
            .Execute '' dbRunAsync
            
'            While .StillExecuting = True
'                t = 1
'            Wend
            
    End With

End Sub

Public Function criarERP()
Dim objERP As Collection: Set objERP = New Collection

    objERP.add "{"
    objERP.add "    ""Chave"": """ & This.ChvAcesso & ""","
    objERP.add "    ""Comando"": """ & This.Comando & """, "
    objERP.add "    ""data"": " & This.dhEmi & ","
    objERP.add "    ""codUsuarioErp"": """ & This.codUsuarioErp & """, "
    objERP.add "    ""nomeUsuarioErp"": """ & This.nomeUsuarioErp & """"
    objERP.add "}"

This.strFileName = This.ChvAcesso & ".erp.json"

execucao objERP, This.strFileName, strApp:="Notepad.exe", pOperacao:=opNone, strFilePath:=This.CaminhoDeProcessados

End Function

Public Function criarManifesto()
Dim objManifesto As Collection: Set objManifesto = New Collection

    objManifesto.add "{"
    objManifesto.add "    ""chaveAcessoNFe"": """ & This.ChvAcesso & ""","
    objManifesto.add "    ""codTipoEvento"": """ & This.codTipoEvento & """, "
    objManifesto.add "    ""data"": " & This.dhEmi & ","
    objManifesto.add "    ""codUsuarioErp"": """ & This.codUsuarioErp & """, "
    objManifesto.add "    ""nomeUsuarioErp"": """ & This.nomeUsuarioErp & """"
    objManifesto.add "}"

This.strFileName = This.ChvAcesso & "-" & codTipoEvento & ".man.json"

execucao objManifesto, This.strFileName, strApp:="Notepad.exe", pOperacao:=opNone, strFilePath:=This.CaminhoDeProcessados

End Function


Public Function NewEnum() As IUnknown
    Set NewEnum = pCol.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set pCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set pCol = Nothing
End Sub

Public Sub add(ByVal rec As clsConexaoNfeCteDados, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As clsConexaoNfeCteDados)
    pCol.add rec
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function clsConexaoNfeCteDados(ByVal i As Variant) As clsConexaoNfeCteDados
    Set clsConexaoNfeCteDados = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As clsConexaoNfeCteDados
    Set Item = pCol(i)
End Property
