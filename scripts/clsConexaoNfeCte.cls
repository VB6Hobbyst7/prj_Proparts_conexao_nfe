VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConexaoNfeCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private FileCollention As New Collection
Private con As ADODB.Connection

Public Enum enumTipoArquivo
    opFlagLancadaERP = 0
    opManifesto = 1
End Enum

'' ########################################
'' #TratamentoDeRegistros - tblOrigemDestino
'' ########################################

Private Const azsUpdateOrigemDestino_tabela_campo As String = _
        "UPDATE tblOrigemDestino SET tblOrigemDestino.tabela = strSplit([Destino],'.',0), tblOrigemDestino.campo = strSplit([Destino],'.',1);"


'' ########################################
'' #TratamentoDeRegistros - PROCESSAMENTO
'' ########################################

'' -- PROCESSAMENTO CONCLUIDO
Private Const qryUpdateProcessamentoConcluido As String = _
            "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 1 WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) " & _
            " AND ((tblDadosConexaoNFeCTe.registroProcessado)=0) AND ((tblDadosConexaoNFeCTe.Chave)='strChave'));"

'' -- CARREGAR TAGs DE VINDAS DO XML
'' #tblOrigemDestino.tabela
Private Const qryTags As String = _
        "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((Len([Tag]))>0) AND ((tblOrigemDestino.tabela) = 'strParametro') AND ((tblOrigemDestino.TagOrigem)=1));"

'' -- SELEÇÃO DE ITENS UNICOS PARA PROCESSAMENTO
'' #tblProcessamento
Private Const sqyProcessamentosPendentes As String = _
        "SELECT DISTINCT pk from tblProcessamento;"


'' -- LIMPAR TABELA
'' #tblProcessamento
Private Const qryDeleteProcessamento As String = _
        "DELETE * FROM tblProcessamento;"


'' -- ATUALIZAR CAMPOS ( Nometabela e NomeCampo ) PARA USO DA TABELA DE PROCESSAMENTO
'' #tblProcessamento.Tabela = [qryOrigemDestino].[Tabela]
'' #tblProcessamento.Campo = [qryOrigemDestino].[Campo]
'' #tblProcessamento.formatacao = [qryOrigemDestino].[formatacao]
Private Const qryUpdateProcessamento As String = _
                "UPDATE (SELECT tblOrigemDestino.Destino " & _
                "           ,tblOrigemDestino.Tag " & _
                "           ,tblOrigemDestino.Tabela " & _
                "           ,tblOrigemDestino.Campo " & _
                "           ,tblOrigemDestino.formatacao " & _
                "       FROM tblOrigemDestino " & _
                "       WHERE tblOrigemDestino.Tabela = 'strParametro' " & _
                "       ) AS qryOrigemDestino INNER JOIN tblProcessamento ON qryOrigemDestino.Tag = tblProcessamento.chave  " & _
                "SET tblProcessamento.nomeTabela = [qryOrigemDestino].[Tabela] " & _
                "   ,tblProcessamento.NomeCampo = [qryOrigemDestino].[Campo] " & _
                "   ,tblProcessamento.formatacao = [qryOrigemDestino].[formatacao];"


'' ------------------------
'' #FormatarCampos
'' ------------------------

Private Const qryUpdateProcessamento_opData As String = _
                "UPDATE tblProcessamento SET tblProcessamento.valor = Mid([tblProcessamento].[valor],1,10) WHERE formatacao = 'opData';"

Private Const qryUpdateProcessamento_opTime As String = _
                "UPDATE tblProcessamento SET tblProcessamento.valor = Mid([tblProcessamento].[valor],12,8) WHERE formatacao = 'opTime';"
'''-- NumberToSql
'Private Const qryUpdateProcessamento_opMoeda As String = _
'                "UPDATE tblProcessamento SET tblProcessamento.valor = NumberToSql([tblProcessamento].[valor]) WHERE formatacao = 'opMoeda';"
'                '' "UPDATE tblProcessamento SET tblProcessamento.valor = Format(Replace([tblProcessamento].[valor], '.', ','), '#,##0.00') WHERE formatacao = 'opMoeda';"


'' ########################################
'' #TratamentoDeRegistros - ARQUIVOS
'' ########################################

'' -- ARQUIVOS - CONSULTA PARA CRIAÇÃO DE ARQUIVOS JSON
'' #tblDadosConexaoNFeCTe
Private Const sqyDadosJson As String = _
        "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi]))>0));"


'' ########################################
'' #TratamentoDeRegistros - DADOS GERAIS
'' ########################################

'' -- SELEÇÃO DE FORNECEDORES VALIDOS
'' #tblDadosConexaoNFeCTe.registroValido
Private Const qryUpdateFornecedoresValidos As String = _
        "UPDATE (SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF FROM tmpClientes) AS qryFornecedoresValidos INNER JOIN tblDadosConexaoNFeCTe ON qryFornecedoresValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit SET tblDadosConexaoNFeCTe.registroValido = 1;"


'' -- SELEÇÃO DE REGISTROS VALIDOS
'' #tblDadosConexaoNFeCTe.registroValido
Private Const qryUpdateRegistrosValidos As String = _
        "UPDATE (SELECT STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF FROM tmpEmpresa) AS qryRegistrosValidos INNER JOIN tblDadosConexaoNFeCTe ON qryRegistrosValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit SET tblDadosConexaoNFeCTe.registroValido = 1;"


'' -- EXCLUSAO DE REGISTROS INVALIDOS
'' #tblDadosConexaoNFeCTe.registroValido
Private Const qryDeleteRegistrosInvalidos As String = _
        "DELETE * FROM tblDadosConexaoNFeCTe WHERE tblDadosConexaoNFeCTe.registroValido = 0;"


'' -- TIPOS DE CADASTRO - tblTipos
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryUpdateIdTipo As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.ID_Tipo = 0 WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) Is Null));"


'' -- FiltroFil
'' #tblDadosConexaoNFeCTe.ID_EMPRESA
Private Const qryUpdateIdEmpresa As String = _
        "UPDATE (SELECT tmpEmpresa.ID_Empresa, STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF FROM tmpEmpresa) AS qryEmpresas INNER JOIN tblDadosConexaoNFeCTe ON qryEmpresas.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit SET tblDadosConexaoNFeCTe.ID_Empresa = qryEmpresas.ID_Empresa;"

'' -- FiltroCFOP
Private Const qryUpdateCFOP_PSC_PES As String = _
            "UPDATE  ( SELECT " & _
            "           tmpNatOp.ID_NatOper, tmpNatOp.Fil_NatOper, tmpNatOp.CFOP_NatOper, qryPscPes.strXMLCFOP, qryPscPes.strEstado  " & _
            "       FROM (SELECT  " & _
            "               strSplit(ValorDoParametro,'|',0) AS strFil_NatOper,  strSplit(ValorDoParametro,'|',1) AS strEstado,  strSplit(ValorDoParametro,'|',2) AS strXMLCFOP,  strSplit(ValorDoParametro,'|',3) AS strCFOP_NatOper  " & _
            "             FROM  " & _
            "               tblParametros  " & _
            "             WHERE  " & _
            "               TipoDeParametro='FiltroFil' And strSplit(ValorDoParametro,'|',0) In ('PSC','PES'))  AS qryPscPes  " & _
            "       INNER JOIN tmpNatOp ON (qryPscPes.strCFOP_NatOper = tmpNatOp.CFOP_NatOper) AND (qryPscPes.strFil_NatOper = tmpNatOp.Fil_NatOper) )  AS tmpPscPes  " & _
            "INNER JOIN  " & _
            "   (   SELECT  *  " & _
            "       FROM  tblDadosConexaoNFeCTe  " & _
            "       WHERE tblDadosConexaoNFeCTe.registroValido IN (SELECT TOP 1 cint(tblParametros.ValorDoParametro) FROM [tblParametros] WHERE TipoDeParametro = 'registroValido')  " & _
            "       AND tblDadosConexaoNFeCTe.FiltroCFOP = 0 )  AS tmpDadosConexaoNFeCTe " & _
            "ON (tmpPscPes.strXMLCFOP = tmpDadosConexaoNFeCTe.CFOP) AND (tmpPscPes.Fil_NatOper = tmpDadosConexaoNFeCTe.ID_Empresa) " & _
            "SET  tmpDadosConexaoNFeCTe.FiltroCFOP = [tmpPscPes].[ID_NatOper];"

'' ------------------------------------------------------------------------------------------------------------------------
'' -- ID_TIPO - APENAS TIPOS COM ID DE VALOR ZERO(0) SERÃO ATUALIZADOS PARA NÃO COMPROMETER OS REGISTROS JÁ PROCESSADOS
'' ------------------------------------------------------------------------------------------------------------------------

'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 4 - NF-e Retorno Armazém
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryUpdateRetornoArmazem As String = _
        "UPDATE (Select ValorDoParametro,TipoDeParametro from tblParametros where TipoDeParametro = 'RetornoArmazem') as tmpRetornoArmazem INNER JOIN ((select top 1 cInt('55') as strMod, '12680452000302' as strCNPJ_CPF, 'RetornoArmazem' as strTipoDeParametro from tblParametros)  AS qryRetornoArmazem INNER JOIN tblDadosConexaoNFeCTe ON (qryRetornoArmazem.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit) AND (qryRetornoArmazem.strMod = tblDadosConexaoNFeCTe.codMod)) ON (tmpRetornoArmazem.TipoDeParametro = qryRetornoArmazem.strTipoDeParametro) AND (tmpRetornoArmazem.TipoDeParametro = qryRetornoArmazem.strTipoDeParametro) SET tblDadosConexaoNFeCTe.ID_Tipo = [tmpRetornoArmazem].[ValorDoParametro] WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)=0));"


'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 6 - NF-e Transferência com código Sisparts
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryUpdateTransferenciaSisparts As String = _
        "UPDATE (Select ValorDoParametro,TipoDeParametro from tblParametros where TipoDeParametro = 'TransferenciaSisparts') as tmpTransferenciaSisparts INNER JOIN ((SELECT CInt('55') AS strMod, 'TransferenciaSisparts' AS strTipoDeParametro, tmpEmpresa.CNPJ_Empresa AS strCNPJ_CPF FROM tmpEmpresa)  AS qryTransferenciaSisparts INNER JOIN tblDadosConexaoNFeCTe ON (qryTransferenciaSisparts.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit) AND (qryTransferenciaSisparts.strMod = tblDadosConexaoNFeCTe.codMod)) ON (tmpTransferenciaSisparts.TipoDeParametro = qryTransferenciaSisparts.strTipoDeParametro) AND (tmpTransferenciaSisparts.TipoDeParametro = qryTransferenciaSisparts.strTipoDeParametro) SET tblDadosConexaoNFeCTe.ID_Tipo = [tmpTransferenciaSisparts].[ValorDoParametro] WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)=0));"


'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 0 - CT-e
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryUpdateCTe As String = _
        "UPDATE (Select ValorDoParametro,TipoDeParametro from tblParametros where TipoDeParametro = 'CTe') as tmpCTe INNER JOIN ((select top 1 cInt('57') as strMod, '68365501000377' as strCNPJ_CPF, 'CTe' as strTipoDeParametro from tblParametros)  AS qryCTe INNER JOIN tblDadosConexaoNFeCTe ON (qryCTe.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit) AND (qryCTe.strMod = tblDadosConexaoNFeCTe.codMod)) ON (tmpCTe.TipoDeParametro = qryCTe.strTipoDeParametro) AND (tmpCTe.TipoDeParametro = qryCTe.strTipoDeParametro) SET tblDadosConexaoNFeCTe.ID_Tipo = [tmpCTe].[ValorDoParametro] WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)=0));"


'' ########################################
'' #TratamentoDeRegistros - COMPRAS
'' ########################################


Private Const qryUpdateFilCompraNF As String = _
        "UPDATE (SELECT tmpEmpresa.ID_Empresa, STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF, tmpEmpresa.CNPJ_Empresa FROM tmpEmpresa WHERE (((tmpEmpresa.CNPJ_Empresa) Is Not Null))) AS qryEmpresas INNER JOIN tblCompraNF ON qryEmpresas.strCNPJ_CPF = tblCompraNF.CNPJ_CPF_CompraNF SET tblCompraNF.Fil_CompraNF = qryEmpresas.ID_Empresa;"




Public Sub FormatarCampos()
Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                    qryUpdateProcessamento_opData, _
                                                    qryUpdateProcessamento_opTime)
    '' FORMATAR CAMPOS
    executarComandos qryProcessos

End Sub


Public Sub TratamentoDeCompras()
Dim qryProcessos() As Variant: qryProcessos = Array(qryUpdateFilCompraNF)

    '' REFINAMENTO DE DADOS
    executarComandos qryProcessos

End Sub

Public Sub TratamentoDeDadosGerais()
Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                qryUpdateFornecedoresValidos, _
                                                qryUpdateRegistrosValidos, _
                                                qryDeleteRegistrosInvalidos, _
                                                qryUpdateIdTipo, _
                                                qryUpdateRetornoArmazem, _
                                                qryUpdateTransferenciaSisparts, _
                                                qryUpdateCTe, _
                                                qryUpdateIdEmpresa, _
                                                qryUpdateCFOP_PSC_PES)

    '' REFINAMENTO DE DADOS
    executarComandos qryProcessos

End Sub


''#######################################################################################
''### EM DESENVOLVIMENTO
''#######################################################################################

Private Sub teste_FiltrarCompraItens()
Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
Dim QRY() As Variant: QRY = Array("chCTe")
Dim Item As Variant
Dim lists As Variant
Dim fieldnode As Variant
Dim childNode As Variant

'' cte
'XDoc.Load "C:\temp\Coleta\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210204884082000569570000039548351039548356-cteproc.xml"

'' nfe
'XDoc.Load "C:\temp\Coleta\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\29210220961864000187550010000001891138200000-nfeproc.xml"

For Each Item In QRY
    Set lists = XDoc.SelectNodes("//" & Item)
    For Each fieldnode In lists
        If (fieldnode.HasChildNodes) Then
            For Each childNode In fieldnode.ChildNodes
                Debug.Print fieldnode.text
            Next childNode
        End If
    Next fieldnode
Next Item

Set XDoc = Nothing

End Sub

''#######################################################################################
''### MÉTODOS
''#######################################################################################

Public Sub ProcessarArquivosXml(pTabelaDestino As String, pArquivos As Collection)
Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
Dim cadastro As clsProcessamento
Dim col As New Collection
Dim colCampos As New Collection
Dim strPk As String
Dim i As Variant
Dim fileName As Variant
Dim Item As Variant
Dim lists As Variant
Dim fieldnode As Variant
Dim childNode As Variant
    
'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()
Dim contadorDeArquivos As Long: contadorDeArquivos = 1

    statusFinal DT_PROCESSO, pTabelaDestino & " - Quantidade de arquivos: " & pArquivos.Count

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdInitMeter, "Processando arquivos...", pArquivos.Count
        
    '' LIMPAR TABELA DE PROCESSAMENTOS
    Application.CurrentDb.Execute qryDeleteProcessamento
        
        
    '' 01.Leitura e identificação do arquivo
    For Each fileName In pArquivos
        XDoc.Load fileName
        
        '' #BARRA_PROGRESSO
        SysCmd acSysCmdUpdateMeter, contadorDeArquivos
        
        '' 01.CRIAR CHAVE UNICA DE REGISTRO PARA CONTROLE DE DADOS
        strPk = Controle & getFileName(CStr(fileName))
        col.add strPk & "|" & "CaminhoDoArquivo" & "|" & fileName
        
        '' 02.CARREGAR CAMPOS DE ORIGEM X DESTINO DO REGISTRO
        For Each Item In carregarParametros(qryTags, pTabelaDestino)
            Set lists = XDoc.SelectNodes("//" & Item)
            For Each fieldnode In lists
                If (fieldnode.HasChildNodes) Then
                    For Each childNode In fieldnode.ChildNodes
                        col.add strPk & "|" & Item & "|" & fieldnode.text
                    Next childNode
                End If
            Next fieldnode
            
            DoEvents
            
        Next Item

        '' 03. REALIZAR CADASTRO DE TODOS OS ITENS COLETADOS NA TABELA DE PROCESSAMENTO
        If (col.Count > 2) Then
            
            '' CADASTRAR REGISTRO
            Set cadastro = New clsProcessamento
            For Each i In col
                With cadastro
                    .pk = Split(i, "|")(0)
                    .Chave = Split(i, "|")(1)
                    .valor = Mid(Split(i, "|")(2), 1, 255)
                    .cadastrar
                End With
                
                DoEvents
            Next i
            
            '' ATUALIZAR CAMPOS DE RELACIONAMENTOS
            Application.CurrentDb.Execute Replace(qryUpdateProcessamento, "strParametro", pTabelaDestino)
            
        End If
        
        '' LIMPAR COLEÇÃO
        ClearCollection col

        '' #BARRA_PROGRESSO
        contadorDeArquivos = contadorDeArquivos + 1
        
    Next fileName
                
                
'    '' ATUALIZAR CAMPOS DE RELACIONAMENTOS
'    Application.CurrentDb.Execute qryUpdateProcessamento
    

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdRemoveMeter

    Set XDoc = Nothing

    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "Processamento - ProcessarArquivosXml"
        
End Sub


'' #TRANSFERIR
Public Sub TransferirDadosProcessados(pDestino As String)

'' CONTROLE DE CADASTRO DE REGISTRO
Dim t As Integer

'' #BANCO_LOCAL
Dim db As dao.Database: Set db = CurrentDb
Dim tmpSql As String: tmpSql = "Select Distinct pk from tblProcessamento where NomeTabela = '" & pDestino & "' Order by pk;"
Dim rstPendentes As dao.Recordset: Set rstPendentes = db.OpenRecordset(tmpSql)
Dim rstOrigem As dao.Recordset

'' #BANCO_DESTINO
tmpSql = "Select * from " & pDestino
Dim rstDestino As dao.Recordset: Set rstDestino = db.OpenRecordset(tmpSql)

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

'' #BARRA_PROGRESSO
Dim contadorDeRegistros As Long: contadorDeRegistros = 1
SysCmd acSysCmdInitMeter, "Transferindo Dados...", rstPendentes.RecordCount

Do While Not rstPendentes.EOF

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdUpdateMeter, contadorDeRegistros

    '' listar itens de registro para cadastro
    Set rstOrigem = db.OpenRecordset("Select * from tblProcessamento where NomeTabela = '" & pDestino & "' and pk = '" & rstPendentes.Fields("pk").value & "' and len(formatacao)>0 Order by ID ")
    Do While Not rstOrigem.EOF
    
        '' CONTROLE DE CADASTRO
        If t = 0 Then rstDestino.AddNew: t = 1
        rstDestino.Fields(rstOrigem.Fields("NomeCampo").value).value = rstOrigem.Fields("Valor").value
        
        rstOrigem.MoveNext
        DoEvents
    Loop
    rstDestino.Update
    t = 0
    
    '' #COMPRAS
    If (pDestino = "tblCompraNF") Then Application.CurrentDb.Execute Replace(qryUpdateProcessamentoConcluido, "strChave", rstPendentes.Fields("pk").value)
    
    '' #DADOS_GERAIS
    '' qryUpdateRegistroValido - Valor padrao
    If (pDestino = "tblDadosConexaoNFeCTe") Then Application.CurrentDb.Execute "Update tblDadosConexaoNFeCTe SET registroValido = 0 where registroValido is null"
    
    '' #BARRA_PROGRESSO
    contadorDeRegistros = contadorDeRegistros + 1
    rstPendentes.MoveNext
    DoEvents
Loop

'' #BARRA_PROGRESSO
SysCmd acSysCmdRemoveMeter

'dbDestino.CloseConnection
db.Close: Set db = Nothing

'' #ANALISE_DE_PROCESSAMENTO
statusFinal DT_PROCESSO, "Processamento - TransferirDadosProcessados"

End Sub

'' ENVIAR DADOS PARA SERVIDOR
Public Sub EnviarDadosParaServidor(pDestino As String)

'' VARIAVEIS AUXILIARES
Dim fldDestino As Variant

'' BANCO LOCAL
Dim db As dao.Database: Set db = CurrentDb
Dim rstOrigem As dao.Recordset: Set rstOrigem = db.OpenRecordset("Select top 1 * from " & pDestino)

'' BANCO DESTINO
Dim strUsuarioNome As String: strUsuarioNome = pegarValorDoParametro(qryParametros, "BancoDados_Usuario")
Dim strUsuarioSenha As String: strUsuarioSenha = pegarValorDoParametro(qryParametros, "BancoDados_Senha")
Dim strOrigem As String: strOrigem = pegarValorDoParametro(qryParametros, "BancoDados_Origem")
Dim strBanco As String: strBanco = pegarValorDoParametro(qryParametros, "BancoDados_Banco")

Dim dbDestino As New Banco: dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
dbDestino.SqlSelect "SELECT * FROM " & pDestino


'' LISTAGEM DE CAMPOS DA TABELA ORIGEM/DESTINO
Dim strCampo As String
Dim rstCampos As dao.Recordset: Set rstCampos = db.OpenRecordset("Select * from tblOrigemDestino where tabela = '" & pDestino & "' and tagOrigem = 1")

Do While Not rstOrigem.EOF

    dbDestino.rs.AddNew
    
    Do While Not rstCampos.EOF
    
        strCampo = strSplit(rstCampos.Fields("destino").value, ".", 1)
    
        For Each fldDestino In dbDestino.rs.Fields
            If fldDestino.Name = strCampo Then
                ' dbDestino.rs(StrCampo).value = rstOrigem.Fields(StrCampo).value
                If rstOrigem.Fields(strCampo).value <> "" Then dbDestino.SqlExecute "Insert into tblCompraNF (" & strCampo & ") values (" & rstOrigem.Fields(strCampo).value & ")"
                Exit For
            End If
        Next fldDestino
        
        rstCampos.MoveNext
    Loop
    
    dbDestino.rs.Update
    
    rstOrigem.MoveNext
    
Loop

dbDestino.CloseConnection
db.Close: Set db = Nothing

End Sub


Public Sub CriarTipoDeArquivo(pArquivo As enumTipoArquivo)
Dim db As dao.Database: Set db = CurrentDb
Dim rst As dao.Recordset: Set rst = db.OpenRecordset(sqyDadosJson)
Dim dados As New clsConexaoNfeCteDados ''Set dados = New clsConexaoNfeCteDados

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()
Dim contadorDeRegistros As Long: contadorDeRegistros = 1
Dim nomeArquivo As String
    
    '' #BARRA_PROGRESSO
    SysCmd acSysCmdInitMeter, "Processando Json...", rst.RecordCount

    '' CRIAÇÃO DE ARQUIVOS
    Do While Not rst.EOF
    
        '' #BARRA_PROGRESSO
        SysCmd acSysCmdUpdateMeter, contadorDeRegistros
    
        '' #CARREGAR_DADOS_ARQUIVO
        dados.ChvAcesso = rst.Fields("ChvAcesso").value
        dados.dhEmi = rst.Fields("dhEmi").value
        dados.codUsuarioErp = pegarValorDoParametro(qryParametros, strUsuarioErpCod)
        dados.nomeUsuarioErp = pegarValorDoParametro(qryParametros, strUsuarioErpNome)
        dados.CaminhoDeProcessados = CreateDir(pegarValorDoParametro(qryParametros, strCaminhoDeProcessados))
                        
        Select Case pArquivo
            Case opFlagLancadaERP
                dados.Comando = pegarValorDoParametro(qryParametros, strComando)
                dados.criarERP
                nomeArquivo = "FlagLancadaERP"
            Case opManifesto
                dados.codTipoEvento = pegarValorDoParametro(qryParametros, strCodTipoEvento)
                dados.criarManifesto
                nomeArquivo = "Manifesto"
            Case Else
        End Select
        
        rst.MoveNext
        DoEvents
        
        '' #BARRA_PROGRESSO
        contadorDeRegistros = contadorDeRegistros + 1
    Loop

    db.Close: Set db = Nothing
    
    '' #BARRA_PROGRESSO
    SysCmd acSysCmdRemoveMeter
    
    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "Processamento - CriarTipoDeArquivo (" & nomeArquivo & ")"
    
End Sub


