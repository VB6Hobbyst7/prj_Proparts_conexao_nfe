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

'' -- CARREGAR TAGs DE VINDAS DO XML
'' #tblOrigemDestino.tabela
Private Const qryTags As String = "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((Len([Tag]))>0) AND ((tblOrigemDestino.tabela) = 'strParametro') AND ((tblOrigemDestino.TagOrigem)=1));"


'' -- SELEÇÃO DE ITENS UNICOS PARA PROCESSAMENTO
'' #tblProcessamento
Private Const sqyProcessamentosPendentes As String = "SELECT DISTINCT pk from tblProcessamento;"


'' -- LIMPAR TABELA
'' #tblProcessamento
Private Const qryDeleteProcessamento As String = "DELETE * FROM tblProcessamento;"


'' -- ATUALIZAR CAMPOS ( Nometabela e NomeCampo ) PARA USO DA TABELA DE PROCESSAMENTO
'' #tblProcessamento.NomeTabela = [qryOrigemDestino].[NomeTabela]
'' #tblProcessamento.NomeCampo = [qryOrigemDestino].[NomeCampo]
Private Const qryUpdateProcessamento As String = "UPDATE (SELECT tblOrigemDestino.Destino, tblOrigemDestino.Tag, strSplit([Destino],'.',0) AS NomeTabela, strSplit([Destino],'.',1) AS NomeCampo FROM tblOrigemDestino WHERE tblOrigemDestino.Tabela = 'strParametro' ) as qryOrigemDestino INNER JOIN tblProcessamento ON qryOrigemDestino.Tag = tblProcessamento.chave SET tblProcessamento.NomeTabela = [qryOrigemDestino].[NomeTabela], tblProcessamento.NomeCampo = [qryOrigemDestino].[NomeCampo];"


'' -- ARQUIVOS - CONSULTA PARA CRIAÇÃO DE ARQUIVOS JSON
'' #tblDadosConexaoNFeCTe
Private Const sqyDadosJson As String = "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi]))>0));"


'' ########################
'' #TratamentoDeRegistros
'' ########################

'' -- SELEÇÃO DE FORNECEDORES VALIDOS
'' #tblDadosConexaoNFeCTe.registroValido
Private Const qryUpdateFornecedoresValidos As String = "UPDATE (SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF FROM tmpClientes) AS qryFornecedoresValidos INNER JOIN tblDadosConexaoNFeCTe ON qryFornecedoresValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit SET tblDadosConexaoNFeCTe.registroValido = 1;"


'' -- SELEÇÃO DE REGISTROS VALIDOS
'' #tblDadosConexaoNFeCTe.registroValido
Private Const qryUpdateRegistrosValidos As String = "UPDATE (SELECT STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF FROM tmpEmpresa) AS qryRegistrosValidos INNER JOIN tblDadosConexaoNFeCTe ON qryRegistrosValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit SET tblDadosConexaoNFeCTe.registroValido = 1;"


'' -- EXCLUSAO DE REGISTROS INVALIDOS
'' #tblDadosConexaoNFeCTe.registroValido
Private Const qryDeleteRegistrosInvalidos As String = "DELETE * FROM tblDadosConexaoNFeCTe WHERE tblDadosConexaoNFeCTe.registroValido = 0;"


'' -- TIPOS DE CADASTRO - tblTipos
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryUpdateIdTipo As String = "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.ID_Tipo = 0 WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) Is Null));"


'' -- FiltroFil
'' #tblDadosConexaoNFeCTe.ID_EMPRESA
Private Const qryUpdateIdEmpresa As String = "UPDATE (SELECT tmpEmpresa.ID_Empresa, STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF FROM tmpEmpresa) AS qryEmpresas INNER JOIN tblDadosConexaoNFeCTe ON qryEmpresas.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit SET tblDadosConexaoNFeCTe.ID_Empresa = qryEmpresas.ID_Empresa;"


'' ########################################################################################################################
'' -- ID_TIPO - APENAS TIPOS COM ID DE VALOR ZERO(0) SERÃO ATUALIZADOS PARA NÃO COMPROMETER OS REGISTROS JÁ PROCESSADOS
'' ########################################################################################################################

'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 4 - NF-e Retorno Armazém
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryUpdateRetornoArmazem As String = "UPDATE (Select ValorDoParametro,TipoDeParametro from tblParametros where TipoDeParametro = 'RetornoArmazem') as tmpRetornoArmazem INNER JOIN ((select top 1 cInt('55') as strMod, '12680452000302' as strCNPJ_CPF, 'RetornoArmazem' as strTipoDeParametro from tblParametros)  AS qryRetornoArmazem INNER JOIN tblDadosConexaoNFeCTe ON (qryRetornoArmazem.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit) AND (qryRetornoArmazem.strMod = tblDadosConexaoNFeCTe.codMod)) ON (tmpRetornoArmazem.TipoDeParametro = qryRetornoArmazem.strTipoDeParametro) AND (tmpRetornoArmazem.TipoDeParametro = qryRetornoArmazem.strTipoDeParametro) SET tblDadosConexaoNFeCTe.ID_Tipo = [tmpRetornoArmazem].[ValorDoParametro] WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)=0));"


'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 6 - NF-e Transferência com código Sisparts
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryUpdateTransferenciaSisparts As String = "UPDATE (Select ValorDoParametro,TipoDeParametro from tblParametros where TipoDeParametro = 'TransferenciaSisparts') as tmpTransferenciaSisparts INNER JOIN ((SELECT CInt('55') AS strMod, 'TransferenciaSisparts' AS strTipoDeParametro, tmpEmpresa.CNPJ_Empresa AS strCNPJ_CPF FROM tmpEmpresa)  AS qryTransferenciaSisparts INNER JOIN tblDadosConexaoNFeCTe ON (qryTransferenciaSisparts.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit) AND (qryTransferenciaSisparts.strMod = tblDadosConexaoNFeCTe.codMod)) ON (tmpTransferenciaSisparts.TipoDeParametro = qryTransferenciaSisparts.strTipoDeParametro) AND (tmpTransferenciaSisparts.TipoDeParametro = qryTransferenciaSisparts.strTipoDeParametro) SET tblDadosConexaoNFeCTe.ID_Tipo = [tmpTransferenciaSisparts].[ValorDoParametro] WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)=0));"


'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 0 - CT-e
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryUpdateCTe As String = "UPDATE (Select ValorDoParametro,TipoDeParametro from tblParametros where TipoDeParametro = 'CTe') as tmpCTe INNER JOIN ((select top 1 cInt('57') as strMod, '68365501000377' as strCNPJ_CPF, 'CTe' as strTipoDeParametro from tblParametros)  AS qryCTe INNER JOIN tblDadosConexaoNFeCTe ON (qryCTe.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit) AND (qryCTe.strMod = tblDadosConexaoNFeCTe.codMod)) ON (tmpCTe.TipoDeParametro = qryCTe.strTipoDeParametro) AND (tmpCTe.TipoDeParametro = qryCTe.strTipoDeParametro) SET tblDadosConexaoNFeCTe.ID_Tipo = [tmpCTe].[ValorDoParametro] WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)=0));"


'' FORMATAÇÃO DE DATA - DD/MM/AAAA
'' #tblDadosConexaoNFeCTe.dhEmi
Private Const qryUpdateDhEmi As String = "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.dhEmi = CDate(Replace(Mid([tblDadosConexaoNFeCTe].[dhEmi],1,10),'-','/'));"


'' COMPRAS
'' -- AJUSTE DE CAMPOS
Private Const qryUpdateDTEmi As String = "UPDATE tblProcessamento SET tblProcessamento.valor = CDate(Replace(Mid([tblProcessamento].[valor],1,10),""-"",""/"")) WHERE (((tblProcessamento.NomeCampo)='DTEmi_CompraNF'));"
Private Const qryUpdateHoraEntd As String = "UPDATE tblProcessamento SET tblProcessamento.valor = Replace(Mid([tblProcessamento].[valor],12,8),'-','/') WHERE (((tblProcessamento.NomeCampo)='HoraEntd_CompraNF'));"
Private Const qryUpdateFilCompraNF As String = "UPDATE (SELECT tmpEmpresa.ID_Empresa, STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF, tmpEmpresa.CNPJ_Empresa FROM tmpEmpresa WHERE (((tmpEmpresa.CNPJ_Empresa) Is Not Null))) AS qryEmpresas INNER JOIN tblCompraNF ON qryEmpresas.strCNPJ_CPF = tblCompraNF.CNPJ_CPF_CompraNF SET tblCompraNF.Fil_CompraNF = qryEmpresas.ID_Empresa;"


Public Enum enumTipoArquivo
    opFlagLancadaERP = 0
    opManifesto = 1
End Enum

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
                                                qryUpdateDhEmi)

    '' REFINAMENTO DE DADOS
    executarComandos qryProcessos

End Sub





''#######################################################################################
''### EM DESENVOLVIMENTO
''#######################################################################################

Private Sub teste_FiltrarCompraItens()
Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
Dim qry() As Variant: qry = Array("chCTe")
Dim Item As Variant
Dim lists As Variant
Dim fieldnode As Variant
Dim childNode As Variant

'' cte
'XDoc.Load "C:\temp\Coleta\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210204884082000569570000039548351039548356-cteproc.xml"

'' nfe
'XDoc.Load "C:\temp\Coleta\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\29210220961864000187550010000001891138200000-nfeproc.xml"

For Each Item In qry
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

Debug.Print "Quantidade de arquivos: " & pArquivos.Count

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

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdRemoveMeter

    Set XDoc = Nothing

    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "Processamento - ProcessarArquivosXml"
        
End Sub


Public Sub CriarTipoDeArquivo(pArquivo As enumTipoArquivo)
Dim db As dao.Database: Set db = CurrentDb
Dim rst As dao.Recordset: Set rst = db.OpenRecordset(sqyDadosJson)
Dim dados As New clsConexaoNfeCteDados ''Set dados = New clsConexaoNfeCteDados

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()
Dim contadorDeRegistros As Long: contadorDeRegistros = 1
    
    '' #BARRA_PROGRESSO
    SysCmd acSysCmdInitMeter, "Processando Json...", rst.RecordCount

    '' CRIAÇÃO DE ARQUIVOS
    Do While Not rst.EOF
    
        '' #BARRA_PROGRESSO
        SysCmd acSysCmdUpdateMeter, contadorDeRegistros
    
        dados.ChvAcesso = rst.Fields("ChvAcesso").value
        dados.dhEmi = rst.Fields("dhEmi").value
        dados.codUsuarioErp = pegarValorDoParametro(qryParametros, strUsuarioErpCod)
        dados.nomeUsuarioErp = pegarValorDoParametro(qryParametros, strUsuarioErpNome)
        dados.CaminhoDeProcessados = CreateDir(pegarValorDoParametro(qryParametros, strCaminhoDeProcessados))
                        
        Select Case pArquivo
            Case opFlagLancadaERP
                dados.Comando = pegarValorDoParametro(qryParametros, strComando)
                dados.criarERP
            Case opManifesto
                dados.codTipoEvento = pegarValorDoParametro(qryParametros, strCodTipoEvento)
                dados.criarManifesto
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
    statusFinal DT_PROCESSO, "Processamento - CriarTipoDeArquivo"
    
End Sub



'Public Sub TransferirDadosConexaoNFeCTe()
'Dim dados As New clsConexaoNfeCteDados
'Dim db As dao.Database: Set db = CurrentDb
'Dim rst As dao.Recordset: Set rst = db.OpenRecordset(sqyProcessamentosPendentes)
'Dim qryProcessos() As Variant: qryProcessos = Array(qryUpdateFornecedoresValidos, qryUpdateRegistrosValidos, qryDeleteRegistrosInvalidos, qryUpdateRetornoArmazem, qryUpdateTransferenciaSisparts, qryUpdateCTe)
'
'    Do While Not rst.EOF
'        '' CADASTRO
'        dados.Comando = pegarValorDoParametro(qryParametros, strComando)
'        dados.codTipoEvento = pegarValorDoParametro(qryParametros, strCodTipoEvento)
'        dados.carregar_dados rst.Fields("pk").value
'        dados.cadastrar
'
'        rst.MoveNext
'
'    Loop
'
'    '' REFINAMENTO DE DADOS
'    executarComandos qryProcessos
'
'    db.Close: Set db = Nothing
'
'    MsgBox "Concluido!", vbOKOnly + vbInformation, "TransferirDadosConexaoNFeCTe"
'
'End Sub

'Public Sub TransferirCompras()
'Dim dados As New clsConexaoNfeCteCompraNF
'Dim db As DAO.Database: Set db = CurrentDb
'Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(sqyProcessamentosPendentes)
'Dim qryProcessos() As Variant: qryProcessos = Array(qryUpdateDTEmi, qryUpdateHoraEntd)
'
'    '' REFINAMENTO DE DADOS
'    executarComandos qryProcessos
'
'    Do While Not rst.EOF
'        dados.carregar_dados rst.Fields("pk").value
'        dados.cadastrar
'
'        Application.CurrentDb.Execute Replace(qryUpdateProcessamentoConcluido, "strChave", rst.Fields("pk").value)
'
'        rst.MoveNext
'    Loop
'
'    db.Close: Set db = Nothing
'
'    MsgBox "Concluido!", vbOKOnly + vbInformation, "TransferirCompras"
'
'End Sub

''#######################################################################################
''### TESTES UNITARIOS
''#######################################################################################



