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

'' #ORIGEM_DESTINO - CARREGAR TAGs DE VINDAS DO XML
Private Const qryTags As String = "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((Len([Tag]))>0) AND ((tblOrigemDestino.tabela) = 'strParametro') AND ((tblOrigemDestino.TagOrigem)=1));"

'' #PROCESSAMENTOS
Private Const sqyProcessamentosPendentes As String = "SELECT DISTINCT pk from tblProcessamento;"
'' -- ATUALIZAR CAMPOS ( Nometabela e NomeCampo ) PARA USO DA TABELA DE PROCESSAMENTO
Private Const qryUpdateProcessamento As String = "UPDATE (SELECT tblOrigemDestino.Destino, tblOrigemDestino.Tag, strSplit([Destino],'.',0) AS NomeTabela, strSplit([Destino],'.',1) AS NomeCampo FROM tblOrigemDestino WHERE tblOrigemDestino.Tabela = 'strParametro' ) as qryOrigemDestino INNER JOIN tblProcessamento ON qryOrigemDestino.Tag = tblProcessamento.chave SET tblProcessamento.NomeTabela = [qryOrigemDestino].[NomeTabela], tblProcessamento.NomeCampo = [qryOrigemDestino].[NomeCampo];"
'' -- LIMPAR TABELA
Private Const qryDeleteProcessamento As String = "DELETE * FROM tblProcessamento;"

'' #ARQUIVOS - CONSULTA PARA CRIAÇÃO DE ARQUIVOS JSON
Private Const sqyDadosJson As String = "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.Comando, tblDadosConexaoNFeCTe.dhEmi, tblDadosConexaoNFeCTe.codTipoEvento FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi]))>0));"
'Private Const sqyDadosJson As String = "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.Comando, tblDadosConexaoNFeCTe.dhEmi, tblDadosConexaoNFeCTe.codTipoEvento FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0));"

'' #CAPTURA_DADOS_GERAIS
'' -- CONTROLE DE FORNECEDORES
Private Const qryUpdateFornecedoresValidos As String = "UPDATE (SELECT STRPontos(Clientes.CNPJ_CPF) AS strCNPJ_CPF FROM Clientes) AS qryFornecedoresValidos INNER JOIN tblDadosConexaoNFeCTe ON qryFornecedoresValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit SET tblDadosConexaoNFeCTe.registroValido = 1;"
Private Const qryUpdateRegistrosValidos As String = "UPDATE (SELECT STRPontos(tblEmpresa.CNPJ_Empresa) AS strCNPJ_CPF FROM tblEmpresa) AS qryRegistrosValidos INNER JOIN tblDadosConexaoNFeCTe ON qryRegistrosValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit SET tblDadosConexaoNFeCTe.registroValido = 1;"
Private Const qryDeleteRegistrosInvalidos As String = "DELETE * FROM tblDadosConexaoNFeCTe WHERE tblDadosConexaoNFeCTe.registroValido = 0;"
'' -- ID_TIPO - APENAS TIPOS COM ID DE VALOR ZERO(0) SERÃO ATUALIZADOS PARA NÃO COMPROMETER OS REGISTROS JÁ PROCESSADOS
Private Const qryUpdateRetornoArmazem As String = "UPDATE (Select ValorDoParametro,TipoDeParametro from tblParametros where TipoDeParametro = 'RetornoArmazem') as tmpRetornoArmazem INNER JOIN ((select top 1 cInt('55') as strMod, '12680452000302' as strCNPJ_CPF, 'RetornoArmazem' as strTipoDeParametro from tblParametros)  AS qryRetornoArmazem INNER JOIN tblDadosConexaoNFeCTe ON (qryRetornoArmazem.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit) AND (qryRetornoArmazem.strMod = tblDadosConexaoNFeCTe.codMod)) ON (tmpRetornoArmazem.TipoDeParametro = qryRetornoArmazem.strTipoDeParametro) AND (tmpRetornoArmazem.TipoDeParametro = qryRetornoArmazem.strTipoDeParametro) SET tblDadosConexaoNFeCTe.ID_Tipo = [tmpRetornoArmazem].[ValorDoParametro] WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)=0));"
Private Const qryUpdateTransferenciaSisparts As String = "UPDATE (Select ValorDoParametro,TipoDeParametro from tblParametros where TipoDeParametro = 'TransferenciaSisparts') as tmpTransferenciaSisparts INNER JOIN ((SELECT CInt('55') AS strMod, 'TransferenciaSisparts' AS strTipoDeParametro, tblEmpresa.CNPJ_Empresa AS strCNPJ_CPF FROM tblEmpresa)  AS qryTransferenciaSisparts INNER JOIN tblDadosConexaoNFeCTe ON (qryTransferenciaSisparts.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit) AND (qryTransferenciaSisparts.strMod = tblDadosConexaoNFeCTe.codMod)) ON (tmpTransferenciaSisparts.TipoDeParametro = qryTransferenciaSisparts.strTipoDeParametro) AND (tmpTransferenciaSisparts.TipoDeParametro = qryTransferenciaSisparts.strTipoDeParametro) SET tblDadosConexaoNFeCTe.ID_Tipo = [tmpTransferenciaSisparts].[ValorDoParametro] WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)=0));"
Private Const qryUpdateCTe As String = "UPDATE (Select ValorDoParametro,TipoDeParametro from tblParametros where TipoDeParametro = 'CTe') as tmpCTe INNER JOIN ((select top 1 cInt('57') as strMod, '68365501000377' as strCNPJ_CPF, 'CTe' as strTipoDeParametro from tblParametros)  AS qryCTe INNER JOIN tblDadosConexaoNFeCTe ON (qryCTe.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit) AND (qryCTe.strMod = tblDadosConexaoNFeCTe.codMod)) ON (tmpCTe.TipoDeParametro = qryCTe.strTipoDeParametro) AND (tmpCTe.TipoDeParametro = qryCTe.strTipoDeParametro) SET tblDadosConexaoNFeCTe.ID_Tipo = [tmpCTe].[ValorDoParametro] WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)=0));"

'' COMPRAS
'' -- AJUSTE DE CAMPOS
Private Const qryUpdateDTEmi As String = "UPDATE tblProcessamento SET tblProcessamento.valor = CDate(Replace(Mid([tblProcessamento].[valor],1,10),""-"",""/"")) WHERE (((tblProcessamento.NomeCampo)='DTEmi_CompraNF'));"
Private Const qryUpdateHoraEntd As String = "UPDATE tblProcessamento SET tblProcessamento.valor = Replace(Mid([tblProcessamento].[valor],12,8),'-','/') WHERE (((tblProcessamento.NomeCampo)='HoraEntd_CompraNF'));"



Public Enum enumTipoArquivo
    opFlagLancadaERP = 0
    opManifesto = 1
End Enum



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
    Dim contadorDeArquivos As Long: contadorDeArquivos = 1
    Dim DT_PROCESSO As Date
    Dim DT_REGISTRO As Date
    
    DT_PROCESSO = Now()
    
    '' LIMPAR TABELA DE PROCESSAMENTOS
    Application.CurrentDb.Execute qryDeleteProcessamento
        
    '' #BARRA_PROGRESSO
    SysCmd acSysCmdInitMeter, "Reading Data...", pArquivos.Count
        
        
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
            
            DT_REGISTRO = Now()
            
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
            
            statusFinal DT_REGISTRO, "Processamento por registro"
            
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

    '' ANALISE DE PROCESSAMENTO
    statusFinal DT_PROCESSO, "Processamento - ProcessarArquivosXml"

    Set XDoc = Nothing
    
    MsgBox "Concluido!", vbOKOnly + vbInformation, "ProcessarArquivosXml"
    
End Sub

Public Sub TransferirDadosConexaoNFeCTe()
Dim dados As New clsConexaoNfeCteDados
Dim db As DAO.Database: Set db = CurrentDb
Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(sqyProcessamentosPendentes)
Dim qryProcessos() As Variant: qryProcessos = Array(qryUpdateFornecedoresValidos, qryUpdateRegistrosValidos, qryDeleteRegistrosInvalidos, qryUpdateRetornoArmazem, qryUpdateTransferenciaSisparts, qryUpdateCTe)

    Do While Not rst.EOF
        '' CADASTRO
        dados.Comando = pegarValorDoParametro(qryParametros, strComando)
        dados.codTipoEvento = pegarValorDoParametro(qryParametros, strCodTipoEvento)
        dados.carregar_dados rst.Fields("pk").value
        dados.cadastrar

        rst.MoveNext
        
    Loop

    '' REFINAMENTO DE DADOS
    executarComandos qryProcessos
    
    db.Close: Set db = Nothing
    
    MsgBox "Concluido!", vbOKOnly + vbInformation, "TransferirDadosConexaoNFeCTe"

End Sub

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


Public Sub CriarTipoDeArquivo(pArquivo As enumTipoArquivo)
Dim db As DAO.Database: Set db = CurrentDb
Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(sqyDadosJson)

Dim dados As New clsConexaoNfeCteDados ''Set dados = New clsConexaoNfeCteDados

    '' CRIAÇÃO DE ARQUIVOS
    Do While Not rst.EOF
        dados.ChvAcesso = rst.Fields("ChvAcesso").value
        dados.dhEmi = rst.Fields("dhEmi").value
        dados.codUsuarioErp = pegarValorDoParametro(qryParametros, strUsuarioErpCod)
        dados.nomeUsuarioErp = pegarValorDoParametro(qryParametros, strUsuarioErpNome)
        dados.CaminhoDeProcessados = CreateDir(pegarValorDoParametro(qryParametros, strCaminhoDeProcessados))
                        
        Select Case pArquivo
            Case opFlagLancadaERP
                dados.Comando = rst.Fields("Comando").value
                dados.criarERP
            Case opManifesto
                dados.codTipoEvento = rst.Fields("codTipoEvento").value
                dados.criarManifesto
            Case Else
        End Select
        
        rst.MoveNext
    Loop

    db.Close
    
    Set db = Nothing
    
'    MsgBox "Concluido!", vbOKOnly + vbInformation, "CriarArquivoFlagLancadaERP"

End Sub


''#######################################################################################
''### TESTES UNITARIOS
''#######################################################################################



