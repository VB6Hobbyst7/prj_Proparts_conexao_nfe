VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConexaoNfeCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'' #####################################
'' 02 PROCESSAMENTO DE DADOS
'' #####################################

''------------------------------------
'' ### REVISÃO
''
'' compras_carregarItensCTe
'' compras_atualizarCampos
''
''
'' ### PRINCIPAIS MÉTODOS
''
'' 01. carregar_DadosGerais
'' 02. ProcessamentoDeArquivo
'' 03. enviar_ComprasParaServidor
'' 00. criarArquivoJson
'' 00. ProcessarArquivosXml
'' 00. ADM_carregarDadosDoServidor
''

'' #00. CRIAÇÃO DE ARQUIVOS JSON
'' #00. CADASTRO DE DADOS PROCESSADOS
'' #01. PROCESSAMENTO DE ARQUIVOS POR TIPO
'' #02. TRANSFERIR DADOS PROCESSAMENTOS PARA A TABELA DE DESTINO
''------------------------------------



Option Compare Database
Option Explicit

Private FileCollention As New Collection
Private con As ADODB.Connection



'' ########################################
'' #tblOrigemDestino
'' ########################################

'' -- CARREGAR TAGs DE VINDAS DO XML
'' #tblOrigemDestino.tabela
Private Const qryTags As String = _
        "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((Len([Tag]))>0) AND ((tblOrigemDestino.tabela) = 'strParametro') AND ((tblOrigemDestino.TagOrigem)=1));"


'' ########################################
'' #tblDadosConexaoNFeCTe - #DADOS_GERAIS
'' ########################################

'' ARQUIVOS VALIDOS
Private Const sqlArquivosValidos As String = _
        "SELECT ID AS ChvAcesso " & _
        "   ,CaminhoDoArquivo AS strOrigem   " & _
        "   ,Replace(DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='caminhoDeColetaProcessados'""), ""empresa"", strSplit(getPath([tblDadosConexaoNFeCTe].[CaminhoDoArquivo]), ""\"", 2)) & getFileNameAndExt([tblDadosConexaoNFeCTe].[CaminhoDoArquivo]) AS strDestino " & _
        "FROM tblDadosConexaoNFeCTe WHERE ((NOT (tblDadosConexaoNFeCTe.CaminhoDoArquivo) IS NULL) AND ((tblDadosConexaoNFeCTe.registroValido) = 1)  AND ((tblDadosConexaoNFeCTe.registroProcessado) = 2));"

'' ATUALIZAR REGISTRO PARA PROCESSADO
'' #tblDadosConexaoNFeCTe.registroProcessado
Private Const qryUpdateProcessado As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 3 " & _
        ",tblDadosConexaoNFeCTe.CaminhoDoArquivo = Replace(DLookUp(""[ValorDoParametro]"",""[tblParametros]"",""[TipoDeParametro]='caminhoDeColetaProcessados'""),""empresa"",strSplit(getPath([tblDadosConexaoNFeCTe].[CaminhoDoArquivo]),""\"",2)) & getFileNameAndExt([tblDadosConexaoNFeCTe].[CaminhoDoArquivo])" & _
        " WHERE tblDadosConexaoNFeCTe.ID = strChave;"

'' ARQUIVOS INVALIDOS
Private Const sqlArquivosInvalidosNaoProcessado As String = _
    "SELECT ID AS ChvAcesso " & _
    "   ,CaminhoDoArquivo AS strOrigem   " & _
    "   ,Replace(DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='caminhoDeColetaExpurgo'""), ""empresa"", strSplit(getPath([tblDadosConexaoNFeCTe].[CaminhoDoArquivo]), ""\"", 2)) & getFileNameAndExt([tblDadosConexaoNFeCTe].[CaminhoDoArquivo]) AS strDestino " & _
    "FROM  tblDadosConexaoNFeCTe " & _
    " WHERE ((NOT (tblDadosConexaoNFeCTe.CaminhoDoArquivo) IS NULL) AND ((tblDadosConexaoNFeCTe.registroValido) = 0) AND ((tblDadosConexaoNFeCTe.registroProcessado) = 0));"

'' ATUALIZAR REGISTRO PARA EXPURGO
'' #tblDadosConexaoNFeCTe.registroProcessado
Private Const qryUpdateExpurgo As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 4  " & _
        ",tblDadosConexaoNFeCTe.CaminhoDoArquivo = Replace(DLookUp(""[ValorDoParametro]"",""[tblParametros]"",""[TipoDeParametro]='caminhoDeColetaExpurgo'""),""empresa"",strSplit(getPath([tblDadosConexaoNFeCTe].[CaminhoDoArquivo]),""\"",2)) & getFileNameAndExt([tblDadosConexaoNFeCTe].[CaminhoDoArquivo])" & _
        " WHERE tblDadosConexaoNFeCTe.ID = strChave ;"

'' tblDadosConexaoNFeCTe.Pendentes
Private Const qrySelecaoDeArquivosPendentes As String = _
    "SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0)) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0) ORDER BY tblDadosConexaoNFeCTe.CaminhoDoArquivo;"

'' SELEÇÃO DE EMPRESAS
Private Const qrySelecaoDeColetaEmpresa As String = _
    "SELECT tblParametros.ValorDoParametro FROM tblParametros WHERE tblParametros.TipoDeParametro = ""caminhoDeColetaEmpresa"";"

'' -- PROCESSAMENTO CONCLUIDO
'' #tblDadosConexaoNFeCTe.registroProcessado = 1
Private Const qryUpdateProcessamentoConcluido As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 1 WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0) AND ((tblDadosConexaoNFeCTe.Chave)='strChave'));"


'' #tblDadosConexaoNFeCTe.registroProcessado = 1
Private Const qryUpdateProcessamentoConcluido_CTE As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 1 WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1) AND ((tblDadosConexaoNFeCTe.ID_Tipo) = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro] = 'Cte'"")));"


'' #tblDadosConexaoNFeCTe.registroProcessado = 1
Private Const qryUpdateProcessamentoConcluido_ItensCompras As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 1 FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0));"


'' -- ARQUIVOS - CONSULTA PARA CRIAÇÃO DE ARQUIVOS JSON
'' #tblDadosConexaoNFeCTe.Distinct
Private Const sqyDadosJson As String = _
        "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi]))>0));"


'' -- SELEÇÃO DE FORNECEDORES VALIDOS
'' #tblDadosConexaoNFeCTe.qryUpdateFornecedoresValidos
'' #tblDadosConexaoNFeCTe.registroValido
Private Const qryUpdateFornecedoresValidos As String = _
        "UPDATE (SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF FROM tmpClientes) AS qryFornecedoresValidos INNER JOIN tblDadosConexaoNFeCTe ON qryFornecedoresValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit SET tblDadosConexaoNFeCTe.registroValido = 1;"


'' -- SELEÇÃO DE REGISTROS VALIDOS.qryRegistrosValidos
'' #tblDadosConexaoNFeCTe.qryUpdateRegistrosValidos
Private Const qryUpdateRegistrosValidos As String = _
        "UPDATE (SELECT STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF FROM tmpEmpresa) AS qryRegistrosValidos INNER JOIN tblDadosConexaoNFeCTe ON qryRegistrosValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit SET tblDadosConexaoNFeCTe.registroValido = 1;"


'' -- EXCLUSAO DE REGISTROS INVALIDOS -- DESCONTINUADO - ( PODEMOS USAR DEPOIS )
'' #tblDadosConexaoNFeCTe.qryDeleteRegistrosInvalidos
'Private Const qryDeleteRegistrosInvalidos As String = _
'        "DELETE * FROM tblDadosConexaoNFeCTe WHERE tblDadosConexaoNFeCTe.registroValido = 0;"


'' -- TIPOS DE CADASTRO - tblTipos
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryUpdateIdTipo As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.ID_Tipo = 0 WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) Is Null));"


'' -- FiltroFil
'' #tblDadosConexaoNFeCTe.ID_Empresa
'' -- CTE
Private Const qryUpdateIdEmpresa As String = _
        "UPDATE (SELECT tmpEmpresa.ID_Empresa, STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF FROM tmpEmpresa) AS qryEmpresas " & _
        " INNER JOIN tblDadosConexaoNFeCTe ON qryEmpresas.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_Rem " & _
        " SET tblDadosConexaoNFeCTe.ID_Empresa = qryEmpresas.ID_Empresa;"


'' #tblDadosConexaoNFeCTe.ID_Empresa
Private Const qryUpdateIdEmpresa_TransferenciaEntreFiliais As String = _
        "UPDATE (SELECT tmpEmpresa.ID_Empresa, STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF FROM tmpEmpresa) AS qryEmpresas " & _
        " INNER JOIN tblDadosConexaoNFeCTe ON qryEmpresas.strCNPJ_CPF = tblDadosConexaoNFeCTe.CPNJ_Dest " & _
        " SET tblDadosConexaoNFeCTe.ID_Empresa = qryEmpresas.ID_Empresa" & _
        " WHERE (((tblDadosConexaoNFeCTe.CFOP) = '6152'));"


'' #20210823_qryUpdateCFOP_PSC_PES -- FiltroCFOP
'' #tblDadosConexaoNFeCTe.FiltroCFOP
'' #tmpDadosConexaoNFeCTe.ID_NatOp_CompraNF
'' RELACIONAMENTO DE CFOP PARA DENTRO E FORA DO ESTADO
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
            "       AND tblDadosConexaoNFeCTe.ID_NatOp_CompraNF IS NULL )  AS tmpDadosConexaoNFeCTe " & _
            "ON (tmpPscPes.strXMLCFOP = tmpDadosConexaoNFeCTe.CFOP) AND (tmpPscPes.Fil_NatOper = tmpDadosConexaoNFeCTe.ID_Empresa) " & _
            "SET  tmpDadosConexaoNFeCTe.ID_NatOp_CompraNF = [tmpPscPes].[ID_NatOper], tmpDadosConexaoNFeCTe.FiltroCFOP = [tmpPscPes].[CFOP_NatOper];"


'' ------------------------------------------------------------------------------------------------------------------------
'' -- ID_TIPO - APENAS TIPOS COM ID DE VALOR ZERO(0) SERÃO ATUALIZADOS PARA NÃO COMPROMETER OS REGISTROS JÁ PROCESSADOS
'' ------------------------------------------------------------------------------------------------------------------------

'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 4 - NF-e Retorno Armazém
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryUpdateRetornoArmazem As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='RetornoArmazem'"") WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = 0) AND ((tblDadosConexaoNFeCTe.codMod) = CInt('55')) AND ((tblDadosConexaoNFeCTe.CNPJ_emit) IN ('12680452000302')));"

'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 6 - NF-e Transferência com código Sisparts
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryUpdateTransferenciaSisparts As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='TransferenciaSisparts'"") WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = 0) AND ((tblDadosConexaoNFeCTe.codMod) = CInt('55')) AND ((tblDadosConexaoNFeCTe.CNPJ_emit) IN (SELECT CNPJ_Empresa FROM [tmpEmpresa])));"

'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 0 - CT-e
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryUpdateCTe As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='CTe'"") WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = 0) AND ((tblDadosConexaoNFeCTe.codMod) = CInt('57')));"


'' ########################################
'' #tblCompraNF - COMPRAS
'' ########################################
 
'' #20210823_qryUpdate_tblCompraNF
'' #20210823_VTotProd_CompraNF
Private Const qryUpdate_tblCompraNF As String = _
        "UPDATE tblDadosConexaoNFeCTe " & _
        "INNER JOIN tblCompraNF ON tblDadosConexaoNFeCTe.ChvAcesso = tblCompraNF.ChvAcesso_CompraNF " & _
        "SET tblCompraNF.CFOP_CompraNF = [tblDadosConexaoNFeCTe].[FiltroCFOP] " & _
        "   ,tblCompraNF.ID_NatOp_CompraNF = [tblDadosConexaoNFeCTe].[ID_NatOp_CompraNF] " & _
        "   ,tblCompraNF.ModeloDoc_CompraNF = [tblDadosConexaoNFeCTe].[codMod] " & _
        "   ,tblCompraNF.Fil_CompraNF = [tblDadosConexaoNFeCTe].[ID_Empresa] " & _
        "   ,tblCompraNF.VTotNF_CompraNF = replace(Nz([tblCompraNF].[VTotNF_CompraNF], 0) / 100, "","", ""."") " & _
        "   ,tblCompraNF.VTotProd_CompraNF = replace(Nz([tblCompraNF].[VTotNF_CompraNF], 0) / 100, "","", ""."") " & _
        "   ,tblCompraNF.BaseCalcICMS_CompraNF = replace(Nz([tblCompraNF].[BaseCalcICMS_CompraNF], 0) / 100, "","", ""."") " & _
        "   ,tblCompraNF.TPNF_CompraNF = 1 " & _
        "   ,tblCompraNF.VTotICMS_CompraNF = 0 " & _
        "   ,tblCompraNF.VTotIPI_CompraNF = 0 " & _
        "WHERE (((tblDadosConexaoNFeCTe.registroValido) = 1) AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1)) AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0); "


'' #20210823_qryUpdate_IDVD
'' #tblCompraNF.IDVD_CompraNF
Private Const qryUpdate_IDVD As String = _
    "UPDATE tblDadosConexaoNFeCTe " & _
    "SET tblDadosConexaoNFeCTe.IDVD_CompraNF = ((Left(Trim(Replace(Replace(tblDadosConexaoNFeCTe.Obs_CompraNF, 'PEDIDO: ', ''), 'PEDIDO ', '')), 6))) " & _
    "WHERE (((Left(tblDadosConexaoNFeCTe.Obs_CompraNF, 6)) = 'PEDIDO ') " & _
    "       AND ((tblDadosConexaoNFeCTe.CNPJ_emit) = '12680452000302') " & _
    "       AND (((tblDadosConexaoNFeCTe.registroValido) = 1)  " & _
    "       AND ((tblDadosConexaoNFeCTe.registroProcessado) = 0))  " & _
    "       AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0));"


'' #20210823_qryUpdateIdFornCompraNF
'' #tblCompraNF.ID_Forn_CompraNF
Private Const qryUpdateIdFornCompraNF As String = _
        "UPDATE (SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF, tmpClientes.CÓDIGOClientes FROM tmpClientes) AS qryClientesFornecedor " & _
        " INNER JOIN tblCompraNF ON tblCompraNF.CNPJ_CPF_CompraNF = qryClientesFornecedor.strCNPJ_CPF " & _
        " SET tblCompraNF.ID_Forn_CompraNF = qryClientesFornecedor.CÓDIGOClientes;"

'' #20210823_qryUpdate_NFe
'' # - Formação de campos com valores padrão com base em código anterior
Private Const qryUpdate_NFe As String = "UPDATE tblCompraNF  " & _
        "INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso " & _
        "SET tblCompraNF.ModeloDoc_CompraNF = 55 " & _
        "   ,tblCompraNF.HoraEntd_CompraNF = NULL " & _
        "   ,tblCompraNF.TPNF_CompraNF = 1 " & _
        "   ,tblCompraNF.Sit_CompraNF = tblDadosConexaoNFeCTe.Sit_CompraNF " & _
        "   ,tblCompraNF.VTotICMS_CompraNF = 0 " & _
        "   ,tblCompraNF.VTotIPI_CompraNF = 0 " & _
        "WHERE (((tblDadosConexaoNFeCTe.codMod) = 55) AND ((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1)); "

'' #20210823_qryUpdate_CTe
'' # - Formação de campos com valores padrão com base em código anterior
Private Const qryUpdate_CTe As String = "UPDATE tblCompraNF   " & _
        "INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso  " & _
        "SET  " & _
        "   tblCompraNF.ModeloDoc_CompraNF = 57  " & _
        "   ,tblCompraNF.DTEntd_CompraNF = DATE () " & _
        "   ,tblCompraNF.HoraEntd_CompraNF = NULL " & _
        "   ,tblCompraNF.TPNF_CompraNF = 1 " & _
        "   ,tblCompraNF.Sit_CompraNF = tblDadosConexaoNFeCTe.Sit_CompraNF " & _
        "   ,tblCompraNF.IDVD_CompraNF = NULL " & _
        "   ,tblCompraNF.VTotIPI_CompraNF = 0 " & _
        "   WHERE (((tblDadosConexaoNFeCTe.codMod) = 57) AND ((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1));"

'' #20210823_qryUpdate_CTeItens
'' #20210823_ID_Prod_CompraNFItem
'' # - Formação de campos com valores padrão com base em código anterior
Private Const qryUpdate_CTeItens As String = "UPDATE tblCompraNFItem INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNFItem.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso " & _
        "SET tblCompraNFItem.Item_CompraNFItem = 1,tblCompraNFItem.ID_Grade_CompraNFItem = 1,tblCompraNFItem.QtdFat_CompraNFItem = 1 " & _
        "   ,tblCompraNFItem.IPI_CompraNFItem = 0,tblCompraNFItem.FlagEst_CompraNFItem = 0,tblCompraNFItem.BaseCalcICMS_CompraNFItem = 100 " & _
        "   ,tblCompraNFItem.IseICMS_CompraNFItem = 0,tblCompraNFItem.BaseCalcIPI_CompraNFItem = 0,tblCompraNFItem.DebIPI_CompraNFItem = 0 " & _
        "   ,tblCompraNFItem.IseIPI_CompraNFItem = 0,tblCompraNFItem.TxMLSubsTrib_CompraNFItem = 0,tblCompraNFItem.BaseCalcICMSSubsTrib_CompraNFItem = 0 " & _
        "   ,tblCompraNFItem.VTotICMSSubsTrib_compranfitem = 0,tblCompraNFItem.VTotDesc_CompraNFItem = 0,tblCompraNFItem.VTotFrete_CompraNFItem = 0 " & _
        "   ,tblCompraNFItem.VTotPIS_CompraNFItem = 0,tblCompraNFItem.VTotBaseCalcPIS_CompraNFItem = 0,tblCompraNFItem.VTotCOFINS_CompraNFItem = 0 " & _
        "   ,tblCompraNFItem.VTotIseICMS_CompraNFItem = 0,tblCompraNFItem.VTotBaseCalcCOFINS_CompraNFItem = 0,tblCompraNFItem.VTotSNCredICMS_CompraNFItem = 0 " & _
        "   ,tblCompraNFItem.VTotSeg_CompraNFItem = 0,tblCompraNFItem.VTotOutDesp_CompraNFItem = 0 " & _
        "   ,tblCompraNFItem.VUnt_CompraNFItem = DLookUp(""[VTotNF_CompraNF]"", ""[tblCompraNF]"", ""[ChvAcesso_CompraNF] = '"" & [tblCompraNFItem].[ChvAcesso_CompraNF] & ""'"") " & _
        "   ,tblCompraNFItem.VTot_CompraNFItem = DLookUp(""[VTotNF_CompraNF]"", ""[tblCompraNF]"", ""[ChvAcesso_CompraNF] = '"" & [tblCompraNFItem].[ChvAcesso_CompraNF] & ""'"") " & _
        "   ,tblCompraNFItem.ID_Prod_CompraNFItem = DLookUp(""[Código]"", ""[tmpProdutos]"", ""[Modelo] = 'TRANSPORTE'"") " & _
        "WHERE (((tblDadosConexaoNFeCTe.codMod) = 57) AND ((tblDadosConexaoNFeCTe.registroValido) = 1) AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1));"

'' ########################################
'' #tblCompraNFItem - COMPRAS ITENS
'' ########################################

'' #20210823_qryUpdateItens_CFOP_CompraNFItem
'' #tblCompraNFItem.CFOP_CompraNFItem
Private Const qryUpdateItens_CFOP_CompraNFItem As String = _
        "UPDATE (tblDadosConexaoNFeCTe INNER JOIN tblCompraNF ON tblDadosConexaoNFeCTe.ChvAcesso = tblCompraNF.ChvAcesso_CompraNF) " & _
        "INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF " & _
        "SET tblCompraNFItem.CFOP_CompraNFItem = [tblCompraNF].[CFOP_CompraNF] " & _
        "WHERE (((tblDadosConexaoNFeCTe.registroValido) = 1) " & _
        "       AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1) " & _
        "       AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0));"

''' UPDATE - ID_Prod_CompraNFItem
'' #tblCompraNFItem.ID_Prod_CompraNFItem
Private Const qryUpdateItens_ID_Prod_CompraNFItem As String = _
        "UPDATE tblDadosConexaoNFeCTe " & _
        "INNER JOIN (tblCompraNF INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF) ON tblDadosConexaoNFeCTe.ChvAcesso = tblCompraNF.ChvAcesso_CompraNF " & _
        "SET tblCompraNFItem.ID_Prod_CompraNFItem = DLookUp(""CodigoProd_Grade"", ""tmpGradeProdutos"", ""CodigoForn_Grade='"" & [tblCompraNFItem].[ID_Prod_CompraNFItem] & ""'"") " & _
        "WHERE (((tblDadosConexaoNFeCTe.registroValido) = 1)    AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1));"

'' #tblCompraNFItem.InsertItensCTe
Private Const qryInsertItensCTe As String = _
        "INSERT INTO tblCompraNFItem ( ChvAcesso_CompraNF   ,VUnt_CompraNFItem  ,Num_CompraNFItem " & _
        "   ,VTot_CompraNFItem  ,DebICMS_CompraNFItem   ,VTotBaseCalcICMS_CompraNFItem " & _
        "   ,ID_NatOp_CompraNFItem  ,Item_CompraNFItem  ,ID_Grade_CompraNFItem " & _
        "   ,QtdFat_CompraNFItem    ,IPI_CompraNFItem   ,FlagEst_CompraNFItem   ,BaseCalcICMS_CompraNFItem  ) " & _
        "SELECT tblCompraNF.ChvAcesso_CompraNF  ,tblCompraNF.VTotNF_CompraNF AS VUnt_CompraNFItem   ,tblCompraNF.NumNF_CompraNF " & _
        "   ,tblCompraNF.VTotNF_CompraNF ,IIf([VTotICMS_CompraNF] <> """", Replace(Nz([tblCompraNF].[BaseCalcICMS_CompraNF], 0) / 100, "","", "".""), 0) AS strVTotICMS ,tblCompraNF.BaseCalcICMS_CompraNF " & _
        "   ,tblCompraNF.ID_NatOp_CompraNF  ,1 AS strItem   ,1 AS strIDGrade " & _
        "   ,1 AS strQtdFat ,0 AS strIPI    ,0 AS strFlag   ,100 AS strBaseCalcICMS " & _
        "FROM tblCompraNF " & _
        "INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso " & _
        "WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='Cte'""))); "

'' #20210823_qryUpdateNumPed_CompraNF
'' #tblCompraNF.NumPed_CompraNF
Private Const qryUpdateNumPed_CompraNF As String = "UPDATE TblCompraNF SET TblCompraNF.NumPed_CompraNF = Format(IIf(IsNull(DLookup(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='NumPed_CompraNF'"")), '000001', DLookup(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='NumPed_CompraNF'"") + 1), '000000') " & _
        "WHERE (((TblCompraNF.ID_CompraNF) IN (SELECT TOP 1 ID_CompraNF FROM TblCompraNF WHERE NumPed_CompraNF IS NULL ORDER BY ID_CompraNF)));"

'' #20210823_qryUpdateNumPed_CompraNF
'' #tblParametros.ValorDoParametro
Private Const qryUpdateNumPed_Contador As String = "UPDATE tblParametros SET tblParametros.ValorDoParametro = Format(IIf(IsNull(DLookup(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='NumPed_CompraNF'"")), '000001', DLookup(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='NumPed_CompraNF'"") + 1), '000000') WHERE tblParametros.TipoDeParametro = ""NumPed_CompraNF"""

'' #20210823_qryUpdateSit_CompraNF
'' #tblDadosConexaoNFeCTe.Sit_CompraNF
'' CONTROLE DE FINALIDADES
Private Const qryUpdateSit_CompraNF As String = "UPDATE  ( SELECT   " & _
        "                 strSplit(ValorDoParametro,'|',0) AS strFinalidade,  strSplit(ValorDoParametro,'|',1) AS strSit_CompraNF   " & _
        "                FROM   " & _
        "                  tblParametros   " & _
        "                WHERE   " & _
        "                  TipoDeParametro='Sit_CompraNF'  " & _
        "              )  AS qrySit_CompraNF   " & _
        "INNER JOIN   " & _
        "   (   SELECT  *   " & _
        "       FROM  tblDadosConexaoNFeCTe   " & _
        "       WHERE tblDadosConexaoNFeCTe.registroValido IN (SELECT TOP 1 cint(tblParametros.ValorDoParametro) FROM [tblParametros] WHERE TipoDeParametro = 'registroValido')   " & _
        "   )  AS tmpDadosConexaoNFeCTe  " & _
        "ON (cint(qrySit_CompraNF.strFinalidade) = tmpDadosConexaoNFeCTe.ID_TIPO)  " & _
        "SET  tmpDadosConexaoNFeCTe.Sit_CompraNF = [qrySit_CompraNF].[strSit_CompraNF];  "


Private Const qrySelecRepositorioDeRelacao As String = _
            "Select * from tmpCompras_ID_CompraNF"

Private Const qrySelecRepositorioDeCompras As String = _
            "SELECT * FROM tblCompraNF ORDER BY ID_CompraNF"
            
''#######################################################################################
''### CONSULTAS DE CONTROLE DE DADOS
''#######################################################################################

'' #20210823_qryUpdateNumPed_CompraNF
Public Function UpdateNumPed_CompraNF()
Dim x As Long
Dim contador As Long: contador = DCount("*", "TblCompraNF", "NumPed_CompraNF is null")
Dim qryProcessos() As Variant: qryProcessos = Array(qryUpdateNumPed_CompraNF, qryUpdateNumPed_Contador)

    For x = 1 To contador
        executarComandos qryProcessos
    Next

End Function

Public Function UpdateProcessamentoConcluido() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='debug'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryUpdateProcessamentoConcluido
    UpdateProcessamentoConcluido = qryUpdateProcessamentoConcluido
                                                    
End Function

Public Function UpdateExpurgo() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='debug'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryUpdateExpurgo
    UpdateExpurgo = qryUpdateExpurgo
                                                    
End Function

Public Function UpdateProcessado() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='debug'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryUpdateProcessado
    UpdateProcessado = qryUpdateProcessado
                                                    
End Function

'' qrySelecRepositorioDeCompras

Public Function SelecRepositorioDeCompras() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='debug'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qrySelecRepositorioDeCompras
    SelecRepositorioDeCompras = qrySelecRepositorioDeCompras
                                                    
End Function


Public Function SelecRepositorioDeRelacao() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='debug'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qrySelecRepositorioDeRelacao
    SelecRepositorioDeRelacao = qrySelecRepositorioDeRelacao
                                                    
End Function

Public Function SelectArquivosInvalidos() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='debug'") Then TextFile_Append CurrentProject.path & "\" & strLog(), sqlArquivosInvalidosNaoProcessado
    SelectArquivosInvalidos = sqlArquivosInvalidosNaoProcessado
                                                    
End Function

Public Function SelectArquivosValidos() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='debug'") Then TextFile_Append CurrentProject.path & "\" & strLog(), sqlArquivosValidos
    SelectArquivosValidos = sqlArquivosValidos
                                                    
End Function

Public Function SelectColetaEmpresa() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='debug'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qrySelecaoDeColetaEmpresa
    SelectColetaEmpresa = qrySelecaoDeColetaEmpresa
                                                    
End Function

Public Function SelectArquivosPendentes() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='debug'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qrySelecaoDeArquivosPendentes
    SelectArquivosPendentes = qrySelecaoDeArquivosPendentes
                                                    
End Function

Private Sub compras_atualizarItensCompras()
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='debug'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryUpdateProcessamentoConcluido_ItensCompras
    Application.CurrentDb.Execute qryUpdateProcessamentoConcluido_ItensCompras

End Sub

'' #AILTON - VALIDAR
'' ATUALIZAR CAMPOS DE COMPRAS
Public Sub compras_atualizarCampos()
If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='debug'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "######## - compras_atualizarCampos()"
Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                    qryUpdate_tblCompraNF, _
                                                    qryUpdateIdFornCompraNF, _
                                                    qryUpdateItens_CFOP_CompraNFItem, _
                                                    qryUpdate_NFe, _
                                                    qryUpdate_CTe, _
                                                    qryUpdate_CTeItens, _
                                                    qryInsertItensCTe, _
                                                    qryUpdateProcessamentoConcluido_CTE, _
                                                    qryUpdateItens_ID_Prod_CompraNFItem): executarComandos qryProcessos
End Sub

Public Sub TratamentoDeDadosGerais()
If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='debug'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "######## - TratamentoDeDadosGerais()"
Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                qryUpdateFornecedoresValidos, _
                                                qryUpdateRegistrosValidos, _
                                                qryUpdateIdTipo, _
                                                qryUpdateRetornoArmazem, _
                                                qryUpdateTransferenciaSisparts, _
                                                qryUpdateSit_CompraNF, _
                                                qryUpdateCTe, _
                                                qryUpdateIdEmpresa, _
                                                qryUpdateIdEmpresa_TransferenciaEntreFiliais, _
                                                qryUpdate_IDVD, _
                                                qryUpdateCFOP_PSC_PES): executarComandos qryProcessos

End Sub

Public Sub CriarRepositorios()
If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='debug'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "######## - CriarRepositorios()"
Dim DadosGerais As New clsConexaoNfeCte

Dim caminhoNovo As Variant
Dim caminhoAntigo As Variant
Dim strCaminhos() As Variant: strCaminhos = Array(DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColetaAcoes'"))

    '' PRINCIPAL
    For Each caminhoNovo In strCaminhos
        CreateDir CStr(caminhoNovo)
    Next
    
    '' EMPRESAS
    For Each caminhoAntigo In Array(DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColeta'"), _
                                    DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColetaProcessados'"), _
                                    DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColetaExpurgo'"))
        For Each caminhoNovo In carregarParametros(DadosGerais.SelectColetaEmpresa)
            CreateDir Replace(caminhoAntigo, "empresa", caminhoNovo)
        Next
    Next
    

Set DadosGerais = Nothing

End Sub


'' REPOSITORIO LOCAL PARA RELACIONAR ID_COMPRA COM CHV_ACESSO
Public Sub criarTabelaTemporariaParaRelacionarIdCompraComChvAcesso()
If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='debug'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "######## - criarTabelaTemporariaParaRelacionarIdCompraComChvAcesso()"

'' #REPOSITORIO_LOCAL - NOME DA TABELA REPOSITORIO
Dim pTabelaNome As String: _
    pTabelaNome = "tmpCompras_ID_CompraNF"

'' #REPOSITORIO_LOCAL - ESTRUTURA
Dim qryCompras_ID_ComprasNF As String: _
    qryCompras_ID_ComprasNF = "SELECT DISTINCT '' AS ID_CompraNF , '' AS ChvAcesso_CompraNF, '' AS NumPed_CompraNF INTO tmpCompras_ID_CompraNF FROM tblCompraNF;"

    '' #REPOSITORIO_LOCAL - CRIAÇÃO
    If IsNull(DLookup("Name", "MSysObjects", "type in(1,6) and Name='" & pTabelaNome & "'")) Then Application.CurrentDb.Execute qryCompras_ID_ComprasNF, dbSeeChanges

    '' #REPOSITORIO_LOCAL - LIMPAR REPOSITORIO
    Application.CurrentDb.Execute "DELETE FROM tmpCompras_ID_CompraNF", dbSeeChanges
    
End Sub

