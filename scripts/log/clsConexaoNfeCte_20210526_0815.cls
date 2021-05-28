VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConexaoNfeCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'' 02. processamentoDeCompras
'' 03. enviar_ComprasParaServidor
'' 00. criar_ArquivoJson
'' 00. ProcessarArquivosXml
'' 00. ADM_carregarDadosDoServidor
''
''------------------------------------



Option Compare Database
Option Explicit

Private FileCollention As New Collection
Private con As ADODB.Connection

Public Enum enumTipoArquivo
    opFlagLancadaERP = 0
    opManifesto = 1
End Enum


'' ########################################
'' #tblProcessamento
'' ########################################

'' #tblProcessamento.Delete
Private Const qryDeleteProcessamento As String = _
        "DELETE * FROM tblProcessamento;"

'' #tblProcessamento.chave
Private Const qryUpdateProcessamento_Chave As String = _
        "UPDATE tblProcessamento SET tblProcessamento.chave = Replace([tblProcessamento].[chave],';','|');"

'' #tblProcessamento.tabela
Private Const qryUpdateProcessamento_RelacaoCamposDeTabelas As String = _
        "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = DLookUp(""tabela"",""tblOrigemDestino"",""tag='"" & [tblProcessamento].[chave] & ""'""), tblProcessamento.NomeCampo = DLookUp(""campo"",""tblOrigemDestino"",""tag='"" & [tblProcessamento].[chave] & ""'""), tblProcessamento.formatacao = DLookUp(""formatacao"",""tblOrigemDestino"",""tag='"" & [tblProcessamento].[chave] & ""'"") ;" ''WHERE (((tblProcessamento.NomeTabela)=""tblCompraNF""))

'' #tblProcessamento.formatacao
Private Const qryUpdateProcessamento_RelacaoCamposDeTabelas_Item_CompraNFItem As String = _
        "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = ""tblCompraNFItem"", tblProcessamento.NomeCampo = [tblProcessamento].[chave], tblProcessamento.formatacao = DLookUp(""formatacao"",""tblOrigemDestino"",""campo='Item_CompraNFItem'"") WHERE (((tblProcessamento.chave)=""Item_CompraNFItem""));"

'' #tblProcessamento.opTexto
Private Const qryUpdateProcessamento_RelacaoCamposDeTabelas_ChvAcesso_CompraNF As String = _
        "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = ""tblCompraNF"", tblProcessamento.NomeCampo = [tblProcessamento].[chave], tblProcessamento.formatacao = ""opTexto"" WHERE (((tblProcessamento.chave)=""ChvAcesso_CompraNF""));"

'' #tblProcessamento
Private Const qryUpdateProcessamento_RelacaoCamposDeTabelas_tblCompraNFItem_ChvAcesso_CompraNF As String = _
        "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = strSplit([tblProcessamento].[chave],'.',0), tblProcessamento.NomeCampo = strSplit([tblProcessamento].[chave],'.',1), tblProcessamento.formatacao = strSplit([tblProcessamento].[chave],'.',2) WHERE (((tblProcessamento.chave)=""tblCompraNFItem.ChvAcesso_CompraNF.opTexto""));"

'' #tblProcessamento.opData
Private Const qryUpdateProcessamento_opData As String = _
        "UPDATE tblProcessamento SET tblProcessamento.valor = Mid([tblProcessamento].[valor],1,10) WHERE formatacao = 'opData';"

'' #tblProcessamento.opTime
Private Const qryUpdateProcessamento_opTime As String = _
        "UPDATE tblProcessamento SET tblProcessamento.valor = Mid([tblProcessamento].[valor],12,8) WHERE formatacao = 'opTime';"

'' #tblProcessamento
Private Const qrySelecaoDeCampos As String = _
        "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((tblOrigemDestino.tabela)='strParametro') AND ((Len([Tag]))>0) AND ((tblOrigemDestino.TagOrigem)=1)) ORDER BY tblOrigemDestino.Tag, tblOrigemDestino.tabela;"
       
'' #tblProcessamento.Pendentes
Private Const qrySelecaoDeArquivosPendentes As String = _
        "SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1)) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0) ORDER BY tblDadosConexaoNFeCTe.CaminhoDoArquivo;"



'' -- SELEÇÃO DE ITENS UNICOS PARA PROCESSAMENTO
'' #tblProcessamento
'Private Const sqyProcessamentosPendentes As String = _
'        "SELECT DISTINCT pk from tblProcessamento;"



'' -- ATUALIZAR CAMPOS ( Nometabela e NomeCampo ) PARA USO DA TABELA DE PROCESSAMENTO
'' #tblProcessamento.Tabela = [qryOrigemDestino].[Tabela]
'' #tblProcessamento.Campo = [qryOrigemDestino].[Campo]
'' #tblProcessamento.formatacao = [qryOrigemDestino].[formatacao]
'Private Const qryUpdateProcessamento As String = _
'        "UPDATE (SELECT tblOrigemDestino.Destino " & _
'        "           ,tblOrigemDestino.Tag " & _
'        "           ,tblOrigemDestino.Tabela " & _
'        "           ,tblOrigemDestino.Campo " & _
'        "           ,tblOrigemDestino.formatacao " & _
'        "       FROM tblOrigemDestino " & _
'        "       WHERE tblOrigemDestino.Tabela = 'strParametro' " & _
'        "       ) AS qryOrigemDestino INNER JOIN tblProcessamento ON qryOrigemDestino.Tag = tblProcessamento.chave  " & _
'        "SET tblProcessamento.nomeTabela = [qryOrigemDestino].[Tabela] " & _
'        "   ,tblProcessamento.NomeCampo = [qryOrigemDestino].[Campo] " & _
'        "   ,tblProcessamento.formatacao = [qryOrigemDestino].[formatacao];"



'' ########################################
'' #tblParametros
'' ########################################

'' #tblParametros
'Private Const qryParametros As String = "SELECT tblParametros.ValorDoParametro FROM tblParametros WHERE (((tblParametros.TipoDeParametro) = 'strParametro'))"
'Private Const strCaminhoDeColeta As String = "caminhoDeColeta"
'Private Const strCaminhoDeProcessados As String = "caminhoDeProcessados"
'Private Const strUsuarioErpCod As String = "UsuarioErpCod"
'Private Const strUsuarioErpNome As String = "UsuarioErpNome"
'Private Const strCodTipoEvento As String = "codTipoEvento"
'Private Const strComando As String = "Comando"

'' ########################################
'' #tblOrigemDestino
'' ########################################

'' #tblOrigemDestino
'Private Const azsUpdateOrigemDestino_tabela_campo As String = _
'        "UPDATE tblOrigemDestino SET tblOrigemDestino.tabela = strSplit([Destino],'.',0), tblOrigemDestino.campo = strSplit([Destino],'.',1);"

'' -- CARREGAR TAGs DE VINDAS DO XML
'' #tblOrigemDestino.tabela
Private Const qryTags As String = _
        "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((Len([Tag]))>0) AND ((tblOrigemDestino.tabela) = 'strParametro') AND ((tblOrigemDestino.TagOrigem)=1));"





'' ########################################
'' #tblDadosConexaoNFeCTe - PROCESSAMENTO
'' ########################################

'' ------------------------
'' #FormatarCamposEmProcessamento
'' ------------------------





'' ########################################
'' #tblDadosConexaoNFeCTe - #DADOS_GERAIS
'' ########################################

'' -- PROCESSAMENTO CONCLUIDO
'' #tblDadosConexaoNFeCTe.registroProcessado
Private Const qryUpdateProcessamentoConcluido As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 1 WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0) AND ((tblDadosConexaoNFeCTe.Chave)='strChave'));"

'' #tblDadosConexaoNFeCTe.registroProcessado
Private Const qryUpdateProcessamentoConcluido_CTE As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 3 WHERE (((tblDadosConexaoNFeCTe.registroValido) = 1) AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1) AND ((tblDadosConexaoNFeCTe.ID_Tipo) = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro] = 'Cte'"")));"

'' #tblDadosConexaoNFeCTe.registroProcessado
Private Const qryUpdateProcessamentoConcluido_ItensCompras As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 3 FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0));"

'' -- ARQUIVOS - CONSULTA PARA CRIAÇÃO DE ARQUIVOS JSON
'' #tblDadosConexaoNFeCTe.Distinct
Private Const sqyDadosJson As String = _
        "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi]))>0));"


'' -- SELEÇÃO DE FORNECEDORES VALIDOS
'' #tblDadosConexaoNFeCTe.qryUpdateFornecedoresValidos
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
'' #tblDadosConexaoNFeCTe.ID_EMPRESA
'' -- CTE
Private Const qryUpdateIdEmpresa As String = _
        "UPDATE (SELECT tmpEmpresa.ID_Empresa, STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF FROM tmpEmpresa) AS qryEmpresas " & _
        " INNER JOIN tblDadosConexaoNFeCTe ON qryEmpresas.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_Rem " & _
        " SET tblDadosConexaoNFeCTe.ID_Empresa = qryEmpresas.ID_Empresa;"

'' #tblDadosConexaoNFeCTe.ID_EMPRESA
Private Const qryUpdateIdEmpresa_TransferenciaEntreFiliais As String = _
        "UPDATE (SELECT tmpEmpresa.ID_Empresa, STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF FROM tmpEmpresa) AS qryEmpresas " & _
        " INNER JOIN tblDadosConexaoNFeCTe ON qryEmpresas.strCNPJ_CPF = tblDadosConexaoNFeCTe.CPNJ_Dest " & _
        " SET tblDadosConexaoNFeCTe.ID_Empresa = qryEmpresas.ID_Empresa" & _
        " WHERE (((tblDadosConexaoNFeCTe.CFOP) = '6152'));"

'' -- FiltroCFOP
'' #tblDadosConexaoNFeCTe.FiltroCFOP
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
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='RetornoArmazem'"") WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = 0) AND ((tblDadosConexaoNFeCTe.codMod) = CInt('55')) AND ((tblDadosConexaoNFeCTe.CNPJ_emit) IN ('12680452000302')));"

'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 6 - NF-e Transferência com código Sisparts
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryUpdateTransferenciaSisparts As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='TransferenciaSisparts'"") WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = 0) AND ((tblDadosConexaoNFeCTe.codMod) = CInt('55')) AND ((tblDadosConexaoNFeCTe.CNPJ_emit) IN (SELECT CNPJ_Empresa FROM [tmpEmpresa])));"

'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 0 - CT-e
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryUpdateCTe As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='CTe'"") WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = 0) AND ((tblDadosConexaoNFeCTe.codMod) = CInt('57')));"

'' #carregarCompras
'' PROCESSAMENTO DAS COMPRAS COM BASE EM REGISTROS VALIDOS PROCESSADOS PELA #DADOS_GERAIS
'' #tblDadosConexaoNFeCTe
'Private Const qrySelectProcessamentoPendente As String = _
'        "SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo, tblDadosConexaoNFeCTe.ID_Tipo FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0));"

'' #tblDadosConexaoNFeCTe
'Private Const qrySelectProcessamentoItensCompras As String = _
'        "SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0));"



'' ########################################
'' #tblCompraNF - COMPRAS
'' ########################################

'' #tblCompraNF.BaseCalcICMS_CompraNF
Private Const qryUpdateBaseCalcICMS As String = _
        "UPDATE tblCompraNF SET tblCompraNF.BaseCalcICMS_CompraNF = replace([tblCompraNF].[BaseCalcICMS_CompraNF] / 100,"","",""."") WHERE (((tblCompraNF.BaseCalcICMS_CompraNF)>""0"")) OR (((tblCompraNF.BaseCalcICMS_CompraNF) Is Not Null));"


'' #tblCompraNF.IDVD_CompraNF
Private Const qryUpdate_IDVD As String = _
        "UPDATE tblCompraNF " & _
        "SET tblCompraNF.IDVD_CompraNF = Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6) " & _
        "WHERE (((Left([Obs_CompraNF], 6)) = 'PEDIDO ')  " & _
        "       AND ((tblCompraNF.CNPJ_CPF_CompraNF) = '12680452000302') " & _
        "       AND ((Val(Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6))) > 0));"

'' #tblCompraNF.CFOP_CompraNF
Private Const qryUpdateCFOP_FilCompra As String = _
        "UPDATE tblCompraNF SET tblCompraNF.CFOP_CompraNF = DLookUp(""[FiltroCFOP]"",""[tblDadosConexaoNFeCTe]"",""[ChvAcesso]='"" & [tblCompraNF].[ChvAcesso_CompraNF] & ""'""), tblCompraNF.Fil_CompraNF = DLookUp(""[ID_EMPRESA]"",""[tblDadosConexaoNFeCTe]"",""[ChvAcesso]='"" & [tblCompraNF].[ChvAcesso_CompraNF] & ""'"");"

'' #tblCompraNF.Fil_CompraNF
Private Const qryUpdateFilCompraNF As String = _
        "UPDATE (SELECT tmpEmpresa.ID_Empresa, STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF, tmpEmpresa.CNPJ_Empresa FROM tmpEmpresa WHERE (((tmpEmpresa.CNPJ_Empresa) Is Not Null))) AS qryEmpresas INNER JOIN tblCompraNF ON qryEmpresas.strCNPJ_CPF = tblCompraNF.CNPJ_CPF_CompraNF SET tblCompraNF.Fil_CompraNF = qryEmpresas.ID_Empresa;"

'' #tblCompraNF.ID_Forn_CompraNF
Private Const qryUpdateIdFornCompraNF As String = _
        "UPDATE (SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF, tmpClientes.CÓDIGOClientes FROM tmpClientes) AS qryClientesFornecedor " & _
        " INNER JOIN tblCompraNF ON tblCompraNF.CNPJ_CPF_CompraNF = qryClientesFornecedor.strCNPJ_CPF " & _
        " SET tblCompraNF.ID_Forn_CompraNF = qryClientesFornecedor.CÓDIGOClientes;"

'' #tblCompraNF.ModeloDoc_CompraNF
'' #tblCompraNF.CFOP_CompraNF
Private Const qryUpdate_ModeloDoc_CFOP As String = _
        "UPDATE tblCompraNF INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso SET tblCompraNF.ModeloDoc_CompraNF = [tblDadosConexaoNFeCTe].[codMod], tblCompraNF.CFOP_CompraNF = [tblDadosConexaoNFeCTe].[FiltroCFOP] WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)>0));"

'' ########################################
'' #tblCompraNFItem - COMPRAS ITENS
'' ########################################

''' UPDATE - ID_Prod_CompraNFItem
Private Const qryUpdateItens_ID_Prod_CompraNFItem As String = _
        "UPDATE tblCompraNFItem SET tblCompraNFItem.ID_Prod_CompraNFItem = DLookUp(""CodigoProd_Grade"",""dbo_tabGradeProdutos"",""CodigoForn_Grade='"" & [tblCompraNFItem].[ID_Prod_CompraNFItem] & ""'"");"

'' #tblCompraNFItem.InsertItensCTe
Private Const qryInsertItensCTe As String = _
        "INSERT INTO tblCompraNFItem ( ChvAcesso_CompraNF, VUnt_CompraNFItem, DTEmi_CompraNFItem, Num_CompraNFItem, VTot_CompraNFItem, DebICMS_CompraNFItem, VTotBaseCalcICMS_CompraNFItem, ID_NatOp_CompraNFItem, Item_CompraNFItem, ID_Grade_CompraNFItem, QtdFat_CompraNFItem, ISS_CompraNFItem, IPI_CompraNFItem, FlagEst_CompraNFItem, BaseCalcICMS_CompraNFItem, IseICMS_CompraNFItem, OutICMS_CompraNFItem, BaseCalcIPI_CompraNFItem, DebIPI_CompraNFItem, IseIPI_CompraNFItem, OutIPI_CompraNFItem, TxMLSubsTrib_CompraNFItem, TxIntSubsTrib_CompraNFItem, TxExtSubsTrib_CompraNFItem, BaseCalcICMSSubsTrib_CompraNFItem, VTotICMSSubsTrib_compranfitem, VTotFrete_CompraNFItem, VTotDesc_CompraNFItem, VTotPIS_CompraNFItem, VTotBaseCalcPIS_CompraNFItem, VTotCOFINS_CompraNFItem, VTotIseICMS_CompraNFItem, VTotBaseCalcCOFINS_CompraNFItem, SNCredICMS_CompraNFItem, VTotSNCredICMS_CompraNFItem, VTotSeg_CompraNFItem, VTotOutDesp_CompraNFItem )  " & _
        "SELECT tblCompraNF.ChvAcesso_CompraNF, tblCompraNF.VTotNF_CompraNF, tblCompraNF.DTEmi_CompraNF, tblCompraNF.NumNF_CompraNF, tblCompraNF.VTotNF_CompraNF, IIf([VTotICMS_CompraNF]<>"""",replace(tblCompraNF.BaseCalcICMS_CompraNF/100,"","","".""),0) AS strVTotICMS, tblCompraNF.BaseCalcICMS_CompraNF, tblCompraNF.ID_NatOp_CompraNF, 1 AS strItem, 1 AS strIDGrade, 1 AS strQtdFat, 0 AS strISS, 0 AS strIPI, 0 AS strFlag, 100 AS strBaseCalcICMS, 0 AS strIseICMS, 0 AS strOutICMS, 0 AS strBaseCalcIPI, 0 AS strDebIPI, 0 AS strIseIPI, 0 AS strOutIPI, 0 AS strTxMLSubsTrib, 0 AS strTxIntSubsTrib, 0 AS strTxExtSubsTrib, 0 AS strBaseCalcICMSSubsTrib, 0 AS strVTotICMSSubsTrib, 0 AS strVTotFrete, 0 AS strVTotDesc, 0 AS strVTotPIS, 0 AS strVTotBaseCalcPIS, 0 AS strVTotCOFINS, 0 AS strVTotIseICMS, 0 AS strVTotBaseCalcCOFINS, 0 AS strSNCredICMS, 0 AS strVTotSNCredICMS, 0 AS strVTotSeg, 0 AS strVTotOutDesp  " & _
        "FROM tblCompraNF INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso  " & _
        "WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)=DLookUp(""[ValorDoParametro]"",""[tblParametros]"",""[TipoDeParametro]='Cte'""))) AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1) ; "

''#######################################################################################
''### CONSULTAS DE CONTROLE DE DADOS
''#######################################################################################


'Private Sub TratamentoDeCompras()
'Dim qryProcessos() As Variant: qryProcessos = Array(qryUpdateFilCompraNF): executarComandos qryProcessos
'
'End Sub

Private Sub compras_atualizarItensCompras()
    Application.CurrentDb.Execute qryUpdateProcessamentoConcluido_ItensCompras

End Sub

'' #AILTON - VALIDAR
'' ATUALIZAR CAMPOS DE COMPRAS
Private Sub compras_atualizarCampos()
Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                    qryUpdateFilCompraNF, _
                                                    qryUpdateCFOP_FilCompra, _
                                                    qryUpdate_IDVD, _
                                                    qryUpdateBaseCalcICMS, _
                                                    qryUpdateIdFornCompraNF, _
                                                    qryUpdate_ModeloDoc_CFOP): executarComandos qryProcessos

End Sub

'' #AILTON
'' CARREGAR ITENS CTE
Private Sub compras_carregarItensCTe()
Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                    qryInsertItensCTe, _
                                                    qryUpdateProcessamentoConcluido_CTE): executarComandos qryProcessos
                                                    
End Sub

Private Sub FormatarCamposEmProcessamento()
Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                    qryUpdateProcessamento_opData, _
                                                    qryUpdateProcessamento_opTime): executarComandos qryProcessos
                                                    
End Sub

Private Sub TratamentoDeDadosGerais()
Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                qryUpdateFornecedoresValidos, _
                                                qryUpdateRegistrosValidos, _
                                                qryUpdateIdTipo, _
                                                qryUpdateRetornoArmazem, _
                                                qryUpdateTransferenciaSisparts, _
                                                qryUpdateCTe, _
                                                qryUpdateIdEmpresa, _
                                                qryUpdateIdEmpresa_TransferenciaEntreFiliais, _
                                                qryUpdateCFOP_PSC_PES): executarComandos qryProcessos

End Sub


''#######################################################################################
''### PRINCIPAL
''#######################################################################################

'' 03.ENVIAR DADOS PARA SERVIDOR
Public Function enviar_ComprasParaServidor()
On Error GoTo adm_Err

'' VARIAVEL DE PARAMETRO
Dim pDestino As String: pDestino = "tblCompraNF"

'' ---------------------
'' VARIAVEIS GERAIS
'' ---------------------

'' LISTAGEM DE CAMPOS DA TABELA ORIGEM/DESTINO
Dim strCampo As String

'' SCRIPT
Dim tmpScript As String
Dim tmp As String

'' ---------------------
'' BANCO LOCAL
'' ---------------------
Dim db As DAO.Database: Set db = CurrentDb
Dim rstOrigem As DAO.Recordset

'' ---------------------
'' BANCO DESTINO
'' ---------------------

Dim strUsuarioNome As String
Dim strUsuarioSenha As String
Dim strOrigem As String
Dim strBanco As String

Dim dbDestino As New Banco

Dim retVal As Variant: retVal = MsgBox("Deseja enviar compras para o servidor?", vbQuestion + vbYesNo, "ADM_enviarComprasParaServidor")

    If retVal = vbYes Then
        strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
        strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
        strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
        strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
    
        dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
        dbDestino.SqlSelect "SELECT * FROM " & pDestino
    
        tmpScript = "Insert into " & pDestino & " ("
    
        '' 1. cabeçalho
        Dim rstCampos As DAO.Recordset
        Set rstCampos = db.OpenRecordset("Select distinct campo,formatacao,valorPadrao from tblOrigemDestino where tblOrigemDestino.tabela = '" & pDestino & "'")
        Do While Not rstCampos.EOF
            tmpScript = tmpScript & rstCampos.Fields("campo").value & ","
            rstCampos.MoveNext
            DoEvents
        Loop
        tmpScript = left(tmpScript, Len(tmpScript) - 1) & ") values ("
        
        
        '' BANCO LOCAL
        Set rstOrigem = db.OpenRecordset("Select * from " & pDestino)
        Do While Not rstOrigem.EOF
            tmp = ""
            
            '' LISTAGEM DE CAMPOS
            rstCampos.MoveFirst
            Do While Not rstCampos.EOF
            
                '' CRIAR SCRIPT DE INCLUSÃO DE DADOS NA TABELA DESTINO
                '' 2. campos x formatação
        
                If rstCampos.Fields("formatacao").value = "opTexto" Then
                    tmp = tmp & "'" & rstOrigem.Fields(rstCampos.Fields("campo").value).value & "',"
                    
                ElseIf rstCampos.Fields("formatacao").value = "opNumero" Or rstCampos.Fields("formatacao").value = "opMoeda" Then
                    tmp = tmp & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", rstOrigem.Fields(rstCampos.Fields("campo").value).value, rstCampos.Fields("valorPadrao").value) & ","
                    
                ElseIf rstCampos.Fields("formatacao").value = "opTime" Or rstCampos.Fields("formatacao").value = "opData" Then
                    tmp = tmp & "'" & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", rstOrigem.Fields(rstCampos.Fields("campo").value).value, rstCampos.Fields("valorPadrao").value) & "',"
                    
                End If
                
                rstCampos.MoveNext
                DoEvents
            Loop
            
            '' BANCO DESTINO
            tmp = left(tmp, Len(tmp) - 1) & ")"
            
            'Debug.Print tmpScript & tmp
            
            rstOrigem.MoveNext
            
            dbDestino.SqlExecute tmpScript & tmp
            
            DoEvents
        Loop
        
        dbDestino.CloseConnection
        db.Close: Set db = Nothing
        
'        MsgBox "Fim!", vbOKOnly + vbExclamation, "enviarComprasParaServidor"
    
    End If

adm_Exit:
    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit


End Function

'' 02.CARREGAR COMPRAS ANTES DE ENVIAR PARA O SERVIDOR

Public Function processamentoDeCompras(pPathFile As String)
On Error Resume Next

Dim s As New clsConexaoNfeCte

'' CHAVES DE CONTROLE
Dim pPK As String: pPK = DLookup("[Chave]", "[tblDadosConexaoNFeCTe]", "[CaminhoDoArquivo]='" & pPathFile & "'")
Dim pChvAcesso As String: pChvAcesso = DLookup("[ChvAcesso]", "[tblDadosConexaoNFeCTe]", "[CaminhoDoArquivo]='" & pPathFile & "'")

'' CARREGAR ARQUIVO
Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
XDoc.Load pPathFile

Dim cont As Integer: cont = XDoc.getElementsByTagName("infNFe/det").Length
Dim Item As Variant

Dim pDados As New Collection
Dim idItem As String: idItem = ""
Dim regiao As String: regiao = ""
Dim subRegiao As String: subRegiao = ""
Dim campo As String: campo = ""
Dim valor As String: valor = ""


    '' IDENTIFICAÇÃO DO ARQUIVO
    pDados.add pPK & "|" & "CaminhoDoArquivo" & "|" & pPathFile

    '' CHAVE DE ACESSO
    pDados.add pPK & "|" & "ChvAcesso_CompraNF" & "|" & pChvAcesso
    
    '' CABEÇALHO DA COMPRA
    For Each Item In carregarParametros(qrySelecaoDeCampos, "tblCompraNF")
        Select Case UBound(Split((Item), "|"))
            Case 1
                regiao = Split((Item), "|")(0)
                campo = Split((Item), "|")(1)
                valor = XDoc.SelectNodes(regiao).Item(0).SelectNodes(campo).Item(0).text
                If valor <> "" Then pDados.add pPK & "|" & Replace(Item, "|", ";") & "|" & valor
            Case Else
        End Select
        
        regiao = ""
        campo = ""
        valor = ""
        DoEvents
    Next Item


'    '' ITENS DA COMPRA
'    For i = 0 To cont - 1
'        '' ID
'        idItem = CStr(XDoc.getElementsByTagName("nfeProc/NFe/infNFe/det").Item(i).Attributes(0).value)
'        pDados.add pPK & "_" & idItem & "|" & "Item_CompraNFItem" & "|" & idItem
'
'        pDados.add pPK & "_" & idItem & "|" & "tblCompraNFItem.ChvAcesso_CompraNF.opTexto" & "|" & pChvAcesso
'
'        For Each Item In carregarParametros(qrySelecaoDeCampos, "tblCompraNFItem")
'
'            Select Case UBound(Split((Item), "|"))
'
'                '' ITEM DE COMPRA
'                Case 1
'                    regiao = Split((Item), "|")(0)
'                    campo = Split((Item), "|")(1)
'                    valor = XDoc.SelectNodes(regiao).Item(i).SelectNodes(campo).Item(0).text
'                    If valor <> "" Then pDados.add pPK & "_" & idItem & "|" & Replace(Item, "|", ";") & "|" & valor
'
'                '' IMPOSTO
'                Case 2
'                    regiao = Split((Item), "|")(0)
'                    subRegiao = Split((Item), "|")(1)
'                    campo = Split((Item), "|")(2)
'                    valor = XDoc.SelectNodes(regiao).Item(i).SelectNodes(subRegiao).Item(0).getElementsByTagName(campo).Item(0).text
'                    If valor <> "" Then pDados.add pPK & "_" & idItem & "|" & Replace(Item, "|", ";") & "|" & valor
'
'                Case Else
'            End Select
'
'            regiao = ""
'            subRegiao = ""
'            campo = ""
'            valor = ""
'            DoEvents
'        Next Item
'        DoEvents
'    Next i

    '' CADASTRAR DADOS
    cadastroProcessamento pDados
    
    '' LIMPAR COLEÇÃO
    ClearCollection pDados
    
    '' AJUSTAR CAMPO CHAVE
    Application.CurrentDb.Execute qryUpdateProcessamento_Chave
    
    '' RELACAO DE CAMPOS DE TABELAS
    Application.CurrentDb.Execute qryUpdateProcessamento_RelacaoCamposDeTabelas
    
    '' RELACAO DE CAMPO Item_CompraNFItem
    Application.CurrentDb.Execute qryUpdateProcessamento_RelacaoCamposDeTabelas_Item_CompraNFItem
    
    '' RELACAO DE CAMPO tblCompraNFItem_ChvAcesso_CompraNF
    Application.CurrentDb.Execute qryUpdateProcessamento_RelacaoCamposDeTabelas_tblCompraNFItem_ChvAcesso_CompraNF

Set XDoc = Nothing

End Function

'' 01.CARREGAR DADOS GERAIS - CONCLUIDO
Public Function carregar_DadosGerais()
On Error GoTo adm_Err

Dim strProcessamento As String: strProcessamento = "tblDadosConexaoNFeCTe"
Dim s As New clsConexaoNfeCte
Dim t As Variant

Dim retVal As Variant: retVal = MsgBox("Deseja carregar os dados gerais dos arquivos XML ?", vbQuestion + vbYesNo, "carregarDadosGerais")

    If retVal = vbYes Then
    
        '' #LIMPAR_BASE_DE_TESTES
        Application.CurrentDb.Execute "DELETE FROM tblDadosConexaoNFeCTe"
        Application.CurrentDb.Execute "DELETE FROM tblCompraNF"
        Application.CurrentDb.Execute "DELETE FROM tblCompraNFItem"
    
        '' #CARREGAR DADOS
        For Each t In Array(strProcessamento)
    
            '' #PROCESSAMENTO DE ARQUIVO - ENVIO DE DADOS PARA tblProcessamento
            ProcessarArquivosXml CStr(t), GetFilesInSubFolders(DLookup("ValorDoParametro", "tblParametros", "TipoDeParametro='caminhoDeColeta'"))
            'ProcessarArquivosXml CStr(t), GetFilesInSubFolders(pegarValorDoParametro(qryParametros, strCaminhoDeColeta)) ''DLookup("ValorDoParametro","tblParametros","TipoDeParametro='caminhoDeColeta'")
    
            '' FORMATAR CAMPOS
            FormatarCamposEmProcessamento
    
            '' #TRANSFERIR DADOS PROCESSADOS - DADOS GERAIS - ENVIO DE DADOS PARA tblDadosConexaoNFeCTe
            TransferirDadosProcessados strProcessamento
    
            '' #TRATAMENTO DE DADOS GERAIS
            TratamentoDeDadosGerais
    
        Next
    
    End If

adm_Exit:
    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit

End Function

'' CRIAÇÃO DE ARQUIVOS JSON
Public Sub criar_ArquivoJson(pArquivo As enumTipoArquivo, sqyDadosJson As String, Optional caminhoDeProcessados As String)
Dim db As DAO.Database: Set db = CurrentDb
Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(sqyDadosJson)
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
        dados.codUsuarioErp = DLookup("ValorDoParametro", "tblParametros", "TipoDeParametro='UsuarioErpCod'")
        dados.nomeUsuarioErp = DLookup("ValorDoParametro", "tblParametros", "TipoDeParametro='UsuarioErpNome'")
        
        '' CAMINHO DE SAIDA DO ARQUIVO
        dados.caminhoDeProcessados = CreateDir(DLookup("ValorDoParametro", "tblParametros", "TipoDeParametro='caminhoDeProcessados'"))
        If caminhoDeProcessados <> "" Then dados.caminhoDeProcessados = caminhoDeProcessados
                        
        Select Case pArquivo
            Case opFlagLancadaERP
                dados.Comando = DLookup("ValorDoParametro", "tblParametros", "TipoDeParametro='Comando'") '' pegarValorDoParametro(qryParametros, strComando)
                dados.criarERP
                nomeArquivo = "FlagLancadaERP"
            Case opManifesto
                dados.codTipoEvento = DLookup("ValorDoParametro", "tblParametros", "TipoDeParametro='codTipoEvento'") '' pegarValorDoParametro(qryParametros, strCodTipoEvento)
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


'' #ADMINISTRACAO - RESPONSAVEL POR TRAZER OS DADOS DO SERVIDOR PARA AUXILIO NO PROCESSAMENTO. QUANDO NECESSARIO
Public Function ADM_carregarDadosDoServidor()
On Error GoTo adm_Err
Dim retVal As Variant: retVal = MsgBox("Deseja carregar dados do servidor?", vbQuestion + vbYesNo, "ADM_carregarDadosDoServidor")

    If retVal = vbYes Then
    
        '' NATUREZA DE OPERAÇÃO
        Application.CurrentDb.Execute "Delete from tmpNatOp"
        ImportarDados "tblNatOp", "tmpNatOp"
        
        '' CADASTRO DE EMPRESA
        Application.CurrentDb.Execute "Delete from tmpEmpresa"
        ImportarDados "tblEmpresa", "tmpEmpresa"
        
'        '' CADASTRO DE CLIENTES
'        Application.CurrentDb.Execute "Delete from tmpClientes"
'        ImportarDados "Clientes", "tmpClientes"
        
        MsgBox "Fim!", vbOKOnly + vbExclamation, "ADM_carregarDadosDoServidor"
    
    End If

adm_Exit:
    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit
    
End Function


''#######################################################################################
''### REPOSITORIO
''#######################################################################################


'' #CARREGAR DADOS DO ARQUIVO COM BASE NA TABELA DE ORIGEM DESTINO
Private Sub ProcessarArquivosXml(pTabelaDestino As String, pArquivos As Collection)
Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
Dim cadastro As clsProcessamento
Dim cRegistros As New Collection
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

    statusFinal DT_PROCESSO, pTabelaDestino & " - Quantidade de arquivos: " & pArquivos.count

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdInitMeter, "Processando arquivos...", pArquivos.count
        
    '' LIMPAR TABELA DE PROCESSAMENTOS
    Application.CurrentDb.Execute qryDeleteProcessamento
        
        
    '' 01.Leitura e identificação do arquivo
    For Each fileName In pArquivos
        XDoc.Load fileName
        
        '' #BARRA_PROGRESSO
        SysCmd acSysCmdUpdateMeter, contadorDeArquivos
        
        '' 01.CRIAR CHAVE UNICA DE REGISTRO PARA CONTROLE DE DADOS
        strPk = IIf(IsNull(DLookup("[chave]", "[tblDadosConexaoNFeCTe]", "[CaminhoDoArquivo]='" & fileName & "'")), getFileName(CStr(fileName)), DLookup("[chave]", "[tblDadosConexaoNFeCTe]", "[CaminhoDoArquivo]='" & fileName & "'"))
        cRegistros.add strPk & "|" & "CaminhoDoArquivo" & "|" & fileName
        
        '' 02.CARREGAR CAMPOS DE ORIGEM X DESTINO DO REGISTRO
        For Each Item In carregarParametros(qryTags, pTabelaDestino)
            Set lists = XDoc.SelectNodes("//" & Item)
            For Each fieldnode In lists
                If (fieldnode.HasChildNodes) Then
                    For Each childNode In fieldnode.ChildNodes
                        cRegistros.add strPk & "|" & Item & "|" & fieldnode.text
                    Next childNode
                End If
            Next fieldnode
            
            DoEvents
            
        Next Item

        '' 03. REALIZAR CADASTRO DE TODOS OS ITENS COLETADOS NA TABELA DE PROCESSAMENTO
        If (cRegistros.count > 2) Then
            
            '' CADASTRAR REGISTRO
             cadastroProcessamento cRegistros
        
        End If
        
        '' LIMPAR COLEÇÃO
        ClearCollection cRegistros

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

'' #TRANSFERIR DADOS PROCESSADOS PARA A TABELA DE TRANSITO - ( tblProcessamento -> tblDadosConexaoNFeCTe/tblCompraNF/tblCompraNFItem )
Public Sub TransferirDadosProcessados(pDestino As String)
On Error Resume Next

'' CONTROLE DE CADASTRO DE REGISTRO
Dim t As Integer
Dim tmpSqlOrigem As String

'' #BANCO_LOCAL
Dim db As DAO.Database: Set db = CurrentDb
Dim tmpSql As String: tmpSql = "Select Distinct pk from tblProcessamento where NomeTabela = '" & pDestino & "' Order by pk;"

If (pDestino = "tblCompraNFItem") Then tmpSql = "SELECT tblProcessamento.pk FROM tblProcessamento WHERE (((tblProcessamento.valor) In (SELECT tblProcessamento.valor as chave FROM tblProcessamento WHERE (((tblProcessamento.NomeCampo)='ChvAcesso_CompraNF') AND ((tblProcessamento.NomeTabela)='tblCompraNF')) ORDER BY tblProcessamento.ID)) AND ((tblProcessamento.NomeTabela)='tblCompraNFItem')) ORDER BY tblProcessamento.ID;"

Dim rstPendentes As DAO.Recordset: Set rstPendentes = db.OpenRecordset(tmpSql)
Dim rstOrigem As DAO.Recordset

'' #BANCO_DESTINO
tmpSql = "Select * from " & pDestino
Dim rstDestino As DAO.Recordset: Set rstDestino = db.OpenRecordset(tmpSql)

'' PROCESSAMENTO PENDENTE
Do While Not rstPendentes.EOF

    '' listar itens de compra
    tmpSqlOrigem = "Select * from tblProcessamento where NomeTabela = '" & pDestino & "' and pk = '" & rstPendentes.Fields("pk").value & "' and len(formatacao)>0  and len(NomeCampo)>0 Order by ID "
    
    '' REGISTROS DO PROCESSAMENTO PENDENTE
    Set rstOrigem = db.OpenRecordset(tmpSqlOrigem)
    Do While Not rstOrigem.EOF
    
        '' CONTROLE DE CADASTRO
        If t = 0 Then rstDestino.AddNew: t = 1
        rstDestino.Fields(rstOrigem.Fields("NomeCampo").value).value = rstOrigem.Fields("Valor").value
        
        '' #CAMPO_CHAVE_RELACIONAMENO
        If (pDestino = "tblDadosConexaoNFeCTe") Then rstDestino.Fields("Chave").value = rstOrigem.Fields("pk").value
    
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
    
    rstPendentes.MoveNext
    DoEvents
Loop

'dbDestino.CloseConnection
db.Close: Set db = Nothing

End Sub

Private Sub ImportarDados(pOrigem As String, pDestino As String)

'' #BANCO_ORIGEM
Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
Dim tmpOrigem As String: tmpOrigem = "Select * from " & pOrigem

Dim dboOrigem As New Banco: dboOrigem.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
dboOrigem.SqlSelect tmpOrigem

'' #BANCO_LOCAL
Dim db As DAO.Database: Set db = CurrentDb
Dim tmpDestino As String: tmpDestino = "Select * from " & pDestino
Dim rstDestino As DAO.Recordset: Set rstDestino = db.OpenRecordset(tmpDestino)
Dim rstOrigem As DAO.Recordset

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()
Dim fld As Variant, t As Variant

'' #BARRA_PROGRESSO
Dim contadorDeRegistros As Long: contadorDeRegistros = 1
SysCmd acSysCmdInitMeter, "Transferindo " & pOrigem & " ...", dboOrigem.rs.RecordCount

Do While Not dboOrigem.rs.EOF

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdUpdateMeter, contadorDeRegistros

    '' listar campos da tabela
    For Each fld In rstDestino.Fields
        If t = 0 Then rstDestino.AddNew: t = 1
        rstDestino.Fields(fld.Name).value = dboOrigem.rs.Fields(fld.Name).value
        DoEvents
    Next
    rstDestino.Update
    t = 0
    
    '' #BARRA_PROGRESSO
    contadorDeRegistros = contadorDeRegistros + 1
    dboOrigem.rs.MoveNext
    DoEvents
Loop

'' #BARRA_PROGRESSO
SysCmd acSysCmdRemoveMeter

'dbDestino.CloseConnection
db.Close: Set db = Nothing

'' #ANALISE_DE_PROCESSAMENTO
statusFinal DT_PROCESSO, "Processamento - ImportarDados ( " & pDestino & " )"

End Sub

Private Function cadastroProcessamento(pRegistros As Collection)
Dim cadastro As New clsProcessamento
Dim i As Variant

    '' CADASTRAR REGISTRO
    For Each i In pRegistros
        With cadastro
            .pk = Split(i, "|")(0)
            .Chave = Split(i, "|")(1)
            .valor = Mid(Split(i, "|")(2), 1, 255)
            .cadastrar
        End With
        
        DoEvents
    Next i

End Function
