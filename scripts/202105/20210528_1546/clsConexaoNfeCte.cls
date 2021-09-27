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

'' #tblOrigemDestino
'Private Const azsUpdateOrigemDestino_tabela_campo As String = _
'        "UPDATE tblOrigemDestino SET tblOrigemDestino.tabela = strSplit([Destino],'.',0), tblOrigemDestino.campo = strSplit([Destino],'.',1);"

'' -- CARREGAR TAGs DE VINDAS DO XML
'' #tblOrigemDestino.tabela
Private Const qryTags As String = _
        "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((Len([Tag]))>0) AND ((tblOrigemDestino.tabela) = 'strParametro') AND ((tblOrigemDestino.TagOrigem)=1));"



'' ########################################
'' #tblDadosConexaoNFeCTe - #DADOS_GERAIS
'' ########################################



'' tblDadosConexaoNFeCTe.Pendentes
Private Const qrySelecaoDeArquivosPendentes As String = _
    "SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0)) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0) ORDER BY tblDadosConexaoNFeCTe.CaminhoDoArquivo;"

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

''' #tblCompraNFItem.ID_Prod_CompraNFItem
'Private Const qryUpdateItens_ID_Prod_CompraNFItem As String = _
'        "UPDATE tblCompraNFItem SET tblCompraNFItem.ID_Prod_CompraNFItem = DLookUp(""CodigoProd_Grade"",""dbo_tabGradeProdutos"",""CodigoForn_Grade='"" & [tblCompraNFItem].[ID_Prod_CompraNFItem] & ""'"");"


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


Public Function SelectArquivosPendentes() As String
    SelectArquivosPendentes = qrySelecaoDeArquivosPendentes
                                                    
End Function

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


Public Sub FormatarItensDeCompras()
    Application.CurrentDb.Execute qryUpdateItens_ID_Prod_CompraNFItem

End Sub


Public Sub TratamentoDeDadosGerais()
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










