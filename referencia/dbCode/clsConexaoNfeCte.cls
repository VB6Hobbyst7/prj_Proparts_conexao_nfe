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


'' ########################################
'' #Compras
'' ########################################

Private Const qryCompras_Select_RegistroValidoPorcessado As String = _
        "SELECT pRepositorio.* FROM pRepositorio INNER JOIN tblDadosConexaoNFeCTe ON pRepositorio.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso WHERE (((tblDadosConexaoNFeCTe.registroValido) = 1) AND (tblDadosConexaoNFeCTe.registroProcessado) = 1)"
        
Private Const qryCompras_Update_RegistroConcluida As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 2 WHERE (((tblDadosConexaoNFeCTe.ChvAcesso)=""strChave"") AND ((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1));"

Private Const qryComprasItens_Update_IDCompraNF As String = _
        "UPDATE tblCompraNFItem SET tblCompraNFItem.ID_CompraNF_CompraNFItem = strID_Compra WHERE  tblCompraNFItem.ChvAcesso_CompraNF = 'strChave'"

Private Const qryComprasCTe_Update_AjustesCampos_LOCAL As String = _
    "UPDATE tblCompraNF INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso SET tblCompraNF.HoraEntd_CompraNF = Null, tblCompraNF.IDVD_CompraNF = Null WHERE (((tblDadosConexaoNFeCTe.codMod)=57));"


'' ########################################
'' #tblParametros
'' ########################################

'' SELEÇÃO DE EMPRESAS
Private Const qryParametros_Select_ColetaEmpresa As String = _
    "SELECT tblParametros.ValorDoParametro FROM tblParametros WHERE tblParametros.TipoDeParametro = ""caminhoDeColetaEmpresa"";"

Private Const qryParametros_Select_CamposNomes As String = _
                            "SELECT distinct   " & _
                            "   tblParametros.TipoDeParametro  " & _
                            "   , tblParametros.ID  " & _
                            "   , tblOrigemDestino.campo  " & _
                            "   , tblOrigemDestino.formatacao  " & _
                            "   , tblOrigemDestino.valorPadrao  " & _
                            "FROM   " & _
                            "   tblParametros INNER JOIN tblOrigemDestino ON tblParametros.ValorDoParametro = tblOrigemDestino.campo  " & _
                            "WHERE (((tblParametros.TipoDeParametro)='pRepositorio') AND ((tblOrigemDestino.TagOrigem)<>0)) ORDER BY tblParametros.ID;"


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

'' ------------------------------------------------------------------------------------------------------------------------
'' -- ID_TIPO - APENAS TIPOS COM ID DE VALOR ZERO(0) SERÃO ATUALIZADOS PARA NÃO COMPROMETER OS REGISTROS JÁ PROCESSADOS
'' ------------------------------------------------------------------------------------------------------------------------

'' -- TIPOS DE CADASTRO - tblTipos
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryDadosGerais_Update_IdTipo As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.ID_Tipo = 0 WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) Is Null));"

'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 0 - CT-e -> ( ID_Tipo: 1 )
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryDadosGerais_Update_CTe As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='CTe'"") WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = 0) AND ((tblDadosConexaoNFeCTe.codMod) = CInt('57')));"

'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 4 - NF-e Retorno Armazém -> ( ID_Tipo: 5 )
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryDadosGerais_Update_RetornoArmazem As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='RetornoArmazem'"") WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = 0) AND ((tblDadosConexaoNFeCTe.codMod) = CInt('55')) AND ((tblDadosConexaoNFeCTe.CNPJ_emit) IN ('12680452000302')));"

'' #tblDadosConexaoNFeCTe.CFOP
Private Const qryDadosGerais_Update_RetornoArmazem_CFOP As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.CFOP = ""5907"" WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)=DLookUp(""id"",""tblTipos"",""Descricao='4 - NF-e Retorno Armazém'"")) AND ((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0));"

'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 6 - NF-e Transferência com código Sisparts -> ( ID_Tipo: 7 )
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryDadosGerais_Update_TransferenciaSisparts As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='TransferenciaSisparts'"") WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = 0) AND ((tblDadosConexaoNFeCTe.codMod) = CInt('55')) AND ((tblDadosConexaoNFeCTe.CNPJ_emit) IN (SELECT CNPJ_Empresa FROM [tmpEmpresa])));"

Private Const qryDadosGerais_Update_TransferenciaSisparts_CFOP As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.CFOP = ""6152"" WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)=DLookUp(""id"",""tblTipos"",""Descricao='6 - NF-e Transferência com código Sisparts'"")) AND ((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0));"

'' #20210823_qryDadosGerais_Update_Sit_CompraNF
'' #tblDadosConexaoNFeCTe.Sit_CompraNF
'' CONTROLE DE FINALIDADES
Private Const qryDadosGerais_Update_Sit_CompraNF As String = "UPDATE  ( SELECT   " & _
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

'' #20210823_qryDadosGerais_Update_IdFornCompraNF
'' #tblDadosConexaoNFeCTe.ID_Forn_CompraNF
Private Const qryDadosGerais_Update_IdFornCompraNF As String = _
        "UPDATE (SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF, tmpClientes.CÓDIGOClientes FROM tmpClientes) AS qryClientesFornecedor " & _
        "INNER JOIN tblDadosConexaoNFeCTe ON tblDadosConexaoNFeCTe.CNPJ_emit = qryClientesFornecedor.strCNPJ_CPF " & _
        "SET tblDadosConexaoNFeCTe.ID_Forn_CompraNF = qryClientesFornecedor.CÓDIGOClientes;"

'' ------------------------------------------------------------------------------------------------------------------------
'' -- registroProcessado - CONTROLE DE PROCESSAMENTO DE REGISTROS
'' ------------------------------------------------------------------------------------------------------------------------

'' -- EXCLUSAO DE REGISTROS INVALIDOS -- DESCONTINUADO - ( PODEMOS USAR DEPOIS )
Private Const qryDadosGerais_Delete_RegistrosInvalidos As String = _
        "DELETE * FROM tblDadosConexaoNFeCTe WHERE tblDadosConexaoNFeCTe.registroValido = 0;"

'' ATUALIZAR PARA REGISTRO PROCESSADO
'' #tblDadosConexaoNFeCTe.registroProcessado = 1
Private Const qryDadosGerais_Update_ProcessamentoConcluido As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 1 WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0) AND ((tblDadosConexaoNFeCTe.chave)='strChave'));"

'' ATUALIZAR PARA REGISTRO PROCESSADO
'' #tblDadosConexaoNFeCTe.registroProcessado = 1
Private Const qryDadosGerais_Update_ProcessamentoConcluido_CTE As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 1 WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1) AND ((tblDadosConexaoNFeCTe.ID_Tipo) = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro] = 'Cte'"")));"

'' ATUALIZAR REGISTRO PARA ENVIADO PARA SERVIDOR
'' #tblDadosConexaoNFeCTe.registroProcessado = 2 - Enviado para servidor
Private Const qryDadosGerais_Update_ProcessamentoConcluido_Servidor As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 2 WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1) AND ((tblDadosConexaoNFeCTe.ChvAcesso)='strChave'));"

'' #20210823_FornecedoresValidos
'' -- SELEÇÃO DE FORNECEDORES VALIDOS
'' #tblDadosConexaoNFeCTe.qryDadosGerais_Update_FornecedoresValidos
'' #tblDadosConexaoNFeCTe.registroValido
Private Const qryDadosGerais_Update_FornecedoresValidos As String = _
        "UPDATE (SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF FROM tmpClientes) AS qryFornecedoresValidos INNER JOIN tblDadosConexaoNFeCTe ON qryFornecedoresValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit SET tblDadosConexaoNFeCTe.registroValido = 1 WHERE ((tblDadosConexaoNFeCTe.registroProcessado) = 0);"


'' -- SELEÇÃO DE REGISTROS VALIDOS.qryRegistrosValidos
'' #tblDadosConexaoNFeCTe.qryDadosGerais_Update_RegistrosValidos
Private Const qryDadosGerais_Update_RegistrosValidos As String = _
        "UPDATE (SELECT STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF FROM tmpEmpresa) AS qryRegistrosValidos INNER JOIN tblDadosConexaoNFeCTe ON qryRegistrosValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit SET tblDadosConexaoNFeCTe.registroValido = 1;"

'' ------------------------------------------------------------------------------------------------------------------------
'' -- CONSULTAS - ...
'' ------------------------------------------------------------------------------------------------------------------------

'' tblDadosConexaoNFeCTe.Pendentes
Private Const qryDadosGerais_Select_ArquivosPendentes As String = _
    "SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0)) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0) ORDER BY tblDadosConexaoNFeCTe.CaminhoDoArquivo;"

'' #20210823_qryDadosGerais_Update_ID_NatOp_CompraNF__FiltroCFOP
'' #tblDadosConexaoNFeCTe.FiltroCFOP
'' #tmpDadosConexaoNFeCTe.ID_NatOp_CompraNF
'' RELACIONAMENTO DE CFOP PARA DENTRO E FORA DO ESTADO
Private Const qryDadosGerais_Update_ID_NatOp_CompraNF__FiltroCFOP As String = _
            "UPDATE  ( SELECT " & _
            "           tmpNatOp.ID_NatOper, tmpNatOp.Fil_NatOper, tmpNatOp.CFOP_NatOper, qryPscPes.strXMLCFOP, qryPscPes.strEstado  " & _
            "       FROM (SELECT  " & _
            "               strSplit(ValorDoParametro,'|',0) AS strFil_NatOper,  strSplit(ValorDoParametro,'|',1) AS strEstado,  strSplit(ValorDoParametro,'|',2) AS strXMLCFOP,  strSplit(ValorDoParametro,'|',3) AS strCFOP_NatOper  " & _
            "             FROM  " & _
            "               tblParametros  " & _
            "             WHERE  " & _
            "               TipoDeParametro='FiltroFil' And strSplit(ValorDoParametro,'|',0) In ('PSC','PES','PSP'))  AS qryPscPes  " & _
            "       INNER JOIN tmpNatOp ON (qryPscPes.strCFOP_NatOper = tmpNatOp.CFOP_NatOper) AND (qryPscPes.strFil_NatOper = tmpNatOp.Fil_NatOper) )  AS tmpPscPes  " & _
            "INNER JOIN  " & _
            "   (   SELECT  *  " & _
            "       FROM  tblDadosConexaoNFeCTe  " & _
            "       WHERE tblDadosConexaoNFeCTe.registroValido IN (SELECT TOP 1 cint(tblParametros.ValorDoParametro) FROM [tblParametros] WHERE TipoDeParametro = 'registroValido')  " & _
            "       AND tblDadosConexaoNFeCTe.ID_NatOp_CompraNF IS NULL )  AS tmpDadosConexaoNFeCTe " & _
            "ON (tmpPscPes.strXMLCFOP = tmpDadosConexaoNFeCTe.CFOP) AND (tmpPscPes.Fil_NatOper = tmpDadosConexaoNFeCTe.ID_Empresa) " & _
            "SET  tmpDadosConexaoNFeCTe.ID_NatOp_CompraNF = [tmpPscPes].[ID_NatOper], tmpDadosConexaoNFeCTe.FiltroCFOP = [tmpPscPes].[CFOP_NatOper];"


'' #20210823_qryDadosGerais_Update_IDVD
'' #tblDadosConexaoNFeCTe.IDVD_CompraNF
Private Const qryDadosGerais_Update_IDVD As String = _
    "UPDATE tblDadosConexaoNFeCTe " & _
    "SET tblDadosConexaoNFeCTe.IDVD_CompraNF = val((Left(Trim(Replace(Replace(tblDadosConexaoNFeCTe.Obs_CompraNF, 'PEDIDO: ', ''), 'PEDIDO ', '')), 6))) " & _
    "WHERE (((Left(tblDadosConexaoNFeCTe.Obs_CompraNF, 6)) = 'PEDIDO ') " & _
    "       AND ((tblDadosConexaoNFeCTe.CNPJ_emit) = '12680452000302') " & _
    "       AND (((tblDadosConexaoNFeCTe.registroValido) = 1)  " & _
    "       AND ((tblDadosConexaoNFeCTe.registroProcessado) = 0))  " & _
    "       AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0));"


'' ########################################
'' #tblCompraNF - COMPRAS
'' ########################################
 
''' #20210823_qryUpdateNumPed_CompraNF
''' #tblCompraNF.NumPed_CompraNF
'Private Const qryUpdateNumPed_CompraNF As String = "UPDATE TblCompraNF SET TblCompraNF.NumPed_CompraNF = Format(IIf(IsNull(DLookup(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='NumPed_CompraNF'"")), '000001', DLookup(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='NumPed_CompraNF'"") + 1), '000000') " & _
'        "WHERE (((TblCompraNF.ID_CompraNF) IN (SELECT TOP 1 ID_CompraNF FROM TblCompraNF WHERE NumPed_CompraNF IS NULL ORDER BY ID_CompraNF)));"
 
'' #20210823_qryUpdateNumPed_CompraNF
'' #tblParametros.ValorDoParametro
Private Const qryUpdateNumPed_Contador As String = _
        "UPDATE tblParametros  " & _
        "SET tblParametros.ValorDoParametro = Format(IIf(IsNull(DLookup(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='NumPed_CompraNF'"")), '000001', DLookup(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='NumPed_CompraNF'"") + 1), '000000')  " & _
        "WHERE tblParametros.TipoDeParametro = ""NumPed_CompraNF"";"
  
'' IDVD_CompraNF
'' #20210823_qryCompras_Update_Dados
'' #20210823_VTotProd_CompraNF
Private Const qryCompras_Update_Dados As String = _
        "UPDATE tblDadosConexaoNFeCTe " & _
        "INNER JOIN tblCompraNF ON tblDadosConexaoNFeCTe.ChvAcesso = tblCompraNF.ChvAcesso_CompraNF " & _
        "SET tblCompraNF.CFOP_CompraNF = [tblDadosConexaoNFeCTe].[FiltroCFOP] " & _
        "   ,tblCompraNF.ID_NatOp_CompraNF = [tblDadosConexaoNFeCTe].[ID_NatOp_CompraNF] " & _
        "   ,tblCompraNF.ID_Forn_CompraNF = [tblDadosConexaoNFeCTe].[ID_Forn_CompraNF] " & _
        "   ,tblCompraNF.ModeloDoc_CompraNF = [tblDadosConexaoNFeCTe].[codMod] " & _
        "   ,tblCompraNF.Sit_CompraNF = IIF([tblDadosConexaoNFeCTe].[codMod]=57,6,IIF([tblDadosConexaoNFeCTe].[ID_Tipo]=5,6,[tblDadosConexaoNFeCTe].[Sit_CompraNF])) " & _
        "   ,tblCompraNF.Fil_CompraNF = [tblDadosConexaoNFeCTe].[ID_Empresa] " & _
        "   ,tblCompraNF.IDVD_CompraNF = IIF([tblDadosConexaoNFeCTe].[codMod]=57,null,[tblDadosConexaoNFeCTe].[IDVD_CompraNF]) " & _
        "   ,tblCompraNF.VTotNF_CompraNF = replace(Nz([tblCompraNF].[VTotNF_CompraNF], 0) / 100, "","", ""."") " & _
        "   ,tblCompraNF.VTotProd_CompraNF = replace(Nz([tblCompraNF].[VTotProd_CompraNF], 0) / 100, "","", ""."") " & _
        "   ,tblCompraNF.BaseCalcICMS_CompraNF = replace(Nz([tblCompraNF].[BaseCalcICMS_CompraNF], 0) / 100, "","", ""."") " & _
        "   ,tblCompraNF.DTEmi_CompraNF = [tblDadosConexaoNFeCTe].[dhEmi]  " & _
        "   ,tblCompraNF.HoraEntd_CompraNF = IIF([tblDadosConexaoNFeCTe].[ID_Tipo]<>5,NULL,tblCompraNF.HoraEntd_CompraNF) " & _
        "   ,tblCompraNF.DTEntd_CompraNF = IIF([tblDadosConexaoNFeCTe].[codMod]=57,Format(DATE(), ""yyyy-mm-dd""),tblCompraNF.DTEmi_CompraNF) " & _
        "   ,tblCompraNF.TPNF_CompraNF = 1 " & _
        "WHERE (((tblDadosConexaoNFeCTe.registroValido) = 1) AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1)) AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0);"

'' "   ,tblCompraNF.VTotIPI_CompraNF = 0 " & _


'' ########################################
'' #tblCompraNFItem - COMPRAS ITENS
'' ########################################

'' #20210823_ID_Prod_CompraNFItem
'' #tblCompraNFItem.InsertItensCTe
Private Const qryComprasItens_Insert_Dados_CTeItens As String = _
        "qryComprasItem_Insert"

Private Const qryComprasItens_Update_Dados_CTeItens As String = _
        "qryComprasItem_Update"

'' #20210823_qryComprasItens_Update_CFOP_CompraNF
'' #tblCompraNFItem.CFOP_CompraNFItem
Private Const qryComprasItens_Update_CFOP_CompraNF As String = _
        "UPDATE (tblDadosConexaoNFeCTe INNER JOIN tblCompraNF ON tblDadosConexaoNFeCTe.ChvAcesso = tblCompraNF.ChvAcesso_CompraNF) " & _
        "INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF " & _
        "SET tblCompraNFItem.CFOP_CompraNFItem = [tblCompraNF].[CFOP_CompraNF] " & _
        "WHERE (((tblDadosConexaoNFeCTe.registroValido) = 1) " & _
        "       AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1) " & _
        "       AND ((tblDadosConexaoNFeCTe.codMod) = 57) " & _
        "       AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0));"

'' #20210823_ID_Prod_CompraNFItem
'' #tblCompraNFItem.ID_Prod_CompraNFItem
Private Const qryComprasItens_Update_Dados_ID_Prod_CompraNFItem As String = _
        "UPDATE tblDadosConexaoNFeCTe " & _
        "INNER JOIN (tblCompraNF INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF) ON tblDadosConexaoNFeCTe.ChvAcesso = tblCompraNF.ChvAcesso_CompraNF " & _
        "SET tblCompraNFItem.ID_Prod_CompraNFItem = DLookUp(""CodigoProd_Grade"", ""tmpGradeProdutos"", ""CodigoForn_Grade='"" & [tblCompraNFItem].[ID_Prod_CompraNFItem] & ""'"") " & _
        "WHERE (((tblDadosConexaoNFeCTe.registroValido) = 1)    AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1));"
            
''#######################################################################################
''### CONSULTAS DE CONTROLE DE DADOS
''#######################################################################################

''' #20210823_qryUpdateNumPed_CompraNF
'Public Function UpdateNumPed_CompraNF()
'Dim x As Long
'Dim contador As Long: contador = DCount("*", "TblCompraNF", "NumPed_CompraNF is null")
'Dim qryProcessos() As Variant: qryProcessos = Array(qryUpdateNumPed_CompraNF, qryUpdateNumPed_Contador)
'
'    For x = 1 To contador
'        executarComandos qryProcessos
'    Next
'
'End Function

Public Function UpdateProcessamentoConcluido() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryDadosGerais_Update_ProcessamentoConcluido
    UpdateProcessamentoConcluido = qryDadosGerais_Update_ProcessamentoConcluido
                                                    
End Function

'' #20210823_qryUpdateNumPed_CompraNF
Public Function UpdateComprasItens_IDCompraNF() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryComprasItens_Update_IDCompraNF
    UpdateComprasItens_IDCompraNF = qryComprasItens_Update_IDCompraNF

End Function

Public Function SelectRegistroValidoPorcessado() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryCompras_Select_RegistroValidoPorcessado
    SelectRegistroValidoPorcessado = qryCompras_Select_RegistroValidoPorcessado
                                                    
End Function

Public Function SelectCamposNomes() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryParametros_Select_CamposNomes
    SelectCamposNomes = qryParametros_Select_CamposNomes
                                                    
End Function

Public Function SelectColetaEmpresa() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryParametros_Select_ColetaEmpresa
    SelectColetaEmpresa = qryParametros_Select_ColetaEmpresa
                                                    
End Function

Public Function SelectArquivosPendentes() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryDadosGerais_Select_ArquivosPendentes
    SelectArquivosPendentes = qryDadosGerais_Select_ArquivosPendentes
                                                    
End Function

Public Function compras_atualizarEnviadoParaServidor() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryDadosGerais_Update_ProcessamentoConcluido_Servidor
    compras_atualizarEnviadoParaServidor = qryDadosGerais_Update_ProcessamentoConcluido_Servidor

End Function

Public Sub compras_atualizarCampos()
If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "######## - compras_atualizarCampos()"
Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                    qryCompras_Update_Dados, _
                                                    qryComprasItens_Insert_Dados_CTeItens, _
                                                    qryComprasItens_Update_Dados_CTeItens, _
                                                    qryComprasItens_Update_CFOP_CompraNF, _
                                                    qryComprasItens_Update_Dados_ID_Prod_CompraNFItem, _
                                                    qryComprasCTe_Update_AjustesCampos_LOCAL, _
                                                    "qryComprasItens_Update_Almox_CompraNFItem", _
                                                    "qryComprasItem_Update_AjustesCampos", _
                                                    "qryComprasItem_Update_STs" _
                                                    ): executarComandos qryProcessos

End Sub

'' #20211127_qryDadosGerais_Update_FiltroFil_DestinatarioNaoProparts
'' #20211127_qryDadosGerais_Update_FiltroFil_DestinatarioProparts
'' #20211127_qryDadosGerais_Update_FiltroFil_DestinatarioProparts_55
'' #20211127_qryDadosGerais_Update_FiltroFil_RemetenteProparts
Public Sub TratamentoDeDadosGerais()
If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "######## - TratamentoDeDadosGerais()"
Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                qryDadosGerais_Update_FornecedoresValidos, _
                                                qryDadosGerais_Update_RegistrosValidos, _
                                                qryDadosGerais_Update_IdTipo, _
                                                qryDadosGerais_Update_RetornoArmazem, _
                                                qryDadosGerais_Update_TransferenciaSisparts, _
                                                qryDadosGerais_Update_CTe, _
                                                qryDadosGerais_Update_Sit_CompraNF, _
                                                qryDadosGerais_Update_IDVD, _
                                                qryDadosGerais_Update_IdFornCompraNF, _
                                                qryDadosGerais_Update_ProcessamentoConcluido_CTE, _
                                                qryDadosGerais_Update_RetornoArmazem_CFOP, _
                                                qryDadosGerais_Update_TransferenciaSisparts_CFOP, _
                                                "qryDadosGerais_Update_FiltroFil_DestinatarioNaoProparts", _
                                                "qryDadosGerais_Update_FiltroFil_DestinatarioProparts", _
                                                "qryDadosGerais_Update_FiltroFil_DestinatarioProparts_55", _
                                                "qryDadosGerais_Update_FiltroFil_RemetenteProparts", _
                                                qryDadosGerais_Update_ID_NatOp_CompraNF__FiltroCFOP, _
                                                "qryDadosGerais_Update_ID_NatOp_CompraNF" _
                                                ): executarComandos qryProcessos

'' qryDadosGerais_Delete_RegistrosInvalidos _

End Sub

Public Sub CriarRepositorios()
If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "######## - CriarRepositorios()"
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
