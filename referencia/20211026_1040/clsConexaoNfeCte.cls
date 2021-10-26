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
        '"Select * from tblCompraNF Where tblCompraNF.ChvAcesso_CompraNF IN (SELECT ChvAcesso FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND (tblDadosConexaoNFeCTe.registroProcessado)=1))"
        
Private Const qryCompras_Update_RegistroConcluida As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 2 WHERE (((tblDadosConexaoNFeCTe.ChvAcesso)=""strChave"") AND ((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1));"

Private Const qryComprasItens_Update_IDCompraNF As String = _
        "UPDATE tblCompraNFItem SET tblCompraNFItem.ID_CompraNF_CompraNFItem = strID_Compra WHERE  tblCompraNFItem.ChvAcesso_CompraNF = 'strChave'"


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


'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 6 - NF-e Transferência com código Sisparts -> ( ID_Tipo: 7 )
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryDadosGerais_Update_TransferenciaSisparts As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='TransferenciaSisparts'"") WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = 0) AND ((tblDadosConexaoNFeCTe.codMod) = CInt('55')) AND ((tblDadosConexaoNFeCTe.CNPJ_emit) IN (SELECT CNPJ_Empresa FROM [tmpEmpresa])));"


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
'' #tblDadosConexaoNFeCTe.qryDeleteRegistrosInvalidos
'Private Const qryDadosGerais_Update_RegistrosInvalidos As String = _
'        "DELETE * FROM tblDadosConexaoNFeCTe WHERE tblDadosConexaoNFeCTe.registroValido = 0;"

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

'' MOVER REGISTRO PARA PASTA DE PROCESSADO
'' #tblDadosConexaoNFeCTe.registroProcessado = 3 - Arquivo(s) movido(s) para pasta de processados
Private Const qryDadosGerais_Update_RegistroProcessado As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 3 " & _
        ",tblDadosConexaoNFeCTe.CaminhoDoArquivo = Replace(DLookUp(""[ValorDoParametro]"",""[tblParametros]"",""[TipoDeParametro]='caminhoDeColetaProcessados'""),""empresa"",strSplit(getPath([tblDadosConexaoNFeCTe].[CaminhoDoArquivo]),""\"",2)) & getFileNameAndExt([tblDadosConexaoNFeCTe].[CaminhoDoArquivo])" & _
        " WHERE tblDadosConexaoNFeCTe.ID = strChave;"

'' MOVER REGISTRO PARA PASTA DE EXPURGO
'' #tblDadosConexaoNFeCTe.registroProcessado = 4 - Arquivo(s) movido(s) para pasta de expurgo
Private Const qryDadosGerais_Update_RegistroExpurgo As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 4  " & _
        ",tblDadosConexaoNFeCTe.CaminhoDoArquivo = Replace(DLookUp(""[ValorDoParametro]"",""[tblParametros]"",""[TipoDeParametro]='caminhoDeColetaExpurgo'""),""empresa"",strSplit(getPath([tblDadosConexaoNFeCTe].[CaminhoDoArquivo]),""\"",2)) & getFileNameAndExt([tblDadosConexaoNFeCTe].[CaminhoDoArquivo])" & _
        " WHERE tblDadosConexaoNFeCTe.ID = strChave ;"

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


'' ARQUIVOS VALIDOS
Private Const qryDadosGerais_Select_ArquivosValidos As String = _
        "SELECT ID AS ChvAcesso " & _
        "   ,CaminhoDoArquivo AS strOrigem   " & _
        "   ,Replace(DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='caminhoDeColetaProcessados'""), ""empresa"", strSplit(getPath([tblDadosConexaoNFeCTe].[CaminhoDoArquivo]), ""\"", 2)) & getFileNameAndExt([tblDadosConexaoNFeCTe].[CaminhoDoArquivo]) AS strDestino " & _
        "FROM tblDadosConexaoNFeCTe WHERE ((NOT (tblDadosConexaoNFeCTe.CaminhoDoArquivo) IS NULL) AND ((tblDadosConexaoNFeCTe.registroValido) = 1)  AND ((tblDadosConexaoNFeCTe.registroProcessado) = 2));"


'' ARQUIVOS INVALIDOS
Private Const qryDadosGerais_Select_ArquivosInvalidosNaoProcessado As String = _
    "SELECT ID AS ChvAcesso " & _
    "   ,CaminhoDoArquivo AS strOrigem   " & _
    "   ,Replace(DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='caminhoDeColetaExpurgo'""), ""empresa"", strSplit(getPath([tblDadosConexaoNFeCTe].[CaminhoDoArquivo]), ""\"", 2)) & getFileNameAndExt([tblDadosConexaoNFeCTe].[CaminhoDoArquivo]) AS strDestino " & _
    "FROM  tblDadosConexaoNFeCTe " & _
    " WHERE ((NOT (tblDadosConexaoNFeCTe.CaminhoDoArquivo) IS NULL) AND ((tblDadosConexaoNFeCTe.registroValido) = 1) AND ((tblDadosConexaoNFeCTe.registroProcessado) = 0));"


'' tblDadosConexaoNFeCTe.Pendentes
Private Const qryDadosGerais_Select_ArquivosPendentes As String = _
    "SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0)) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0) ORDER BY tblDadosConexaoNFeCTe.CaminhoDoArquivo;"


''' -- ARQUIVOS - CONSULTA PARA CRIAÇÃO DE ARQUIVOS JSON
''' #tblDadosConexaoNFeCTe.Distinct
'Private Const sqyDadosJson As String = _
'        "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi]))>0));"


'' -- FiltroFil
'' #tblDadosConexaoNFeCTe.ID_Empresa
'' -- CTE
Private Const qryDadosGerais_Update_IdEmpresa As String = _
        "UPDATE (SELECT tmpEmpresa.ID_Empresa, STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF FROM tmpEmpresa) AS qryEmpresas " & _
        " INNER JOIN tblDadosConexaoNFeCTe ON qryEmpresas.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_Rem " & _
        " SET tblDadosConexaoNFeCTe.ID_Empresa = qryEmpresas.ID_Empresa;"


'' #tblDadosConexaoNFeCTe.ID_Empresa
Private Const qryDadosGerais_Update_IdEmpresa_TransferenciaEntreFiliais As String = _
        "UPDATE (SELECT tmpEmpresa.ID_Empresa, STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF FROM tmpEmpresa) AS qryEmpresas " & _
        " INNER JOIN tblDadosConexaoNFeCTe ON qryEmpresas.strCNPJ_CPF = tblDadosConexaoNFeCTe.CPNJ_Dest " & _
        " SET tblDadosConexaoNFeCTe.ID_Empresa = qryEmpresas.ID_Empresa" & _
        " WHERE (((tblDadosConexaoNFeCTe.CFOP) = '6152'));"


'' #20210823_qryDadosGerais_Update_ID_NatOp_CompraNF__FiltroCFOP -- FiltroCFOP
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
            "               TipoDeParametro='FiltroFil' And strSplit(ValorDoParametro,'|',0) In ('PSC','PES'))  AS qryPscPes  " & _
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
 
'' #20210823_qryUpdateNumPed_CompraNF
'' #tblCompraNF.NumPed_CompraNF
Private Const qryUpdateNumPed_CompraNF As String = "UPDATE TblCompraNF SET TblCompraNF.NumPed_CompraNF = Format(IIf(IsNull(DLookup(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='NumPed_CompraNF'"")), '000001', DLookup(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='NumPed_CompraNF'"") + 1), '000000') " & _
        "WHERE (((TblCompraNF.ID_CompraNF) IN (SELECT TOP 1 ID_CompraNF FROM TblCompraNF WHERE NumPed_CompraNF IS NULL ORDER BY ID_CompraNF)));"
 

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
        "   ,tblCompraNF.Sit_CompraNF = [tblDadosConexaoNFeCTe].[Sit_CompraNF] " & _
        "   ,tblCompraNF.Fil_CompraNF = [tblDadosConexaoNFeCTe].[ID_Empresa] " & _
        "   ,tblCompraNF.IDVD_CompraNF = [tblDadosConexaoNFeCTe].[IDVD_CompraNF] " & _
        "   ,tblCompraNF.VTotNF_CompraNF = replace(Nz([tblCompraNF].[VTotNF_CompraNF], 0) / 100, "","", ""."") " & _
        "   ,tblCompraNF.VTotProd_CompraNF = replace(Nz([tblCompraNF].[VTotNF_CompraNF], 0) / 100, "","", ""."") " & _
        "   ,tblCompraNF.BaseCalcICMS_CompraNF = replace(Nz([tblCompraNF].[BaseCalcICMS_CompraNF], 0) / 100, "","", ""."") " & _
        "   ,tblCompraNF.HoraEntd_CompraNF = NULL " & _
        "   ,tblCompraNF.DTEntd_CompraNF = IIF([tblDadosConexaoNFeCTe].[codMod]=57,Format(DATE(), ""yyyy-mm-dd""),tblCompraNF.DTEntd_CompraNF) " & _
        "   ,tblCompraNF.TPNF_CompraNF = 1 " & _
        "   ,tblCompraNF.VTotICMS_CompraNF = 0 " & _
        "   ,tblCompraNF.VTotIPI_CompraNF = 0 " & _
        "WHERE (((tblDadosConexaoNFeCTe.registroValido) = 1) AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1)) AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0);"


'' ########################################
'' #tblCompraNFItem - COMPRAS ITENS
'' ########################################

'' #20210823_ID_Prod_CompraNFItem
'' #tblCompraNFItem.InsertItensCTe
Private Const qryComprasItens_Insert_Dados_CTeItens As String = _
        "INSERT INTO tblCompraNFItem (ChvAcesso_CompraNF, Num_CompraNFItem, DebICMS_CompraNFItem, VTotBaseCalcICMS_CompraNFItem, ID_NatOp_CompraNFItem " & _
        "   ,Item_CompraNFItem, ID_Grade_CompraNFItem, QtdFat_CompraNFItem, IPI_CompraNFItem, FlagEst_CompraNFItem " & _
        "   ,BaseCalcICMS_CompraNFItem, IseICMS_CompraNFItem, BaseCalcIPI_CompraNFItem, DebIPI_CompraNFItem, IseIPI_CompraNFItem " & _
        "   ,TxMLSubsTrib_CompraNFItem, BaseCalcICMSSubsTrib_CompraNFItem, VTotICMSSubsTrib_compranfitem, VTotDesc_CompraNFItem, VTotFrete_CompraNFItem " & _
        "   ,VTotPIS_CompraNFItem, VTotBaseCalcPIS_CompraNFItem, VTotCOFINS_CompraNFItem, VTotIseICMS_CompraNFItem, VTotBaseCalcCOFINS_CompraNFItem " & _
        "   ,VTotSNCredICMS_CompraNFItem, VTotSeg_CompraNFItem, VTotOutDesp_CompraNFItem, VUnt_CompraNFItem, VTot_CompraNFItem, ID_Prod_CompraNFItem) " & _
        "SELECT tblCompraNF.ChvAcesso_CompraNF, tblCompraNF.NumNF_CompraNF,IIf([VTotICMS_CompraNF] <> """", Replace(Nz([tblCompraNF].[BaseCalcICMS_CompraNF], 0) / 100, "","", "".""), 0) AS strVTotICMS " & _
        "   ,tblCompraNF.BaseCalcICMS_CompraNF, tblCompraNF.ID_NatOp_CompraNF, 1 AS str_Item_CompraNFItem, 1 AS str_ID_Grade_CompraNFItem, 1 AS str_QtdFat_CompraNFItem " & _
        "   ,0 AS str_IPI_CompraNFItem, 0 AS str_FlagEst_CompraNFItem, 100 AS str_BaseCalcICMS_CompraNFItem, 0 AS str_IseICMS_CompraNFItem, 0 AS str_BaseCalcIPI_CompraNFItem " & _
        "   ,0 AS str_DebIPI_CompraNFItem, 0 AS str_IseIPI_CompraNFItem, 0 AS str_TxMLSubsTrib_CompraNFItem, 0 AS str_BaseCalcICMSSubsTrib_CompraNFItem, 0 AS str_VTotICMSSubsTrib_compranfitem " & _
        "   ,0 AS str_VTotDesc_CompraNFItem, 0 AS str_VTotFrete_CompraNFItem, 0 AS str_VTotPIS_CompraNFItem, 0 AS str_VTotBaseCalcPIS_CompraNFItem, 0 AS str_VTotCOFINS_CompraNFItem " & _
        "   ,0 AS str_VTotIseICMS_CompraNFItem, 0 AS str_VTotBaseCalcCOFINS_CompraNFItem, 0 AS str_VTotSNCredICMS_CompraNFItem, 0 AS str_VTotSeg_CompraNFItem, 0 AS str_VTotOutDesp_CompraNFItem " & _
        "   ,[tblCompraNF].[VTotNF_CompraNF] AS str_VUnt_CompraNFItem, [tblCompraNF].[VTotNF_CompraNF] AS str_VTot_CompraNFItem " & _
        "   ,DLookUp(""[Código]"", ""[tmpProdutos]"", ""[Modelo] = 'TRANSPORTE'"") AS str_ID_Prod_CompraNFItem " & _
        "FROM tblCompraNF " & _
        "INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso " & _
        "WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='Cte'"")) " & _
        "   AND ((tblDadosConexaoNFeCTe.registroValido) = 1) AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1));"


'' #20210823_qryComprasItens_Update_CFOP_CompraNF
'' #tblCompraNFItem.CFOP_CompraNFItem
Private Const qryComprasItens_Update_CFOP_CompraNF As String = _
        "UPDATE (tblDadosConexaoNFeCTe INNER JOIN tblCompraNF ON tblDadosConexaoNFeCTe.ChvAcesso = tblCompraNF.ChvAcesso_CompraNF) " & _
        "INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF " & _
        "SET tblCompraNFItem.CFOP_CompraNFItem = [tblCompraNF].[CFOP_CompraNF] " & _
        "WHERE (((tblDadosConexaoNFeCTe.registroValido) = 1) " & _
        "       AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1) " & _
        "       AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0));"

'' #20210823_ID_Prod_CompraNFItem
''' UPDATE - ID_Prod_CompraNFItem
'' #tblCompraNFItem.ID_Prod_CompraNFItem
Private Const qryComprasItens_Update_Dados_ID_Prod_CompraNFItem As String = _
        "UPDATE tblDadosConexaoNFeCTe " & _
        "INNER JOIN (tblCompraNF INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF) ON tblDadosConexaoNFeCTe.ChvAcesso = tblCompraNF.ChvAcesso_CompraNF " & _
        "SET tblCompraNFItem.ID_Prod_CompraNFItem = DLookUp(""CodigoProd_Grade"", ""tmpGradeProdutos"", ""CodigoForn_Grade='"" & [tblCompraNFItem].[ID_Prod_CompraNFItem] & ""'"") " & _
        "WHERE (((tblDadosConexaoNFeCTe.registroValido) = 1)    AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1));"



'Private Const qrySelectRepositorioDeRelacao As String = _
'            "Select * from tmpCompras_ID_CompraNF"


'Private Const qrySelectRepositorioDeCompras As String = _
'            "SELECT * FROM tblCompraNF ORDER BY ID_CompraNF"
            
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
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryDadosGerais_Update_ProcessamentoConcluido
    UpdateProcessamentoConcluido = qryDadosGerais_Update_ProcessamentoConcluido
                                                    
End Function

Public Function UpdateExpurgo() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryDadosGerais_Update_RegistroExpurgo
    UpdateExpurgo = qryDadosGerais_Update_RegistroExpurgo
                                                    
End Function

Public Function UpdateProcessado() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryDadosGerais_Update_RegistroProcessado
    UpdateProcessado = qryDadosGerais_Update_RegistroProcessado
                                                    
End Function

'Public Function SelectRepositorioDeCompras() As String
'    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qrySelectRepositorioDeCompras
'    SelectRepositorioDeCompras = qrySelectRepositorioDeCompras
'
'End Function
'
'Public Function SelectRepositorioDeRelacao() As String
'    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qrySelectRepositorioDeRelacao
'    SelectRepositorioDeRelacao = qrySelectRepositorioDeRelacao
'
'End Function


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


Public Function SelectArquivosInvalidos() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryDadosGerais_Select_ArquivosInvalidosNaoProcessado
    SelectArquivosInvalidos = qryDadosGerais_Select_ArquivosInvalidosNaoProcessado
                                                    
End Function

Public Function SelectArquivosValidos() As String
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryDadosGerais_Select_ArquivosValidos
    SelectArquivosValidos = qryDadosGerais_Select_ArquivosValidos
                                                    
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


'' #AILTON - VALIDAR
'' ATUALIZAR CAMPOS DE COMPRAS
Public Sub compras_atualizarCampos()
If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "######## - compras_atualizarCampos()"
Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                    qryCompras_Update_Dados, _
                                                    qryComprasItens_Insert_Dados_CTeItens, _
                                                    qryComprasItens_Update_CFOP_CompraNF, _
                                                    qryComprasItens_Update_Dados_ID_Prod_CompraNFItem, _
                                                    qryDadosGerais_Update_ProcessamentoConcluido_CTE): executarComandos qryProcessos
                                                    
End Sub

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
                                                qryDadosGerais_Update_IdEmpresa, _
                                                qryDadosGerais_Update_IdEmpresa_TransferenciaEntreFiliais, _
                                                qryDadosGerais_Update_IDVD, _
                                                qryDadosGerais_Update_IdFornCompraNF, _
                                                qryDadosGerais_Update_ID_NatOp_CompraNF__FiltroCFOP): executarComandos qryProcessos

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


'' REPOSITORIO LOCAL PARA RELACIONAR ID_COMPRA COM CHV_ACESSO
Public Sub criarTabelaTemporariaParaRelacionarIdCompraComChvAcesso()
If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "######## - criarTabelaTemporariaParaRelacionarIdCompraComChvAcesso()"

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

