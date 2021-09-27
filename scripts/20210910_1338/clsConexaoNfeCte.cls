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
'' ### REVIS�O
''
'' compras_carregarItensCTe
'' compras_atualizarCampos
''
''
'' ### PRINCIPAIS M�TODOS
''
'' 01. carregar_DadosGerais
'' 02. ProcessamentoDeArquivo
'' 03. enviar_ComprasParaServidor
'' 00. criarArquivoJson
'' 00. ProcessarArquivosXml
'' 00. ADM_carregarDadosDoServidor
''

'' #00. CRIA��O DE ARQUIVOS JSON
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
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 1 WHERE (((tblDadosConexaoNFeCTe.registroValido) = 1) AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1) AND ((tblDadosConexaoNFeCTe.ID_Tipo) = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro] = 'Cte'"")));"


'' #tblDadosConexaoNFeCTe.registroProcessado
Private Const qryUpdateProcessamentoConcluido_ItensCompras As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 1 FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0));"


'' -- ARQUIVOS - CONSULTA PARA CRIA��O DE ARQUIVOS JSON
'' #tblDadosConexaoNFeCTe.Distinct
Private Const sqyDadosJson As String = _
        "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi]))>0));"


'' -- SELE��O DE FORNECEDORES VALIDOS
'' #tblDadosConexaoNFeCTe.qryUpdateFornecedoresValidos
Private Const qryUpdateFornecedoresValidos As String = _
        "UPDATE (SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF FROM tmpClientes) AS qryFornecedoresValidos INNER JOIN tblDadosConexaoNFeCTe ON qryFornecedoresValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit SET tblDadosConexaoNFeCTe.registroValido = 1;"


'' -- SELE��O DE REGISTROS VALIDOS.qryRegistrosValidos
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


'' #20210823_qryUpdateCFOP_PSC_PES -- FiltroCFOP
'' #tblDadosConexaoNFeCTe.FiltroCFOP
'' #tmpDadosConexaoNFeCTe.ID_NatOp_CompraNF
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
'' -- ID_TIPO - APENAS TIPOS COM ID DE VALOR ZERO(0) SER�O ATUALIZADOS PARA N�O COMPROMETER OS REGISTROS J� PROCESSADOS
'' ------------------------------------------------------------------------------------------------------------------------

'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 4 - NF-e Retorno Armaz�m
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryUpdateRetornoArmazem As String = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='RetornoArmazem'"") WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = 0) AND ((tblDadosConexaoNFeCTe.codMod) = CInt('55')) AND ((tblDadosConexaoNFeCTe.CNPJ_emit) IN ('12680452000302')));"

'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 6 - NF-e Transfer�ncia com c�digo Sisparts
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
        "UPDATE tblCompraNF SET tblCompraNF.BaseCalcICMS_CompraNF = replace(Nz([tblCompraNF].[BaseCalcICMS_CompraNF], 0) / 100,"","",""."") WHERE (((tblCompraNF.BaseCalcICMS_CompraNF)>""0"")) OR (((tblCompraNF.BaseCalcICMS_CompraNF) Is Not Null));"


'' #20210823_qryUpdate_IDVD
'' #tblCompraNF.IDVD_CompraNF
Private Const qryUpdate_IDVD As String = _
        "UPDATE tblCompraNF " & _
        "SET tblCompraNF.IDVD_CompraNF = Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6) " & _
        "WHERE (((Left([Obs_CompraNF], 6)) = 'PEDIDO ')  " & _
        "       AND ((tblCompraNF.CNPJ_CPF_CompraNF) = '12680452000302') " & _
        "       AND ((Val(Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6))) > 0));"


'' #20210823_qryUpdateID_NatOp_CompraNF
'' #tblCompraNF.CFOP_CompraNF
'' #tblCompraNF.ID_NatOp_CompraNF
'' #tblCompraNF.Sit_CompraNF
Private Const qryUpdateID_NatOp_CompraNF As String = "UPDATE tblDadosConexaoNFeCTe INNER JOIN tblCompraNF ON tblDadosConexaoNFeCTe.ChvAcesso = tblCompraNF.ChvAcesso_CompraNF " & _
            "SET tblCompraNF.CFOP_CompraNF = [tblDadosConexaoNFeCTe].[FiltroCFOP], tblCompraNF.ID_NatOp_CompraNF = [tblDadosConexaoNFeCTe].[ID_NatOp_CompraNF], tblCompraNF.Sit_CompraNF = [tblDadosConexaoNFeCTe].[Sit_CompraNF];"


'' #20210823_qryUpdateCFOP_FilCompra
'' #tblCompraNF.CFOP_CompraNF
'' #tblCompraNF.Fil_CompraNF
Private Const qryUpdateCFOP_FilCompra As String = _
        "UPDATE tblCompraNF SET tblCompraNF.CFOP_CompraNF = DLookUp(""[FiltroCFOP]"",""[tblDadosConexaoNFeCTe]"",""[ChvAcesso]='"" & [tblCompraNF].[ChvAcesso_CompraNF] & ""'""), tblCompraNF.Fil_CompraNF = DLookUp(""[ID_EMPRESA]"",""[tblDadosConexaoNFeCTe]"",""[ChvAcesso]='"" & [tblCompraNF].[ChvAcesso_CompraNF] & ""'"");"


'' #20210823_qryUpdate_ModeloDoc_CFOP
'' #tblCompraNF.ModeloDoc_CompraNF
'' #tblCompraNF.CFOP_CompraNF
Private Const qryUpdate_ModeloDoc_CFOP As String = _
        "UPDATE tblCompraNF INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso SET tblCompraNF.ModeloDoc_CompraNF = [tblDadosConexaoNFeCTe].[codMod], tblCompraNF.CFOP_CompraNF = [tblDadosConexaoNFeCTe].[FiltroCFOP] WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)>0));"


'' #20210823_qryUpdateFilCompraNF
'' #tblCompraNF.Fil_CompraNF
Private Const qryUpdateFilCompraNF As String = _
        "UPDATE (SELECT tmpEmpresa.ID_Empresa, STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF, tmpEmpresa.CNPJ_Empresa FROM tmpEmpresa WHERE (((tmpEmpresa.CNPJ_Empresa) Is Not Null))) AS qryEmpresas INNER JOIN tblCompraNF ON qryEmpresas.strCNPJ_CPF = tblCompraNF.CNPJ_CPF_CompraNF SET tblCompraNF.Fil_CompraNF = qryEmpresas.ID_Empresa;"


'' #20210823_qryUpdateIdFornCompraNF
'' #tblCompraNF.ID_Forn_CompraNF
Private Const qryUpdateIdFornCompraNF As String = _
        "UPDATE (SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF, tmpClientes.C�DIGOClientes FROM tmpClientes) AS qryClientesFornecedor " & _
        " INNER JOIN tblCompraNF ON tblCompraNF.CNPJ_CPF_CompraNF = qryClientesFornecedor.strCNPJ_CPF " & _
        " SET tblCompraNF.ID_Forn_CompraNF = qryClientesFornecedor.C�DIGOClientes;"



'' ########################################
'' #tblCompraNFItem - COMPRAS ITENS
'' ########################################


''' UPDATE - ID_Prod_CompraNFItem
'' #tblCompraNFItem.ID_Prod_CompraNFItem
Private Const qryUpdateItens_ID_Prod_CompraNFItem As String = _
        "UPDATE tblCompraNFItem SET tblCompraNFItem.ID_Prod_CompraNFItem = DLookUp(""CodigoProd_Grade"",""dbo_tabGradeProdutos"",""CodigoForn_Grade='"" & [tblCompraNFItem].[ID_Prod_CompraNFItem] & ""'"");"


'' #tblCompraNFItem.InsertItensCTe
Private Const qryInsertItensCTe As String = _
        "INSERT INTO tblCompraNFItem (  ChvAcesso_CompraNF  ,VUnt_CompraNFItem  ,Num_CompraNFItem   ,VTot_CompraNFItem  ,DebICMS_CompraNFItem " & _
        "   ,VTotBaseCalcICMS_CompraNFItem  ,ID_NatOp_CompraNFItem  ,Item_CompraNFItem  ,ID_Grade_CompraNFItem  ,QtdFat_CompraNFItem " & _
        "   ,IPI_CompraNFItem   ,FlagEst_CompraNFItem   ,BaseCalcICMS_CompraNFItem  ) " & _
        "SELECT tblCompraNF.ChvAcesso_CompraNF  ,tblCompraNF.VTotNF_CompraNF AS VUnt_CompraNFItem   ,tblCompraNF.NumNF_CompraNF " & _
        "   ,tblCompraNF.VTotNF_CompraNF    ,IIf([VTotICMS_CompraNF] <> """", Replace(Nz([tblCompraNF].[BaseCalcICMS_CompraNF], 0) / 100, "","", "".""), 0) AS strVTotICMS " & _
        "   ,tblCompraNF.BaseCalcICMS_CompraNF  ,tblCompraNF.ID_NatOp_CompraNF  ,1 AS strItem   ,1 AS strIDGrade    ,1 AS strQtdFat ,0 AS strIPI    ,0 AS strFlag   ,100 AS strBaseCalcICMS  " & _
        "FROM tblCompraNF " & _
        "INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso " & _
        "WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='Cte'"")));"


'' #20210823_qryUpdateNumPed_CompraNF
'' #tblCompraNF.NumPed_CompraNF
Private Const qryUpdateNumPed_CompraNF As String = "UPDATE TblCompraNF SET TblCompraNF.NumPed_CompraNF = Format(IIf(IsNull(DMax('NumPed_CompraNF', 'TblCompraNF')), '000001', DMax('NumPed_CompraNF', 'TblCompraNF') + 1), '000000') " & _
        "WHERE (((TblCompraNF.ID_CompraNF) IN (SELECT TOP 1 ID_CompraNF FROM TblCompraNF WHERE NumPed_CompraNF IS NULL ORDER BY ID_CompraNF)));"


'' #20210823_qryUpdateSit_CompraNF
'' #tblDadosConexaoNFeCTe.Sit_CompraNF
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


''#######################################################################################
''### CONSULTAS DE CONTROLE DE DADOS
''#######################################################################################


'Private Sub TratamentoDeCompras()
'Dim qryProcessos() As Variant: qryProcessos = Array(qryUpdateFilCompraNF): executarComandos qryProcessos
'
'End Sub

'' #20210823_UpdateNumPed_CompraNF
Public Function UpdateNumPed_CompraNF()
Dim x As Long
Dim contador As Long: contador = DCount("*", "TblCompraNF", "NumPed_CompraNF is null")

    For x = 1 To contador
        Application.CurrentDb.Execute qryUpdateNumPed_CompraNF
    Next

End Function


Public Function UpdateProcessamentoConcluido() As String
    UpdateProcessamentoConcluido = qryUpdateProcessamentoConcluido
                                                    
End Function

Public Function SelectArquivosPendentes() As String
    SelectArquivosPendentes = qrySelecaoDeArquivosPendentes
                                                    
End Function

Private Sub compras_atualizarItensCompras()
    Application.CurrentDb.Execute qryUpdateProcessamentoConcluido_ItensCompras

End Sub

'' #AILTON - VALIDAR
'' ATUALIZAR CAMPOS DE COMPRAS
Public Sub compras_atualizarCampos()
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
Public Sub compras_carregarItensCTe()
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
                                                qryUpdateCFOP_PSC_PES, _
                                                qryUpdateID_NatOp_CompraNF, _
                                                qryUpdateSit_CompraNF): executarComandos qryProcessos

End Sub









