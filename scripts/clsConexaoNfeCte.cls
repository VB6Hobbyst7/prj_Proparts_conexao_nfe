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

'' 35210320147617002276570010055234421994476558

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
'' DESCONTINUADO - ( PODEMOS USAR DEPOIS )
'' #tblDadosConexaoNFeCTe.registroValido
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

Private Const qryUpdateIdEmpresa_TransferenciaEntreFiliais As String = _
        "UPDATE (SELECT tmpEmpresa.ID_Empresa, STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF FROM tmpEmpresa) AS qryEmpresas " & _
        " INNER JOIN tblDadosConexaoNFeCTe ON qryEmpresas.strCNPJ_CPF = tblDadosConexaoNFeCTe.CPNJ_Dest " & _
        " SET tblDadosConexaoNFeCTe.ID_Empresa = qryEmpresas.ID_Empresa" & _
        " WHERE (((tblDadosConexaoNFeCTe.CFOP) = '6152'));"


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
        "UPDATE (Select ValorDoParametro,TipoDeParametro from tblParametros where TipoDeParametro = 'RetornoArmazem') as tmpRetornoArmazem INNER JOIN ((select top 1 cInt('55') as strMod, '12680452000302' as strCNPJ_CPF, 'RetornoArmazem' as strTipoDeParametro from tblParametros)  AS qryRetornoArmazem INNER JOIN tblDadosConexaoNFeCTe ON (qryRetornoArmazem.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit) AND (qryRetornoArmazem.strMod = tblDadosConexaoNFeCTe.codMod)) ON (tmpRetornoArmazem.TipoDeParametro = qryRetornoArmazem.strTipoDeParametro) AND (tmpRetornoArmazem.TipoDeParametro = qryRetornoArmazem.strTipoDeParametro) SET tblDadosConexaoNFeCTe.ID_Tipo = [tmpRetornoArmazem].[ValorDoParametro] WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)=0))  AND (((tblDadosConexaoNFeCTe.CFOP) = '5907'));"


'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 6 - NF-e Transferência com código Sisparts
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryUpdateTransferenciaSisparts As String = _
        "UPDATE (Select ValorDoParametro,TipoDeParametro from tblParametros where TipoDeParametro = 'TransferenciaSisparts') as tmpTransferenciaSisparts INNER JOIN ((SELECT CInt('55') AS strMod, 'TransferenciaSisparts' AS strTipoDeParametro, tmpEmpresa.CNPJ_Empresa AS strCNPJ_CPF FROM tmpEmpresa)  AS qryTransferenciaSisparts INNER JOIN tblDadosConexaoNFeCTe ON (qryTransferenciaSisparts.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit) AND (qryTransferenciaSisparts.strMod = tblDadosConexaoNFeCTe.codMod)) ON (tmpTransferenciaSisparts.TipoDeParametro = qryTransferenciaSisparts.strTipoDeParametro) AND (tmpTransferenciaSisparts.TipoDeParametro = qryTransferenciaSisparts.strTipoDeParametro) SET tblDadosConexaoNFeCTe.ID_Tipo = [tmpTransferenciaSisparts].[ValorDoParametro] WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)=0));"


'' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 0 - CT-e
'' #tblDadosConexaoNFeCTe.ID_Tipo
Private Const qryUpdateCTe As String = _
        "UPDATE (SELECT ValorDoParametro,TipoDeParametro FROM tblParametros WHERE TipoDeParametro = 'CTe') AS tmpCTe INNER JOIN ((SELECT TOP 1 cInt('57') AS strMod,'CTe' AS strTipoDeParametro FROM tblParametros) AS qryCTe INNER JOIN tblDadosConexaoNFeCTe ON (qryCTe.strMod = tblDadosConexaoNFeCTe.codMod)) ON (tmpCTe.TipoDeParametro = qryCTe.strTipoDeParametro) AND (tmpCTe.TipoDeParametro = qryCTe.strTipoDeParametro) SET tblDadosConexaoNFeCTe.ID_Tipo = [tmpCTe].[ValorDoParametro] WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = 0));"


'' ########################################
'' #TratamentoDeRegistros - COMPRAS
'' ########################################

Private Const qryUpdateBaseCalcICMS As String = _
    "UPDATE tblCompraNF SET tblCompraNF.BaseCalcICMS_CompraNF = replace([tblCompraNF].[BaseCalcICMS_CompraNF] / 100,"","",""."") WHERE (((tblCompraNF.BaseCalcICMS_CompraNF)>""0"")) OR (((tblCompraNF.BaseCalcICMS_CompraNF) Is Not Null));"
    ''"UPDATE tblCompraNF SET tblCompraNF.BaseCalcICMS_CompraNF = [tblCompraNF].[BaseCalcICMS_CompraNF]/100 WHERE (((tblCompraNF.BaseCalcICMS_CompraNF)>""0"")) OR (((tblCompraNF.BaseCalcICMS_CompraNF) Is Not Null));"

Private Const qryUpdate_IDVD As String = _
                "UPDATE tblCompraNF " & _
                "SET tblCompraNF.IDVD_CompraNF = Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6) " & _
                "WHERE (((Left([Obs_CompraNF], 6)) = 'PEDIDO ')  " & _
                "       AND ((tblCompraNF.CNPJ_CPF_CompraNF) = '12680452000302') " & _
                "       AND ((Val(Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6))) > 0));"

Private Const qryUpdateCFOP_FilCompra As String = _
    "UPDATE tblCompraNF SET tblCompraNF.CFOP_CompraNF = DLookUp(""[FiltroCFOP]"",""[tblDadosConexaoNFeCTe]"",""[ChvAcesso]='"" & [tblCompraNF].[ChvAcesso_CompraNF] & ""'""), tblCompraNF.Fil_CompraNF = DLookUp(""[ID_EMPRESA]"",""[tblDadosConexaoNFeCTe]"",""[ChvAcesso]='"" & [tblCompraNF].[ChvAcesso_CompraNF] & ""'"");"

Private Const qryUpdateFilCompraNF As String = _
        "UPDATE (SELECT tmpEmpresa.ID_Empresa, STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF, tmpEmpresa.CNPJ_Empresa FROM tmpEmpresa WHERE (((tmpEmpresa.CNPJ_Empresa) Is Not Null))) AS qryEmpresas INNER JOIN tblCompraNF ON qryEmpresas.strCNPJ_CPF = tblCompraNF.CNPJ_CPF_CompraNF SET tblCompraNF.Fil_CompraNF = qryEmpresas.ID_Empresa;"

Private Const qryUpdateIdFornCompraNF As String = _
        "UPDATE (SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF, tmpClientes.CÓDIGOClientes FROM tmpClientes) AS qryClientesFornecedor " & _
        " INNER JOIN tblCompraNF ON tblCompraNF.CNPJ_CPF_CompraNF = qryClientesFornecedor.strCNPJ_CPF " & _
        " SET tblCompraNF.ID_Forn_CompraNF = qryClientesFornecedor.CÓDIGOClientes;"

Private Const qryUpdate_ModeloDoc_CFOP As String = _
        "UPDATE tblCompraNF INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso SET tblCompraNF.ModeloDoc_CompraNF = [tblDadosConexaoNFeCTe].[codMod], tblCompraNF.CFOP_CompraNF = [tblDadosConexaoNFeCTe].[FiltroCFOP] WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)>0));"
        ''"UPDATE tblCompraNF INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.NumNF_CompraNF = tblDadosConexaoNFeCTe.NumNF SET tblCompraNF.ModeloDoc_CompraNF = [tblDadosConexaoNFeCTe].[codMod], tblCompraNF.CFOP_CompraNF = [tblDadosConexaoNFeCTe].[FiltroCFOP];"

Private Const qryInsertItensCTe As String = _
    "INSERT INTO tblCompraNFItem ( ChvAcesso_CompraNF, VUnt_CompraNFItem, DTEmi_CompraNFItem, Num_CompraNFItem, VTot_CompraNFItem, DebICMS_CompraNFItem, VTotBaseCalcICMS_CompraNFItem, ID_NatOp_CompraNFItem, Item_CompraNFItem, ID_Grade_CompraNFItem, QtdFat_CompraNFItem, ISS_CompraNFItem, IPI_CompraNFItem, FlagEst_CompraNFItem, BaseCalcICMS_CompraNFItem, IseICMS_CompraNFItem, OutICMS_CompraNFItem, BaseCalcIPI_CompraNFItem, DebIPI_CompraNFItem, IseIPI_CompraNFItem, OutIPI_CompraNFItem, TxMLSubsTrib_CompraNFItem, TxIntSubsTrib_CompraNFItem, TxExtSubsTrib_CompraNFItem, BaseCalcICMSSubsTrib_CompraNFItem, VTotICMSSubsTrib_compranfitem, VTotFrete_CompraNFItem, VTotDesc_CompraNFItem, VTotPIS_CompraNFItem, VTotBaseCalcPIS_CompraNFItem, VTotCOFINS_CompraNFItem, VTotIseICMS_CompraNFItem, VTotBaseCalcCOFINS_CompraNFItem, SNCredICMS_CompraNFItem, VTotSNCredICMS_CompraNFItem, VTotSeg_CompraNFItem, VTotOutDesp_CompraNFItem )  " & _
    "SELECT tblCompraNF.ChvAcesso_CompraNF, tblCompraNF.VTotNF_CompraNF, tblCompraNF.DTEmi_CompraNF, tblCompraNF.NumNF_CompraNF, tblCompraNF.VTotNF_CompraNF, IIf([VTotICMS_CompraNF]<>"""",replace(tblCompraNF.BaseCalcICMS_CompraNF/100,"","","".""),0) AS strVTotICMS, tblCompraNF.BaseCalcICMS_CompraNF, tblCompraNF.ID_NatOp_CompraNF, 1 AS strItem, 1 AS strIDGrade, 1 AS strQtdFat, 0 AS strISS, 0 AS strIPI, 0 AS strFlag, 100 AS strBaseCalcICMS, 0 AS strIseICMS, 0 AS strOutICMS, 0 AS strBaseCalcIPI, 0 AS strDebIPI, 0 AS strIseIPI, 0 AS strOutIPI, 0 AS strTxMLSubsTrib, 0 AS strTxIntSubsTrib, 0 AS strTxExtSubsTrib, 0 AS strBaseCalcICMSSubsTrib, 0 AS strVTotICMSSubsTrib, 0 AS strVTotFrete, 0 AS strVTotDesc, 0 AS strVTotPIS, 0 AS strVTotBaseCalcPIS, 0 AS strVTotCOFINS, 0 AS strVTotIseICMS, 0 AS strVTotBaseCalcCOFINS, 0 AS strSNCredICMS, 0 AS strVTotSNCredICMS, 0 AS strVTotSeg, 0 AS strVTotOutDesp  " & _
    "FROM tblCompraNF INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso  " & _
    "WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)=DLookUp(""[ValorDoParametro]"",""[tblParametros]"",""[TipoDeParametro]='Cte'""))); "


''#######################################################################################
''### MÉTODOS
''#######################################################################################

Public Sub compras_atualizarCampos()
Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                    qryUpdateCFOP_FilCompra, _
                                                    qryUpdate_IDVD, _
                                                    qryUpdateBaseCalcICMS, _
                                                    qryUpdateIdFornCompraNF, _
                                                    qryUpdate_ModeloDoc_CFOP, _
                                                    qryInsertItensCTe)
    '' ATUALIZAR CAMPOS
    executarComandos qryProcessos

End Sub

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
                                                qryUpdateIdTipo, _
                                                qryUpdateRetornoArmazem, _
                                                qryUpdateTransferenciaSisparts, _
                                                qryUpdateCTe, _
                                                qryUpdateIdEmpresa, _
                                                qryUpdateIdEmpresa_TransferenciaEntreFiliais, _
                                                qryUpdateCFOP_PSC_PES)

    '' REFINAMENTO DE DADOS
    executarComandos qryProcessos

End Sub

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
        strPk = IIf(IsNull(DLookup("[chave]", "[tblDadosConexaoNFeCTe]", "[CaminhoDoArquivo]='" & fileName & "'")), Controle & getFileName(CStr(fileName)), DLookup("[chave]", "[tblDadosConexaoNFeCTe]", "[CaminhoDoArquivo]='" & fileName & "'"))
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


'' #TRANSFERIR DADOS PROCESSADOS PARA A TABELA DE TRANSITO - ( tblProcessamento -> tblDadosConexaoNFeCTe/tblCompraNF/tblCompraNFItem )
Public Sub TransferirDadosProcessados(pDestino As String)

'' CONTROLE DE CADASTRO DE REGISTRO
Dim t As Integer
Dim tmpSqlOrigem As String

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

''' #VALIDAR_INFORMACAO
'Debug.Print "--------------------------------"

Do While Not rstPendentes.EOF

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdUpdateMeter, contadorDeRegistros

    '' listar itens de registro para cadastro
    tmpSqlOrigem = "Select * from tblProcessamento where NomeTabela = '" & pDestino & "' and pk = '" & rstPendentes.Fields("pk").value & "' and len(formatacao)>0 Order by ID "
        
'    '' #VALIDAR_INFORMACAO
'    Debug.Print tmpSqlOrigem
    
    Set rstOrigem = db.OpenRecordset(tmpSqlOrigem)
    Do While Not rstOrigem.EOF
    
        '' CONTROLE DE CADASTRO
        If t = 0 Then rstDestino.AddNew: t = 1
        
'        '' #VALIDAR_INFORMACAO
'        Debug.Print rstOrigem.Fields("NomeCampo").value
'        Debug.Print rstOrigem.Fields("Valor").value
        
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


'' CRIAÇÃO DE ARQUIVOS JSON
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


