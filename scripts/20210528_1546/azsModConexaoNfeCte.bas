Attribute VB_Name = "azsModConexaoNfeCte"
Option Compare Database

''----------------------------

'' #AILTON - VALIDAR
'' #ARQUIVOS - GERAR ARQUIVOS | PROCESSAMENTO POR ARQUIVO(S)

'' #05_XML_ICMS                         - REVISÃO / FERNANDA
'' #05_XML_ICMS_Orig                    - REVISÃO / FERNANDA
'' #05_XML_ICMS_CST                     - REVISÃO / FERNANDA
'' #05_XML_ICMS_CST_VICMS               - REVISÃO / FERNANDA
'' #05_XML_IPI                          - REVISÃO / FERNANDA
'' #AILTON - qryInsertCompraItens       - REVISÃO / FERNANDA
'' #AILTON - qryInsertProdutoConsumo    - REVISÃO / FERNANDA

''----------------------------




'Private Const qryDeleteProcessamento As String = _
'        "DELETE * FROM tblProcessamento;"
'
'Private Const qryUpdateProcessamento_Chave As String = _
'        "UPDATE tblProcessamento SET tblProcessamento.chave = Replace([tblProcessamento].[chave],';','|');"
'
'Private Const qryUpdateProcessamento_RelacaoCamposDeTabelas As String = _
'        "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = DLookUp(""tabela"",""tblOrigemDestino"",""tag='"" & [tblProcessamento].[chave] & ""'""), tblProcessamento.NomeCampo = DLookUp(""campo"",""tblOrigemDestino"",""tag='"" & [tblProcessamento].[chave] & ""'""), tblProcessamento.formatacao = DLookUp(""formatacao"",""tblOrigemDestino"",""tag='"" & [tblProcessamento].[chave] & ""'"");"
'
'Private Const qryUpdateProcessamento_RelacaoCamposDeTabelas_Item_CompraNFItem As String = _
'        "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = ""tblCompraNFItem"", tblProcessamento.NomeCampo = [tblProcessamento].[chave], tblProcessamento.formatacao = DLookUp(""formatacao"",""tblOrigemDestino"",""campo='Item_CompraNFItem'"") WHERE (((tblProcessamento.chave)=""Item_CompraNFItem""));"
'
'Private Const qryUpdateProcessamento_RelacaoCamposDeTabelas_ChvAcesso_CompraNF As String = _
'        "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = ""tblCompraNF"", tblProcessamento.NomeCampo = [tblProcessamento].[chave], tblProcessamento.formatacao = ""opTexto"" WHERE (((tblProcessamento.chave)=""ChvAcesso_CompraNF""));"
'
'Private Const qryUpdateProcessamento_RelacaoCamposDeTabelas_tblCompraNFItem_ChvAcesso_CompraNF As String = _
'        "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = strSplit([tblProcessamento].[chave],'.',0), tblProcessamento.NomeCampo = strSplit([tblProcessamento].[chave],'.',1), tblProcessamento.formatacao = strSplit([tblProcessamento].[chave],'.',2) WHERE (((tblProcessamento.chave)=""tblCompraNFItem.ChvAcesso_CompraNF.opTexto""));"
'
'Private Const qrySelecaoDeCampos As String = _
'        "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((tblOrigemDestino.tabela)='strParametro') AND ((Len([Tag]))>0) AND ((tblOrigemDestino.TagOrigem)=1)) ORDER BY tblOrigemDestino.Tag, tblOrigemDestino.tabela;"
'
'Private Const qrySelecaoDeArquivosPendentes As String = _
'        "SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1)) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0) ORDER BY tblDadosConexaoNFeCTe.CaminhoDoArquivo;"
'
'Private Const qrySelectRegistroValido As String = _
'        "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi]))>0) AND ((tblDadosConexaoNFeCTe.registroValido)=1));"



Function exemplo_MODELO()
On Error GoTo adm_Err
Dim s As New clsProcessamentoDados
Dim DadosGerais As New clsConexaoNfeCte

Dim strRepositorio As String: strRepositorio = "tblDadosConexaoNFeCTe"

'    s.ProcessamentoTransferir strRepositorio

'    DadosGerais.FormatarItensDeCompras

adm_Exit:
    Set s = Nothing
    Set DadosGerais = Nothing

    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit

End Function


''' Using DoCmd.OutputTo Method
''' DoCmd.OutputTo ObjectType:=acOutputQuery, ObjectName:=”Query1", OutputFormat:=acFormatXLS, Outputfile:=”C:\test\test.xls”

'Function proc_01_Dados_Gerais()
'On Error GoTo adm_Err
'Dim s As New clsConexaoNfeCte
'
''' 01.CARREGAR DADOS GERAIS - CONCLUIDO
''    s.carregar_DadosGerais
'
'
'    '' LIMPAR TABELA DE PROCESSAMENTOS
'    Application.CurrentDb.Execute qryDeleteProcessamento
'
'    '' CARREGAR ARQUIVOS
'    For Each Item In GetFilesInSubFolders(DLookup("ValorDoParametro", "tblParametros", "TipoDeParametro='caminhoDeColeta'"))
'
'        Debug.Print CStr(Item)
'        s.ProcessamentoDeArquivo CStr(Item), opDadosGerais
'
'    Next Item
'
'
'
'    s.processamento_IdentificarCamposTabela "tblDadosConexaoNFeCTe"
'
'
'
'    MsgBox "Fim!", vbOKOnly + vbExclamation, "proc_01"
'
'
'adm_Exit:
'    Exit Function
'
'adm_Err:
'    MsgBox Error$
'    Resume adm_Exit
'
'End Function



'Sub proc_02_cadastro_de_teste_unico_por_tipos_de_arquivos_para_processamento()
'Dim s As New clsConexaoNfeCte
'
'
'    '' LIMPAR TABELA DE PROCESSAMENTOS
'    Application.CurrentDb.Execute qryDeleteProcessamento
'
'
'    '' Retorno simbólico de mercadoria depositada em depósito fecha
'    s.ProcessamentoDeArquivo "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210212680452000302550020000886301507884230-nfeproc.xml", opCompras
'
'
'    '' TRANSF. DE MERCADORIAS
''    s.ProcessamentoDeArquivo "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210368365501000296550000000638811001361356-nfeproc.xml", opCompras
'
'
''    '' #MODELO 57 - TIPO 01
''    ''##########################
''
''    '' TRANSPORTE RODOVIARIO
''    s.ProcessamentoDeArquivo "C:\temp\Coleta\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210248740351015359570000000309211914301218-cteproc.xml", opCompras
''
''    '' PREST. SERV. TRANSPORTE A ESTABELECIMENTO COMERCIAL
''    s.ProcessamentoDeArquivo "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210220147617000494570010009539201999046070-cteproc.xml", opCompras
'
'
'End Sub


'Sub proc_03_cadastro_de_teste_unico()
'Dim s As New clsConexaoNfeCte
'
'''' UPDATE - ID_Prod_CompraNFItem
'Dim qryUpdateItens_ID_Prod_CompraNFItem As String: qryUpdateItens_ID_Prod_CompraNFItem = "UPDATE tblCompraNFItem SET tblCompraNFItem.ID_Prod_CompraNFItem = DLookUp(""CodigoProd_Grade"",""dbo_tabGradeProdutos"",""CodigoForn_Grade='"" & [tblCompraNFItem].[ID_Prod_CompraNFItem] & ""'"");"
'
'
'    '' LIMPAR TABELAS DE TESTES
'    Application.CurrentDb.Execute "DELETE * FROM tblCompraNF;"
'    Application.CurrentDb.Execute "DELETE * FROM tblCompraNFItem;"
'
'    '' TRANSFERIR DADOS PROCESSADOS PARA TABELAS DE TESTES
'    s.ProcessamentoTransferir "tblCompraNF"
'    s.ProcessamentoTransferir "tblCompraNFItem"
'
'    '' ATUALISAR CAMPOS [ID_Prod_CompraNFItem] NA TABELA [tblCompraNFItem]
'    Application.CurrentDb.Execute qryUpdateItens_ID_Prod_CompraNFItem
'
'
'End Sub


        
'Sub TESTE_CARREGAR_REGISTRO_LISTA()
'Dim s As New clsConexaoNfeCte
'Dim Item As Variant
'Dim DT_PROCESSO As Date: DT_PROCESSO = Now()
'Dim DT_PROCESSO_ITEM As Date
'Dim contadorDeRegistros As Long: contadorDeRegistros = 1
'SysCmd acSysCmdInitMeter, "Processando...", DCount("*", "tblDadosConexaoNFeCTe", "(((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0))")
'
'    '' LIMPAR TABELA DE PROCESSAMENTOS
'    Application.CurrentDb.Execute qryDeleteProcessamento
'
'    '' CARREGAR_DADOS
'    For Each Item In carregarParametros(qrySelecaoDeArquivosPendentes)
'        DT_PROCESSO_ITEM = Now()
'
'        SysCmd acSysCmdUpdateMeter, contadorDeRegistros
'        s.ProcessamentoDeArquivo CStr(Item), opCompras
'        contadorDeRegistros = contadorDeRegistros + 1
'
'        statusFinal DT_PROCESSO_ITEM, "Processamento - " & CStr(Item)
'        DoEvents
'    Next Item
'
'    '' FORMATAR CAMPOS
''    s.FormatarCamposEmProcessamento
'
'    '' #TRANSFERIR DADOS PROCESSADOS - COMPRAS
'    s.ProcessamentoTransferir "tblCompraNF"
'
'    '' #TRANSFERIR DADOS PROCESSADOS - COMPRAS ITENS
'    s.ProcessamentoTransferir "tblCompraNFItem"
'
'    '' #TRATAMENTO
''    s.TratamentoDeCompras
''    s.compras_atualizarCampos
'
'
'    SysCmd acSysCmdRemoveMeter
'    statusFinal DT_PROCESSO, "Processamento - teste_listarArquivos"
'
'End Sub


'' ############################################################################################################################



'' ############################################################################################################################

