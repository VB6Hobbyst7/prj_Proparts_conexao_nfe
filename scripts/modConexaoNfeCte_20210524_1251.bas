Attribute VB_Name = "modConexaoNfeCte"
Option Compare Database

Private Const qryDeleteProcessamento As String = _
        "DELETE * FROM tblProcessamento;"

Private Const qryUpdateProcessamento_Chave As String = _
        "UPDATE tblProcessamento SET tblProcessamento.chave = Replace([tblProcessamento].[chave],';','|');"

Private Const qryUpdateProcessamento_RelacaoCamposDeTabelas As String = _
        "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = DLookUp(""tabela"",""tblOrigemDestino"",""tag='"" & [tblProcessamento].[chave] & ""'""), tblProcessamento.NomeCampo = DLookUp(""campo"",""tblOrigemDestino"",""tag='"" & [tblProcessamento].[chave] & ""'""), tblProcessamento.formatacao = DLookUp(""formatacao"",""tblOrigemDestino"",""tag='"" & [tblProcessamento].[chave] & ""'"");"

Private Const qryUpdateProcessamento_RelacaoCamposDeTabelas_Item_CompraNFItem As String = _
        "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = ""tblCompraNFItem"", tblProcessamento.NomeCampo = [tblProcessamento].[chave], tblProcessamento.formatacao = DLookUp(""formatacao"",""tblOrigemDestino"",""campo='Item_CompraNFItem'"") WHERE (((tblProcessamento.chave)=""Item_CompraNFItem""));"

Private Const qryUpdateProcessamento_RelacaoCamposDeTabelas_ChvAcesso_CompraNF As String = _
        "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = ""tblCompraNF"", tblProcessamento.NomeCampo = [tblProcessamento].[chave], tblProcessamento.formatacao = ""opTexto"" WHERE (((tblProcessamento.chave)=""ChvAcesso_CompraNF""));"

Private Const qryUpdateProcessamento_RelacaoCamposDeTabelas_tblCompraNFItem_ChvAcesso_CompraNF As String = _
        "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = strSplit([tblProcessamento].[chave],'.',0), tblProcessamento.NomeCampo = strSplit([tblProcessamento].[chave],'.',1), tblProcessamento.formatacao = strSplit([tblProcessamento].[chave],'.',2) WHERE (((tblProcessamento.chave)=""tblCompraNFItem.ChvAcesso_CompraNF.opTexto""));"

Private Const qrySelecaoDeCampos As String = _
        "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((tblOrigemDestino.tabela)='strParametro') AND ((Len([Tag]))>0) AND ((tblOrigemDestino.TagOrigem)=1)) ORDER BY tblOrigemDestino.Tag, tblOrigemDestino.tabela;"
       
Private Const qrySelecaoDeArquivosPendentes As String = _
        "SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1)) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0) ORDER BY tblDadosConexaoNFeCTe.CaminhoDoArquivo;"

Private Const qrySelectRegistroValido As String = _
        "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi]))>0) AND ((tblDadosConexaoNFeCTe.registroValido)=1));"


Sub teste_arquivos_json()
Dim s As New clsConexaoNfeCte

    s.criar_ArquivoJson opFlagLancadaERP, qrySelectRegistroValido '', "C:\temp\20210524\"
    s.criar_ArquivoJson opManifesto, qrySelectRegistroValido '', "C:\temp\20210524\"
    
    Set s = Nothing

    MsgBox "Concluido!", vbOKOnly + vbInformation, "teste_arquivos_json"

End Sub


Sub teste_TRANFERIR_DADOS_PROCESSADOS_PARA_TABELA__LOCAL()
Dim s As New clsConexaoNfeCte

    s.TransferirDadosProcessados "tblCompraNF"
    s.TransferirDadosProcessados "tblCompraNFItem"

End Sub

Sub TESTE_CARREGAR_REGISTRO_UNICO()
Dim s As New clsConexaoNfeCte

    '' LIMPAR TABELA DE PROCESSAMENTOS
    Application.CurrentDb.Execute qryDeleteProcessamento
       

    '' VENDA MERCADORIAS ADQUIRIDAS E/OU RECEB TERCEIROS
    '' NUMERO_NF: 629140
    '' NUMERO_ITENS: 01
    s.processamentoDeCompras "C:\temp\Coleta\68.365.5010001-05 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\26210324073694000155550010006291401018935070-nfeproc.xml"

    '' SAIDA DE VENDA
    '' NUMERO_NF: 8745405
    '' NUMERO_ITENS: 10
    s.processamentoDeCompras "C:\temp\Coleta\68.365.5010001-05 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\35210343283811001202550010087454051410067364-nfeproc.xml"


End Sub
        
Sub TESTE_CARREGAR_REGISTRO_LISTA()
Dim s As New clsConexaoNfeCte
Dim Item As Variant
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()
Dim DT_PROCESSO_ITEM As Date
Dim contadorDeRegistros As Long: contadorDeRegistros = 1
SysCmd acSysCmdInitMeter, "Processando...", DCount("*", "tblDadosConexaoNFeCTe", "(((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0))")

    '' LIMPAR TABELA DE PROCESSAMENTOS
    Application.CurrentDb.Execute qryDeleteProcessamento

    '' CARREGAR_DADOS
    For Each Item In carregarParametros(qrySelecaoDeArquivosPendentes)
        DT_PROCESSO_ITEM = Now()
        
        SysCmd acSysCmdUpdateMeter, contadorDeRegistros
        s.processamentoDeCompras CStr(Item)
        contadorDeRegistros = contadorDeRegistros + 1
        
        statusFinal DT_PROCESSO_ITEM, "Processamento - " & CStr(Item)
        DoEvents
    Next Item
    
    '' FORMATAR CAMPOS
'    s.FormatarCamposEmProcessamento
    
    '' #TRANSFERIR DADOS PROCESSADOS - COMPRAS
    s.TransferirDadosProcessados "tblCompraNF"
    
    '' #TRANSFERIR DADOS PROCESSADOS - COMPRAS ITENS
    s.TransferirDadosProcessados "tblCompraNFItem"
        
    '' #TRATAMENTO
'    s.TratamentoDeCompras
'    s.compras_atualizarCampos
    
    
    SysCmd acSysCmdRemoveMeter
    statusFinal DT_PROCESSO, "Processamento - teste_listarArquivos"

End Sub





'' ############################################################################################################################


Function proc_01()
On Error GoTo adm_Err
Dim s As New clsConexaoNfeCte

'' #ANALISE_DE_PROCESSAMENTO

'' #ADMINISTRACAO - RESPONSAVEL POR TRAZER OS DADOS DO SERVIDOR PARA AUXILIO NO PROCESSAMENTO. QUANDO NECESSARIO
'    s.ADM_carregarDadosDoServidor

'' 01.CARREGAR DADOS GERAIS - CONCLUIDO
    s.carregar_DadosGerais

'' 02.CARREGAR COMPRAS ANTES DE ENVIAR PARA O SERVIDOR
'''''''    s.processamentoDeCompras

'' 03.ENVIAR DADOS PARA SERVIDOR
'    s.enviar_ComprasParaServidor

    MsgBox "Fim!", vbOKOnly + vbExclamation, "proc_01"


adm_Exit:
    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit

End Function

'' ############################################################################################################################

Sub TESTE_CARREGAR_REGISTRO_UNICO__20210520_1600()
Dim s As New clsConexaoNfeCte

    '' LIMPAR TABELA DE PROCESSAMENTOS
    Application.CurrentDb.Execute qryDeleteProcessamento
       

    '' VENDA MERCADORIAS ADQUIRIDAS E/OU RECEB TERCEIROS
    '' NUMERO_NF: 629140
    '' NUMERO_ITENS: 01
    s.processamentoDeCompras "C:\temp\Coleta\68.365.5010001-05 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\26210324073694000155550010006291401018935070-nfeproc.xml"

    '' SAIDA DE VENDA
    '' NUMERO_NF: 8745405
    '' NUMERO_ITENS: 10
    s.processamentoDeCompras "C:\temp\Coleta\68.365.5010001-05 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\35210343283811001202550010087454051410067364-nfeproc.xml"

    '' Retorno simbólico de mercadoria depositada em depósito fecha
'    processamentoDeCompras "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210312680452000302550020000896201269925336-nfeproc.xml"


    '' TRANSPORTE RODOVIARIO
'    processamentoDeCompras "C:\temp\Coleta\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210248740351015359570000000309211914301218-cteproc.xml"


    '' ARQUIVO - CERTIFICADO
'    processamentoDeCompras "C:\temp\Coleta\68.365.5010001-05 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\35210365833410000169550000006711211238251650-nfeproc.xml"


End Sub
