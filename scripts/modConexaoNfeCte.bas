Attribute VB_Name = "modConexaoNfeCte"
Option Compare Database

'Sub teste_cadastroProcessamento()
'
'Dim strTemp As String: strTemp = _
'    "INSERT INTO tblProcessamento (valor, NomeTabela, NomeCampo, formatacao, NomeCampo) " & _
'    "SELECT tblorigemDestino.valorPadrao    ,tblorigemDestino.tabela ,tblorigemDestino.campo ,tblorigemDestino.formatacao ,tblorigemDestino.campo " & _
'    "FROM tblorigemDestino " & _
'    "WHERE (((tblorigemDestino.tabela) = 'tblCompraNF')  " & _
'    "AND ((tblorigemDestino.formatacao) = 'opMoeda')  " & _
'    "AND ((tblorigemDestino.campo)  " & _
'    "NOT IN (SELECT DISTINCT tblProcessamento.nomecampo,tblProcessamento.pk FROM tblProcessamento WHERE (((tblProcessamento.nomecampo) IS NOT NULL) AND ((tblProcessamento.[pk]) EXISTS (SELECT DISTINCT tblProcessamento.pk FROM tblProcessamento))))));"
'
'Application.CurrentDb.Execute strTemp
'
'End Sub

Sub TESTE_TransferirDadosProcessados()
Dim strProcessamento As String: strProcessamento = "tblCompraNF"
Dim s As New clsConexaoNfeCte

    '' #CARREGAR DADOS
    For Each t In Array(strProcessamento)
        
        '' #TRANSFERIR DADOS PROCESSADOS - COMPRAS
        s.TransferirDadosProcessados strProcessamento

    Next

    '' #VALIDAR_DADOS
    criarConsultasParaTestes
    
    MsgBox "Fim!", vbOKOnly + vbExclamation, "carregarCompras"

End Sub

'' 02.CARREGAR COMPRAS ANTES DE VENVIAR PARA O SERVIDOR
Sub carregarCompras()
Dim strProcessamento As String: strProcessamento = "tblCompraNF"
Dim s As New clsConexaoNfeCte
Dim t As Variant

    '' #CARREGAR DADOS
    For Each t In Array(strProcessamento)
    
        '' PROCESSAR APENAS ARQUIVOS VALIDOS
        s.ProcessarArquivosXml CStr(t), carregarParametros(qrySelectProcessamentoPendente)
        
        '' FORMATAR CAMPOS
        s.FormatarCampos
        
        '' #TRATAMENTO
        s.TratamentoDeCompras
        
        '' #TRANSFERIR DADOS PROCESSADOS - COMPRAS
        s.TransferirDadosProcessados strProcessamento

    Next

    '' #VALIDAR_DADOS
    criarConsultasParaTestes
    
    MsgBox "Fim!", vbOKOnly + vbExclamation, "carregarCompras"

End Sub

'' 01.CARREGAR DADOS GERAIS - CONCLUIDO
Sub carregarDadosGerais()
Dim strProcessamento As String: strProcessamento = "tblDadosConexaoNFeCTe"
Dim s As New clsConexaoNfeCte
Dim t As Variant

    '' #LIMPAR_BASE_DE_TESTES
    Application.CurrentDb.Execute "DELETE FROM tblDadosConexaoNFeCTe"
    Application.CurrentDb.Execute "DELETE FROM tblCompraNF"
'    Application.CurrentDb.Execute "DELETE FROM tblCompraNFItem"

    '' #CARREGAR DADOS
    For Each t In Array(strProcessamento)

        '' #PROCESSAMENTO DE ARQUIVO - ENVIO DE DADOS PARA tblProcessamento
        s.ProcessarArquivosXml CStr(t), GetFilesInSubFolders(pegarValorDoParametro(qryParametros, strCaminhoDeColeta))

        '' FORMATAR CAMPOS
        s.FormatarCampos

        '' #TRANSFERIR DADOS PROCESSADOS - DADOS GERAIS - ENVIO DE DADOS PARA tblDadosConexaoNFeCTe
        s.TransferirDadosProcessados strProcessamento

        '' #TRATAMENTO
        s.TratamentoDeDadosGerais

        '' #ARQUIVOS - GERAR ARQUIVOS
        s.CriarTipoDeArquivo opFlagLancadaERP
        s.CriarTipoDeArquivo opManifesto

    Next

    MsgBox "Fim!", vbOKOnly + vbExclamation, "carregarDadosGerais"

End Sub

'' #ADMINISTRACAO
Sub ADM_criarTabelas()

    ''tblCompraNF
    'Application.CurrentDb.Execute "DROP TABLE tblCompraNF"
    Application.CurrentDb.Execute createTable("tblCompraNF")

End Sub

'' #ADMINISTRACAO - RESPONSAVEL POR TRAZER OS DADOS DO SERVIDOR PARA AUXILIO NO PROCESSAMENTO. QUANDO NECESSARIO
Sub ADM_carregarDadosDoServidor()
    
    '' NATUREZA DE OPERAÇÃO
    ImportarDados "tblNatOp", "tmpNatOp"
    
    '' CADASTRO DE EMPRESA
    ImportarDados "tblEmpresa", "tmpEmpresa"
    
    '' CADASTRO DE CLIENTES
    ImportarDados "Clientes", "tmpClientes"

End Sub
