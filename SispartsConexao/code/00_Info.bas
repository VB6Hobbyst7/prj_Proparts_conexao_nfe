Attribute VB_Name = "00_Info"
Option Compare Database

Sub teste_FuncionamentoGeralDeProcessamentoDeArquivos()
Dim strCaminhoAcoes As String: strCaminhoAcoes = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColetaAcoes'")
Dim dataBaseClear As Boolean: dataBaseClear = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoClear'")
Dim dataBaseReplay As Boolean: dataBaseReplay = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoReplay'")
    
    ''==================================================
    '' REPOSITORIO GERAL
    ''==================================================

    '' LIMPAR REPOSITORIO GERAL
    If dataBaseClear Then Application.CurrentDb.Execute _
            "Delete from tblDadosConexaoNFeCTe"

    '' Carregar todos os arquivos para processamento.
    processarDadosGerais


    ''==================================================
    '' REPOSITORIOS DE COMPRAS
    ''==================================================

    '' LIMPAR REPOSITORIO DE ITENS DE COMPRAS
    If dataBaseClear Then Application.CurrentDb.Execute _
            "Delete from tblCompraNFItem"

    '' LIMPAR REPOSITORIO DE COMPRAS
    If dataBaseClear Then Application.CurrentDb.Execute _
            "Delete from tblCompraNF"

    '' Processamento de arquivos pendentes da pasta de coleta.
    processarArquivosPendentes

    ''==================================================
    '' EXPORTAR DADOS PARA O SERVIDOR
    ''==================================================

    '' EXPORTAÇÃO DE DADOS
    CadastroDeComprasEmServidor

    ''==================================================
    '' PROCESSAMENTO DE ARQUIVOS
    ''==================================================

'    '' Transferir Arquivos Validos para pasta de processados
'    tratamentoDeArquivosValidos
'
'    '' Transferir Arquivos Invalidos para pasta de Expurgo
'    tratamentoDeArquivosInvalidos
'
'
'
'    '' #### GERAR ARQUIVOS DE LANÇAMENTO E MANIFESTO
'    '' LANÇAMENTO
'    gerarArquivosJson opFlagLancadaERP, , strCaminhoAcoes
'
'    '' MANIFESTO
'    gerarArquivosJson opManifesto, , strCaminhoAcoes
    
    
Debug.Print "### Concluido! - testeDeFuncionamentoGeral"
TextFile_Append CurrentProject.path & "\" & strLog(), "Concluido! - testeDeFuncionamentoGeral"

End Sub

