Attribute VB_Name = "00_Info"
Option Compare Database

Sub teste_FuncionamentoGeralDeProcessamentoDeArquivos()
Dim strCaminhoAcoes As String: strCaminhoAcoes = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColetaAcoes'")
    
    ''==================================================
    '' REPOSITORIO GERAL
    ''==================================================

    '' LIMPAR REPOSITORIO GERAL
    Application.CurrentDb.Execute "Delete from tblDadosConexaoNFeCTe"

    '' Carregar todos os arquivos para processamento.
    processarDadosGerais

'    '' REPROCESSAR ARQUIVOS VALIDOS
'    Application.CurrentDb.Execute _
'            "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado=0 WHERE tblDadosConexaoNFeCTe.registroValido=1 AND tblDadosConexaoNFeCTe.ID_Tipo>0"

    ''==================================================
    '' REPOSITORIOS DE COMPRAS
    ''==================================================

    '' ZERAR CONTADOR DE NUMERO DE PEDIDOS
    Application.CurrentDb.Execute _
            "UPDATE tblParametros SET tblParametros.ValorDoParametro = 0 WHERE (((tblParametros.TipoDeParametro)=""NumPed_CompraNF""));"

    '' LIMPAR REPOSITORIO DE ITENS DE COMPRAS
    Application.CurrentDb.Execute _
            "Delete from tblCompraNFItem"

    '' LIMPAR REPOSITORIO DE COMPRAS
    Application.CurrentDb.Execute _
            "Delete from tblCompraNF"

    '' Processamento de arquivos pendentes da pasta de coleta.
    processarArquivosPendentes

'    ''==================================================
'    '' EXPORTAR DADOS PARA O SERVIDOR
'    ''==================================================
'
'    '' EXPORTAÇÃO DE DADOS
'    enviarDadosServidor
'
'    ''==================================================
'    '' PROCESSAMENTO DE ARQUIVOS
'    ''==================================================
'
'    '' #### ARQUIVAR DADOS PROCESSADOS
'    '' Transferir Arquivos Validos para pasta de processados
'    tratamentoDeArquivosValidos
'
'    '' Transferir Arquivos Invalidos para pasta de Expurgo
'    tratamentoDeArquivosInvalidos
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

