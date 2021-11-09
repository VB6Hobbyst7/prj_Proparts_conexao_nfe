Attribute VB_Name = "00_Processamento"
Option Compare Database

'' 01. Seleção
Sub teste_CarregarArquivosDeColeta()
Dim arquivo As Variant
Dim Processamento As New clsProcessamentoDados
Dim DadosGerais As New clsConexaoNfeCte

Dim strRepositorio As String: strRepositorio = "tblDadosConexaoNFeCTe"

    For Each arquivo In CarregarArquivosDeColeta
        Debug.Print arquivo
        TextFile_Append CurrentProject.path & "\CarregarArquivosDeColeta.log", CStr(arquivo)
                
        '' LIMPAR TABELA DE PROCESSAMENTOS
        Processamento.DeleteProcessamento
        
        '' PROCESSAMENTO
        Processamento.ProcessamentoDeArquivo CStr(arquivo), opDadosGerais
        
        '' IDENTIFICAR CAMPOS
        Processamento.UpdateProcessamentoIdentificarCampos strRepositorio
                
        '' FORMATAR DADOS
        Processamento.UpdateProcessamentoFormatarDados
        
        '' TRANSFERENCIA DE DADOS
        Processamento.ProcessamentoTransferir strRepositorio
    
        '' CLASSIFICAR DADOS GERAIS
        DadosGerais.TratamentoDeDadosGerais
    
    
        DoEvents
    Next

End Sub

'' 02. Classificação
Sub teste_CarregarArquivosParaProcessamento()
Dim arquivo As Variant

    For Each arquivo In CarregarArquivosParaProcessamento
        Debug.Print arquivo
        TextFile_Append CurrentProject.path & "\CarregarArquivosParaProcessamento.log", CStr(arquivo)
        
        
        
        
        
        
        DoEvents
    Next

End Sub



'' #############################################
'' ### LIB'S
'' #############################################


Function CarregarArquivosParaProcessamento() As Collection: Set CarregarArquivosParaProcessamento = New Collection
Dim DadosGerais As New clsConexaoNfeCte

    For Each item In carregarParametros(DadosGerais.SelectArquivosPendentes)
        CarregarArquivosParaProcessamento.add CStr(item)
    Next

Set DadosGerais = Nothing
End Function


Function CarregarArquivosDeColeta() As Collection: Set CarregarArquivosDeColeta = New Collection
Dim DadosGerais As New clsConexaoNfeCte

'' REPOSITORIOS
For Each caminhoAntigo In Array(DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColeta'"))
    For Each caminhoNovo In carregarParametros(DadosGerais.SelectColetaEmpresa)
        For Each item In GetFilesInSubFolders(CStr(Replace(Replace(caminhoAntigo, "empresa", caminhoNovo), "recebimento\", "")))
            CarregarArquivosDeColeta.add CStr(item)
        Next
    Next
Next

End Function


