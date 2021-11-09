Attribute VB_Name = "00_Processamento"
Option Compare Database

'' 32210268365501000296550000000637741001351624 | 55
'' 32210304884082000569570000040073831040073834 | 57


Sub teste_TransferirProcessamentoParaRepositorios()
Dim Processamento As New clsProcessamentoDados
Dim DadosGerais As New clsConexaoNfeCte

Dim strChave As String: strChave = "32210268365501000296550000000637741001351624" ' 55
'Dim strChave As String: strChave = "32210304884082000569570000040073831040073834" ' 57
Dim strRepositorio As String: strRepositorio = "tblCompraNF"

Dim strArquivo As String: strArquivo = DLookup("[CaminhoDoArquivo]", "[tblDadosConexaoNFeCTe]", "[ChvAcesso]='" & strChave & "'")

    '' LIMPAR TABELA DE PROCESSAMENTOS
    Processamento.DeleteProcessamento
    
    '' #IMPORTACAO - CADASTRO DE DADOS DO ARQUIVO NA TABELA DE PROCESSAMENTO -
    Processamento.ProcessamentoDeArquivo strArquivo, opCompras
        
    '' #IDENTIFICA��O - CAMPOS
    Processamento.UpdateProcessamentoIdentificarCampos "tblCompraNF"
    
    '' #CORRE��O DE DADOS MARCADOS ERRADOS EM ITENS DE COMPRAS
    Processamento.UpdateProcessamentoLimparItensMarcadosErrados

    '' #IDENTIFICA��O - CAMPOS
    Processamento.UpdateProcessamentoIdentificarCampos "tblCompraNFItem"
            
    '' FORMATAR DADOS
    Processamento.UpdateProcessamentoFormatarDados
            
    '' #CLASSIFICAR
    DadosGerais.TratamentoDeDadosGerais
    DadosGerais.compras_atualizarCampos
    
    
    
    
    
            
    '' TRANSFERIR
'    TransferirProcessamentoParaRepositorios strChave


Set Processamento = Nothing

End Sub



'' 01. Sele��o
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

'' 02. Classifica��o
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


