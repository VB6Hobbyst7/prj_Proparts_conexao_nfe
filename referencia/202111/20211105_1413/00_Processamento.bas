Attribute VB_Name = "00_Processamento"
Option Compare Database


Sub teste_CarregarArquivos()
Dim arquivo As Variant

For Each arquivo In CarregarArquivosDeColeta
    Debug.Print arquivo
    TextFile_Append CurrentProject.path & "\ListagemDeArquivosColetados.txt", CStr(arquivo)
Next


Set DadosGerais = Nothing

End Sub



Function CarregarArquivosDeColeta() As Collection: Set CarregarArquivosDeColeta = New Collection
Dim arquivos As Collection: Set arquivos = New Collection

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


