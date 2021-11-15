'' 01.CARREGAR DADOS GERAIS - CONCLUIDO
Public Function carregar_DadosGerais()
On Error GoTo adm_Err

Dim strProcessamento As String: strProcessamento = "tblDadosConexaoNFeCTe"
Dim s As New clsConexaoNfeCte
Dim t As Variant

Dim retVal As Variant: retVal = MsgBox("Deseja carregar os dados gerais dos arquivos XML ?", vbQuestion + vbYesNo, "carregarDadosGerais")

    If retVal = vbYes Then
    
        '' #LIMPAR_BASE_DE_TESTES
        Application.CurrentDb.Execute "DELETE FROM tblDadosConexaoNFeCTe"
        Application.CurrentDb.Execute "DELETE FROM tblCompraNF"
        Application.CurrentDb.Execute "DELETE FROM tblCompraNFItem"
    
        '' #CARREGAR DADOS
        For Each t In Array(strProcessamento)
    
            '' #PROCESSAMENTO DE ARQUIVO - ENVIO DE DADOS PARA tblProcessamento
            ProcessarArquivosXml CStr(t), GetFilesInSubFolders(DLookup("ValorDoParametro", "tblParametros", "TipoDeParametro='caminhoDeColeta'"))
    
            '' FORMATAR CAMPOS
            FormatarCamposEmProcessamento
    
            '' #TRANSFERIR DADOS PROCESSADOS - DADOS GERAIS - ENVIO DE DADOS PARA tblDadosConexaoNFeCTe
            s.ProcessamentoTransferir strProcessamento
    
            '' #TRATAMENTO DE DADOS GERAIS
            TratamentoDeDadosGerais
    
        Next
    
    End If

adm_Exit:
    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit

End Function