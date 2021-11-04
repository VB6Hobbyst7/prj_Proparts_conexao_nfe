Attribute VB_Name = "00_Info"
Option Compare Database

'' LIMPAR TODA A BASE DE DADOS
Public Const dataBaseClear As Boolean = True

'' REPROCESSAR ARQUIVOS PENDENTES
Public Const dataBaseReplay As Boolean = False

'' EXPORTAR DADOS PARA SERVIDOR
Public Const dataBaseExportarDados As Boolean = False

'' PROCESSAMENTO DE ARQUIVOS
Public Const dataBaseTratamentoDeArquivos As Boolean = False
Public Const dataBaseGerarLancamentoManifesto As Boolean = False



Sub ProcessamentoTransferir_teste()
Dim s As New clsProcessamentoDados: s.ProcessamentoTransferir "tblCompraNF"
Dim item  As Variant





For Each item In Array("ChvAcesso_CompraNF|42210300634453001303570010001139451001171544", "CPNJ_Dest|68365501000105", "emit_UF|SC", "CNPJ_CPF_CompraNF|00634453001303", "NomeCompleto_CompraNF|T S V TRANSPORTES RAPIDOS LTDA  ITJ", "CFOP|6353", "HoraEntd_CompraNF|20:35:00", "codMod|57", "NumNF_CompraNF|113945", "Serie_CompraNF|1", "TPNF_CompraNF|0", "BaseCalcICMS_CompraNF|119.36", "VTotICMS_CompraNF|14.32", "rem_UF|SC", "CNPJ_Rem_CompraNF|68365501000377", "VTotNF_CompraNF|119.36", "ChvAcesso|42210300634453001303570010001139451001171544")

    Debug.Print CStr(Split(item, "|")(1))

Next


Set s = Nothing
End Sub


Sub teste_FuncionamentoGeralDeProcessamentoDeArquivos()
Dim strCaminhoAcoes As String: strCaminhoAcoes = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColetaAcoes'")
    
    ''==================================================
    '' REPOSITORIO GERAL
    ''==================================================

    '' LIMPAR TODA A BASE DE DADOS
    If dataBaseClear Then
    
        '' Limpar toda a base de dados
'        Application.CurrentDb.Execute "Delete from tblDadosConexaoNFeCTe"

        '' Limpar repositorio de itens de compras
        Application.CurrentDb.Execute _
                "Delete from tblCompraNFItem"
    
        '' Limpar repositorio de compras
        Application.CurrentDb.Execute _
                "Delete from tblCompraNF"

        '' Carregar todos os arquivos para processamento.
        processarDadosGerais
        
    Else
        
        '' Carregar todos os arquivos para processamento.
        processarDadosGerais
    
    
    End If

    ''==================================================
    '' REPOSITORIOS DE COMPRAS
    ''==================================================
    
    '' REPROCESSAR ARQUIVOS VALIDOS
    If dataBaseReplay Then
    
        '' Ajustar marcação de registro
        Application.CurrentDb.Execute _
            "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado=0 WHERE tblDadosConexaoNFeCTe.registroValido=1 AND tblDadosConexaoNFeCTe.ID_Tipo>0"
        
        '' Limpar repositorio de itens de compras
        Application.CurrentDb.Execute _
                "Delete from tblCompraNFItem"
    
        '' Limpar repositorio de compras
        Application.CurrentDb.Execute _
                "Delete from tblCompraNF"

        '' Processamento de arquivos pendentes da pasta de coleta.
        processarArquivosPendentes
            
    Else
    
        '' Processamento de arquivos pendentes da pasta de coleta.
        processarArquivosPendentes
    
    End If


    ''==================================================
    '' EXPORTAR DADOS PARA O SERVIDOR
    ''==================================================

    '' EXPORTAÇÃO DE DADOS
    If dataBaseExportarDados Then _
            CadastroDeComprasEmServidor

    ''==================================================
    '' PROCESSAMENTO DE ARQUIVOS
    ''==================================================

    '' #### TRANSFERENCIAS DE ARQUIVOS
    If dataBaseTratamentoDeArquivos Then _

        '' Transferir Arquivos Validos para pasta de processados
        tratamentoDeArquivosValidos
    
        '' Transferir Arquivos Invalidos para pasta de Expurgo
        tratamentoDeArquivosInvalidos

    End If

    '' #### GERAR ARQUIVOS DE LANÇAMENTO E MANIFESTO
    If dataBaseGerarLancamentoManifesto Then
    
        '' LANÇAMENTO
        gerarArquivosJson opFlagLancadaERP, , strCaminhoAcoes
    
        '' MANIFESTO
        gerarArquivosJson opManifesto, , strCaminhoAcoes
        
    End If
    
Debug.Print "### Concluido! - testeDeFuncionamentoGeral"
TextFile_Append CurrentProject.path & "\" & strLog(), "Concluido! - testeDeFuncionamentoGeral"

End Sub

