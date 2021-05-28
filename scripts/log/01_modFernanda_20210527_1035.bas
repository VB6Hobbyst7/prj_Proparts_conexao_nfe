Attribute VB_Name = "01_modFernanda"
Option Compare Database




''----------------------------
'' ### EXEMPLOS DE FUNÇÕES
''
'' exemplo_processamento_dados_gerais
'' exemplos_criacao_arquivos_json
''
''----------------------------

Function exemplo_processamento_dados_gerais()
On Error GoTo adm_Err
Dim s As New clsProcessamentoDados

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

'' #CONTADOR
Dim contadorDeRegistros As Long: contadorDeRegistros = 1

    '' LIMPAR TABELA DE PROCESSAMENTOS
    s.DeleteProcessamento

    '' CARREGAR DADOS DE ARQUIVOS NA TABELA DE PROCESSAMENTO
    For Each Item In GetFilesInSubFolders(DLookup("ValorDoParametro", "tblParametros", "TipoDeParametro='caminhoDeColeta'"))
        s.ProcessamentoDeArquivo CStr(Item), opDadosGerais
        
        '' #CONTADOR
        contadorDeRegistros = contadorDeRegistros + 1
        Debug.Print contadorDeRegistros
        
        DoEvents
    Next Item

    '' IDENTIFICAR CAMPOS
    s.UpdateProcessamentoIdentificarCampos "tblDadosConexaoNFeCTe"
    
    '' FORMATAR DADOS
    s.UpdateProcessamentoFormatarDados


    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "Processamento - ImportarDados ( Quantidade de registros: " & contadorDeRegistros & " )"

    MsgBox "Fim!", vbOKOnly + vbExclamation, "exemplo_processamento_dados_gerais"


adm_Exit:
    Set s = Nothing
    
    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit

End Function

Sub exemplos_criacao_arquivos_json()
Dim s As New clsCriarArquivos
Dim qrySelectRegistroValido As String: qrySelectRegistroValido = _
    "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi]))>0) AND ((tblDadosConexaoNFeCTe.registroValido)=1));"

Dim strCaminhoDeSaida As String: strCaminhoDeSaida = "C:\temp\20210527\"
CreateDir strCaminhoDeSaida
    
    '' NO PROCESSAMENTO DO ARQUIVO DE XML
    s.criarArquivoJson opFlagLancadaERP, qrySelectRegistroValido, strCaminhoDeSaida

    '' SELEÇÃO PELO USUARIO
    s.criarArquivoJson opManifesto, qrySelectRegistroValido, strCaminhoDeSaida


    MsgBox "Concluido!", vbOKOnly + vbInformation, "teste_arquivos_json"

Cleanup:

    Set s = Nothing

End Sub
