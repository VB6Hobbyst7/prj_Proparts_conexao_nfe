Attribute VB_Name = "01_modFernanda"
Option Compare Database


''----------------------------
'' ### EXEMPLOS DE FUNÇÕES
''
'' exemplo_processamento_dados_gerais
'' exemplos_criacao_arquivos_json
''
''----------------------------

Sub id_testes()
Dim s As New clsProcessamentoDados

s.UpdateProcessamentoIdentificarCampos "tblCompraNFItem"


End Sub


Function exemplo_processamento_compras()
On Error GoTo adm_Err
Dim s As New clsProcessamentoDados
Dim DadosGerais As New clsConexaoNfeCte

Dim strRepositorio As String: strRepositorio = "tblCompraNF"

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

'' #CONTADOR
Dim contadorDeRegistros As Long: contadorDeRegistros = 0

    '' LIMPAR TABELA DE PROCESSAMENTOS
    s.DeleteProcessamento

    ''#######################################################################################
    ''### TESTES UNITARIOS
    ''#######################################################################################

    Dim arquivos As Collection: Set arquivos = New Collection

    ''' RETORNO SIMBÓLICO DE MERCADORIA DEPOSITADA EM DEPÓSITO FECHA
'    arquivos.add "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210212680452000302550020000886301507884230-nfeproc.xml"

'    ''' TRANSF. DE MERCADORIAS
'    arquivos.add "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210368365501000296550000000638811001361356-nfeproc.xml"
'
'    ''' #TIPO 01 - CTE
'    ''' TRANSPORTE RODOVIARIO
    arquivos.add "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210220147617000494570010009539201999046070-cteproc.xml" '' ---> Pendente
'
'    ''' PREST. SERV. TRANSPORTE A ESTABELECIMENTO COMERCIAL
'    arquivos.add "C:\temp\Coleta\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210304884082000569570000040073831040073834-cteproc.xml"

    For Each Item In arquivos
        s.ProcessamentoDeArquivo CStr(Item), opCompras
        
        '' #CONTADOR
        contadorDeRegistros = contadorDeRegistros + 1
        Debug.Print contadorDeRegistros

        DoEvents
    Next Item

    '' IDENTIFICAR CAMPOS
    s.UpdateProcessamentoIdentificarCampos strRepositorio
    
    '' CORREÇÃO DE DADOS MARCADOS ERRADOS EM ITENS DE COMPRAS
    s.UpdateProcessamentoLimparItensMarcadosErrados
    
    '' IDENTIFICAR CAMPOS DE ITENS DE COMPRAS
    s.UpdateProcessamentoIdentificarCampos strRepositorio & "Item"
    
    '' FORMATAR DADOS
    s.UpdateProcessamentoFormatarDados

    '' TRANSFERIR DADOS PROCESSADOS
'    s.ProcessamentoTransferir strRepositorio
'    s.ProcessamentoTransferir strRepositorio & "Item"
    
    '' FORMATAR ITENS DE COMPRA
'    DadosGerais.FormatarItensDeCompras
    
    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "Processamento - Importar Dados Gerais ( Quantidade de registros: " & contadorDeRegistros & " )"


    MsgBox "Concluido!", vbOKOnly + vbInformation, strRepositorio

adm_Exit:
    Set s = Nothing

    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit

End Function




Function exemplo_processamento_dados_gerais()
On Error GoTo adm_Err
Dim s As New clsProcessamentoDados
Dim DadosGerais As New clsConexaoNfeCte
Dim Item As Variant

Dim strRepositorio As String: strRepositorio = "tblDadosConexaoNFeCTe"

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

'' #CONTADOR
Dim contadorDeRegistros As Long: contadorDeRegistros = 0

    '' LIMPAR TABELA DE PROCESSAMENTOS
    s.DeleteProcessamento
    
    ''#######################################################################################
    ''### CARREGAR DADOS DE ARQUIVOS NA TABELA DE PROCESSAMENTO
    ''#######################################################################################

    '' CARREGAR DADOS DE ARQUIVOS NA TABELA DE PROCESSAMENTO
    For Each Item In GetFilesInSubFolders(DLookup("ValorDoParametro", "tblParametros", "TipoDeParametro='caminhoDeColeta'"))
        s.ProcessamentoDeArquivo CStr(Item), opDadosGerais

        '' #CONTADOR
        contadorDeRegistros = contadorDeRegistros + 1
        Debug.Print contadorDeRegistros

        DoEvents
    Next Item


    '' IDENTIFICAR CAMPOS
    s.UpdateProcessamentoIdentificarCampos strRepositorio
    
    '' FORMATAR DADOS
    s.UpdateProcessamentoFormatarDados
    
    '' TRANSFERIR DADOS PROCESSADOS
    s.ProcessamentoTransferir strRepositorio
    
    '' CLASSIFICAR DADOS GERAIS
    DadosGerais.TratamentoDeDadosGerais
    

    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "Processamento - Importar Dados Gerais ( Quantidade de registros: " & contadorDeRegistros & " )"


    MsgBox "Concluido!", vbOKOnly + vbInformation, strRepositorio


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




'Function exemplo_MODELO()
'On Error GoTo adm_Err
'Dim s As New clsProcessamentoDados
'
'Dim strRepositorio As String: strRepositorio = "tblDadosConexaoNFeCTe"
'
''    s.ProcessamentoTransferir strRepositorio
'
'adm_Exit:
'    Set s = Nothing
'
'    Exit Function
'
'adm_Err:
'    MsgBox Error$
'    Resume adm_Exit
'
'End Function

