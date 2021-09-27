Attribute VB_Name = "01_modFernanda"
Option Compare Database


''----------------------------
'' ### EXEMPLOS DE FUNÇÕES
''
'' 01. processamento_dados_gerais
'' 02. processamento_compras_por_arquivos_pendentes
'' 03. exemplos_criacao_arquivos_json
''
'' 99. exemplo_processamento_compras_por_arquivos_unicos
''----------------------------


''----------------------------

'' #AILTON - VALIDAR
'' #ARQUIVOS - GERAR ARQUIVOS | PROCESSAMENTO POR ARQUIVO(S)

'' #05_XML_ICMS                         - REVISÃO / FERNANDA
'' #05_XML_ICMS_Orig                    - REVISÃO / FERNANDA
'' #05_XML_ICMS_CST                     - REVISÃO / FERNANDA
'' #05_XML_ICMS_CST_VICMS               - REVISÃO / FERNANDA
'' #05_XML_IPI                          - REVISÃO / FERNANDA
'' #AILTON - qryInsertCompraItens       - REVISÃO / FERNANDA
'' #AILTON - qryInsertProdutoConsumo    - REVISÃO / FERNANDA
''
'' #PENDENTE - Processamento de arquivos CTE Inclusão de itens
'' #PENDENTE - Validação de campos de compras e itens
'' #PENDENTE - Teste de inclusão em banco SQL com todas as compras

''----------------------------

'' INFO 05/27/2021 17:10:29 - Processamento - Importar Dados Gerais ( Quantidade de registros: 1087 ) - 00:13:52
'' INFO 05/28/2021 15:19:27 - Processamento - Importar Registros Validos ( Quantidade de registros: 562 ) - 00:44:49


Function processamento_dados_gerais()
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
        'Debug.Print contadorDeRegistros

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
    Set DadosGerais = Nothing
    
    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit

End Function

Function processamento_compras_por_arquivos_pendentes()
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
    ''### TESTES COM TODOS OS ARQUIVOS PENDENTES DA BASE DE DADOS GERAIS
    ''#######################################################################################

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdInitMeter, "Pendentes ...", rst.RecordCount

    '' CARREGAR DADOS DE ARQUIVOS NA TABELA DE PROCESSAMENTO
    For Each Item In carregarParametros(DadosGerais.SelectArquivosPendentes)
    
        '' #BARRA_PROGRESSO
        SysCmd acSysCmdUpdateMeter, contadorDeRegistros
    
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
    s.ProcessamentoTransferir strRepositorio
    s.ProcessamentoTransferir strRepositorio & "Item"
    
    '' #BARRA_PROGRESSO
    SysCmd acSysCmdRemoveMeter
    
    '' FORMATAR ITENS DE COMPRA
    DadosGerais.FormatarItensDeCompras
    
    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "Processamento - Importar Registros Validos ( Quantidade de registros: " & contadorDeRegistros & " )"


    MsgBox "Concluido!", vbOKOnly + vbInformation, strRepositorio

adm_Exit:
    Set s = Nothing
    Set DadosGerais = Nothing

    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit

End Function

Function exemplo_processamento_compras_por_arquivos_unicos()
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
    ''### TESTES COM ARQUIVOS PRÉ SELECIONADOS DE TIPOS DIFERENTES
    ''#######################################################################################

    Dim arquivos As Collection: Set arquivos = New Collection

    ''' RETORNO SIMBÓLICO DE MERCADORIA DEPOSITADA EM DEPÓSITO FECHA
    arquivos.add "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210212680452000302550020000886301507884230-nfeproc.xml"

    ''' TRANSF. DE MERCADORIAS
    arquivos.add "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210368365501000296550000000638811001361356-nfeproc.xml"

    ''' #TIPO 01 - CTE
    ''' TRANSPORTE RODOVIARIO
    arquivos.add "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210220147617000494570010009539201999046070-cteproc.xml" '' ---> Pendente testes com itens de compras

    ''' PREST. SERV. TRANSPORTE A ESTABELECIMENTO COMERCIAL
    arquivos.add "C:\temp\Coleta\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210304884082000569570000040073831040073834-cteproc.xml" '' ---> Pendente testes com itens de compras

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
    s.ProcessamentoTransferir strRepositorio
    s.ProcessamentoTransferir strRepositorio & "Item"
    
    '' FORMATAR ITENS DE COMPRA
'    DadosGerais.FormatarItensDeCompras
    
    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "Processamento - Importar Dados Gerais ( Quantidade de registros: " & contadorDeRegistros & " )"


    MsgBox "Concluido!", vbOKOnly + vbInformation, strRepositorio

adm_Exit:
    Set s = Nothing
    Set DadosGerais = Nothing

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





