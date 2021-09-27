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
Dim Item As Variant

Dim strRepositorio As String: strRepositorio = "tblDadosConexaoNFeCTe"

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

'' #CONTADOR
Dim contadorDeRegistros As Long: contadorDeRegistros = 1

Dim strArquivoTeste As String


''' RETORNO SIMBÓLICO DE MERCADORIA DEPOSITADA EM DEPÓSITO FECHA
strArquivoTeste = "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210212680452000302550020000886301507884230-nfeproc.xml" '' --> OK

''' TRANSF. DE MERCADORIAS
'strArquivoTeste = "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210368365501000296550000000638811001361356-nfeproc.xml" '' --> OK
'
''' TRANSPORTE RODOVIARIO
'strArquivoTeste = "C:\temp\Coleta\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210248740351015359570000000309211914301218-cteproc.xml" '' --> OK
'
''' PREST. SERV. TRANSPORTE A ESTABELECIMENTO COMERCIAL
'strArquivoTeste = "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210220147617000494570010009539201999046070-cteproc.xml" '' --> OK
'strArquivoTeste = "C:\temp\Coleta\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210304884082000569570000040073831040073834-cteproc.xml" '' --> OK


    '' LIMPAR TABELA DE PROCESSAMENTOS
    s.DeleteProcessamento
    
    s.ProcessamentoDeArquivo strArquivoTeste, opDadosGerais

'    '' CARREGAR DADOS DE ARQUIVOS NA TABELA DE PROCESSAMENTO
'    For Each Item In GetFilesInSubFolders(DLookup("ValorDoParametro", "tblParametros", "TipoDeParametro='caminhoDeColeta'"))
'        s.ProcessamentoDeArquivo CStr(Item), opDadosGerais
'
'        '' #CONTADOR
'        contadorDeRegistros = contadorDeRegistros + 1
'        Debug.Print contadorDeRegistros
'
'        DoEvents
'    Next Item

    '' IDENTIFICAR CAMPOS
    s.UpdateProcessamentoIdentificarCampos strRepositorio
    
    '' FORMATAR DADOS
    s.UpdateProcessamentoFormatarDados


    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "Processamento - ImportarDados ( Quantidade de registros: " & contadorDeRegistros & " )"


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
