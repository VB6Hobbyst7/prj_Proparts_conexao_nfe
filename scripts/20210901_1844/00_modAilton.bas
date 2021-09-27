Attribute VB_Name = "00_modAilton"
Option Compare Database

''----------------------------
'' ### EXEMPLOS DE FUNÇÕES
''
'' 01. processamento_dados_gerais
'' 02. processamento_compras_por_arquivos_pendentes
'' 03. enviar_compras_para_servidor
'' 04. exemplos_criacao_arquivos_json
''
''----------------------------

Sub testeUnitario()
Dim arquivos As Collection: Set arquivos = New Collection

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

'' #CONTADOR
Dim contadorDeRegistros As Long: contadorDeRegistros = 1


    ''#######################################################################################
    ''### BASE DE TESTES
    ''#######################################################################################

    ''' RETORNO SIMBÓLICO DE MERCADORIA DEPOSITADA EM DEPÓSITO FECHA
    arquivos.add "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210212680452000302550020000886301507884230-nfeproc.xml"

    ''' TRANSF. DE MERCADORIAS
    arquivos.add "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210368365501000296550000000638811001361356-nfeproc.xml"

    ''' #TIPO 01 - CTE
    ''' TRANSPORTE RODOVIARIO
    arquivos.add "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210220147617000494570010009539201999046070-cteproc.xml"

    ''' PREST. SERV. TRANSPORTE A ESTABELECIMENTO COMERCIAL
    arquivos.add "C:\temp\Coleta\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210304884082000569570000040073831040073834-cteproc.xml"

    
    ''#######################################################################################
    ''### PROCESSAMENTO
    ''#######################################################################################

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdInitMeter, "Pendentes ...", arquivos.count

    For Each Item In arquivos
        
        carregarDadosGerais CStr(Item)

        '' #CONTADOR
        contadorDeRegistros = contadorDeRegistros + 1

        '' #BARRA_PROGRESSO
        SysCmd acSysCmdUpdateMeter, contadorDeRegistros

        DoEvents
    Next Item


    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "Processamento - Importar Dados Gerais ( Quantidade de registros: " & contadorDeRegistros & " )"
    
    '' #BARRA_PROGRESSO
    SysCmd acSysCmdRemoveMeter

    MsgBox "Concluido!", vbOKOnly + vbInformation, strRepositorio

End Sub


Function carregarDadosGerais(strArquivo As String)
On Error GoTo adm_Err
Dim s As New clsProcessamentoDados
Dim DadosGerais As New clsConexaoNfeCte
Dim Item As Variant

Dim strRepositorio As String: strRepositorio = "tblDadosConexaoNFeCTe"

    '' LIMPAR TABELA DE PROCESSAMENTOS
    s.DeleteProcessamento
    
    ''#######################################################################################
    ''### CARREGAR DADOS DE ARQUIVOS PARA TABELA DE PROCESSAMENTO
    ''#######################################################################################

    '' PROCESSAMENTO
    s.ProcessamentoDeArquivo CStr(Item), opDadosGerais

    '' IDENTIFICAR CAMPOS
    s.UpdateProcessamentoIdentificarCampos strRepositorio
    
    '' FORMATAR DADOS
    s.UpdateProcessamentoFormatarDados
        
    '' TRANSFERIR DADOS PROCESSADOS PARA REPOSITORIO
    s.ProcessamentoTransferir strRepositorio
    
    ''#######################################################################################
    ''### CLASSIFICAR DADOS EM TABELA DE DADOS GERAIS
    ''#######################################################################################
    
    '' CLASSIFICAR DADOS GERAIS
    DadosGerais.TratamentoDeDadosGerais


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

Dim db As DAO.Database: Set db = CurrentDb
Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(DadosGerais.SelectArquivosPendentes)

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

'' #CONTADOR
Dim contadorDeRegistros As Long: contadorDeRegistros = 0

    '' LIMPAR TABELA DE PROCESSAMENTOS
    s.DeleteProcessamento

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdInitMeter, "Pendentes ...", rst.RecordCount

    ''#######################################################################################
    ''### CARREGAR DADOS DE ARQUIVOS COM BASE EM ITENS (PENDENTES) DA TABELA DE DADOS GERAIS
    ''#######################################################################################

    '' PROCESSAMENTO
    For Each Item In carregarParametros(DadosGerais.SelectArquivosPendentes)
    
        '' #BARRA_PROGRESSO
        SysCmd acSysCmdUpdateMeter, contadorDeRegistros
    
        s.ProcessamentoDeArquivo CStr(Item), opCompras

        '' #CONTADOR
        contadorDeRegistros = contadorDeRegistros + 1

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
    
    '' COMPRAS ATUALIAR CAMPOS
    DadosGerais.compras_atualizarCampos
    
    '' COMPRAS ITENS CTE
    DadosGerais.compras_carregarItensCTe

    ''#######################################################################################
    ''### FORMATAR DADOS PROCESSADOS
    ''#######################################################################################

    '' FORMATAR ITENS DE COMPRA
    DadosGerais.FormatarItensDeCompras
    
    '' CADASTRO DE NUMERO DE PEDIDOS
    DadosGerais.UpdateNumPed_CompraNF
        
    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "Processamento - Importar Registros Validos ( Quantidade de registros: " & contadorDeRegistros & " )"

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdRemoveMeter

    MsgBox "Concluido!", vbOKOnly + vbInformation, strRepositorio

adm_Exit:
    Set s = Nothing
    Set DadosGerais = Nothing

    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit

End Function

Sub enviar_compras_para_servidor()

'' CADASTRO DE CABEÇALHO DE COMPRAS
enviar_ComprasParaServidor "tblCompraNF"

'' RELACIONAMENTO DE ID_COMPRAS COM CHAVES DE ACESSO CADASTRADAS DO SERVIDOR
criarTabelaTemporariaParaRelacionarIdCompraComChvAcesso
relacionarIdCompraComChvAcesso

'' CADASTRO DE ITENS DE COMPRAS
enviar_ComprasParaServidor "tblCompraNFItem"

End Sub


Sub exemplos_criacao_arquivos_json()
Dim s As New clsCriarArquivos
Dim qrySelectRegistroValido As String: qrySelectRegistroValido = _
            "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi]))>0) AND ((tblDadosConexaoNFeCTe.registroValido)=1));"

Dim strCaminhoDeSaida As String: strCaminhoDeSaida = "C:\temp\" & strControle
CreateDir strCaminhoDeSaida
    
    '' NO PROCESSAMENTO DO ARQUIVO DE XML
    s.criarArquivoJson opFlagLancadaERP, qrySelectRegistroValido, strCaminhoDeSaida

    '' SELEÇÃO PELO USUARIO
    s.criarArquivoJson opManifesto, qrySelectRegistroValido, strCaminhoDeSaida


    MsgBox "Concluido!", vbOKOnly + vbInformation, "teste_arquivos_json"

Cleanup:

    Set s = Nothing

End Sub



