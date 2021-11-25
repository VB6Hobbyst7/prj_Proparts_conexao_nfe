Attribute VB_Name = "99_Testes_Dev"
Option Compare Database

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
