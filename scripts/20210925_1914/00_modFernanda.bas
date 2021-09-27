Attribute VB_Name = "00_modFernanda"
Option Compare Database

''#AILTON - DESATIVADO PARA TESTES DE PROCESSAMENTO
''#AILTON - formatarDados


''----------------------------
'' ### EXEMPLOS DE FUNÇÕES
''
'' 01. testeUnitario_carregarDadosGerais
'' 02. testeUnitario_carregarArquivosPendentes
'' 03. testeUnitario_criacaoArquivosJson
'' 04. testeUnitario_enviarDadosServidor
'' 00. testeUnitario_TratamentoDeArquivosValidos
'' 00. testeUnitario_TratamentoDeArquivosInvalidos
''
'' 99. FUNÇÃO_AUXILIAR: carregarDadosGerais(strArquivo As String)
'' 99. FUNÇÃO_AUXILIAR: carregarArquivosPendentes(strArquivo As String)
'' 99. FUNÇÃO_AUXILIAR: azsMoverArquivo(strArquivo As String, strOrigem As String, strDestino As String)
'' 99. FUNÇÃO_AUXILIAR: azsProcessamentoDeArquivos(sqlArquivos As String, qryUpdate As String, strOrigem As String, strDestino As String)
''
''----------------------------


'' ### TO-DO ###
''
'' #20210823_XML_FORMULARIO | Não encontrei um formulário com os XML´s que não foram processados e o motivo. | <<< ATENÇÃO - NÃO DEFINIMOS COMO CLASSIFICAREMOS OS MOTIMOS DE NÃO PROCESSAMENTO DE ARQUIVOS >>>
'' ID_Prod_CompraNFItem | VALIDAR ATUALIZAÇÃO DESSE CAMPO
'' #20210823_VTotProd_CompraNF


'' ### DONE ###
''
'' Consultas
'' #20210823_EXPORTACAO_LIMITE
'' #20210823_qryUpdateCFOP_PSC_PES -- FiltroCFOP
'' #20210823_qryUpdate_IDVD
'' #20210823_qryUpdateID_NatOp_CompraNF
'' #20210823_qryUpdateCFOP_FilCompra
'' #20210823_qryUpdate_ModeloDoc_CFOP
'' #20210823_qryUpdateFilCompraNF
'' #20210823_qryUpdateIdFornCompraNF
'' #20210823_qryUpdateNumPed_CompraNF
'' #20210823_qryUpdateSit_CompraNF
'' #20210823_XML_CONTROLE | Quando importar cada XML, precisa recortar o arquivo da pasta da empresa e colar dentro de uma pasta chama “Processados”, porém dentro de cada pasta de cada empresa, pois não podemos misturar os XML´s de cada empresa.



'' 01. CARREGAR DADOS GERAIS
Sub testeUnitario_carregarDadosGerais()
Dim Item As Variant


''#######################################################################################
''### REPOSITORIO
''#######################################################################################
Dim DadosGerais As New clsConexaoNfeCte

'' REPOSITORIO
Dim arquivos As Collection: Set arquivos = New Collection

'' DADOS
For Each Item In GetFilesInSubFolders(DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColeta'"))
    arquivos.add CStr(Item)
Next

'''' RETORNO SIMBÓLICO DE MERCADORIA DEPOSITADA EM DEPÓSITO FECHA
'arquivos.add "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210212680452000302550020000886301507884230-nfeproc.xml"
'
'''' TRANSF. DE MERCADORIAS
'arquivos.add "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210368365501000296550000000638811001361356-nfeproc.xml"
'
'''' #TIPO 01 - CTE - TRANSPORTE RODOVIARIO
'arquivos.add "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210220147617000494570010009539201999046070-cteproc.xml"
'
'''' PREST. SERV. TRANSPORTE A ESTABELECIMENTO COMERCIAL
'arquivos.add "C:\temp\Coleta\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210304884082000569570000040073831040073834-cteproc.xml"

''#######################################################################################
''### PROCESSAMENTO
''#######################################################################################

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

'' #CONTADOR
Dim contadorDeRegistros As Long: contadorDeRegistros = 1
Dim totalDeRegistros As Long: totalDeRegistros = arquivos.count

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdInitMeter, "Pendentes ...", totalDeRegistros

    For Each Item In arquivos
    
        carregarDadosGerais CStr(Item)

        '' #BARRA_PROGRESSO
        contadorDeRegistros = contadorDeRegistros + 1
        SysCmd acSysCmdUpdateMeter, contadorDeRegistros
        
        Debug.Print "carregarDadosGerais " & contadorDeRegistros & " - " & CStr(totalDeRegistros)

        DoEvents
    Next Item

    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "carregarDadosGerais - Importar Dados Gerais ( Quantidade de registros: " & contadorDeRegistros & " )"
    
    '' #BARRA_PROGRESSO
    SysCmd acSysCmdRemoveMeter

End Sub

'' 02. PROCESSAR ARQUIVOS VALIDOS E PENDENTES
Sub testeUnitario_carregarArquivosPendentes()
Dim Item As Variant

''#######################################################################################
''### REPOSITORIO
''#######################################################################################
Dim DadosGerais As New clsConexaoNfeCte

'' REPOSITORIO
Dim arquivos As Collection: Set arquivos = New Collection

'' DADOS
For Each Item In carregarParametros(DadosGerais.SelectArquivosPendentes)
    arquivos.add CStr(Item)
Next

''#######################################################################################
''### PROCESSAMENTO
''#######################################################################################

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

'' #CONTADOR
Dim contadorDeRegistros As Long: contadorDeRegistros = 1
Dim totalDeRegistros As Long: totalDeRegistros = arquivos.count

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdInitMeter, "Pendentes ...", totalDeRegistros

    For Each Item In arquivos
    
        carregarArquivosPendentes CStr(Item)

        '' #BARRA_PROGRESSO
        contadorDeRegistros = contadorDeRegistros + 1
        SysCmd acSysCmdUpdateMeter, contadorDeRegistros
        
        Debug.Print "carregarArquivosPendentes " & contadorDeRegistros & " - " & totalDeRegistros

        DoEvents
    Next Item

''#######################################################################################
''### FORMATAR DADOS PROCESSADOS
''#######################################################################################

    '' COMPRAS ATUALIAR CAMPOS
    DadosGerais.compras_atualizarCampos

    '' COMPRAS ITENS CTE
    DadosGerais.compras_carregarItensCTe

    '' FORMATAR ITENS DE COMPRA
    DadosGerais.FormatarItensDeCompras

    '' CADASTRO DE NUMERO DE PEDIDOS
    DadosGerais.UpdateNumPed_CompraNF
    

    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "carregarArquivosPendentes - Importar arquivos pendentes ( Quantidade de registros: " & contadorDeRegistros & " )"
    
    '' #BARRA_PROGRESSO
    SysCmd acSysCmdRemoveMeter

End Sub

'' 03. GERAR ARQUIVOS JSONs
Sub testeUnitario_criacaoArquivosJson()
Dim s As New clsCriarArquivos
Dim qrySelectRegistroValido As String: qrySelectRegistroValido = _
            "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi]))>0) AND ((tblDadosConexaoNFeCTe.registroValido)=1));"

Dim strCaminhoDeSaida As String: strCaminhoDeSaida = "C:\temp\" & strControle
CreateDir strCaminhoDeSaida
    
    '' NO PROCESSAMENTO DO ARQUIVO DE XML
    s.criarArquivoJson opFlagLancadaERP, qrySelectRegistroValido, strCaminhoDeSaida

    '' SELEÇÃO PELO USUARIO
    s.criarArquivoJson opManifesto, qrySelectRegistroValido, strCaminhoDeSaida

    Debug.Print "Concluido! - criacaoArquivosJson"

Cleanup:
    Set s = Nothing

End Sub

'' 04. ENVIAR DADOS PARA SERVIDOR
Sub testeUnitario_enviarDadosServidor()

''#######################################################################################
''### PROCESSAMENTO
''#######################################################################################

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

    '' CADASTRO DE CABEÇALHO DE COMPRAS
    enviar_ComprasParaServidor "tblCompraNF"

    '' RELACIONAMENTO DE ID_COMPRAS COM CHAVES DE ACESSO CADASTRADAS DO SERVIDOR
    criarTabelaTemporariaParaRelacionarIdCompraComChvAcesso
    relacionarIdCompraComChvAcesso
    
    '' CADASTRO DE ITENS DE COMPRAS
    enviar_ComprasParaServidor "tblCompraNFItem"
    
    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "enviarDadosServidor"

    Debug.Print "Concluido! - enviarDadosServidor"

End Sub


'' #20210823_XML_CONTROLE
Sub testeUnitario_TratamentoDeArquivosValidos()

''==================================================
'' PROCESSAMENTO DE ARQUVOS VALIDOS
''==================================================

'' Listagem de arquivos
Dim sqlArquivosValidos As String: sqlArquivosValidos = _
    "SELECT ChvAcesso, CaminhoDoArquivo FROM  tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=2));"
Debug.Print sqlArquivosValidos


'' Atualização de processamento
Dim qryUpdateProcessado As String: qryUpdateProcessado = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 3 WHERE tblDadosConexaoNFeCTe.ChvAcesso =""strChave"";"
Debug.Print qryUpdateProcessado


'' Caminhos para tipos de arquivos
Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColeta'")
Dim strDestino As String: strDestino = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeProcessados'")


    azsProcessamentoDeArquivos sqlArquivosValidos, qryUpdateProcessado, strOrigem, strDestino

End Sub


'' #20210823_XML_CONTROLE
Sub testeUnitario_TratamentoDeArquivosInvalidos()

''==================================================
'' PROCESSAMENTO DE ARQUVOS INVALIDOS - EXPURGO
''==================================================

'' Listagem de arquivos
Dim sqlArquivosInvalidosNaoProcessado As String: sqlArquivosInvalidosNaoProcessado = _
    "SELECT ID AS ChvAcesso, CaminhoDoArquivo FROM  tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=0) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0));"
Debug.Print sqlArquivosInvalidosNaoProcessado


'' Atualização de expurgo
Dim qryUpdateExpurgo As String: qryUpdateExpurgo = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 4 WHERE tblDadosConexaoNFeCTe.ID =strChave;"
Debug.Print qryUpdateExpurgo


'' Caminhos para tipos de arquivos
Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColeta'")
Dim strDestino As String: strDestino = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeExpurgo'")


    azsProcessamentoDeArquivos sqlArquivosInvalidosNaoProcessado, qryUpdateExpurgo, strOrigem, strDestino

End Sub


''=======================================================================================================
'' LIB
''=======================================================================================================

Function carregarDadosGerais(strArquivo As String)
On Error GoTo adm_Err

Dim s As New clsProcessamentoDados
Dim DadosGerais As New clsConexaoNfeCte
Dim Item As Variant
Dim strRepositorio As String

    ''#######################################################################################
    ''### ENVIAR DADOS DE ARQUIVOS PARA TABELA DE PROCESSAMENTO
    ''#######################################################################################
    
    '' REPOSITORIO
    strRepositorio = "tblDadosConexaoNFeCTe"

    '' LIMPAR TABELA DE PROCESSAMENTOS
    s.DeleteProcessamento

    '' PROCESSAMENTO
    s.ProcessamentoDeArquivo strArquivo, opDadosGerais

    '' IDENTIFICAR CAMPOS
    s.UpdateProcessamentoIdentificarCampos strRepositorio
    
    '' FORMATAR DADOS
    s.UpdateProcessamentoFormatarDados
        
        
    ''#######################################################################################
    ''### TRANSFERIR DADOS PROCESSADOS PARA REPOSITORIO
    ''#######################################################################################
        
    '' TRANSFERENCIA DE DADOS
    s.ProcessamentoTransferir strRepositorio
    
    ''#######################################################################################
    ''### TRATAMENTO DE DADOS IMPORTADOS
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

Function carregarArquivosPendentes(strArquivo As String)
On Error GoTo adm_Err

Dim s As New clsProcessamentoDados
'Dim DadosGerais As New clsConexaoNfeCte
Dim strRepositorio As String
    
    ''#######################################################################################
    ''### ENVIAR DADOS DE ARQUIVOS PARA TABELA DE PROCESSAMENTO
    ''#######################################################################################

    '' REPOSITORIO
    strRepositorio = "tblCompraNF"

    '' LIMPAR TABELA DE PROCESSAMENTOS
    s.DeleteProcessamento

    '' PROCESSAMENTO
    s.ProcessamentoDeArquivo strArquivo, opCompras

    '' IDENTIFICAR CAMPOS
    s.UpdateProcessamentoIdentificarCampos strRepositorio
    
    '' CORREÇÃO DE DADOS MARCADOS ERRADOS EM ITENS DE COMPRAS
    s.UpdateProcessamentoLimparItensMarcadosErrados
    
    '' IDENTIFICAR CAMPOS DE ITENS DE COMPRAS
    s.UpdateProcessamentoIdentificarCampos strRepositorio & "Item"
    
    '' FORMATAR DADOS
    s.UpdateProcessamentoFormatarDados
    
    ''#######################################################################################
    ''### TRANSFERIR DADOS PROCESSADOS PARA REPOSITORIO
    ''#######################################################################################

    '' TRANSFERIR DADOS PROCESSADOS
    s.ProcessamentoTransferir strRepositorio
    s.ProcessamentoTransferir strRepositorio & "Item"

adm_Exit:
    Set s = Nothing
    Set DadosGerais = Nothing

    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit

End Function

Function azsMoverArquivo(strArquivo As String, strOrigem As String, strDestino As String)
On Error GoTo adm_Err

Dim objFSO As Object: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim tmpDestino As String

    '' 01. CRIAR PASTA DESTINO
    tmpDestino = getPath(Replace(strArquivo, strOrigem, strDestino))
    CreateDir tmpDestino
    
    '' 02. MOVER ARQUIVO ORIGINAL PARA A PASTA DESTINO
    If (Dir(tmpDestino & getFileNameAndExt(strArquivo)) <> "") Then Kill tmpDestino & getFileNameAndExt(strArquivo)
    objFSO.CopyFile strArquivo, tmpDestino & getFileNameAndExt(strArquivo)
    
    '' 03. REMOVER ARQUIVO ORIGINAL
    If (Dir(strArquivo) <> "") Then Kill strArquivo

adm_Exit:
    Exit Function

adm_Err:
    Debug.Print Error$ & " - " & strArquivo
    Resume adm_Exit
    
End Function

Function azsProcessamentoDeArquivos(sqlArquivos As String, qryUpdate As String, strOrigem As String, strDestino As String)

Dim db As DAO.Database: Set db = CurrentDb
Dim rstArquivos As DAO.Recordset

Dim qryTemp As String

    Set rstArquivos = db.OpenRecordset(sqlArquivos)
    Do While Not rstArquivos.EOF
        
        azsMoverArquivo rstArquivos.Fields("CaminhoDoArquivo").value, strOrigem, strDestino
        
        qryTemp = Replace(qryUpdate, "strChave", rstArquivos.Fields("ChvAcesso").value)
        Debug.Print qryTemp
        
        Application.CurrentDb.Execute qryTemp
        
        rstArquivos.MoveNext
        DoEvents
    Loop
            
    rstArquivos.Close
    db.Close
    
    Set rstArquivos = Nothing
    Set db = Nothing

End Function


'''#AILTON - formatarDados
'Sub formatarDados()
'Dim DadosGerais As New clsConexaoNfeCte
'
''' #ANALISE_DE_PROCESSAMENTO
'Dim DT_PROCESSO As Date: DT_PROCESSO = Now()
'
'    '' COMPRAS ATUALIAR CAMPOS
'    Debug.Print "#################### DadosGerais.compras_atualizarCampos"
'    DadosGerais.compras_atualizarCampos
'
'    '' COMPRAS ITENS CTE
'    Debug.Print "#################### DadosGerais.compras_carregarItensCTe"
'    DadosGerais.compras_carregarItensCTe
'
'    '' FORMATAR ITENS DE COMPRA
'    Debug.Print "#################### DadosGerais.FormatarItensDeCompras"
'    DadosGerais.FormatarItensDeCompras
'
'    '' CADASTRO DE NUMERO DE PEDIDOS
'    Debug.Print "#################### DadosGerais.UpdateNumPed_CompraNF"
'    DadosGerais.UpdateNumPed_CompraNF
'
'    '' #ANALISE_DE_PROCESSAMENTO
'    statusFinal DT_PROCESSO, "formatarDados"
'
'
'End Sub

