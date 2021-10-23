Attribute VB_Name = "01_modFernanda"
Option Compare Database


'' ### TO-DO ###
''
'' #20210823_NumPed_CompraNF

'' ### DONE ###
''
'' Consultas
'' #20210823_EXPORTACAO_LIMITE
'' #20210823_qryDadosGerais_Update_ID_NatOp_CompraNF__FiltroCFOP -- FiltroCFOP
'' #20210823_qryDadosGerais_Update_IDVD
'' #20210823_qryUpdateID_NatOp_CompraNF
'' #20210823_qryUpdateCFOP_FilCompra
'' #20210823_qryUpdate_ModeloDoc_CFOP
'' #20210823_qryUpdateFilCompraNF
'' #20210823_qryDadosGerais_Update_IdFornCompraNF
'' #20210823_qryUpdateNumPed_CompraNF
'' #20210823_qryDadosGerais_Update_Sit_CompraNF
'' #20210823_XML_CONTROLE | Quando importar cada XML, precisa recortar o arquivo da pasta da empresa e colar dentro de uma pasta chama �Processados�, por�m dentro de cada pasta de cada empresa, pois n�o podemos misturar os XML�s de cada empresa.
'' #20210823_XML_FORMULARIO | N�o encontrei um formul�rio com os XML�s que n�o foram processados e o motivo. | <<< ATEN��O - N�O DEFINIMOS COMO CLASSIFICAREMOS OS MOTIMOS DE N�O PROCESSAMENTO DE ARQUIVOS >>>
'' #20210823_VTotProd_CompraNF
'' #20210823_ID_Prod_CompraNFItem


''----------------------------
'' ### EXEMPLOS DE FUN��ES
''
'' 01. processarDadosGerais
'' 02. processarArquivosPendentes
'' 04. enviarDadosServidor
'' 05. tratamentoDeArquivosValidos
'' 06. tratamentoDeArquivosInvalidos
'' 07. criacaoArquivosJson
''
'' 99. FUN��O_AUXILIAR: carregarDadosGerais(strArquivo As String)
'' 99. FUN��O_AUXILIAR: carregarArquivosPendentes(strArquivo As String)
'' 99. FUN��O_AUXILIAR: azsProcessamentoDeArquivos(sqlArquivos As String, qryUpdate As String, strOrigem As String, strDestino As String)
'' 99. FUN��O_AUXILIAR: tratamentoDeArquivosValidos()
'' 99. FUN��O_AUXILIAR: tratamentoDeArquivosInvalidos()
''
''----------------------------


'' 01. PROCESSAR DADOS GERAIS
Sub processarDadosGerais()
On Error GoTo adm_Err

Dim DadosGerais As New clsConexaoNfeCte
Dim arquivos As Collection: Set arquivos = New Collection

Dim caminhoNovo As Variant
Dim caminhoAntigo As Variant
Dim item As Variant

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

'' #CONTADOR
Dim contadorDeRegistros As Long: contadorDeRegistros = 0
Dim totalDeRegistros As Long

'' #REPOSITORIOS
DadosGerais.CriarRepositorios

''#######################################################################################
''### REPOSITORIO
''#######################################################################################

'' REPOSITORIOS
For Each caminhoAntigo In Array(DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColeta'"))
    For Each caminhoNovo In carregarParametros(DadosGerais.SelectColetaEmpresa)
        For Each item In GetFilesInSubFolders(CStr(Replace(Replace(caminhoAntigo, "empresa", caminhoNovo), "recebimento\", "")))
            arquivos.add CStr(item)
        Next
    Next
Next

'''' RETORNO SIMB�LICO DE MERCADORIA DEPOSITADA EM DEP�SITO FECHA
'arquivos.add "C:\xmls\68.365.5010003-77 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210212680452000302550020000886301507884230-nfeproc.xml"
'
'''' TRANSF. DE MERCADORIAS
'arquivos.add "C:\xmls\68.365.5010003-77 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\recebimento\32210368365501000296550000000638811001361356-nfeproc.xml"
'
'''' #TIPO 01 - CTE - TRANSPORTE RODOVIARIO
'arquivos.add "C:\xmls\68.365.5010003-77 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210220147617000494570010009539201999046070-cteproc.xml"
'
'''' PREST. SERV. TRANSPORTE A ESTABELECIMENTO COMERCIAL
'arquivos.add "C:\xmls\68.365.5010002-96 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\recebimento\32210304884082000569570000040073831040073834-cteproc.xml"

''#######################################################################################
''### PROCESSAMENTO
''#######################################################################################
totalDeRegistros = arquivos.count

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdInitMeter, "Pendentes ...", totalDeRegistros

    For Each item In arquivos
    
        carregarDadosGerais CStr(item)

        '' #BARRA_PROGRESSO
        contadorDeRegistros = contadorDeRegistros + 1
        SysCmd acSysCmdUpdateMeter, contadorDeRegistros
        
        Debug.Print "carregarDadosGerais " & contadorDeRegistros & " - " & CStr(totalDeRegistros)
        If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "carregarDadosGerais " & contadorDeRegistros & " - " & CStr(totalDeRegistros)

        DoEvents
    Next item
    
    '' CLASSIFICAR DADOS GERAIS
    DadosGerais.TratamentoDeDadosGerais

    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "carregarDadosGerais - Importar Dados Gerais ( Quantidade de registros: " & contadorDeRegistros & " )"
    
    '' #BARRA_PROGRESSO
    SysCmd acSysCmdRemoveMeter

adm_Exit:
    Set DadosGerais = Nothing
    Set arquivos = Nothing
    
    Exit Sub

adm_Err:
    Debug.Print "processarDadosGerais() - " & Err.Description
    Resume adm_Exit
    
End Sub

'' 02. PROCESSAR ARQUIVOS VALIDOS E PENDENTES
Sub processarArquivosPendentes()
On Error GoTo adm_Err

Dim DadosGerais As New clsConexaoNfeCte
Dim arquivos As Collection: Set arquivos = New Collection

Dim item As Variant

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

'' #CONTADOR
Dim contadorDeRegistros As Long: contadorDeRegistros = 0
Dim totalDeRegistros As Long

''#######################################################################################
''### REPOSITORIO
''#######################################################################################

'' REPOSITORIO
For Each item In carregarParametros(DadosGerais.SelectArquivosPendentes)
    arquivos.add CStr(item)
Next

''#######################################################################################
''### PROCESSAMENTO
''#######################################################################################
totalDeRegistros = arquivos.count

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdInitMeter, "Pendentes ...", totalDeRegistros

    For Each item In arquivos
    
        carregarArquivosPendentes CStr(item)

        '' #BARRA_PROGRESSO
        contadorDeRegistros = contadorDeRegistros + 1
        SysCmd acSysCmdUpdateMeter, contadorDeRegistros
        
        Debug.Print "carregarArquivosPendentes " & contadorDeRegistros & " - " & totalDeRegistros
        If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "carregarDadosGerais " & contadorDeRegistros & " - " & CStr(totalDeRegistros)

        DoEvents
    Next item

''#######################################################################################
''### FORMATAR DADOS PROCESSADOS
''#######################################################################################

    '' COMPRAS ATUALIAR CAMPOS
    DadosGerais.compras_atualizarCampos

    '' CADASTRO DE NUMERO DE PEDIDOS
    DadosGerais.UpdateNumPed_CompraNF
       
    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "carregarArquivosPendentes - Importar arquivos pendentes ( Quantidade de registros: " & contadorDeRegistros & " )"
    
    '' #BARRA_PROGRESSO
    SysCmd acSysCmdRemoveMeter
        
adm_Exit:
    Set DadosGerais = Nothing
    Set arquivos = Nothing
    
    Exit Sub

adm_Err:
    Debug.Print "processarArquivosPendentes() - " & Err.Description
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "processarArquivosPendentes() - " & Err.Description
    Resume adm_Exit

End Sub


'' 04. ENVIAR DADOS PARA SERVIDOR
Sub enviarDadosServidor()

''==================================================
''### PROCESSAMENTO
''==================================================

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

    '' CADASTRO DE CABE�ALHO DE COMPRAS
    enviar_ComprasParaServidor "tblCompraNF"

    '' RELACIONAMENTO DE ID_COMPRAS COM CHAVES DE ACESSO CADASTRADAS DO SERVIDOR
    criarTabelaTemporariaParaRelacionarIdCompraComChvAcesso
    relacionarIdCompraComChvAcesso
    
    '' CADASTRO DE ITENS DE COMPRAS
    enviar_ComprasParaServidor "tblCompraNFItem"
    
    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "enviarDadosServidor"

    Debug.Print "Concluido! - enviarDadosServidor"
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "Concluido! - enviarDadosServidor"

End Sub

'' #20210823_XML_CONTROLE
'' 05. TRATAMENTO DE ARQUIVOS VALIDOS
Sub tratamentoDeArquivosValidos()
Dim DadosGerais As New clsConexaoNfeCte

''==================================================
''### PROCESSAMENTO DE ARQUVOS VALIDOS
''==================================================

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

    azsProcessamentoDeArquivos DadosGerais.SelectArquivosValidos, DadosGerais.UpdateProcessado

    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "TratamentoDeArquivosValidos"

    Set DadosGerais = Nothing

End Sub


'' #20210823_XML_CONTROLE
'' 06. TRATAMENTO DE ARQUIVOS INVALIDOS
Sub tratamentoDeArquivosInvalidos()
Dim DadosGerais As New clsConexaoNfeCte

''==================================================
'' PROCESSAMENTO DE ARQUVOS INVALIDOS - EXPURGO
''==================================================

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

    azsProcessamentoDeArquivos DadosGerais.SelectArquivosInvalidos, DadosGerais.UpdateExpurgo

    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "TratamentoDeArquivosInvalidos"

    Set DadosGerais = Nothing

End Sub

'' 07. GERAR ARQUIVOS JSONs
Sub gerarArquivosJson(pArquivo As enumTipoArquivo, Optional strConsulta As String, Optional strCaminho As String)
Dim s As New clsCriarArquivos
Dim strCaminhoDeSaida As String

Dim qrySelectRegistroValido As String: qrySelectRegistroValido = _
    "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi]))>0) AND ((tblDadosConexaoNFeCTe.registroValido)=1))"

    '' SELE��O DE REGISTRO
    If strConsulta <> "" Then
        qrySelectRegistroValido = "SELECT * FROM (" & qrySelectRegistroValido & ") AS tmpSelecao WHERE tmpSelecao.ChvAcesso =  '" & strConsulta & "';"
    Else
        qrySelectRegistroValido = _
                    "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi]))>0) AND ((tblDadosConexaoNFeCTe.registroValido)=1));"
    End If
    
    Debug.Print qrySelectRegistroValido
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qrySelectRegistroValido

    '' CAMINHO DE SAIDA DO ARQUIVO
    If strCaminho <> "" Then
        strCaminhoDeSaida = _
            strCaminho
    Else
        strCaminhoDeSaida = _
            DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColetaAcoes'")
    End If
    CreateDir strCaminhoDeSaida
    
    '' EXECUCAO
    s.criarArquivoJson pArquivo, qrySelectRegistroValido, strCaminhoDeSaida

'    '' SELE��O PELO USUARIO
'    s.criarArquivoJson opManifesto, qrySelectRegistroValido, strCaminhoDeSaida

    Debug.Print "Concluido! - criacaoArquivosJson"
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "Concluido! - criacaoArquivosJson"

Cleanup:
    Set s = Nothing

End Sub


''=======================================================================================================
'' LIB
''=======================================================================================================

Function carregarDadosGerais(strArquivo As String)
On Error GoTo adm_Err

Dim s As New clsProcessamentoDados
Dim DadosGerais As New clsConexaoNfeCte
Dim item As Variant
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
Dim strRepositorio As String
    
    ''#######################################################################################
    ''### ENVIAR DADOS DE ARQUIVOS PARA TABELA DE PROCESSAMENTO
    ''#######################################################################################

    '' REPOSITORIO
    strRepositorio = "tblCompraNF"

'' #AZS - DESLIGAR
    '' LIMPAR TABELA DE PROCESSAMENTOS
    s.DeleteProcessamento

    '' PROCESSAMENTO
    s.ProcessamentoDeArquivo strArquivo, opCompras

    '' IDENTIFICAR CAMPOS
    s.UpdateProcessamentoIdentificarCampos strRepositorio
    
    '' CORRE��O DE DADOS MARCADOS ERRADOS EM ITENS DE COMPRAS
    s.UpdateProcessamentoLimparItensMarcadosErrados
    
    '' IDENTIFICAR CAMPOS DE ITENS DE COMPRAS
    s.UpdateProcessamentoIdentificarCampos strRepositorio & "Item"
    
    '' FORMATAR DADOS
    s.UpdateProcessamentoFormatarDados
    
    ''#######################################################################################
    ''### TRANSFERIR DADOS PROCESSADOS PARA REPOSITORIO
    ''#######################################################################################

'' #AZS - DESLIGAR
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

