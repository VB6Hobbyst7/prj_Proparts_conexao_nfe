<<<<<<< HEAD
Attribute VB_Name = "01_modFernanda"
Option Compare Database

Private Const DATE_TIME_FORMAT               As String = "yyyy/mm/dd hh:mm:ss"
Private Const DATE_FORMAT                    As String = "yyyy/mm/dd"

'' 01. PROCESSAR DADOS GERAIS
Sub processarDadosGerais(Optional pArquivos As Collection)
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
    If pArquivos.count > 0 Then
        For Each item In pArquivos
            arquivos.add CStr(item)
        Next
        
    Else
        For Each caminhoAntigo In Array(DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColeta'"))
            For Each caminhoNovo In carregarParametros(DadosGerais.SelectColetaEmpresa)
                For Each item In GetFilesInSubFolders(CStr(Replace(Replace(caminhoAntigo, "empresa", caminhoNovo), "recebimento\", "")))
                    arquivos.add CStr(item)
                Next
            Next
        Next
        
    End If

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


'' 03. ENVIAR DADOS PARA SERVIDOR
'' #20210823_CadastroDeComprasEmServidor
Sub CadastroDeComprasEmServidor()

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

'' BANCO_DESTINO
Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
Dim dbDestino As New Banco

'' BANCO_LOCAL
Dim Scripts As New clsConexaoNfeCte
Dim db As DAO.Database: Set db = CurrentDb
Dim rstChvAcesso As DAO.Recordset: Set rstChvAcesso = db.OpenRecordset(Replace(Scripts.SelectRegistroValidoPorcessado, "pRepositorio", "tblCompraNF"))

''' #AILTON - TESTES
'Dim TMP2 As String: TMP2 = "select * from (" & Replace(Scripts.SelectRegistroValidoPorcessado, "pRepositorio", "tblCompraNF") & ") as tmp where tmp.ChvAcesso_CompraNF = '42210300634453001303570010001139451001171544'"
'Dim rstChvAcesso As DAO.Recordset: Set rstChvAcesso = db.OpenRecordset(TMP2)

Dim qryCompras_Insert_Compras As String
Dim qryComprasItens_Update_IDCompraNF As String

'' CONTROLE DE "NumPed_CompraNF" RELACIONADO COM REGISTROS DO SERVIDOR
Dim contador As Long: contador = 1

'' CONTROLE DE REPOSITORIOS x CHAVES DE ACESSO
Dim item As Variant
Dim pChvAcesso As String
Dim pRepositorio As String

'' SCRIPT DE INCLUSÃO DE DADOS NO SERVIDOR
Dim tmpCamposNomes As String
Dim strCamposNomes As String
Dim strCamposNomesTmp As String
Dim strRepositorio As String
Dim tmpScript As String: _
    tmpScript = "INSERT INTO pRepositorio ( strCamposNomes ) SELECT strCamposTmp FROM ( VALUES strCamposValores ) AS TMP ( strCamposTmp ) LEFT JOIN pRepositorio ON pRepositorio.ChvAcesso_CompraNF = tmp.ChvAcesso WHERE pRepositorio.ChvAcesso_CompraNF IS NULL;"

Dim tmpScriptItens As String: _
    tmpScriptItens = "INSERT INTO pRepositorio ( strCamposNomes ) SELECT strCamposTmp FROM ( VALUES strCamposValores ) AS TMP ( strCamposTmp );"

Dim qryComprasCTe_Update_AjustesCampos As String: _
    qryComprasCTe_Update_AjustesCampos = "UPDATE tblCompraNF SET tblCompraNF.HoraEntd_CompraNF = NULL ,tblCompraNF.IDVD_CompraNF = NULL WHERE (((tblCompraNF.ChvAcesso_CompraNF) IN (pLista_ChvAcesso_CompraNF)));"

'' #CONTADOR
Dim contadorDeRegistros As Long: contadorDeRegistros = 1
Dim totalDeRegistros As Long

'' VALIDAR CONCILIAÇÃO
Dim TMP As String

    '' BANCO_DESTINO
    dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
    
    '' #BARRA_PROGRESSO
    totalDeRegistros = rstChvAcesso.RecordCount
    SysCmd acSysCmdInitMeter, "Exportação...", totalDeRegistros

    '' CADASTRO
    Do While Not rstChvAcesso.EOF
        pRepositorio = "tblCompraNF"
        pChvAcesso = rstChvAcesso.Fields("ChvAcesso_CompraNF").value
        
        tmpCamposNomes = carregarCamposNomes(pRepositorio)
        strCamposNomes = Replace(tmpScript, "strCamposNomes", tmpCamposNomes)
        strCamposNomesTmp = Replace(Replace(tmpScript, "strCamposNomes", tmpCamposNomes), "strCamposTmp", Replace(tmpCamposNomes, "_" & right(pRepositorio, Len(pRepositorio) - 3), ""))
        strRepositorio = Replace(strCamposNomesTmp, "pRepositorio", pRepositorio)
        
        '' CONTADOR
        dbDestino.SqlSelect "SELECT max(NumPed_CompraNF)+1 as contador from tblCompraNF"
        contador = IIf(IsNull(dbDestino.rs.Fields("contador").value), 1, dbDestino.rs.Fields("contador").value)
        
        '' CADASTRO DE COMPRAS
        For Each item In carregarCamposValores(pRepositorio, pChvAcesso)
            TMP = Replace(Replace(strRepositorio, "strCamposValores", item), "strNumPed_CompraNF", contador)
            Debug.Print TMP
            dbDestino.SqlExecute TMP
        Next item
               
        '' RELACIONAR ITENS DE COMPRAS COM COMPRAS JÁ CADASTRADAS NO SERVIDOR
        dbDestino.SqlSelect "SELECT ChvAcesso_CompraNF,ID_CompraNF FROM tblCompraNF where ChvAcesso_CompraNF = '" & pChvAcesso & "';"
        qryComprasItens_Update_IDCompraNF = Replace(Replace(Scripts.UpdateComprasItens_IDCompraNF, "strChave", pChvAcesso), "strID_Compra", dbDestino.rs.Fields("ID_CompraNF").value)
        If Not dbDestino.rs.EOF Then
            pRepositorio = "tblCompraNFItem"
            
            tmpCamposNomes = carregarCamposNomes(pRepositorio)
            strCamposNomes = Replace(tmpScriptItens, "strCamposNomes", tmpCamposNomes)
            strCamposNomesTmp = Replace(Replace(tmpScriptItens, "strCamposNomes", tmpCamposNomes), "strCamposTmp", Replace(tmpCamposNomes, "_" & right(pRepositorio, Len(pRepositorio) - 3), ""))
            strRepositorio = Replace(strCamposNomesTmp, "pRepositorio", pRepositorio)

            '' #20210823_qryUpdateNumPed_CompraNF
            Application.CurrentDb.Execute qryComprasItens_Update_IDCompraNF
            
            '' CADASTRO DE ITENS DE COMPRAS
            For Each item In carregarCamposValores(pRepositorio, pChvAcesso)
                TMP = Replace(Replace(strRepositorio, "strCamposValores", item), "strNumPed_CompraNF", contador)
                Debug.Print TMP
                dbDestino.SqlExecute TMP
            Next item

        End If
                
        '' MUDAR STATUS DO REGISTRO
        Application.CurrentDb.Execute Replace(Scripts.compras_atualizarEnviadoParaServidor, "strChave", rstChvAcesso.Fields("chvAcesso_CompraNF").value)

        rstChvAcesso.MoveNext
        contador = contador + 1
        contadorDeRegistros = contadorDeRegistros + 1
        
        '' #BARRA_PROGRESSO
        SysCmd acSysCmdUpdateMeter, contadorDeRegistros
    
        DoEvents
    Loop

    '' #20211122_AjusteDeCampos_CTe
    dbDestino.SqlExecute Replace(qryComprasCTe_Update_AjustesCampos, "pLista_ChvAcesso_CompraNF", carregarComprasCTe)
    dbDestino.SqlExecute "UPDATE tblCompraNF SET HoraEntd_CompraNF = NULL, IDVD_CompraNF = NULL WHERE tblCompraNF.IDVD_CompraNF=0;"
    
    '' #20211128_LimparRepositorios
    '' Limpar repositorio de itens de compras
    Application.CurrentDb.Execute _
            "Delete from tblCompraNFItem"

    '' Limpar repositorio de compras
    Application.CurrentDb.Execute _
            "Delete from tblCompraNF"
    

rstChvAcesso.Close
dbDestino.CloseConnection
db.Close

Set Scripts = Nothing
Set rstChvAcesso = Nothing
Set db = Nothing

'' #ANALISE_DE_PROCESSAMENTO
statusFinal DT_PROCESSO, "CadastroDeComprasEmServidor - Exportar compras ( Quantidade de registros: " & contador & " )"

'' #BARRA_PROGRESSO
SysCmd acSysCmdRemoveMeter

Debug.Print "Concluido!"

End Sub

'' 04. GERAR ARQUIVOS JSONs
Sub gerarArquivosJson(pArquivo As enumTipoArquivo, Optional strConsulta As String, Optional strCaminho As String)
Dim s As New clsCriarArquivos
Dim strCaminhoDeSaida As String

Dim sql_Select_tblDadosConexaoNFeCTe_registroValido As String: sql_Select_tblDadosConexaoNFeCTe_registroValido = _
    "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi]))>0) AND ((tblDadosConexaoNFeCTe.registroValido)=1))"

    '' SELEÇÃO DE REGISTRO
    If strConsulta <> "" Then
        sql_Select_tblDadosConexaoNFeCTe_registroValido = "SELECT * FROM (" & sql_Select_tblDadosConexaoNFeCTe_registroValido & ") AS tmpSelecao WHERE tmpSelecao.ChvAcesso =  '" & strConsulta & "';"
    Else
        sql_Select_tblDadosConexaoNFeCTe_registroValido = _
                    "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi]))>0) AND ((tblDadosConexaoNFeCTe.registroValido)=1));"
    End If
    
    Debug.Print sql_Select_tblDadosConexaoNFeCTe_registroValido
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), sql_Select_tblDadosConexaoNFeCTe_registroValido

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
    s.criarArquivoJson pArquivo, sql_Select_tblDadosConexaoNFeCTe_registroValido, strCaminhoDeSaida

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

Function carregarCamposValores(pRepositorio As String, pChvAcesso As String) As Collection
Set carregarCamposValores = New Collection
Dim Scripts As New clsConexaoNfeCte

'Dim pRepositorio As String: pRepositorio = "tblCompraNFItem"
'Dim pChvAcesso As String: pChvAcesso = "32210368365501000296550000000638791001361285"

'' BANCO_LOCAL
Dim db As DAO.Database: Set db = CurrentDb

Dim rstCampos As DAO.Recordset: Set rstCampos = db.OpenRecordset(Replace(Scripts.SelectCamposNomes, "pRepositorio", pRepositorio))
Dim rstOrigem As DAO.Recordset

'' VALIDAR CONCILIAÇÃO
Dim tmpScript As String
Dim tmpValidarCampo As String: tmpValidarCampo = right(pRepositorio, Len(pRepositorio) - 3)

Dim sqlOrigem As String: sqlOrigem = _
    "Select * from (" & Replace(Scripts.SelectRegistroValidoPorcessado, "pRepositorio", pRepositorio) & ") as tmpRepositorio where tmpRepositorio.ChvAcesso_CompraNF = '" & pChvAcesso & "'"
    
    Set rstOrigem = db.OpenRecordset(sqlOrigem)
    
    rstOrigem.MoveLast
    rstOrigem.MoveFirst
    Do While Not rstOrigem.EOF
        tmpScript = "("
        
        '' LISTAGEM DE CAMPOS
        rstCampos.MoveFirst
        Do While Not rstCampos.EOF
    
            '' CRIAR SCRIPT DE INCLUSAO DE DADOS NA TABELA DESTINO
            '' 2. campos x formatacao
            If InStr(rstCampos.Fields("campo").value, tmpValidarCampo) Then
    
                If InStr(rstCampos.Fields("campo").value, "NumPed_CompraNF") Then tmpScript = tmpScript & "strNumPed_CompraNF,": GoTo pulo
    
                If rstCampos.Fields("formatacao").value = "opTexto" Then
                    tmpScript = tmpScript & "'" & rstOrigem.Fields(rstCampos.Fields("campo").value).value & "',"
    
                ElseIf rstCampos.Fields("formatacao").value = "opNumero" Or rstCampos.Fields("formatacao").value = "opMoeda" Then
                    tmpScript = tmpScript & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", rstOrigem.Fields(rstCampos.Fields("campo").value).value, rstCampos.Fields("valorPadrao").value) & ","
    
                ElseIf rstCampos.Fields("formatacao").value = "opTime" Then
                    tmpScript = tmpScript & "'" & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", Format(rstOrigem.Fields(rstCampos.Fields("campo").value).value, DATE_TIME_FORMAT), rstCampos.Fields("valorPadrao").value) & "',"
    
                ElseIf rstCampos.Fields("formatacao").value = "opData" Then
                    tmpScript = tmpScript & "'" & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", Format(rstOrigem.Fields(rstCampos.Fields("campo").value).value, DATE_FORMAT), rstCampos.Fields("valorPadrao").value) & "',"
    
                End If
    
            End If
            
pulo:
            rstCampos.MoveNext
            DoEvents
        Loop
        
        carregarCamposValores.add left(tmpScript, Len(tmpScript) - 1) & ")"
        rstOrigem.MoveNext
        DoEvents
    Loop

    Set Scripts = Nothing
    rstCampos.Close
    rstOrigem.Close
    db.Close

End Function

Function carregarCamposNomes(pRepositorio As String) As String
Dim Scripts As New clsConexaoNfeCte

'' BANCO_LOCAL
Dim db As DAO.Database: Set db = CurrentDb
Dim rstCampos As DAO.Recordset

Dim tmpValidarCampo As String: tmpValidarCampo = right(pRepositorio, Len(pRepositorio) - 3)

'' VALIDAR CONCILIAÇÃO
Dim tmpScript As String

    '' MONTAR STRING DE NOME DE COLUNAS
    Set rstCampos = db.OpenRecordset(Replace(Scripts.SelectCamposNomes, "pRepositorio", pRepositorio))
    Do While Not rstCampos.EOF
        If InStr(rstCampos.Fields("campo").value, tmpValidarCampo) Then
            tmpScript = tmpScript & rstCampos.Fields("campo").value & ","
        End If
        rstCampos.MoveNext
        DoEvents
    Loop

    Set Scripts = Nothing
    rstCampos.Close
    db.Close

    carregarCamposNomes = left(tmpScript, Len(tmpScript) - 1)

End Function

'' #20211122_AjusteDeCampos_CTe
Function carregarComprasCTe() As String
Dim Scripts As New clsConexaoNfeCte

'' BANCO_LOCAL
Dim db As DAO.Database: Set db = CurrentDb
Dim rstCampos As DAO.Recordset

'' VALIDAR CONCILIAÇÃO
Dim tmpScript As String

    '' MONTAR STRING DE NOME DE COLUNAS
    Set rstCampos = db.OpenRecordset("qryCompras_CTe_Select_AjustesCampos")
    Do While Not rstCampos.EOF
        tmpScript = tmpScript & "'" & rstCampos.Fields("ChvAcesso_CompraNF").value & "',"
        
        rstCampos.MoveNext
        DoEvents
    Loop

    Set Scripts = Nothing
    rstCampos.Close
    db.Close

    carregarComprasCTe = left(tmpScript, Len(tmpScript) - 1)

End Function

'' #20211128_MoverArquivosProcessados
Sub MoverArquivosProcessados()
Dim db As DAO.Database: Set db = CurrentDb
Dim rst As DAO.Recordset

''#registroProcessado
'' 0|1 - caminhoDeColeta
'' 2 - caminhoDeColeta
'' 3 - caminhoDeColetaProcessados
'' 4 - caminhoDeColetaExpurgo
'' 5 - caminhoDeColetaAcoes
'' 8 - MOVER ARQUIVOS            -- CONTROLE INTERNO
'' 9 - FINAL                     -- CONTROLE INTERNO

'Dim sql_Update_tblDadosConexaoNFeCTe_registroProcessado_Reclassificar As String: sql_Update_tblDadosConexaoNFeCTe_registroProcessado_Reclassificar = _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 1;"
'    Application.CurrentDb.Execute Replace(sql_Update_tblDadosConexaoNFeCTe_registroProcessado_Reclassificar, 1, 5)


'' Classificação de Processados
'' 1. Classificar como "registroProcessado(3) - Processados" onde ...
'' 1.1 Onde temos arquivo para processar - "CaminhoDoArquivo" e
'' 1.2 Registro é valido "registroValido(1) - Registro Válido" e
'' 1.3 Registro processado é "registroProcessado (2) - CadastroDeComprasEmServidor()".
Dim sql_Update_registroProcessado_Processados As String: sql_Update_registroProcessado_Processados = _
    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 3 WHERE ((Not (tblDadosConexaoNFeCTe.CaminhoDoArquivo) Is Null) AND ((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=2));"
    Application.CurrentDb.Execute sql_Update_registroProcessado_Processados
    
'' Expurgo
'' 1. Classificar como "registroProcessado(4) - Expurgo" onde ...
'' 1.1 Registro é valido "registroValido(1) - Registro Válido" e não foi processado "registroProcessado(0) - Não foi processado"
'' 2 Registro é invalido "registroValido(0) - Registro Inválido"
Dim sql_Update_registroProcessado_Expurgo() As Variant: sql_Update_registroProcessado_Expurgo = Array( _
    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 4 WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0));", _
    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 4 WHERE (((tblDadosConexaoNFeCTe.registroValido)=0));")
    executarComandos sql_Update_registroProcessado_Expurgo
    
'' Classificar como arquivos finalizados - "registroProcessado(9) - Finalizados"
Dim sql_Update_CopyFile_Final As String: sql_Update_CopyFile_Final = _
    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 9, tblDadosConexaoNFeCTe.CaminhoDoArquivo = [tblDadosConexaoNFeCTe].[CaminhoDestino], tblDadosConexaoNFeCTe.CaminhoDestino = Null WHERE (((tblDadosConexaoNFeCTe.[registroProcessado])=8))"
        
'' Atualização do caminho de destino
Dim sql_Update_CaminhoDestino As String: sql_Update_CaminhoDestino = _
    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.CaminhoDestino = strCaminhoDestino([tblDadosConexaoNFeCTe].[CaminhoDoArquivo]), tblDadosConexaoNFeCTe.registroProcessado = 8 WHERE (((tblDadosConexaoNFeCTe.registroProcessado)<8));"
    Application.CurrentDb.Execute sql_Update_CaminhoDestino
    
'' Seleção de arquivos para movimentação de pastas
Dim sql_Select_CaminhoDestino As String: sql_Select_CaminhoDestino = _
    "SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo, tblDadosConexaoNFeCTe.CaminhoDestino FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroProcessado)=8));"
        
    '' MOVER ARQUIVOS
    Set rst = db.OpenRecordset(sql_Select_CaminhoDestino)
    Do While Not rst.EOF
        If (Dir(rst.Fields("CaminhoDoArquivo").value) <> "") Then
            FileCopy rst.Fields("CaminhoDoArquivo").value, rst.Fields("CaminhoDestino").value
            Kill rst.Fields("CaminhoDoArquivo").value
        End If
        rst.MoveNext
    Loop

    '' ATUALIZAÇÃO - registroProcessado ( FINAL )
    Application.CurrentDb.Execute sql_Update_CopyFile_Final

db.Close

Set db = Nothing
Set rst = Nothing

End Sub
=======
Attribute VB_Name = "01_modFernanda"
Option Compare Database


'' ### TO-DO ###
''
'' #20211123_  ''tmpCompraNF, mod 55, Tipo 4 - NF-e Retorno Armazém
'' #20211123_  ''tmpCompraNFItem, mod55, Tipo 4 - NF-e Retorno Armazém
'' #20211123_  ''tmpCompraNF, mod 55, Tipo 6 - NF-e Transferência com código Sisparts
'' #20211123_  ''tmpCompraNFItem, mod 55, Tipo 6 - NF-e Transferência com código Sisparts

'' ### DONE ###
''
'' #20211122_AjusteDeCampos_CTe_tblCompraNFItem
'' #20211122_AjusteDeCampos_CTe
'' Consultas
'' #20210823_EXPORTACAO_LIMITE
'' #20210823_qryDadosGerais_Update_ID_NatOp_CompraNF__FiltroCFOP -- FiltroCFOP
'' #20210823_qryDadosGerais_Update_IDVD
'' #20210823_qryUpdateID_NatOp_CompraNF
'' #20210823_qryUpdateCFOP_FilCompra
'' #20210823_qryUpdate_ModeloDoc_CFOP
'' #20210823_qryUpdateFilCompraNF
'' #20210823_qryDadosGerais_Update_IdFornCompraNF
'' #20210823_qryDadosGerais_Update_Sit_CompraNF
'' #20210823_XML_CONTROLE | Quando importar cada XML, precisa recortar o arquivo da pasta da empresa e colar dentro de uma pasta chama “Processados”, porém dentro de cada pasta de cada empresa, pois não podemos misturar os XML´s de cada empresa.
'' #20210823_XML_FORMULARIO | Não encontrei um formulário com os XML´s que não foram processados e o motivo. | <<< ATENÇÃO - NÃO DEFINIMOS COMO CLASSIFICAREMOS OS MOTIMOS DE NÃO PROCESSAMENTO DE ARQUIVOS >>>
'' #20210823_VTotProd_CompraNF
'' #20210823_ID_Prod_CompraNFItem
'' #20210823_BaseCalcICMS_CompraNF
'' #20210823_VTotICMS_CompraNF
'' #20210823_CadastroDeComprasEmServidor
'' #20210823_qryUpdateNumPed_CompraNF
'' #20210823_FornecedoresValidos


''----------------------------
'' ### EXEMPLOS DE FUNÇÕES
''
'' 01. processarDadosGerais
'' 02. processarArquivosPendentes
'' 04. CadastroDeComprasEmServidor
'' 05. tratamentoDeArquivosValidos
'' 06. tratamentoDeArquivosInvalidos
'' 07. criacaoArquivosJson
''
'' 99. FUNÇÃO_AUXILIAR: carregarDadosGerais(strArquivo As String)
'' 99. FUNÇÃO_AUXILIAR: carregarArquivosPendentes(strArquivo As String)
'' 99. FUNÇÃO_AUXILIAR: azsProcessamentoDeArquivos(sqlArquivos As String, qryUpdate As String, strOrigem As String, strDestino As String)
'' 99. FUNÇÃO_AUXILIAR: tratamentoDeArquivosValidos()
'' 99. FUNÇÃO_AUXILIAR: tratamentoDeArquivosInvalidos()
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

'' #20211116_1839
'arquivos.add "C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210348740351012767570000021186731952977908-cteproc.xml"
'arquivos.add "C:\Sisparts\SispartsConexao\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210348740351012767570000021186731952977908-cteproc.xml"

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


'' #20211116_1839
'arquivos.add "C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210348740351012767570000021186731952977908-cteproc.xml"
''    arquivos.add "C:\Sisparts\SispartsConexao\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210348740351012767570000021186731952977908-cteproc.xml"
''    arquivos.add "C:\Sisparts\SispartsConexao\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\32210368365501000296550000000638791001361285-nfeproc.xml"

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


''' 04. ENVIAR DADOS PARA SERVIDOR
''' #20210823_CadastroDeComprasEmServidor
'Sub enviarDadosServidor()
'
'''==================================================
'''### PROCESSAMENTO
'''==================================================
'
'''' #ANALISE_DE_PROCESSAMENTO
''Dim DT_PROCESSO As Date: DT_PROCESSO = Now()
''
''    '' CADASTRO DE CABEÇALHO DE COMPRAS
''    enviar_ComprasParaServidor "tblCompraNF"
''
''    '' RELACIONAMENTO DE ID_COMPRAS COM CHAVES DE ACESSO CADASTRADAS DO SERVIDOR
''    criarTabelaTemporariaParaRelacionarIdCompraComChvAcesso
''    relacionarIdCompraComChvAcesso
''
''    '' CADASTRO DE ITENS DE COMPRAS
''    enviar_ComprasParaServidor "tblCompraNFItem"
''
''    '' #ANALISE_DE_PROCESSAMENTO
''    statusFinal DT_PROCESSO, "enviarDadosServidor"
''
''    Debug.Print "Concluido! - enviarDadosServidor"
''    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "Concluido! - enviarDadosServidor"
'
'End Sub

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

    '' SELEÇÃO DE REGISTRO
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

'    '' SELEÇÃO PELO USUARIO
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

>>>>>>> f4084cb29d769387d25e7b837853d2119e0da429
