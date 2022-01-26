Attribute VB_Name = "01_modConexaoNfeCte"
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
''### LEITURA DE REPOSITORIOS
''#######################################################################################

    '' REPOSITORIOS
    If pArquivos.count > 0 Then
        For Each item In pArquivos
            If (IsNull(DLookup("ID", "tblDadosConexaoNFeCTe", "Chave='" & getFileName(CStr(item)) & "'"))) Then arquivos.add CStr(item)
        Next
        
    Else
        For Each caminhoAntigo In Array(DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColeta'"))
            For Each caminhoNovo In carregarParametros(DadosGerais.SelectColetaEmpresa)
                For Each item In GetFilesInSubFolders(CStr(Replace(Replace(caminhoAntigo, "empresa", caminhoNovo), "recebimento\", "")))
                    If (IsNull(DLookup("Chave", "logArquivosProcessados", "Chave='" & getFileName(CStr(item)) & "'"))) Then arquivos.add CStr(item)
                Next
            Next
        Next
        
    End If

''#######################################################################################
''### PROCESSAMENTO DE ARQUIVOS COLETADOS
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
''### LEITURA DE REPOSITORIOS
''#######################################################################################

'' REPOSITORIO
For Each item In carregarParametros(DadosGerais.SelectArquivosPendentes)
    arquivos.add CStr(item)
Next

''#######################################################################################
''### PROCESSAMENTO DE ARQUIVOS COLETADOS
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
Dim scripts As New clsConexaoNfeCte
Dim db As DAO.Database: Set db = CurrentDb
Dim rstChvAcesso As DAO.Recordset: Set rstChvAcesso = db.OpenRecordset(Replace(scripts.SelectRegistroValidoPorcessado, "pRepositorio", "tblCompraNF"))

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

Dim sql_comprasItens_count As String:
    sql_comprasItens_count = "SELECT COUNT(*) as contador FROM tblCompraNFItem where ID_CompraNF_CompraNFItem = (SELECT ID_CompraNF FROM tblCompraNF where ChvAcesso_CompraNF = 'strID_Compra')"


Dim sql_comprasItens_update_IdProd As String:
    sql_comprasItens_update_IdProd = "UPDATE tblCompraNFItem SET ID_Prod_CompraNFItem = tbProdutos.[código] " & _
                                        "FROM tblCompraNFItem AS tbItens " & _
                                        "INNER JOIN tblCompraNF as tbCompras ON tbCompras.ID_CompraNF = tbItens.ID_CompraNF_CompraNFItem " & _
                                        "INNER join [Cadastro de Produtos] as tbProdutos on tbProdutos.modelo = 'Transporte'  " & _
                                        "WHERE tbCompras.Sit_CompraNF = 6 and tbItens.ID_Prod_CompraNFItem=0;"
    
    
    
Dim sql_comprasItens_Update_FlagEst As String:
    sql_comprasItens_Update_FlagEst = "UPDATE tblCompraNFItem SET tblCompraNFItem.FlagEst_CompraNFItem = 1 WHERE (((tblCompraNFItem.FlagEst_CompraNFItem)=0));"

    
Dim strCaminhoAcoes As String: strCaminhoAcoes = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColetaAcoes'")
    
       
'' #CONTADOR
Dim contadorDeRegistros As Long: contadorDeRegistros = 1

'' VALIDAR CONCILIAÇÃO
Dim TMP As String

    If (rstChvAcesso.RecordCount > 0) Then
        '' BANCO_DESTINO
        dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
        
        '' #BARRA_PROGRESSO
        rstChvAcesso.MoveLast
        rstChvAcesso.MoveFirst
        SysCmd acSysCmdInitMeter, "Exportação...", rstChvAcesso.RecordCount
    
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
                dbDestino.SqlExecute TMP
            Next item
                   
            '' CONTROLE DE RECADASTRO
            dbDestino.SqlSelect Replace(sql_comprasItens_count, "strID_Compra", pChvAcesso)
            If (dbDestino.rs.Fields("contador").value = 0) Then
            
                '' RELACIONAR ITENS DE COMPRAS COM COMPRAS JÁ CADASTRADAS NO SERVIDOR
                dbDestino.SqlSelect "SELECT ChvAcesso_CompraNF,ID_CompraNF FROM tblCompraNF where ChvAcesso_CompraNF = '" & pChvAcesso & "';"
                qryComprasItens_Update_IDCompraNF = Replace(Replace(scripts.UpdateComprasItens_IDCompraNF, "strChave", pChvAcesso), "strID_Compra", dbDestino.rs.Fields("ID_CompraNF").value)
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
                        dbDestino.SqlExecute TMP
                    Next item
        
                End If
                        
                '' MUDAR STATUS DO REGISTRO
                Application.CurrentDb.Execute Replace(scripts.compras_atualizarEnviadoParaServidor, "strChave", rstChvAcesso.Fields("chvAcesso_CompraNF").value)
        
                '' #20211122_AjusteDeCampos_CTe
                If (Len(carregarComprasCTe) > 0) Then dbDestino.SqlExecute Replace(qryComprasCTe_Update_AjustesCampos, "pLista_ChvAcesso_CompraNF", carregarComprasCTe)
                dbDestino.SqlExecute "UPDATE tblCompraNF SET HoraEntd_CompraNF = NULL, IDVD_CompraNF = NULL WHERE tblCompraNF.IDVD_CompraNF=0;"
    
            End If
            
            contador = contador + 1
            contadorDeRegistros = contadorDeRegistros + 1
            
            '' #BARRA_PROGRESSO
            SysCmd acSysCmdUpdateMeter, contadorDeRegistros
            
            Debug.Print contadorDeRegistros
            
            rstChvAcesso.MoveNext
            DoEvents
        Loop
        
'        '' #20211202_update_Almox_CompraNFItem
'        dbDestino.SqlExecute sql_comprasItens_update_Almox_CompraNFItem
        
        '' #20220106_update_IdProd_CompraNFItem
        dbDestino.SqlExecute sql_comprasItens_update_IdProd
        
        '' #20220111_update_FlagEst_CompraNFItem
        dbDestino.SqlExecute sql_comprasItens_Update_FlagEst
            
        '' #20211128_LimparRepositorios
        '' Limpar repositorio de itens de compras
        Application.CurrentDb.Execute _
                "Delete from tblCompraNFItem"

        '' Limpar repositorio de compras
        Application.CurrentDb.Execute _
                "Delete from tblCompraNF"
        
        '' #20211128_MoverArquivosProcessados
        MoverArquivosProcessados
                
        '' #20220119_GerarArquivosDeLancamentos_e_Manifestos
        '' LANÇAMENTO
        gerarArquivosJson opFlagLancadaERP, , strCaminhoAcoes
        
        '' MANIFESTO
        gerarArquivosJson opManifesto, , strCaminhoAcoes
        
        '' Limpar repositorio de dados gerais
        Application.CurrentDb.Execute _
                scripts.InsertLogProcessados
        Application.CurrentDb.Execute _
                "Delete from tblDadosConexaoNFeCTe"
        
    End If
    

rstChvAcesso.Close
dbDestino.CloseConnection
db.Close

Set scripts = Nothing
Set rstChvAcesso = Nothing
Set db = Nothing

'' #ANALISE_DE_PROCESSAMENTO
statusFinal DT_PROCESSO, "CadastroDeComprasEmServidor - Exportar compras ( Quantidade de registros: " & contador & " )"

'' #BARRA_PROGRESSO
SysCmd acSysCmdRemoveMeter

Debug.Print "CadastroDeComprasEmServidor() - Concluido!"

End Sub


''=======================================================================================================
'' LIB
''=======================================================================================================

'' GERAR ARQUIVOS JSONs
Sub gerarArquivosJson(pArquivo As enumTipoArquivo, Optional strConsulta As String, Optional strCaminho As String)
Dim s As New clsCriarArquivos
Dim strCaminhoDeSaida As String

Dim sql_Select_tblDadosConexaoNFeCTe_registroValido As String: sql_Select_tblDadosConexaoNFeCTe_registroValido = _
    "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi_copia FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi_copia]))>0) AND ((tblDadosConexaoNFeCTe.registroValido)=1))"

    '' SELEÇÃO DE REGISTRO
    If strConsulta <> "" Then
        sql_Select_tblDadosConexaoNFeCTe_registroValido = "SELECT * FROM (" & sql_Select_tblDadosConexaoNFeCTe_registroValido & ") AS tmpSelecao WHERE tmpSelecao.ChvAcesso =  '" & strConsulta & "';"
    Else
        sql_Select_tblDadosConexaoNFeCTe_registroValido = _
                    "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi_copia FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi_copia]))>0) AND ((tblDadosConexaoNFeCTe.registroValido)=1));"
    End If
    
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

    Debug.Print "gerarArquivosJson() - Concluido!"
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "Concluido! - criacaoArquivosJson"

Cleanup:
    Set s = Nothing

End Sub

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
Dim scripts As New clsConexaoNfeCte

'Dim pRepositorio As String: pRepositorio = "tblCompraNFItem"
'Dim pChvAcesso As String: pChvAcesso = "32210368365501000296550000000638791001361285"

'' BANCO_LOCAL
Dim db As DAO.Database: Set db = CurrentDb

Dim rstCampos As DAO.Recordset: Set rstCampos = db.OpenRecordset(Replace(scripts.SelectCamposNomes, "pRepositorio", pRepositorio))
Dim rstOrigem As DAO.Recordset

'' VALIDAR CONCILIAÇÃO
Dim tmpScript As String
Dim tmpValidarCampo As String: tmpValidarCampo = right(pRepositorio, Len(pRepositorio) - 3)

Dim sqlOrigem As String: sqlOrigem = _
    "Select * from (" & Replace(scripts.SelectRegistroValidoPorcessado, "pRepositorio", pRepositorio) & ") as tmpRepositorio where tmpRepositorio.ChvAcesso_CompraNF = '" & pChvAcesso & "'"
    
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

    Set scripts = Nothing
    rstCampos.Close
    rstOrigem.Close
    db.Close

End Function

Function carregarCamposNomes(pRepositorio As String) As String
Dim scripts As New clsConexaoNfeCte

'' BANCO_LOCAL
Dim db As DAO.Database: Set db = CurrentDb
Dim rstCampos As DAO.Recordset

Dim tmpValidarCampo As String: tmpValidarCampo = right(pRepositorio, Len(pRepositorio) - 3)

'' VALIDAR CONCILIAÇÃO
Dim tmpScript As String

    '' MONTAR STRING DE NOME DE COLUNAS
    Set rstCampos = db.OpenRecordset(Replace(scripts.SelectCamposNomes, "pRepositorio", pRepositorio))
    Do While Not rstCampos.EOF
        If InStr(rstCampos.Fields("campo").value, tmpValidarCampo) Then
            tmpScript = tmpScript & rstCampos.Fields("campo").value & ","
        End If
        rstCampos.MoveNext
        DoEvents
    Loop

    Set scripts = Nothing
    rstCampos.Close
    db.Close

    carregarCamposNomes = left(tmpScript, Len(tmpScript) - 1)

End Function

'' #20211122_AjusteDeCampos_CTe
Function carregarComprasCTe() As String
Dim scripts As New clsConexaoNfeCte

'' BANCO_LOCAL
Dim db As DAO.Database: Set db = CurrentDb
Dim rstCampos As DAO.Recordset

'' VALIDAR CONCILIAÇÃO
Dim tmpScript As String

Dim sql_Compras_CTe_Select_AjustesCampos As String: sql_Compras_CTe_Select_AjustesCampos = _
    "SELECT tblCompraNF.ChvAcesso_CompraNF FROM tblCompraNF INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso WHERE (((tblDadosConexaoNFeCTe.codMod)=57) AND ((tblDadosConexaoNFeCTe.registroProcessado)=2));"

    '' MONTAR STRING DE NOME DE COLUNAS
    Set rstCampos = db.OpenRecordset(sql_Compras_CTe_Select_AjustesCampos)
    Do While Not rstCampos.EOF
        tmpScript = tmpScript & "'" & rstCampos.Fields("ChvAcesso_CompraNF").value & "',"
        rstCampos.MoveNext
        DoEvents
    Loop

    Set scripts = Nothing
    rstCampos.Close
    db.Close

    If (Len(tmpScript) > 0) Then carregarComprasCTe = left(tmpScript, Len(tmpScript) - 1)

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
    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 4 WHERE (((tblDadosConexaoNFeCTe.registroValido)=0) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0));")
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
        
    Debug.Print "############################"
        
    '' MOVER ARQUIVOS
    Set rst = db.OpenRecordset(sql_Select_CaminhoDestino)
    Do While Not rst.EOF
        If (Dir(rst.Fields("CaminhoDoArquivo").value) <> "") Then
            Debug.Print "# ORIGEM"
            Debug.Print rst.Fields("CaminhoDoArquivo").value
            
            Debug.Print "# DESTINO"
            Debug.Print rst.Fields("CaminhoDestino").value
            
            
            Kill rst.Fields("CaminhoDestino").value
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


