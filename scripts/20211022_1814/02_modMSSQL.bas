Attribute VB_Name = "02_modMSSQL"
Public Const DATE_TIME_FORMAT               As String = "yyyy/mm/dd hh:mm:ss"
Public Const DATE_FORMAT                    As String = "yyyy/mm/dd"

Function carregarScript_Insert(pRepositorio As String, pChvAcesso As String) As String
'Dim pRepositorio As String: pRepositorio = "tblCompraNF"
'Dim pChvAcesso As String: pChvAcesso = "32210304884082000569570000040073831040073834"

Dim strCamposNomes As String: _
    strCamposNomes = carregarCamposNomes(pRepositorio)

Dim strCamposNomesTmp As String: _
    strCamposNomesTmp = Replace(strCamposNomes, "_CompraNF", "")

Dim strCamposValores As String: _
    strCamposValores = carregarCamposValores(pRepositorio, pChvAcesso)

Dim tmpScript As String: _
    tmpScript = "INSERT INTO tblCompraNF ( " & strCamposNomes & " ) SELECT " & strCamposNomesTmp & " FROM ( VALUES ( " & strCamposValores & " ) ) AS TMP ( " & strCamposNomesTmp & " ) LEFT JOIN tblCompraNF ON tblCompraNF.ChvAcesso_CompraNF = tmp.ChvAcesso WHERE tblCompraNF.ChvAcesso_CompraNF IS NULL;"
    
    carregarScript_Insert = tmpScript

End Function

Function carregarCamposValores(pRepositorio As String, pChvAcesso As String) As String
Dim Scripts As New clsConexaoNfeCte
Dim db As DAO.Database: Set db = CurrentDb
Dim rstCampos As DAO.Recordset: Set rstCampos = db.OpenRecordset(Replace(Scripts.SelectCamposNomes, "pRepositorio", pRepositorio))
Dim rstOrigem As DAO.Recordset

Dim tmpScript As String
Dim tmp As String: tmp = right(pRepositorio, Len(pRepositorio) - 3)

    Set rstOrigem = db.OpenRecordset("Select * from (" & Replace(Scripts.SelectRegistroValidoPorcessado, "pRepositorio", pRepositorio) & ") as tmpRepositorio where tmpRepositorio.ChvAcesso_CompraNF = '" & pChvAcesso & "'")
    
    Do While Not rstOrigem.EOF
        tmpScript = ""
    
        '' LISTAGEM DE CAMPOS
        rstCampos.MoveFirst
        Do While Not rstCampos.EOF
    
            '' CRIAR SCRIPT DE INCLUSAO DE DADOS NA TABELA DESTINO
            '' 2. campos x formatacao
            If InStr(rstCampos.Fields("campo").value, tmp) Then
    
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
    
            rstCampos.MoveNext
            DoEvents
        Loop
        
        rstOrigem.MoveNext
        DoEvents
    Loop

    Set Scripts = Nothing
    rstCampos.Close
    rstOrigem.Close
    db.Close

    carregarCamposValores = left(tmpScript, Len(tmpScript) - 1)

End Function

Function carregarCamposNomes(pRepositorio As String) As String
Dim Scripts As New clsConexaoNfeCte
Dim db As DAO.Database: Set db = CurrentDb
Dim rstCampos As DAO.Recordset
Dim tmpScript As String
Dim tmp As String

    '' 1. cabecalho
    Set rstCampos = db.OpenRecordset(Replace(Scripts.SelectCamposNomes, "pRepositorio", pRepositorio))
    Do While Not rstCampos.EOF
        tmpScript = tmpScript & rstCampos.Fields("campo").value & ","
        rstCampos.MoveNext
        DoEvents
    Loop

    Set Scripts = Nothing
    rstCampos.Close
    db.Close

    carregarCamposNomes = left(tmpScript, Len(tmpScript) - 1)

End Function


Sub CadastroDeComprasEmServidor()

'' BANCO_DESTINO
Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
Dim dbDestino As New Banco

'' BANCO_ORIGEM
Dim Scripts As New clsConexaoNfeCte
Dim db As DAO.Database: Set db = CurrentDb
Dim rstChvAcesso As DAO.Recordset: Set rstChvAcesso = db.OpenRecordset(Scripts.SelectRegistroValidoPorcessado)

Dim contador As Long


    '' BANCO_DESTINO
    dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer

    '' CADASTRO
    Do While Not rstChvAcesso.EOF
        
'        '' CONTADOR
'        dbDestino.SqlSelect "SELECT max(NumPed_CompraNF)+1 as contador from tblCompraNF"
'        contador = IIf(IsNull(dbDestino.rs.Fields("contador").value), 1, dbDestino.rs.Fields("contador").value)
        
        dbDestino.SqlExecute carregarScript_Insert("tblCompraNF", rstChvAcesso.Fields("ChvAcesso_CompraNF").value)
        
        rstChvAcesso.MoveNext
'        contador = contador + 1
    Loop

rstChvAcesso.Close
db.Close

dbDestino.CloseConnection

Debug.Print "Concluido!"

End Sub




'' #20210823_qryUpdateNumPed_CompraNF
Sub CadastroDeComprasComControleDeNumeroDePedidos()

Dim qryDadosGerais_Select_ChvAcesso_RegistrosValidosProcessados As String: qryDadosGerais_Select_ChvAcesso_RegistrosValidosProcessados = "Select Distinct ChvAcesso from tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido) = 1) AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1) AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0)) AND tblDadosConexaoNFeCTe.ChvAcesso IS NOT NULL;"

'' BANCO_ORIGEM
Dim contador As Long
Dim sqlItens As String
Dim sqlScript As String
Dim sqlScriptTemplate As String: sqlScriptTemplate = _
        "INSERT INTO tblCompraNF (ChvAcesso_CompraNF,NumPed_CompraNF) " & _
        "SELECT ChvAcesso, NumPed FROM (VALUES ('sqlItens',contador)) AS tmp(ChvAcesso, NumPed) " & _
        "LEFT JOIN tblCompraNF ON tblCompraNF.ChvAcesso_CompraNF = tmp.ChvAcesso WHERE tblCompraNF.ChvAcesso_CompraNF IS NULL;"

Dim db As DAO.Database: Set db = CurrentDb
Dim rstChvAcesso As DAO.Recordset: Set rstChvAcesso = db.OpenRecordset(qryDadosGerais_Select_ChvAcesso_RegistrosValidosProcessados)

'' BANCO_DESTINO
Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
Dim dbDestino As New Banco

dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer

    Do While Not rstChvAcesso.EOF
        
        '' CONTADOR
        dbDestino.SqlSelect "SELECT max(NumPed_CompraNF)+1 as contador from tblCompraNF"
        contador = IIf(IsNull(dbDestino.rs.Fields("contador").value), 1, dbDestino.rs.Fields("contador").value)
                
        '' CADASTRO
        sqlScript = Replace(Replace(sqlScriptTemplate, "sqlItens", rstChvAcesso.Fields("ChvAcesso").value), "contador", Format(contador, "000000"))
        Debug.Print rstChvAcesso.Fields("ChvAcesso").value
        Debug.Print sqlScript
        
        dbDestino.SqlExecute sqlScript
        
        rstChvAcesso.MoveNext
        'dbDestino.rs.Close
        contador = contador + 1
        sqlItens = ""
    Loop

rstChvAcesso.Close
db.Close

dbDestino.CloseConnection

Debug.Print "Concluido!"

End Sub

''' 03. Enviar Dados Para Servidor
'Public Function enviar_ComprasParaServidor(pDestino As String)
'On Error Resume Next
''On Error GoTo adm_Err
'
''' ---------------------
''' VARIAVEIS GERAIS
''' ---------------------
'
''' LISTAGEM DE CAMPOS DA TABELA ORIGEM/DESTINO
''Dim strCampo As String
'
'Dim qryCampos As String: qryCampos = _
'                            "SELECT distinct   " & _
'                            "   tblParametros.TipoDeParametro  " & _
'                            "   , tblParametros.ID  " & _
'                            "   , tblOrigemDestino.campo  " & _
'                            "   , tblOrigemDestino.formatacao  " & _
'                            "   , tblOrigemDestino.valorPadrao  " & _
'                            "FROM   " & _
'                            "   tblParametros INNER JOIN tblOrigemDestino ON tblParametros.ValorDoParametro = tblOrigemDestino.campo  " & _
'                            "WHERE (((tblParametros.TipoDeParametro)='pDestino') AND ((tblOrigemDestino.TagOrigem)<>0)) ORDER BY tblParametros.ID;"
'
'Dim qryUpdateRegistroProcessado As String: qryUpdateRegistroProcessado = _
'        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 2 WHERE (((tblDadosConexaoNFeCTe.ChvAcesso)=""strChave"") AND ((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1));"
'
'
''' SCRIPT
'Dim tmpScript As String
'Dim tmpSelecaoDeRegistros As String
'Dim tmp As String
'
''' ---------------------
''' BANCO LOCAL
''' ---------------------
'Dim db As DAO.Database: Set db = CurrentDb
'Dim rstOrigem As DAO.Recordset
'Dim rstCampos As DAO.Recordset
'
''' ---------------------
''' BANCO DESTINO
''' ---------------------
'Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
'Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
'Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
'Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
'Dim sqlCampos As String: sqlCampos = Replace(qryCampos, "pDestino", pDestino)
'
'Debug.Print sqlCampos
'If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), sqlCampos
'
'Dim dbDestino As New Banco
'
'
''' #ANALISE_DE_PROCESSAMENTO
'Dim DT_PROCESSO As Date: DT_PROCESSO = Now()
'
''' #CONTADOR
'Dim contadorDeRegistros As Long: contadorDeRegistros = 1
'Dim totalDeRegistros As Long
'
'
'    dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
'    dbDestino.SqlSelect "SELECT * FROM " & pDestino
'
'    tmpScript = "Insert into " & pDestino & " ("
'
'    '' 1. cabecalho
'    Set rstCampos = db.OpenRecordset(sqlCampos)
'    Do While Not rstCampos.EOF
'        If InStr(rstCampos.Fields("campo").value, right(pDestino, Len(pDestino) - 3), vbTextCompare) Then
'            tmpScript = tmpScript & rstCampos.Fields("campo").value & ","
'        End If
'
'        rstCampos.MoveNext
'        DoEvents
'    Loop
'    tmpScript = left(tmpScript, Len(tmpScript) - 1) & ") values ("
''    Debug.Print tmpScript
'
'    '' #azs - testes
'    '' #20210823_EXPORTACAO_LIMITE - LIMITE
'    '' BANCO LOCAL
'    tmpSelecaoDeRegistros = "Select * from " & pDestino & " Where ChvAcesso_CompraNF IN (SELECT ChvAcesso FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND (tblDadosConexaoNFeCTe.registroProcessado)=1))"
'    Debug.Print tmpSelecaoDeRegistros
'    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), tmpSelecaoDeRegistros
'
'    Set rstOrigem = db.OpenRecordset(tmpSelecaoDeRegistros)
'
'    '' #BARRA_PROGRESSO
'    totalDeRegistros = rstOrigem.RecordCount
'    SysCmd acSysCmdInitMeter, pDestino, totalDeRegistros
'
'    Do While Not rstOrigem.EOF
'        tmp = ""
'
'        '' LISTAGEM DE CAMPOS
'        rstCampos.MoveFirst
'        Do While Not rstCampos.EOF
'
'            '' CRIAR SCRIPT DE INCLUSAO DE DADOS NA TABELA DESTINO
'            '' 2. campos x formatacao
'            If InStr(rstCampos.Fields("campo").value, right(pDestino, Len(pDestino) - 3)) Then
'
'                If rstCampos.Fields("formatacao").value = "opTexto" Then
'                    tmp = tmp & "'" & rstOrigem.Fields(rstCampos.Fields("campo").value).value & "',"
'
'                ElseIf rstCampos.Fields("formatacao").value = "opNumero" Or rstCampos.Fields("formatacao").value = "opMoeda" Then
'                    tmp = tmp & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", rstOrigem.Fields(rstCampos.Fields("campo").value).value, rstCampos.Fields("valorPadrao").value) & ","
'
'                ElseIf rstCampos.Fields("formatacao").value = "opTime" Then
'                    tmp = tmp & "'" & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", Format(rstOrigem.Fields(rstCampos.Fields("campo").value).value, DATE_TIME_FORMAT), rstCampos.Fields("valorPadrao").value) & "',"
'
'                ElseIf stCampos.Fields("formatacao").value = "opData" Then
'                    tmp = tmp & "'" & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", Format(rstOrigem.Fields(rstCampos.Fields("campo").value).value, DATE_FORMAT), rstCampos.Fields("valorPadrao").value) & "',"
'
'                End If
'
'            End If
'
'            rstCampos.MoveNext
'            DoEvents
'        Loop
'
'        '' BANCO DESTINO
'        tmp = left(tmp, Len(tmp) - 1) & ")"
'
'        Debug.Print tmpScript & tmp
'        If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), tmpScript & tmp
'
'
'        Debug.Print tmpScript & tmp
'
'        dbDestino.SqlExecute tmpScript & tmp
'
'        '' Terminio de operacao
'        If (pDestino = "tblCompraNFItem") Then
'            tmp = Replace(qryUpdateRegistroProcessado, "strChave", rstOrigem.Fields("ChvAcesso_CompraNF").value)
'            Debug.Print tmp
'            If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), tmp
'
'            Application.CurrentDb.Execute tmp
'        End If
'
'        '' #BARRA_PROGRESSO
'        contadorDeRegistros = contadorDeRegistros + 1
'        SysCmd acSysCmdUpdateMeter, contadorDeRegistros
'
'        rstOrigem.MoveNext
'        DoEvents
'    Loop
'
'    dbDestino.CloseConnection
'    db.Close: Set db = Nothing
'
'    '' RELACIONAR ITENS DE COMPRAS COM COMPRAS JÁ CADASTRADAS
'    If (pDestino = "tblCompraNF") Then relacionarIdCompraComChvAcesso
'
'    '' #ANALISE_DE_PROCESSAMENTO
'    statusFinal DT_PROCESSO, "enviar_ComprasParaServidor - Exportar registros pendentes ( Quantidade de registros: " & contadorDeRegistros & " )"
'
'    '' #BARRA_PROGRESSO
'    SysCmd acSysCmdRemoveMeter
'
'adm_Exit:
'    Exit Function
'
''adm_Err:
''    MsgBox Error$
''    Resume adm_Exit
'
'
'End Function
