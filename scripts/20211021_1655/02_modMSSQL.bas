Attribute VB_Name = "02_modMSSQL"
Public Const DATE_TIME_FORMAT               As String = "yyyy/mm/dd hh:mm:ss"
Public Const DATE_FORMAT                    As String = "yyyy/mm/dd"

Sub teste()
Dim sqlItens As String
Dim sqlScript As String
Dim sqlScriptTemplate As String: sqlScriptTemplate = _
        "INSERT INTO tblCompraNF (ChvAcesso_CompraNF,NumPed_CompraNF) " & _
        "SELECT ChvAcesso, NumPed FROM (VALUES ('sqlItens',contador)) AS tmp(ChvAcesso, NumPed) " & _
        "LEFT JOIN tblCompraNF ON tblCompraNF.ChvAcesso_CompraNF = tmp.ChvAcesso WHERE tblCompraNF.ChvAcesso_CompraNF IS NULL;"

Dim qryDadosGerais_Select_ChvAcesso_RegistrosValidosProcessados As String: qryDadosGerais_Select_ChvAcesso_RegistrosValidosProcessados = "Select Distinct ChvAcesso from tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido) = 1) AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1) AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0)) AND tblDadosConexaoNFeCTe.ChvAcesso IS NOT NULL;"

Dim db As DAO.Database: Set db = CurrentDb
Dim rstChvAcesso As DAO.Recordset: Set rstChvAcesso = db.OpenRecordset(qryDadosGerais_Select_ChvAcesso_RegistrosValidosProcessados)

Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")

Dim dbDestino As New Banco

Dim contador As Long

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

'dbDestino.CloseConnection

Debug.Print "Concluido!"

End Sub


'' 01. Criar Tabela Temporaria Para Relacionar IdCompra Com ChvAcesso
Public Function criarTabelaTemporariaParaRelacionarIdCompraComChvAcesso()

Dim pTabelaNome As String: pTabelaNome = "tmpCompras_ID_CompraNF"
Dim qryCompras_ID_ComprasNF As String: qryCompras_ID_ComprasNF = _
                                            "SELECT DISTINCT " & _
                                            "   '' AS ID_CompraNF " & _
                                            "   ,ChvAcesso_CompraNF " & _
                                            "INTO tmpCompras_ID_CompraNF " & _
                                            "FROM tblCompraNF; "

    '' EXCLUIR CASO EXISTA
    If Not IsNull(DLookup("Name", "MSysObjects", "type in(1,6) and Name='" & pTabelaNome & "'")) Then Application.CurrentDb.Execute "DROP TABLE " & pTabelaNome
    
    '' CRIAR TABELA
    Application.CurrentDb.Execute qryCompras_ID_ComprasNF, dbSeeChanges
    
End Function

'' 02. Relacionar IdCompra Com ChvAcesso
Public Function relacionarIdCompraComChvAcesso()

Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")


Dim qryUpdateItens_ID_CompraNF As String: qryUpdateItens_ID_CompraNF = _
                                            "UPDATE tmpCompras_ID_CompraNF " & _
                                            "INNER JOIN tblCompraNFItem ON tmpCompras_ID_CompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF " & _
                                            "SET tblCompraNFItem.ID_CompraNF_CompraNFItem = [tmpCompras_ID_CompraNF].[ID_CompraNF]; "

Dim dbDestino As New Banco

Dim db As DAO.Database: Set db = CurrentDb
Dim rst As DAO.Recordset
Set rst = db.OpenRecordset("Select * from tmpCompras_ID_CompraNF")

        dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
        dbDestino.SqlSelect "SELECT * FROM tblCompraNF"
                
        Do While Not dbDestino.rs.EOF
        
            Do While Not rst.EOF
                If dbDestino.rs.Fields("ChvAcesso_CompraNF").value = rst.Fields("ChvAcesso_CompraNF").value Then
                    rst.Edit
                    rst.Fields("ID_CompraNF").value = dbDestino.rs.Fields("ID_CompraNF").value
                    rst.Update
                    Exit Do
                End If
                rst.MoveNext
            Loop
                        
            dbDestino.rs.MoveNext
            rst.MoveFirst
            DoEvents
        Loop

    '' RELACIONAR ITENS DE COMPRAS COM COMPRAS JÁ CADASTRADAS
    Application.CurrentDb.Execute qryUpdateItens_ID_CompraNF

dbDestino.CloseConnection
rst.Close

End Function


'' 03. Enviar Dados Para Servidor
Public Function enviar_ComprasParaServidor(pDestino As String)
On Error Resume Next
'On Error GoTo adm_Err

'' ---------------------
'' VARIAVEIS GERAIS
'' ---------------------

'' LISTAGEM DE CAMPOS DA TABELA ORIGEM/DESTINO
'Dim strCampo As String

Dim qryCampos As String: qryCampos = _
                            "SELECT distinct   " & _
                            "   tblParametros.TipoDeParametro  " & _
                            "   , tblParametros.ID  " & _
                            "   , tblOrigemDestino.campo  " & _
                            "   , tblOrigemDestino.formatacao  " & _
                            "   , tblOrigemDestino.valorPadrao  " & _
                            "FROM   " & _
                            "   tblParametros INNER JOIN tblOrigemDestino ON tblParametros.ValorDoParametro = tblOrigemDestino.campo  " & _
                            "WHERE (((tblParametros.TipoDeParametro)='pDestino') AND ((tblOrigemDestino.TagOrigem)<>0)) ORDER BY tblParametros.ID;"

Dim qryUpdateRegistroProcessado As String: qryUpdateRegistroProcessado = _
        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 2 WHERE (((tblDadosConexaoNFeCTe.ChvAcesso)=""strChave"") AND ((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1));"


'' SCRIPT
Dim tmpScript As String
Dim tmpSelecaoDeRegistros As String
Dim tmp As String

'' ---------------------
'' BANCO LOCAL
'' ---------------------
Dim db As DAO.Database: Set db = CurrentDb
Dim rstOrigem As DAO.Recordset

'' ---------------------
'' BANCO DESTINO
'' ---------------------
Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
Dim sqlCampos As String: sqlCampos = Replace(qryCampos, "pDestino", pDestino)

Debug.Print sqlCampos
If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), sqlCampos

Dim dbDestino As New Banco
Dim rstCampos As DAO.Recordset

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

'' #CONTADOR
Dim contadorDeRegistros As Long: contadorDeRegistros = 1
Dim totalDeRegistros As Long


    dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
    dbDestino.SqlSelect "SELECT * FROM " & pDestino

    tmpScript = "Insert into " & pDestino & " ("

    '' 1. cabecalho
    Set rstCampos = db.OpenRecordset(sqlCampos)
    Do While Not rstCampos.EOF
        If InStr(rstCampos.Fields("campo").value, right(pDestino, Len(pDestino) - 3), vbTextCompare) Then
            tmpScript = tmpScript & rstCampos.Fields("campo").value & ","
        End If

        rstCampos.MoveNext
        DoEvents
    Loop
    tmpScript = left(tmpScript, Len(tmpScript) - 1) & ") values ("
'    Debug.Print tmpScript

    '' #azs - testes
    '' #20210823_EXPORTACAO_LIMITE - LIMITE
    '' BANCO LOCAL
     tmpSelecaoDeRegistros = "Select * from " & pDestino & " Where ChvAcesso_CompraNF IN (SELECT ChvAcesso FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND (tblDadosConexaoNFeCTe.registroProcessado)=1))"
    Debug.Print tmpSelecaoDeRegistros
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), tmpSelecaoDeRegistros

    Set rstOrigem = db.OpenRecordset(tmpSelecaoDeRegistros)

    '' #BARRA_PROGRESSO
    totalDeRegistros = rstOrigem.RecordCount
    SysCmd acSysCmdInitMeter, pDestino, totalDeRegistros

    Do While Not rstOrigem.EOF
        tmp = ""

        '' LISTAGEM DE CAMPOS
        rstCampos.MoveFirst
        Do While Not rstCampos.EOF

            '' CRIAR SCRIPT DE INCLUSAO DE DADOS NA TABELA DESTINO
            '' 2. campos x formatacao
            If InStr(rstCampos.Fields("campo").value, right(pDestino, Len(pDestino) - 3)) Then

                If rstCampos.Fields("formatacao").value = "opTexto" Then
                    tmp = tmp & "'" & rstOrigem.Fields(rstCampos.Fields("campo").value).value & "',"

                ElseIf rstCampos.Fields("formatacao").value = "opNumero" Or rstCampos.Fields("formatacao").value = "opMoeda" Then
                    tmp = tmp & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", rstOrigem.Fields(rstCampos.Fields("campo").value).value, rstCampos.Fields("valorPadrao").value) & ","

                ElseIf rstCampos.Fields("formatacao").value = "opTime" Then
                    tmp = tmp & "'" & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", Format(rstOrigem.Fields(rstCampos.Fields("campo").value).value, DATE_TIME_FORMAT), rstCampos.Fields("valorPadrao").value) & "',"

                ElseIf stCampos.Fields("formatacao").value = "opData" Then
                    tmp = tmp & "'" & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", Format(rstOrigem.Fields(rstCampos.Fields("campo").value).value, DATE_FORMAT), rstCampos.Fields("valorPadrao").value) & "',"

                End If

            End If

            rstCampos.MoveNext
            DoEvents
        Loop

        '' BANCO DESTINO
        tmp = left(tmp, Len(tmp) - 1) & ")"

        Debug.Print tmpScript & tmp
        If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), tmpScript & tmp

        dbDestino.SqlExecute tmpScript & tmp

        '' Terminio de operacao
        If (pDestino = "tblCompraNFItem") Then
            tmp = Replace(qryUpdateRegistroProcessado, "strChave", rstOrigem.Fields("ChvAcesso_CompraNF").value)
            Debug.Print tmp
            If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), tmp

            Application.CurrentDb.Execute tmp
        End If

        '' #BARRA_PROGRESSO
        contadorDeRegistros = contadorDeRegistros + 1
        SysCmd acSysCmdUpdateMeter, contadorDeRegistros

        rstOrigem.MoveNext
        DoEvents
    Loop

    dbDestino.CloseConnection
    db.Close: Set db = Nothing

    '' RELACIONAR ITENS DE COMPRAS COM COMPRAS JÁ CADASTRADAS
    If (pDestino = "tblCompraNF") Then relacionarIdCompraComChvAcesso

    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "enviar_ComprasParaServidor - Exportar registros pendentes ( Quantidade de registros: " & contadorDeRegistros & " )"

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdRemoveMeter

adm_Exit:
    Exit Function

'adm_Err:
'    MsgBox Error$
'    Resume adm_Exit


End Function
