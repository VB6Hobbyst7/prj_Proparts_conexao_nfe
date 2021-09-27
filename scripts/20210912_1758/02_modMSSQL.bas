Attribute VB_Name = "02_modMSSQL"





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
Dim strCampo As String

Dim qryCampos As String: qryCampos = _
                            "SELECT distinct   " & _
                            "   tblParametros.TipoDeParametro  " & _
                            "   , tblOrigemDestino.campo  " & _
                            "   , tblOrigemDestino.formatacao  " & _
                            "   , tblOrigemDestino.valorPadrao  " & _
                            "FROM   " & _
                            "   tblParametros INNER JOIN tblOrigemDestino ON tblParametros.ValorDoParametro = tblOrigemDestino.campo  " & _
                            "WHERE (((tblParametros.TipoDeParametro)='pDestino'));"

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
Dim strUsuarioNome As String
Dim strUsuarioSenha As String
Dim strOrigem As String
Dim strBanco As String
Dim sqlCampos As String: sqlCampos = Replace(qryCampos, "pDestino", pDestino)

Dim dbDestino As New Banco

'Dim retVal As Variant: retVal = MsgBox("Deseja enviar compras para o servidor?", vbQuestion + vbYesNo, "ADM_enviarComprasParaServidor")
'
'    If retVal = vbYes Then
        strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
        strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
        strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
        strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
    
        dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
        dbDestino.SqlSelect "SELECT * FROM " & pDestino
    
        tmpScript = "Insert into " & pDestino & " ("
    
        '' 1. cabecalho
        Dim rstCampos As DAO.Recordset
        Set rstCampos = db.OpenRecordset(sqlCampos)
        Do While Not rstCampos.EOF
            If InStr(rstCampos.Fields("campo").value, right(pDestino, Len(pDestino) - 3), vbTextCompare) Then
                tmpScript = tmpScript & rstCampos.Fields("campo").value & ","
            End If
            
            rstCampos.MoveNext
            DoEvents
        Loop
        tmpScript = left(tmpScript, Len(tmpScript) - 1) & ") values ("
        
        
        '' #20210823_EXPORTACAO_LIMITE - LIMITE
        '' BANCO LOCAL
        tmpSelecaoDeRegistros = "Select * from " & pDestino & " Where ChvAcesso_CompraNF NOT IN (SELECT tmpCompras_ID_CompraNF.ChvAcesso_CompraNF FROM tmpCompras_ID_CompraNF)"
        Set rstOrigem = db.OpenRecordset(tmpSelecaoDeRegistros)
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
                        
                    ElseIf rstCampos.Fields("formatacao").value = "opTime" Or rstCampos.Fields("formatacao").value = "opData" Then
                        tmp = tmp & "'" & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", rstOrigem.Fields(rstCampos.Fields("campo").value).value, rstCampos.Fields("valorPadrao").value) & "',"
                        
                    End If
                
                End If
                
                rstCampos.MoveNext
                DoEvents
            Loop
            
            '' BANCO DESTINO
            tmp = left(tmp, Len(tmp) - 1) & ")"
            
            Debug.Print tmpScript & tmp
            
            rstOrigem.MoveNext
            
            dbDestino.SqlExecute tmpScript & tmp
                        
            DoEvents
        Loop
        
        dbDestino.CloseConnection
        db.Close: Set db = Nothing
        
        
        '' RELACIONAR ITENS DE COMPRAS COM COMPRAS JÁ CADASTRADAS
        If (pDestino = "tblCompraNF") Then relacionarIdCompraComChvAcesso
        
'    End If

adm_Exit:
    Exit Function

'adm_Err:
'    MsgBox Error$
'    Resume adm_Exit


End Function
