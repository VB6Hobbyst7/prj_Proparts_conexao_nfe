'' 03.ENVIAR DADOS PARA SERVIDOR
Public Function enviar_ComprasParaServidor()
On Error GoTo adm_Err

'' VARIAVEL DE PARAMETRO
Dim pDestino As String: pDestino = "tblCompraNF"

'' ---------------------
'' VARIAVEIS GERAIS
'' ---------------------

'' LISTAGEM DE CAMPOS DA TABELA ORIGEM/DESTINO
Dim strCampo As String

'' SCRIPT
Dim tmpScript As String
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

Dim dbDestino As New Banco

Dim retVal As Variant: retVal = MsgBox("Deseja enviar compras para o servidor?", vbQuestion + vbYesNo, "ADM_enviarComprasParaServidor")

    If retVal = vbYes Then
        strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
        strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
        strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
        strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
    
        dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
        dbDestino.SqlSelect "SELECT * FROM " & pDestino
    
        tmpScript = "Insert into " & pDestino & " ("
    
        '' 1. cabeçalho
        Dim rstCampos As DAO.Recordset
        Set rstCampos = db.OpenRecordset("Select distinct campo,formatacao,valorPadrao from tblOrigemDestino where tblOrigemDestino.tabela = '" & pDestino & "'")
        Do While Not rstCampos.EOF
            tmpScript = tmpScript & rstCampos.Fields("campo").value & ","
            rstCampos.MoveNext
            DoEvents
        Loop
        tmpScript = left(tmpScript, Len(tmpScript) - 1) & ") values ("
        
        
        '' BANCO LOCAL
        Set rstOrigem = db.OpenRecordset("Select * from " & pDestino)
        Do While Not rstOrigem.EOF
            tmp = ""
            
            '' LISTAGEM DE CAMPOS
            rstCampos.MoveFirst
            Do While Not rstCampos.EOF
            
                '' CRIAR SCRIPT DE INCLUSÃO DE DADOS NA TABELA DESTINO
                '' 2. campos x formatação
        
                If rstCampos.Fields("formatacao").value = "opTexto" Then
                    tmp = tmp & "'" & rstOrigem.Fields(rstCampos.Fields("campo").value).value & "',"
                    
                ElseIf rstCampos.Fields("formatacao").value = "opNumero" Or rstCampos.Fields("formatacao").value = "opMoeda" Then
                    tmp = tmp & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", rstOrigem.Fields(rstCampos.Fields("campo").value).value, rstCampos.Fields("valorPadrao").value) & ","
                    
                ElseIf rstCampos.Fields("formatacao").value = "opTime" Or rstCampos.Fields("formatacao").value = "opData" Then
                    tmp = tmp & "'" & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", rstOrigem.Fields(rstCampos.Fields("campo").value).value, rstCampos.Fields("valorPadrao").value) & "',"
                    
                End If
                
                rstCampos.MoveNext
                DoEvents
            Loop
            
            '' BANCO DESTINO
            tmp = left(tmp, Len(tmp) - 1) & ")"
            
            'Debug.Print tmpScript & tmp
            
            rstOrigem.MoveNext
            
            dbDestino.SqlExecute tmpScript & tmp
            
            DoEvents
        Loop
        
        dbDestino.CloseConnection
        db.Close: Set db = Nothing
        
'        MsgBox "Fim!", vbOKOnly + vbExclamation, "enviarComprasParaServidor"
    
    End If

adm_Exit:
    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit


End Function