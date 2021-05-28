'' 03.ENVIAR DADOS PARA SERVIDOR
Sub TESTE_EnviarDadosParaServidor()

'' VARIAVEL DE PARAMETRO
Dim pDestino As String: pDestino = "tblCompraNF"

'' ---------------------
'' BANCO LOCAL
'' ---------------------
Dim db As dao.Database: Set db = CurrentDb
Dim rstOrigem As dao.Recordset

'' ---------------------
'' BANCO DESTINO
'' ---------------------
Dim strUsuarioNome As String: strUsuarioNome = pegarValorDoParametro(qryParametros, "BancoDados_Usuario")
Dim strUsuarioSenha As String: strUsuarioSenha = pegarValorDoParametro(qryParametros, "BancoDados_Senha")
Dim strOrigem As String: strOrigem = pegarValorDoParametro(qryParametros, "BancoDados_Origem")
Dim strBanco As String: strBanco = pegarValorDoParametro(qryParametros, "BancoDados_Banco")

Dim dbDestino As New Banco
dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
dbDestino.SqlSelect "SELECT * FROM " & pDestino

'' ---------------------
'' VARIAVEIS GERAIS
'' ---------------------

'' LISTAGEM DE CAMPOS DA TABELA ORIGEM/DESTINO
Dim strCampo As String

'' SCRIPT
Dim tmpScript As String: tmpScript = "Insert into " & pDestino & " ("


'' 1. cabeçalho
Dim rstCampos As dao.Recordset
Set rstCampos = db.OpenRecordset("Select campo,formatacao,valorPadrao from tblOrigemDestino where tblOrigemDestino.tabela = '" & pDestino & "' and tagOrigem = 1 order by tblOrigemDestino.id")
Do While Not rstCampos.EOF
    tmpScript = tmpScript & rstCampos.Fields("campo").value & ","
    rstCampos.MoveNext
    DoEvents
Loop
tmpScript = left(tmpScript, Len(tmpScript) - 1) & ") values ("
rstCampos.MoveFirst

'' BANCO LOCAL
Set rstOrigem = db.OpenRecordset("Select * from " & pDestino)
Do While Not rstOrigem.EOF

    
    '' LISTAGEM DE CAMPOS
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
    dbDestino.SqlExecute tmpScript & tmp
    
'    dbDestino.rs.Update
    
    rstOrigem.MoveNext
    DoEvents
Loop

dbDestino.CloseConnection
db.Close: Set db = Nothing

End Sub