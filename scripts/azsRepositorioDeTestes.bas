Attribute VB_Name = "azsRepositorioDeTestes"
Option Compare Database

'' #####################################################################
'' ### #DESENVOLVIMENTO
'' #####################################################################

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
Set rstCampos = db.OpenRecordset("Select distinct * from tblOrigemDestino where tabela = '" & pDestino & "' and tagOrigem = 1")
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

    '' BANCO DESTINO
'    dbDestino.rs.AddNew
    
    '' LISTAGEM DE CAMPOS
    Do While Not rstCampos.EOF
    
        '' CRIAR SCRIPT DE INCLUSÃO DE DADOS NA TABELA DESTINO
        '' 2. campos x formatação

        If rstCampos.Fields("formatacao").value = "opTexto" Then
            tmp = tmp & "'" & rstOrigem.Fields(rstCampos.Fields("campo").value).value & "',"
        ElseIf rstCampos.Fields("formatacao").value = "opNumero" Or rstCampos.Fields("formatacao").value = "opMoeda" Then
            tmp = tmp & IIf(IsNull(rstOrigem.Fields(rstCampos.Fields("campo").value).value), 0, rstOrigem.Fields(rstCampos.Fields("campo").value).value) & ","
        ElseIf rstCampos.Fields("formatacao").value = "opTime" Or rstCampos.Fields("formatacao").value = "opData" Then
            tmp = tmp & rstOrigem.Fields(rstCampos.Fields("campo").value).value & ","
        End If
        
        rstCampos.MoveNext
        DoEvents
    Loop
    
    tmp = left(tmp, Len(tmp) - 1) & ")"
    dbDestino.SqlExecute tmpScript & tmp
    
'    dbDestino.rs.Update
    
    rstOrigem.MoveNext
    DoEvents
Loop

dbDestino.CloseConnection
db.Close: Set db = Nothing

End Sub

'' #####################################################################
'' ### #TESTES
'' #####################################################################


Sub TESTE_IDVD()
Dim db As dao.Database: Set db = CurrentDb
Dim tmpSql As String: tmpSql = "Select * from tblCompraNF ORDER BY ID_CompraNF;"
Dim rstPendentes As dao.Recordset: Set rstPendentes = db.OpenRecordset(tmpSql)
Dim parts() As String

Do While Not rstPendentes.EOF
    
    rstPendentes.Edit
    
    If rstPendentes.Fields("IDVD_CompraNF").value <> "" Then
        rstPendentes.Fields("IDVD_CompraNF").value = Replace(parts(LBound(Split((rstPendentes.Fields("IDVD_CompraNF").value), ","))), "Pedido", "")
    Else
        rstPendentes.Fields("IDVD_CompraNF").value = 0
    End If
    
    rstPendentes.Update
    rstPendentes.MoveNext
Loop

db.Close: Set db = Nothing

End Sub



'' Teste de conn com SQL
Sub teste_cnn()
Dim b As New Banco

    '' INICIO
    b.Start "sa", "41L70N@@", "WIN-VE2KJO1LP3\SQLEXPRESS", "SispartsConexao", drSqlServer
    
    '' SELECT
    b.SqlSelect "Select * from tblCompraNF"
    
    '' INSERT
    b.SqlExecute "Insert into tblCompraNF (DTEntd_CompraNF) values ('2020-02-15')"
    
    '' COUNT
    Debug.Print b.rs.RecordCount
    
    '' FIM
    b.CloseConnection

End Sub

'' Progress
Sub ProgressMeter()
   Dim MyDB As dao.Database, MyTable As dao.Recordset
   Dim Count As Long
   Dim Progress_Amount As Integer
    
   Set MyDB = CurrentDb()
   Set MyTable = MyDB.OpenRecordset("tblProcessamento")
 
   ' Move to last record of the table to get the total number of records.
   MyTable.MoveLast
   Count = MyTable.RecordCount
 
   ' Move back to first record.
   MyTable.MoveFirst
 
   ' Initialize the progress meter.
    SysCmd acSysCmdInitMeter, "Reading Data...", Count
 
   ' Enumerate through all the records.
   For Progress_Amount = 1 To Count
     ' Update the progress meter.
      SysCmd acSysCmdUpdateMeter, Progress_Amount
       
     'Print the contact name and number of orders in the Immediate window.
      Debug.Print MyTable![pk] ''; Count("[OrderID]", "Orders", "[CustomerID]='" & MyTable![CustomerID] & "'")
                   
     ' Go to the next record.
      MyTable.MoveNext
   Next Progress_Amount
 
   ' Remove the progress meter.
   SysCmd acSysCmdRemoveMeter
         
End Sub
