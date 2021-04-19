
''' #TRANSFERIR
'Sub TransferirDadosProcessados(pDestino As String)
'
''' #BANCO_LOCAL
'Dim db As dao.Database: Set db = CurrentDb
'Dim tmpSql As String: tmpSql = "Select Distinct pk from tblProcessamento where NomeTabela = '" & pDestino & "' Order by pk;"
'Dim rstPendentes As dao.Recordset: Set rstPendentes = db.OpenRecordset(tmpSql)
'Dim rstOrigem As dao.Recordset
'
''' #BANCO_DESTINO
'tmpSql = "Select * from " & pDestino
'Dim rstDestino As dao.Recordset: Set rstDestino = db.OpenRecordset(tmpSql)
'
''' #ANALISE_DE_PROCESSAMENTO
'Dim DT_PROCESSO As Date: DT_PROCESSO = Now()
'
''' #BARRA_PROGRESSO
'Dim contadorDeRegistros As Long: contadorDeRegistros = 1
'SysCmd acSysCmdInitMeter, "Transferindo Dados...", rstPendentes.RecordCount
'
'Do While Not rstPendentes.EOF
'
'    '' #BARRA_PROGRESSO
'    SysCmd acSysCmdUpdateMeter, contadorDeRegistros
'
'    '' listar itens de registro para cadastro
'    Set rstOrigem = db.OpenRecordset("Select * from tblProcessamento where NomeTabela = '" & pDestino & "' and pk = '" & rstPendentes.Fields("pk").value & "' Order by ID ")
'    Do While Not rstOrigem.EOF
'
'        '' CONTROLE DE CADASTRO
'        If t = 0 Then rstDestino.AddNew: t = 1
'
'        rstDestino.Fields(rstOrigem.Fields("NomeCampo").value).value = rstOrigem.Fields("Valor").value
'
'        rstOrigem.MoveNext
'        DoEvents
'    Loop
'    rstDestino.Update
'    t = 0
'
'    '' #COMPRAS
'    If (pDestino = "tblCompraNF") Then Application.CurrentDb.Execute Replace(qryUpdateProcessamentoConcluido, "strChave", rstPendentes.Fields("pk").value)
'
'    '' #DADOS_GERAIS
'    '' qryUpdateRegistroValido - Valor padrao
'    If (pDestino = "tblDadosConexaoNFeCTe") Then Application.CurrentDb.Execute "Update tblDadosConexaoNFeCTe SET registroValido = 0 where registroValido is null"
'
'    '' #BARRA_PROGRESSO
'    contadorDeRegistros = contadorDeRegistros + 1
'    rstPendentes.MoveNext
'    DoEvents
'Loop
'
''' #BARRA_PROGRESSO
'SysCmd acSysCmdRemoveMeter
'
''dbDestino.CloseConnection
'db.Close: Set db = Nothing
'
''' #ANALISE_DE_PROCESSAMENTO
'statusFinal DT_PROCESSO, "Processamento - TransferirDadosProcessados"
'
'End Sub