Attribute VB_Name = "00_AreaDeTestes"
Option Compare Database

'' file:///C:/XMLS/68.365.5010003-77%20-%20Proparts%20Com%C3%A9rcio%20de%20Artigos%20Esportivos%20e%20Tecnologia%20Ltda/recebimento/42210300634453001303570010001139451001171544-cteproc.xml


Sub TransferirProcessamentoParaRepositorios(pChave As String)
    
Dim qryProcessamento_Select_CompraComItens As String: _
    qryProcessamento_Select_CompraComItens = _
    "SELECT DISTINCT tblProcessamento.NomeTabela " & _
        "   ,tblProcessamento.pk " & _
        "   ,tblProcessamento.NomeCampo " & _
        "   ,tblProcessamento.valor " & _
        "   ,tblProcessamento.formatacao " & _
        "FROM tblParametros " & _
        "INNER JOIN tblProcessamento ON (tblParametros.ValorDoParametro = tblProcessamento.NomeCampo) AND (tblParametros.TipoDeParametro = tblProcessamento.NomeTabela) " & _
        "WHERE (((tblProcessamento.pk) LIKE 'pChave*') AND ((tblProcessamento.NomeCampo) IS NOT NULL)) " & _
        "ORDER BY tblProcessamento.NomeTabela,tblProcessamento.pk"

Dim qryCompras_Insert_Processamento As String: _
    qryCompras_Insert_Processamento = "INSERT INTO pRepositorio (strCamposNomes) SELECT strCamposValores  "

Dim DadosGerais As New clsConexaoNfeCte
Dim db As DAO.Database: Set db = CurrentDb
Dim rstProcessamento As DAO.Recordset: Set rstProcessamento = db.OpenRecordset(Replace(qryProcessamento_Select_CompraComItens, "pChave", pChave))
Dim rstRegistros As DAO.Recordset: Set rstRegistros = db.OpenRecordset("select distinct tmp.pk from (" & Replace(qryProcessamento_Select_CompraComItens, "pChave", pChave) & ") as tmp order by tmp.pk;")
Dim rstFiltered As DAO.Recordset

Dim item As Variant

Dim pRepositorio As String
Dim strChave As String
Dim strCamposNomes As String
Dim strCamposValores As String
Dim strCompras_Insert_Processamento As String
    
    '' REGISTROS
    Do While Not rstRegistros.EOF
        
        '' VARIAVEIS
        strCamposNomes = ""
        strCamposValores = ""
        strCompras_Insert_Processamento = qryCompras_Insert_Processamento
        
        '' SELE��O DE ITEM
        strChave = rstRegistros.Fields("pk").value
        rstProcessamento.Filter = "pk = '" & strChave & "'"
            
        '' PROCESSAMENTO DO ITEM
        Set rstFiltered = rstProcessamento.OpenRecordset
        Do While Not rstFiltered.EOF
        
            pRepositorio = rstFiltered.Fields("NomeTabela").value

            '' NOME DAS COLUNAS
            strCamposNomes = strCamposNomes & rstFiltered.Fields("NomeCampo").value & ","
            
            '' VALOR DAS COLUNAS
            If rstFiltered.Fields("formatacao").value = "opTexto" Then
                strCamposValores = strCamposValores & "'" & rstFiltered.Fields("Valor").value & "',"
                
            ElseIf rstFiltered.Fields("formatacao").value = "opNumero" Or rstFiltered.Fields("formatacao").value = "opMoeda" Then
                strCamposValores = strCamposValores & rstFiltered.Fields("Valor").value & ","
            
            ElseIf rstFiltered.Fields("formatacao").value = "opTime" Then
                strCamposValores = strCamposValores & "'" & Format(rstFiltered.Fields("Valor").value, DATE_TIME_FORMAT) & "',"
            
            ElseIf rstFiltered.Fields("formatacao").value = "opData" Then
                strCamposValores = strCamposValores & "'" & Format(rstFiltered.Fields("Valor").value, DATE_FORMAT) & "',"
            
            End If
       
            rstFiltered.MoveNext
            DoEvents
        Loop
        
        '' EXIT SCRIPT
        strCamposNomes = left(strCamposNomes, Len(strCamposNomes) - 1)
        strCamposValores = left(strCamposValores, Len(strCamposValores) - 1)
        
        ''Debug.Print
        Application.CurrentDb.Execute Replace(Replace(Replace(strCompras_Insert_Processamento, "strCamposNomes", strCamposNomes), "strCamposValores", strCamposValores), "pRepositorio", pRepositorio)
        
        rstRegistros.MoveNext
        DoEvents
    Loop
        
    '' registroProcessado
    Application.CurrentDb.Execute Replace(DadosGerais.UpdateProcessamentoConcluido, "strChave", strChave)
        
        
'' Cleanup
rstFiltered.Close
rstProcessamento.Close
rstRegistros.Close

Set rstFiltered = Nothing
Set rstProcessamento = Nothing
Set rstRegistros = Nothing

End Sub

'' #20211103_Ailton
'Sub testeUnitario()
'
'    processarArquivosPendentes
'
'    ProcessamentoTransferirParaRepositorios
'
'End Sub


'' #20211102_Ailton
'Sub ProcessamentoTransferirParaRepositorios()
'Dim pChave As String: _
'    pChave = DLookup("[Chave]", "[tblDadosConexaoNFeCTe]", "[CaminhoDoArquivo]='C:\xmls\68.365.5010003-77 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210348740351012767570000021186731952977908-cteproc.xml'") '"42210348740351012767570000021186731952977908-cteproc"
'    'pChave = "42210368365501000377550000000064281001362494-nfeproc"
'
'Dim qryProcessamento_Select_CompraComItens As String: _
'    qryProcessamento_Select_CompraComItens = _
'    "SELECT DISTINCT tblProcessamento.NomeTabela " & _
'        "   ,tblProcessamento.pk " & _
'        "   ,tblProcessamento.NomeCampo " & _
'        "   ,tblProcessamento.valor " & _
'        "   ,tblProcessamento.formatacao " & _
'        "FROM tblParametros " & _
'        "INNER JOIN tblProcessamento ON (tblParametros.ValorDoParametro = tblProcessamento.NomeCampo) AND (tblParametros.TipoDeParametro = tblProcessamento.NomeTabela) " & _
'        "WHERE (((tblProcessamento.pk) LIKE 'pChave*') AND ((tblProcessamento.NomeCampo) IS NOT NULL)) " & _
'        "ORDER BY tblProcessamento.NomeTabela,tblProcessamento.pk"
'
'Dim qryCompras_Insert_Processamento As String: _
'    qryCompras_Insert_Processamento = "INSERT INTO pRepositorio (strCamposNomes) SELECT strCamposValores  "
'
'
'Dim db As DAO.Database: Set db = CurrentDb
'Dim rstProcessamento As DAO.Recordset: Set rstProcessamento = db.OpenRecordset(Replace(qryProcessamento_Select_CompraComItens, "pChave", pChave))
'Dim rstRegistros As DAO.Recordset: Set rstRegistros = db.OpenRecordset("select distinct tmp.pk from (" & Replace(qryProcessamento_Select_CompraComItens, "pChave", pChave) & ") as tmp order by tmp.pk;")
'Dim rstFiltered As DAO.Recordset
'
'Dim item As Variant
'
'Dim pRepositorio As String
'Dim strCamposNomes As String
'Dim strCamposValores As String
'Dim strCompras_Insert_Processamento As String
'
'    '' REGISTROS
'    Do While Not rstRegistros.EOF
'
'        '' VARIAVEIS
'        strCamposNomes = ""
'        strCamposValores = ""
'        strCompras_Insert_Processamento = qryCompras_Insert_Processamento
'
'        '' SELE��O DE ITEM
'        rstProcessamento.Filter = "pk = '" & rstRegistros.Fields("pk").value & "'"
''        Debug.Print rstRegistros.Fields("pk").value
'
'        '' PROCESSAMENTO DO ITEM
'        Set rstFiltered = rstProcessamento.OpenRecordset
'        Do While Not rstFiltered.EOF
'
'            pRepositorio = rstFiltered.Fields("NomeTabela").value
'
'            '' NOME DAS COLUNAS
'            strCamposNomes = strCamposNomes & rstFiltered.Fields("NomeCampo").value & ","
'
'            '' VALOR DAS COLUNAS
'            If rstFiltered.Fields("formatacao").value = "opTexto" Then
'                strCamposValores = strCamposValores & "'" & rstFiltered.Fields("Valor").value & "',"
'
'            ElseIf rstFiltered.Fields("formatacao").value = "opNumero" Or rstFiltered.Fields("formatacao").value = "opMoeda" Then
'                strCamposValores = strCamposValores & rstFiltered.Fields("Valor").value & ","
'
'            ElseIf rstFiltered.Fields("formatacao").value = "opTime" Then
'                strCamposValores = strCamposValores & "'" & Format(rstFiltered.Fields("Valor").value, DATE_TIME_FORMAT) & "',"
'
'            ElseIf rstFiltered.Fields("formatacao").value = "opData" Then
'                strCamposValores = strCamposValores & "'" & Format(rstFiltered.Fields("Valor").value, DATE_FORMAT) & "',"
'
'            End If
'
'            rstFiltered.MoveNext
'            DoEvents
'        Loop
'
'        '' EXIT SCRIPT
'        strCamposNomes = left(strCamposNomes, Len(strCamposNomes) - 1)
'        strCamposValores = left(strCamposValores, Len(strCamposValores) - 1)
'
'        ''Debug.Print
'        Application.CurrentDb.Execute Replace(Replace(Replace(strCompras_Insert_Processamento, "strCamposNomes", strCamposNomes), "strCamposValores", strCamposValores), "pRepositorio", pRepositorio)
'
'        rstRegistros.MoveNext
'        DoEvents
'    Loop
'
''' Cleanup
'rstFiltered.Close
'rstProcessamento.Close
'rstRegistros.Close
'
'Set rstFiltered = Nothing
'Set rstProcessamento = Nothing
'Set rstRegistros = Nothing
'
'End Sub
