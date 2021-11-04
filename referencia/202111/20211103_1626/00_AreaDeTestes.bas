Attribute VB_Name = "00_AreaDeTestes"
Option Compare Database

'' file:///C:/XMLS/68.365.5010003-77%20-%20Proparts%20Com%C3%A9rcio%20de%20Artigos%20Esportivos%20e%20Tecnologia%20Ltda/recebimento/42210300634453001303570010001139451001171544-cteproc.xml

'' #20211102_Ailton
Sub teste_ProcessamentoTransferir()
Dim pChave As String: _
    pChave = "42210368365501000377550000000064281001362494-nfeproc" '"42210348740351012767570000021186731952977908-cteproc"
    
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


Dim db As DAO.Database: Set db = CurrentDb
Dim rstProcessamento As DAO.Recordset: Set rstProcessamento = db.OpenRecordset(Replace(qryProcessamento_Select_CompraComItens, "pChave", pChave))
Dim tmp As String: tmp = "select distinct tmp.pk from (" & Replace(qryProcessamento_Select_CompraComItens, "pChave", pChave) & ") as tmp order by tmp.pk; "
Debug.Print tmp

Dim rstRegistros As DAO.Recordset: Set rstRegistros = db.OpenRecordset(tmp)

Dim item As Variant

Dim tmpChave As String: tmpChave = ""
Dim pRepositorio As String: pRepositorio = ""
Dim strCamposNomes As String: strCamposNomes = ""
Dim strCamposValores As String: strCamposValores = ""
'Dim strpk As String
    
    Do While Not rstRegistros.EOF
    
'        strpk = rstRegistros.Fields("pk").value
        
        rstProcessamento.Filter = "pk = '" & rstRegistros.Fields("pk").value & "'"
        Do While Not rstProcessamento.EOF
            
            '' ATRIBUIÇÃO DE CHAVE INICIAL
            If tmpChave <> rstProcessamento.Fields("pk").value Then
                tmpChave = rstProcessamento.Fields("pk").value
                pRepositorio = rstProcessamento.Fields("NomeTabela").value
                strCamposNomes = ""
                strCamposValores = ""
            End If
            
            '' SELEÇÃO DE VALORES DOS CAMPOS
            strCamposNomes = strCamposNomes & rstProcessamento.Fields("NomeCampo").value & ","
            
            If rstProcessamento.Fields("formatacao").value = "opTexto" Then
                strCamposValores = strCamposValores & "'" & rstProcessamento.Fields("Valor").value & "',"
                
            ElseIf rstProcessamento.Fields("formatacao").value = "opNumero" Or rstProcessamento.Fields("formatacao").value = "opMoeda" Then
                strCamposValores = strCamposValores & rstProcessamento.Fields("Valor").value & ","
            
            ElseIf rstProcessamento.Fields("formatacao").value = "opTime" Then
                strCamposValores = strCamposValores & "'" & Format(rstProcessamento.Fields("Valor").value, DATE_TIME_FORMAT) & "',"
            
            ElseIf rstProcessamento.Fields("formatacao").value = "opData" Then
                strCamposValores = strCamposValores & "'" & Format(rstProcessamento.Fields("Valor").value, DATE_FORMAT) & "',"
            
            End If
            
            rstProcessamento.MoveNext
            DoEvents
            
        Loop
    
        '' EXIT SCRIPT
        strCamposNomes = left(strCamposNomes, Len(strCamposNomes) - 1)
        strCamposValores = left(strCamposValores, Len(strCamposValores) - 1)
        
        ''Application.CurrentDb.Execute
        Debug.Print _
            Replace(Replace(Replace(qryCompras_Insert_Processamento, "strCamposNomes", strCamposNomes), "strCamposValores", strCamposValores), "pRepositorio", pRepositorio)
        
        rstRegistros.MoveNext
        DoEvents
    Loop
    
   

End Sub

