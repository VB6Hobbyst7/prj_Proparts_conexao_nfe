Attribute VB_Name = "azsRepositorioDeTestes"
Option Compare Database


'' #####################################################################
'' ### #DESENVOLVIMENTO
'' #####################################################################




'' #####################################################################
'' ### #TESTES
'' #####################################################################

'' Teste de conn com SQL
Sub teste_cnn()
Dim b As New Banco

    '' INICIO
    b.Start "sa", "41L70N@@", "WIN-VE2KJO1LP3\SQLEXPRESS", "SispartsConexao", drSqlServer
    
    '' SELECT
    b.SqlSelect "Select * from tblCompraNF"
    
    '' INSERT
    b.SqlExecute "Insert into tblCompraNF (HoraEntd_CompraNF) values ('15:26:00')"
    
    '' COUNT
    Debug.Print b.rs.RecordCount
    
    '' FIM
    b.CloseConnection

End Sub

Private Sub teste_FiltrarCompraItens()
Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
Dim QRY() As Variant: QRY = Array("chCTe")
Dim Item As Variant
Dim lists As Variant
Dim fieldnode As Variant
Dim childNode As Variant

'' cte
'XDoc.Load "C:\temp\Coleta\68.365.5010002-96 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\32210204884082000569570000039548351039548356-cteproc.xml"

'' nfe
'XDoc.Load "C:\temp\Coleta\68.365.5010002-96 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\29210220961864000187550010000001891138200000-nfeproc.xml"

For Each Item In QRY
    Set lists = XDoc.SelectNodes("//" & Item)
    For Each fieldnode In lists
        If (fieldnode.HasChildNodes) Then
            For Each childNode In fieldnode.ChildNodes
                Debug.Print fieldnode.text
            Next childNode
        End If
    Next fieldnode
Next Item

Set XDoc = Nothing

End Sub


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



Function teste_getConsultarSeRetornoArmazemParaRecuperarNumeroDePedido() As String
'' 1.  Verificar se retorno do armazem (5)
'' 1.2 Recuperar numero de pedido
'' 2   Caso contrario
'' 2.1 Retorno (0)


'' valor padrao
Dim tRetorno As String: tRetorno = 0

'' Codigo do Retorno de armazem
Dim tTipo As String: tTipo = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='RetornoArmazem'")


'' chave
Dim pChvAcesso As String: pChvAcesso = "42210312680452000302550020000897611617746185"
'' tipo de cadastro
Dim tTipoCadastro As String: tTipoCadastro = DLookup("[ID_Tipo]", "[tblDadosConexaoNFeCTe]", "[ChvAcesso]='" & pChvAcesso & "'")

'' ----------------------[pChvAcesso]

'' dados
Dim pDados As String: pDados = "PEDIDO: 322382, . RETORNO SIMBOLICO DE ARMAZENAGEM DE SUA(S) NF-E(S) 4027 SERIE 0 DE 08/01/2021 CHAVE: 42210168365501000377550000000040271001314575 LOTE: , , 5359 SERIE 0 DE 04/02/2021. REFERENTE SUA(S) NF-E(S) DE Venda NUMERO 5735 SERIE 0 DE 15/02/2021, "

'' ----------------------[pDados]


'' limpar dado inicial
Dim tValor() As Variant: tValor = Array("PEDIDO:", "PEDIDO ")

    
    If tTipoCadastro = tTipo Then
        tRetorno = left(Trim(Replace(Replace(pDados, tValor(0), ""), tValor(1), "")), 6)
    End If
    
    
    Debug.Print tRetorno

End Function

Sub TESTE_PEDIDO()

Dim tmp As String
Dim tDados As String
Dim tValor() As Variant: tValor = Array("PEDIDO:", "PEDIDO ")

tDados = "PEDIDO: 322382, . RETORNO SIMBOLICO DE ARMAZENAGEM DE SUA(S) NF-E(S) 4027 SERIE 0 DE 08/01/2021 CHAVE: 42210168365501000377550000000040271001314575 LOTE: , , 5359 SERIE 0 DE 04/02/2021. REFERENTE SUA(S) NF-E(S) DE Venda NUMERO 5735 SERIE 0 DE 15/02/2021, "

tmp = left(Trim(Replace(Replace(tDados, tValor(0), ""), tValor(1), "")), 6)

Debug.Print tmp


End Sub

'Sub teste_cadastroProcessamento()
'
'Dim strTemp As String: strTemp = _
'    "INSERT INTO tblProcessamento (valor, NomeTabela, NomeCampo, formatacao, NomeCampo) " & _
'    "SELECT tblorigemDestino.valorPadrao    ,tblorigemDestino.tabela ,tblorigemDestino.campo ,tblorigemDestino.formatacao ,tblorigemDestino.campo " & _
'    "FROM tblorigemDestino " & _
'    "WHERE (((tblorigemDestino.tabela) = 'tblCompraNF')  " & _
'    "AND ((tblorigemDestino.formatacao) = 'opMoeda')  " & _
'    "AND ((tblorigemDestino.campo)  " & _
'    "NOT IN (SELECT DISTINCT tblProcessamento.nomecampo,tblProcessamento.pk FROM tblProcessamento WHERE (((tblProcessamento.nomecampo) IS NOT NULL) AND ((tblProcessamento.[pk]) EXISTS (SELECT DISTINCT tblProcessamento.pk FROM tblProcessamento))))));"
'
'Application.CurrentDb.Execute strTemp
'
'End Sub



