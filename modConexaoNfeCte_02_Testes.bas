Attribute VB_Name = "modConexaoNfeCte_02_Testes"
Option Compare Database

Private Const qryFields As String = "SELECT strSplit([Destino],'.',1) AS strCampo FROM tblOrigemDestino WHERE (((tblOrigemDestino.tabela)='strParametro'));"
Private Const sqyProcessamentosPendentes As String = "SELECT DISTINCT pk from tblProcessamento;"
Private Const qryProcessamento As String = "Select nomecampo, valor from tblProcessamento where pk = 'strChave' and len(nomecampo)>0;"

Private Const sqyCompras As String = "SELECT * FROM tblCompraNF"



'' #####################################################################
'' ### #PENDENCIAS
'' #####################################################################

'' #PENDENTE - IDVD_CompraNF
'UPDATE tblProcessamento SET tblProcessamento.valor = TRIM(Replace(Replace([tblProcessamento].[valor],'PEDIDO',''),';','')) WHERE (((tblProcessamento.NomeCampo)='IDVD_CompraNF'));
'UPDATE tblProcessamento SET tblProcessamento.valor = TRIM(Replace(parts(LBound(Split(([tblProcessamento].[valor]), ","))), "Pedido:", "")) WHERE (((tblProcessamento.NomeCampo)='IDVD_CompraNF'));

Sub TESTES_20210329_0754()
Dim s As New clsConexaoNfeCte
Dim c As New Collection
Dim t As Variant

    '' ###########################
    '' #LIMPAR_BASE_DE_TESTES
    '' ###########################
'    Application.CurrentDb.Execute "DELETE FROM tblDadosConexaoNFeCTe"
    Application.CurrentDb.Execute "DELETE FROM tblCompraNF"
    Application.CurrentDb.Execute "DELETE FROM tblCompraNFItem"

    '' ###########################
    '' #CAPTURA_DADOS_GERAIS
    '' ###########################
'    For Each t In Array("tblDadosConexaoNFeCTe")
'
'        '' PROCESSAMENTO DE ARQUIVO
'        s.ProcessarArquivosXml CStr(t), GetFilesInSubFolders(pegarValorDoParametro(qryParametros, strCaminhoDeColeta))
'
'        '' TRANSFERIR DADOS
'        s.TransferirDadosConexaoNFeCTe
'
'        '' GERAR ARQUIVOS
'        s.CriarTipoDeArquivo opFlagLancadaERP
'        s.CriarTipoDeArquivo opManifesto
'
'    Next



    '' ###########################
    '' #CAPTURA_COMPRAS
    '' ###########################
    For Each t In Array("tblCompraNF")
        '' PROCESSAR APENAS ARQUIVOS VALIDOS
        s.ProcessarArquivosXml CStr(t), carregarParametros(qrySelectProcessamentoPendente)

    Next

    '' #FORMATAR
    formatarCampos
    MsgBox "Concluido!", vbOKOnly + vbInformation, "formatarCampos"

    '' #TRANSFERIR
    TransferirDados
    MsgBox "Concluido!", vbOKOnly + vbInformation, "TransferirDados_COMPRAS"

    '' #VALIDAR_DADOS
    criarConsultasParaTestes
    MsgBox "Concluido!", vbOKOnly + vbInformation, "criarConsultasParaTestes"
    
    MsgBox "Terminio!", vbOKOnly + vbExclamation, "TESTES_20210329_0754"

End Sub




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


'' #TRANSFERIR
Sub TransferirDados()

'' BANCO LOCAL
Dim db As DAO.Database: Set db = CurrentDb
Dim rstOrigem As DAO.Recordset: Set rstOrigem = db.OpenRecordset("Select * from tblCompraNF")

'' BANCO DESTINO
Dim dbDestino As New Banco: dbDestino.Start "sa", "41L70N@@", "WIN-VE2KJO1LP3\SQLEXPRESS", "SispartsConexao", drSqlServer
dbDestino.SqlSelect sqyCompras


'' LISTAGEM DE CAMPOS DA TABELA ORIGEM/DESTINO
Dim rstCampos As DAO.Recordset: Set rstCampos = db.OpenRecordset("Select * from tblOrigemDestino where tabela = 'tblCompraNF' and tagOrigem = 1")


Dim StrCampo As String

Do While Not rstOrigem.EOF

    dbDestino.rs.AddNew
    
    Do While Not rstCampos.EOF
    
        StrCampo = strSplit(rstCampos.Fields("destino").value, ".", 1)
    
        For Each fldDestino In dbDestino.rs.Fields
            If fldDestino.Name = StrCampo Then
                ' dbDestino.rs(StrCampo).value = rstOrigem.Fields(StrCampo).value
                If rstOrigem.Fields(StrCampo).value <> "" Then dbDestino.SqlExecute "Insert into tblCompraNF (" & StrCampo & ") values (" & rstOrigem.Fields(StrCampo).value & ")"
                Exit For
            End If
        Next fldDestino
        
        rstCampos.MoveNext
    Loop
    
    dbDestino.rs.Update
    
    rstOrigem.MoveNext
    
Loop

dbDestino.CloseConnection
db.Close: Set db = Nothing

End Sub

'' Progress
Sub ProgressMeter()
   Dim MyDB As DAO.Database, MyTable As DAO.Recordset
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


'' #######################################################################################################################

'Sub TransferirDados_BKP()
'Dim db As DAO.Database: Set db = CurrentDb
'Dim rstProcessamentosPendentes As DAO.Recordset: Set rstProcessamentosPendentes = db.OpenRecordset(sqyProcessamentosPendentes)
'Dim rstDestino As DAO.Recordset: Set rstDestino = db.OpenRecordset(sqyCompras)
'
'Dim rstOrigem As DAO.Recordset
'
'Do While Not rstProcessamentosPendentes.EOF
'
'    Set rstOrigem = db.OpenRecordset(Replace(qryProcessamento, "strChave", rstProcessamentosPendentes.Fields("pk").value))
'    rstDestino.AddNew
'
'    Do While Not rstOrigem.EOF
'        rstDestino(rstOrigem.Fields("nomeCampo").value).value = rstOrigem.Fields("valor").value
'        rstOrigem.MoveNext
'    Loop
'
'    rstDestino.Update
'    rstProcessamentosPendentes.MoveNext
'Loop
'
'db.Close: Set db = Nothing
'
'End Sub

'' #######################################################################################################################

'Sub TransferirDados_bkp_01()
'Dim dbDestino As New Banco: dbDestino.Start "sa", "41L70N@@", "WIN-VE2KJO1LP3\SQLEXPRESS", "SispartsConexao", drSqlServer
'
'Dim db As DAO.Database: Set db = CurrentDb
'Dim rstProcessamentosPendentes As DAO.Recordset: Set rstProcessamentosPendentes = db.OpenRecordset(sqyProcessamentosPendentes)
''Dim rstDestino As DAO.Recordset: Set rstDestino = db.OpenRecordset(sqyCompras)
'
'dbDestino.SqlSelect sqyCompras
'
'Dim rstOrigem As DAO.Recordset
'
'Do While Not rstProcessamentosPendentes.EOF
'
'    Set rstOrigem = db.OpenRecordset(Replace(qryProcessamento, "strChave", rstProcessamentosPendentes.Fields("pk").value))
'
'    dbDestino.rs.AddNew
'    Do While Not rstOrigem.EOF
'
'        For Each fld In dbDestino.rs.Fields
'            If fld.Name = rstOrigem("nomeCampo").value Then
'                dbDestino.rs(fld.Name) = 0 'rstOrigem.Fields("valor")
'                Debug.Print dbDestino.rs(fld.Name).Name
'                Debug.Print rstOrigem.Fields("valor")
'                rstOrigem.MoveNext
'                Exit For
'            End If
'        Next
'
'        '' dbDestino.rs.Fields(rstOrigem("nomeCampo").value).value = rstOrigem.Fields("valor").value
'        rstOrigem.MoveNext
'    Loop
'
'    dbDestino.rs.Update
'    rstProcessamentosPendentes.MoveNext
'Loop
'
'db.Close: Set db = Nothing
'dbDestino.CloseConnection: Set dbDestino = Nothing
'
'End Sub
