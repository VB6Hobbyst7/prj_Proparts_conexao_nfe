Attribute VB_Name = "modConexaoNfeCte_02_Testes"
Option Compare Database

Private Const qryFields As String = "SELECT strSplit([Destino],'.',1) AS strCampo FROM tblOrigemDestino WHERE (((tblOrigemDestino.tabela)='strParametro'));"
Private Const sqyProcessamentosPendentes As String = "SELECT DISTINCT pk from tblProcessamento;"
Private Const qryProcessamento As String = "Select nomecampo, valor from tblProcessamento where pk = 'strChave' and len(nomecampo)>0;"

Private Const sqyCompras As String = "SELECT * FROM tblCompraNF"

'' COMPRAS
'' -- AJUSTE DE CAMPOS
Private Const qryUpdateCurrecy As String = "UPDATE tblProcessamento SET tblProcessamento.valor = Format(Replace([tblProcessamento].[valor], '.', ','), '#,##0.00') WHERE (((tblProcessamento.NomeCampo)='strCampo'));"
Private Const qryUpdateDate As String = "UPDATE tblProcessamento SET tblProcessamento.valor = CDate(Replace(Mid([tblProcessamento].[valor],1,10),'-','/')) WHERE (((tblProcessamento.NomeCampo)='strCampo'));"
Private Const qryUpdateTime As String = "UPDATE tblProcessamento SET tblProcessamento.valor = Replace(Mid([tblProcessamento].[valor],12,8),'-','/') WHERE (((tblProcessamento.NomeCampo)='strCampo'));"


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
    Application.CurrentDb.Execute "DELETE FROM tblDadosConexaoNFeCTe"
    Application.CurrentDb.Execute "DELETE FROM tblCompraNF"
    Application.CurrentDb.Execute "DELETE FROM tblCompraNFItem"

    '' ###########################
    '' #CAPTURA_DADOS_GERAIS
    '' ###########################
    For Each t In Array("tblDadosConexaoNFeCTe")

        '' PROCESSAMENTO DE ARQUIVO
        s.ProcessarArquivosXml CStr(t), GetFilesInSubFolders(pegarValorDoParametro(qryParametros, strCaminhoDeColeta))

        '' TRANSFERIR DADOS
        s.TransferirDadosConexaoNFeCTe

        '' GERAR ARQUIVOS
        s.CriarTipoDeArquivo opFlagLancadaERP
        s.CriarTipoDeArquivo opManifesto

    Next

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

'' #VALIDAR_DADOS
Sub criarConsultasParaTestes()
Dim db As DAO.Database: Set db = CurrentDb
Dim rstOrigem As DAO.Recordset
Dim strSql As String
Dim qrySelectTabelas As String: qrySelectTabelas = "Select Distinct tabela from tblOrigemDestino order by tabela"
Dim tabela As Variant

'' CRIAR CONSULTA PARA VALIDAR DADOS PROCESSADOS
For Each tabela In carregarParametros(qrySelectTabelas)
    strSql = "Select "
    Set rstOrigem = db.OpenRecordset("Select distinct Destino from tblOrigemDestino where tabela = '" & tabela & "'")
    Do While Not rstOrigem.EOF
        strSql = strSql & strSplit(rstOrigem.Fields("Destino").Value, ".", 1) & ","
        rstOrigem.MoveNext
    Loop
    strSql = left(strSql, Len(strSql) - 1) & " from " & tabela
    qryExists "qry_" & tabela
    qryCreate "qry_" & tabela, strSql
Next tabela

db.Close: Set db = Nothing

End Sub

'' #FORMATAR_CAMPOS
Sub formatarCampos()
Dim t As Variant
Dim s As String

'' MOEDA
For Each t In Array("BaseCalcICMSSubsTrib_CompraNF", "BaseCalcICMS_CompraNF", "VTotICMS_CompraNF", "VTotServ_CompraNF", "VTotProd_CompraNF", "VTotNF_CompraNF", "VTotICMSSubsTrib_CompraNF", "VTotFrete_CompraNF", "VTotSeguro_CompraNF", "VTotOutDesp_CompraNF", "VTotIPI_CompraNF", "VTotISS_CompraNF", "TxDesc_CompraNF", "VTotDesc_CompraNF")
    Application.CurrentDb.Execute Replace(qryUpdateCurrecy, "strCampo", t)
Next t

'' DATAS
For Each t In Array("DTEmi_CompraNF", "DTEntd_CompraNF")
    Application.CurrentDb.Execute Replace(qryUpdateDate, "strCampo", t)
Next t

'' HORAS
For Each t In Array("HoraEntd_CompraNF")
    Application.CurrentDb.Execute Replace(qryUpdateTime, "strCampo", t)
Next t

End Sub

'' #TRANSFERIR
Sub TransferirDados()
Dim db As DAO.Database: Set db = CurrentDb
Dim rstProcessamentosPendentes As DAO.Recordset: Set rstProcessamentosPendentes = db.OpenRecordset(sqyProcessamentosPendentes)
Dim rstDestino As DAO.Recordset: Set rstDestino = db.OpenRecordset(sqyCompras)

Dim rstOrigem As DAO.Recordset

Do While Not rstProcessamentosPendentes.EOF
    
    Set rstOrigem = db.OpenRecordset(Replace(qryProcessamento, "strChave", rstProcessamentosPendentes.Fields("pk").Value))
    rstDestino.AddNew
    
    Do While Not rstOrigem.EOF
        rstDestino(rstOrigem.Fields("nomeCampo").Value).Value = rstOrigem.Fields("valor").Value
        rstOrigem.MoveNext
    Loop
       
    rstDestino.Update
    rstProcessamentosPendentes.MoveNext
Loop

db.Close: Set db = Nothing

End Sub
