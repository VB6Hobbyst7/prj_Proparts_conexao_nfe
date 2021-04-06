Attribute VB_Name = "modConexaoNfeCte"
Option Compare Database

'Private Const qryFields As String = "SELECT strSplit([Destino],'.',1) AS strCampo FROM tblOrigemDestino WHERE (((tblOrigemDestino.tabela)='strParametro'));"
'Private Const sqyProcessamentosPendentes As String = "SELECT DISTINCT pk from tblProcessamento;"
'Private Const qryProcessamento As String = "Select nomecampo, valor from tblProcessamento where pk = 'strChave' and len(nomecampo)>0;"

Private Const sqyCompras As String = "SELECT * FROM tblCompraNF"


'' #####################################################################
'' ### #PENDENCIAS
'' #####################################################################

'' #PENDENTE - IDVD_CompraNF
'UPDATE tblProcessamento SET tblProcessamento.valor = TRIM(Replace(Replace([tblProcessamento].[valor],'PEDIDO',''),';','')) WHERE (((tblProcessamento.NomeCampo)='IDVD_CompraNF'));
'UPDATE tblProcessamento SET tblProcessamento.valor = TRIM(Replace(parts(LBound(Split(([tblProcessamento].[valor]), ","))), "Pedido:", "")) WHERE (((tblProcessamento.NomeCampo)='IDVD_CompraNF'));

Sub carregarDadosGerais()
Dim strProcessamento As String: strProcessamento = "tblDadosConexaoNFeCTe"
Dim s As New clsConexaoNfeCte
Dim t As Variant

    '' #LIMPAR_BASE_DE_TESTES
    Application.CurrentDb.Execute "DELETE FROM tblDadosConexaoNFeCTe"
    Application.CurrentDb.Execute "DELETE FROM tblCompraNF"
    Application.CurrentDb.Execute "DELETE FROM tblCompraNFItem"

    '' #CARREGAR DADOS
    For Each t In Array(strProcessamento)

        '' #PROCESSAMENTO DE ARQUIVO - ENVIO DE DADOS PARA tblProcessamento
        s.ProcessarArquivosXml CStr(t), GetFilesInSubFolders(pegarValorDoParametro(qryParametros, strCaminhoDeColeta))

        '' #TRANSFERIR DADOS PROCESSADOS - DADOS GERAIS - ENVIO DE DADOS PARA tblDadosConexaoNFeCTe
        TransferirDadosProcessados strProcessamento

        '' #TRATAMENTO
        s.TratamentoDeDadosGerais
                
        '' #ARQUIVOS - GERAR ARQUIVOS
        s.CriarTipoDeArquivo opFlagLancadaERP
        s.CriarTipoDeArquivo opManifesto

    Next


    MsgBox "Fim!", vbOKOnly + vbExclamation, "carregarDadosGerais"

End Sub


Sub carregarCompras()
Dim strProcessamento As String: strProcessamento = "tblCompraNF"
Dim s As New clsConexaoNfeCte
Dim t As Variant

    '' #CARREGAR DADOS
    For Each t In Array(strProcessamento)
    
        '' PROCESSAR APENAS ARQUIVOS VALIDOS
        s.ProcessarArquivosXml CStr(t), carregarParametros(qrySelectProcessamentoPendente)
        
        '' #TRANSFERIR DADOS PROCESSADOS - COMPRAS
        TransferirDadosProcessados strProcessamento
        
        
        '' #TRATAMENTO
        s.TratamentoDeCompras
    Next

    '' #FORMATAR
    formatarCampos

    '' #VALIDAR_DADOS
    criarConsultasParaTestes
    
    MsgBox "Fim!", vbOKOnly + vbExclamation, "carregarCompras"

End Sub



'' #ENVIAR_PARA_SERVIDOR
Sub EnviarDadosParaServidor()

'' BANCO LOCAL
Dim db As dao.Database: Set db = CurrentDb
Dim rstOrigem As dao.Recordset: Set rstOrigem = db.OpenRecordset("Select top 1 * from tblCompraNF")

'' BANCO DESTINO
Dim dbDestino As New Banco: dbDestino.Start "sa", "41L70N@@", "WIN-VE2KJO1LP3\SQLEXPRESS", "SispartsConexao", drSqlServer
dbDestino.SqlSelect sqyCompras


'' LISTAGEM DE CAMPOS DA TABELA ORIGEM/DESTINO
Dim rstCampos As dao.Recordset: Set rstCampos = db.OpenRecordset("Select * from tblOrigemDestino where tabela = 'tblCompraNF' and tagOrigem = 1")


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

