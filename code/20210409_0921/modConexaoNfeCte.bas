Attribute VB_Name = "modConexaoNfeCte"
Option Compare Database

'Private Const qryFields As String = "SELECT strSplit([Destino],'.',1) AS strCampo FROM tblOrigemDestino WHERE (((tblOrigemDestino.tabela)='strParametro'));"
'Private Const sqyProcessamentosPendentes As String = "SELECT DISTINCT pk from tblProcessamento;"
'Private Const qryProcessamento As String = "Select nomecampo, valor from tblProcessamento where pk = 'strChave' and len(nomecampo)>0;"

Private Const sqyCompras As String = "SELECT * FROM tblCompraNF"


'' #PENDENTE - IDVD_CompraNF
'UPDATE tblProcessamento SET tblProcessamento.valor = TRIM(Replace(Replace([tblProcessamento].[valor],'PEDIDO',''),';','')) WHERE (((tblProcessamento.NomeCampo)='IDVD_CompraNF'));
'UPDATE tblProcessamento SET tblProcessamento.valor = TRIM(Replace(parts(LBound(Split(([tblProcessamento].[valor]), ","))), "Pedido:", "")) WHERE (((tblProcessamento.NomeCampo)='IDVD_CompraNF'));



'' #####################################################################
'' ### #TESTES
'' #####################################################################

'' #ADMINISTRACAO - RESPONSAVEL POR TRAZER OS DADOS DO SERVIDOR PARA AUXILIO NO PROCESSAMENTO. QUANDO NECESSARIO
Sub teste_ImportarDados()

    ImportarDados "tblNatOp", "tmpNatOp"
    ImportarDados "tblEmpresa", "tmpEmpresa"
    ImportarDados "Clientes", "tmpClientes"

End Sub

Sub teste_IDVD()
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


'' #FORMATAR_CAMPOS
Public Sub formatarCampos()
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



'' 01.CARREGAR DADOS GERAIS
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


'' 02.CARREGAR COMPRAS ANTES DE VENVIAR PARA O SERVIDOR
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


'' 03.ENVIAR DADOS PARA SERVIDOR
Sub EnviarDadosParaServidor()

'' BANCO LOCAL
Dim db As dao.Database: Set db = CurrentDb
Dim rstOrigem As dao.Recordset: Set rstOrigem = db.OpenRecordset("Select top 1 * from tblCompraNF")

'' BANCO DESTINO
Dim strUsuarioNome As String: strUsuario = pegarValorDoParametro(qryParametros, "BancoDados_Usuario")
Dim strUsuarioSenha As String: strUsuarioSenha = pegarValorDoParametro(qryParametros, "BancoDados_Senha")
Dim strOrigem As String: strOrigem = pegarValorDoParametro(qryParametros, "BancoDados_Origem")
Dim strBanco As String: strBanco = pegarValorDoParametro(qryParametros, "BancoDados_Banco")

Dim dbDestino As New Banco: dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
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

