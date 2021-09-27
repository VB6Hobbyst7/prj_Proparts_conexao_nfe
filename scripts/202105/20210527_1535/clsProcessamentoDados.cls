VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProcessamentoDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Enum enumTipoProcessamento
    opDadosGerais = 0
    opCompras = 1
End Enum


'' #tblProcessamento
Private Const qrySelecaoDeCampos As String = _
    "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((tblOrigemDestino.tabela)='strParametro') AND ((Len([Tag]))>0) AND ((tblOrigemDestino.TagOrigem)=1)) ORDER BY tblOrigemDestino.Tag, tblOrigemDestino.tabela;"
       
'' #tblProcessamento.Pendentes
Private Const qrySelecaoDeArquivosPendentes As String = _
    "SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1)) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0) ORDER BY tblDadosConexaoNFeCTe.CaminhoDoArquivo;"


'' #tblProcessamento.Delete
Private Const qryDeleteProcessamento As String = _
    "DELETE * FROM tblProcessamento;"

'' #tblProcessamento.chave
Private Const qryUpdateProcessamento_Chave As String = _
    "UPDATE tblProcessamento SET tblProcessamento.chave = Replace([tblProcessamento].[chave],';','|');"

'' #tblProcessamento.NomeTabela
Private Const qryUpdateProcessamento_NomeTabela As String = _
    "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = ""tblRepositorio"";"

'' #tblProcessamento.NomeCampo
Private Const qryUpdateProcessamento_NomeCampo As String = _
    "UPDATE tblProcessamento SET tblProcessamento.NomeCampo = DLookUp(""campo"",""tblOrigemDestino"",""tag='"" & [tblProcessamento].[chave] & ""' and Tabela= 'tblRepositorio'"") WHERE (((tblProcessamento.NomeTabela)=""tblRepositorio""));"

'' #tblProcessamento.formatacao
Private Const qryUpdateProcessamento_Formatacao As String = _
    "UPDATE tblProcessamento SET tblProcessamento.formatacao = DLookUp(""formatacao"",""tblOrigemDestino"",""tag='"" & [tblProcessamento].[chave] & ""' and Tabela= 'tblRepositorio'"") WHERE (((tblProcessamento.NomeTabela)=""tblRepositorio""));"

'' #tblProcessamento.formatacao_Item_CompraNFItem
Private Const qryUpdateProcessamento_RelacaoCamposDeTabelas_Item_CompraNFItem As String = _
        "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = ""tblCompraNFItem"", tblProcessamento.NomeCampo = [tblProcessamento].[chave], tblProcessamento.formatacao = DLookUp(""formatacao"",""tblOrigemDestino"",""campo='Item_CompraNFItem'"") WHERE (((tblProcessamento.chave)=""Item_CompraNFItem""));"

'' #tblProcessamento.opTexto_ChvAcesso_CompraNF
Private Const qryUpdateProcessamento_RelacaoCamposDeTabelas_ChvAcesso_CompraNF As String = _
        "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = ""tblCompraNF"", tblProcessamento.NomeCampo = [tblProcessamento].[chave], tblProcessamento.formatacao = ""opTexto"" WHERE (((tblProcessamento.chave)=""ChvAcesso_CompraNF""));"

'' #tblProcessamento
Private Const qryUpdateProcessamento_RelacaoCamposDeTabelas_tblCompraNFItem_ChvAcesso_CompraNF As String = _
        "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = strSplit([tblProcessamento].[chave],'.',0), tblProcessamento.NomeCampo = strSplit([tblProcessamento].[chave],'.',1), tblProcessamento.formatacao = strSplit([tblProcessamento].[chave],'.',2) WHERE (((tblProcessamento.chave)=""tblCompraNFItem.ChvAcesso_CompraNF.opTexto""));"

'' #tblProcessamento.opData
Private Const qryUpdateProcessamento_opData As String = _
        "UPDATE tblProcessamento SET tblProcessamento.valor = Mid([tblProcessamento].[valor],1,10) WHERE formatacao = 'opData';"

'' #tblProcessamento.opTime
Private Const qryUpdateProcessamento_opTime As String = _
        "UPDATE tblProcessamento SET tblProcessamento.valor = Mid([tblProcessamento].[valor],12,8) WHERE formatacao = 'opTime';"
        
Public Sub DeleteProcessamento()
    Application.CurrentDb.Execute qryDeleteProcessamento
                                                    
End Sub
        
Public Sub UpdateProcessamentoFormatarDados()
Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                    qryUpdateProcessamento_opData, _
                                                    qryUpdateProcessamento_opTime): executarComandos qryProcessos
End Sub

Public Sub UpdateProcessamentoIdentificarCampos(pRepositorio As String)
Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                Replace(qryUpdateProcessamento_NomeTabela, "tblRepositorio", pRepositorio), _
                                                Replace(qryUpdateProcessamento_NomeCampo, "tblRepositorio", pRepositorio), _
                                                Replace(qryUpdateProcessamento_Formatacao, "tblRepositorio", pRepositorio)): executarComandos qryProcessos
End Sub


'' #01. PROCESSAMENTO DE ARQUIVOS POR TIPO
Public Function ProcessamentoDeArquivo(pPathFile As String, pTipoProcessamento As enumTipoProcessamento)
On Error Resume Next

Dim s As New clsConexaoNfeCte

'' CHAVES DE CONTROLE
Dim pPK As String: pPK = ""
Dim pChvAcesso As String: pChvAcesso = ""

'' CARREGAR ARQUIVO
Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
XDoc.Load pPathFile

Dim cont As Integer: cont = XDoc.getElementsByTagName("infNFe/det").Length
Dim Item As Variant

Dim pDados As New Collection
Dim idItem As String: idItem = ""
Dim regiao As String: regiao = ""
Dim subRegiao As String: subRegiao = ""
Dim campo As String: campo = ""
Dim valor As String: valor = ""

    Select Case pTipoProcessamento
        
        '' DADOS GERAIS
        Case opDadosGerais
            
            '' IDENTIFICAÇÃO DO ARQUIVO
            pPK = getFileName(pPathFile)
            pDados.add pPK & "|" & "Chave" & "|" & pPK
            pDados.add pPK & "|" & "CaminhoDoArquivo" & "|" & pPathFile
            
            For Each Item In carregarParametros(qrySelecaoDeCampos, "tblDadosConexaoNFeCTe")
                Select Case UBound(Split((Item), "|"))
                    Case 1
                        regiao = Split((Item), "|")(0)
                        campo = Split((Item), "|")(1)
                        valor = XDoc.SelectNodes(regiao).Item(0).SelectNodes(campo).Item(0).text
                        If valor <> "" Then pDados.add pPK & "|" & Replace(Item, "|", ";") & "|" & valor
                    Case Else
                End Select
                
                regiao = ""
                campo = ""
                valor = ""
                DoEvents
            Next Item
        
        '' COMPRAS
        Case opCompras
        
            pPK = DLookup("[Chave]", "[tblDadosConexaoNFeCTe]", "[CaminhoDoArquivo]='" & pPathFile & "'")
            pChvAcesso = DLookup("[ChvAcesso]", "[tblDadosConexaoNFeCTe]", "[CaminhoDoArquivo]='" & pPathFile & "'")

            '' IDENTIFICAÇÃO DO ARQUIVO
            pDados.add pPK & "|" & "CaminhoDoArquivo" & "|" & pPathFile
        
            '' CHAVE DE ACESSO
            pDados.add pPK & "|" & "ChvAcesso_CompraNF" & "|" & pChvAcesso
            
            '' CABEÇALHO DA COMPRA
            For Each Item In carregarParametros(qrySelecaoDeCampos, "tblCompraNF")
                Select Case UBound(Split((Item), "|"))
                    Case 1
                        regiao = Split((Item), "|")(0)
                        campo = Split((Item), "|")(1)
                        valor = XDoc.SelectNodes(regiao).Item(0).SelectNodes(campo).Item(0).text
                        If valor <> "" Then pDados.add pPK & "|" & Replace(Item, "|", ";") & "|" & valor
                    Case Else
                End Select
                
                regiao = ""
                campo = ""
                valor = ""
                DoEvents
            Next Item


            '' ITENS DA COMPRA
            For i = 0 To cont - 1
                '' ID
                idItem = CStr(XDoc.getElementsByTagName("nfeProc/NFe/infNFe/det").Item(i).Attributes(0).value)
                pDados.add pPK & "_" & idItem & "|" & "Item_CompraNFItem" & "|" & idItem
        
                pDados.add pPK & "_" & idItem & "|" & "tblCompraNFItem.ChvAcesso_CompraNF.opTexto" & "|" & pChvAcesso
        
                For Each Item In carregarParametros(qrySelecaoDeCampos, "tblCompraNFItem")
        
                    Select Case UBound(Split((Item), "|"))
        
                        '' ITEM DE COMPRA
                        Case 1
                            regiao = Split((Item), "|")(0)
                            campo = Split((Item), "|")(1)
                            valor = XDoc.SelectNodes(regiao).Item(i).SelectNodes(campo).Item(0).text
                            If valor <> "" Then pDados.add pPK & "_" & idItem & "|" & Replace(Item, "|", ";") & "|" & valor
        
                        '' IMPOSTO
                        Case 2
                            regiao = Split((Item), "|")(0)
                            subRegiao = Split((Item), "|")(1)
                            campo = Split((Item), "|")(2)
                            valor = XDoc.SelectNodes(regiao).Item(i).SelectNodes(subRegiao).Item(0).getElementsByTagName(campo).Item(0).text
                            If valor <> "" Then pDados.add pPK & "_" & idItem & "|" & Replace(Item, "|", ";") & "|" & valor
        
                        Case Else
                    End Select
        
                    regiao = ""
                    subRegiao = ""
                    campo = ""
                    valor = ""
                    DoEvents
                Next Item
                DoEvents
            Next i

        Case Else
    End Select

    '' CADASTRAR DADOS
    cadastroProcessamento pDados
    
    '' LIMPAR COLEÇÃO
    ClearCollection pDados
    
    '' AJUSTAR CAMPO CHAVE
    Application.CurrentDb.Execute qryUpdateProcessamento_Chave

Set XDoc = Nothing

End Function

'' #02. TRANSFERIR DADOS PROCESSAMENTOS PARA A TABELA DE DESTINO
Public Sub ProcessamentoTransferir(pDestino As String)
On Error Resume Next

'' CONTROLE DE CADASTRO DE REGISTRO
Dim t As Integer
Dim tmpSqlOrigem As String

'' #BANCO_LOCAL
Dim db As DAO.Database: Set db = CurrentDb
Dim tmpSql As String: tmpSql = "Select Distinct pk from tblProcessamento where NomeTabela = '" & pDestino & "' Order by pk;"

If (pDestino = "tblCompraNFItem") Then tmpSql = "SELECT tblProcessamento.pk FROM tblProcessamento WHERE (((tblProcessamento.valor) In (SELECT tblProcessamento.valor as chave FROM tblProcessamento WHERE (((tblProcessamento.NomeCampo)='ChvAcesso_CompraNF') AND ((tblProcessamento.NomeTabela)='tblCompraNF')) ORDER BY tblProcessamento.ID)) AND ((tblProcessamento.NomeTabela)='tblCompraNFItem')) ORDER BY tblProcessamento.ID;"

Dim rstPendentes As DAO.Recordset: Set rstPendentes = db.OpenRecordset(tmpSql)
Dim rstOrigem As DAO.Recordset

'' #BANCO_DESTINO
tmpSql = "Select * from " & pDestino
Dim rstDestino As DAO.Recordset: Set rstDestino = db.OpenRecordset(tmpSql)

'' PROCESSAMENTO PENDENTE
Do While Not rstPendentes.EOF

    '' listar itens de compra
    tmpSqlOrigem = "Select * from tblProcessamento where NomeTabela = '" & pDestino & "' and pk = '" & rstPendentes.Fields("pk").value & "' and len(formatacao)>0  and len(NomeCampo)>0 Order by ID "
    
    '' REGISTROS DO PROCESSAMENTO PENDENTE
    Set rstOrigem = db.OpenRecordset(tmpSqlOrigem)
    Do While Not rstOrigem.EOF
    
        '' CONTROLE DE CADASTRO
        If t = 0 Then rstDestino.AddNew: t = 1
        rstDestino.Fields(rstOrigem.Fields("NomeCampo").value).value = rstOrigem.Fields("Valor").value
        
        '' #CAMPO_CHAVE_RELACIONAMENO
        If (pDestino = "tblDadosConexaoNFeCTe") Then rstDestino.Fields("Chave").value = rstOrigem.Fields("pk").value
    
        rstOrigem.MoveNext
        DoEvents
    Loop
        
    rstDestino.Update
    t = 0
    
'    '' #COMPRAS
'    If (pDestino = "tblCompraNF") Then Application.CurrentDb.Execute Replace(qryUpdateProcessamentoConcluido, "strChave", rstPendentes.Fields("pk").value)
'
'    '' #DADOS_GERAIS
'    '' qryUpdateRegistroValido - Valor padrao
'    If (pDestino = "tblDadosConexaoNFeCTe") Then Application.CurrentDb.Execute "Update tblDadosConexaoNFeCTe SET registroValido = 0 where registroValido is null"
    
    rstPendentes.MoveNext
    DoEvents
Loop

'dbDestino.CloseConnection
db.Close: Set db = Nothing

End Sub

'' #00. CADASTRO DE DADOS PROCESSADOS
Private Function cadastroProcessamento(pRegistros As Collection)
Dim cadastro As New clsProcessamento
Dim i As Variant
    
    For Each i In pRegistros
        With cadastro
            .pk = Split(i, "|")(0)
            .Chave = Split(i, "|")(1)
            .valor = Mid(Split(i, "|")(2), 1, 255)
            .cadastrar
        End With
        
        DoEvents
    Next i

End Function

