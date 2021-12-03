<<<<<<< HEAD
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProcessamentoDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'' #####################################
'' 01 CARREGAR DE DADOS
'' #####################################

Option Compare Database

Public Enum enumTipoProcessamento
    opDadosGerais = 0
    opCompras = 1
End Enum

'' #tblProcessamento
Private Const qrySelecaoDeCampos As String = _
    "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((tblOrigemDestino.tabela)='strParametro') AND ((Len([Tag]))>0) AND ((tblOrigemDestino.TagOrigem)=1)) ORDER BY tblOrigemDestino.Tag, tblOrigemDestino.tabela;"

'' #tblProcessamento.Delete
Private Const qryDeleteProcessamento As String = _
    "DELETE * FROM tblProcessamento;"

'' #tblProcessamento.chave
Private Const qryUpdateProcessamento_Chave As String = _
    "UPDATE tblProcessamento SET tblProcessamento.chave = Replace([tblProcessamento].[chave],';','|');"

'' #tblProcessamento.NomeTabela
Private Const qryUpdateProcessamento_NomeTabela As String = _
    "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = ""tblRepositorio"" where tblProcessamento.NomeTabela is null;"

'' #tblProcessamento.NomeTabela - Limpar Itens Marcados Errados
Private Const qryUpdateProcessamento_LimparItensMarcadosErrados As String = _
    "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = Null WHERE (((classificacao([tblProcessamento].[pk]))=1));"

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
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "######## - DeleteProcessamento()"
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryDeleteProcessamento
    Application.CurrentDb.Execute qryDeleteProcessamento
                                                    
End Sub
        
Public Sub UpdateProcessamentoLimparItensMarcadosErrados()
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "######## - UpdateProcessamentoLimparItensMarcadosErrados()"
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryUpdateProcessamento_LimparItensMarcadosErrados
    Application.CurrentDb.Execute qryUpdateProcessamento_LimparItensMarcadosErrados
                                                    
End Sub
        
        
Public Sub UpdateProcessamentoFormatarDados()
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "######## - UpdateProcessamentoFormatarDados()"
    Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                        qryUpdateProcessamento_opData, _
                                                        qryUpdateProcessamento_opTime): executarComandos qryProcessos
End Sub

Public Sub UpdateProcessamentoIdentificarCampos(pRepositorio As String)
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "######## - UpdateProcessamentoIdentificarCampos(pRepositorio As String)"
    Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                    Replace(qryUpdateProcessamento_NomeTabela, "tblRepositorio", pRepositorio), _
                                                    Replace(qryUpdateProcessamento_NomeCampo, "tblRepositorio", pRepositorio), _
                                                    Replace(qryUpdateProcessamento_Formatacao, "tblRepositorio", pRepositorio), _
                                                    qryUpdateProcessamento_RelacaoCamposDeTabelas_Item_CompraNFItem, _
                                                    qryUpdateProcessamento_RelacaoCamposDeTabelas_ChvAcesso_CompraNF, _
                                                    qryUpdateProcessamento_RelacaoCamposDeTabelas_tblCompraNFItem_ChvAcesso_CompraNF): executarComandos qryProcessos
                                                
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

'Dim cont As Integer: cont = XDoc.getElementsByTagName("infNFe/det").Length
Dim item As Variant

Dim pDados As New Collection
Dim idItem As String: idItem = ""
Dim regiao As String: regiao = ""
Dim subRegiao As String: subRegiao = ""
Dim campo As String: campo = ""
Dim valor As String: valor = ""

'' DADOS GERAIS
Dim repositorioDadosGerais As String: repositorioDadosGerais = "tblDadosConexaoNFeCTe"

'' COMPRAS
Dim repositorioCompras As String: repositorioCompras = "tblCompraNF"
Dim repositorioComprasItens As String: repositorioComprasItens = "tblCompraNFItem"
    
    Select Case pTipoProcessamento
        
        '' DADOS GERAIS
        Case opDadosGerais
            
            '' IDENTIFICAÇÃO DO ARQUIVO
            pPK = getFileName(pPathFile)
            pDados.add pPK & "|" & "Chave" & "|" & pPK
            pDados.add pPK & "|" & "CaminhoDoArquivo" & "|" & pPathFile
            
            For Each item In carregarParametros(qrySelecaoDeCampos, repositorioDadosGerais)
                Select Case UBound(Split((item), "|"))
                    
                    Case 0
'                        Debug.Print CStr(item)
                        
                        valor = XDoc.getElementsByTagName(Split((item), "|")(0)).item(0).Text
                        If valor <> "" Then pDados.add pPK & "|" & Replace(item, "|", ";") & "|" & valor
                        
                    Case 1
'                        Debug.Print CStr(item)
                        
                        regiao = Split((item), "|")(0)
                        campo = Split((item), "|")(1)
                        valor = XDoc.SelectNodes(regiao).item(0).SelectNodes(campo).item(0).Text
                        If valor <> "" Then pDados.add pPK & "|" & Replace(item, "|", ";") & "|" & valor
                        
                    Case Else
                End Select
                
                regiao = ""
                campo = ""
                valor = ""
                DoEvents
            Next item
        
        '' COMPRAS
        Case opCompras
        
            pPK = DLookup("[Chave]", "[tblDadosConexaoNFeCTe]", "[CaminhoDoArquivo]='" & pPathFile & "'")
            pChvAcesso = DLookup("[ChvAcesso]", "[tblDadosConexaoNFeCTe]", "[CaminhoDoArquivo]='" & pPathFile & "'")

            '' IDENTIFICAÇÃO DO ARQUIVO
            pDados.add pPK & "|" & "CaminhoDoArquivo" & "|" & pPathFile
        
            '' CHAVE DE ACESSO
            pDados.add pPK & "|" & "ChvAcesso_CompraNF" & "|" & pChvAcesso
            
            '' CABEÇALHO DA COMPRA
            For Each item In carregarParametros(qrySelecaoDeCampos, repositorioCompras)
                Select Case UBound(Split((item), "|"))
                    Case 1
                        regiao = Split((item), "|")(0)
                        campo = Split((item), "|")(1)
                        valor = XDoc.SelectNodes(regiao).item(0).SelectNodes(campo).item(0).Text
                        If valor <> "" Then pDados.add pPK & "|" & Replace(item, "|", ";") & "|" & valor
                    Case Else
                End Select
                
                regiao = ""
                campo = ""
                valor = ""
                DoEvents
            Next item


            '' ITENS DA COMPRA
            For i = 0 To XDoc.getElementsByTagName("infNFe/det").Length - 1
                '' ID
                idItem = CStr(XDoc.getElementsByTagName("nfeProc/NFe/infNFe/det").item(i).Attributes(0).value)
                pDados.add pPK & "_" & idItem & "|" & "Item_CompraNFItem" & "|" & idItem
        
                pDados.add pPK & "_" & idItem & "|" & "tblCompraNFItem.ChvAcesso_CompraNF.opTexto" & "|" & pChvAcesso
        
                For Each item In carregarParametros(qrySelecaoDeCampos, repositorioComprasItens)
        
                    Select Case UBound(Split((item), "|"))
        
                        '' ITEM DE COMPRA
                        Case 1
                            regiao = Split((item), "|")(0)
                            campo = Split((item), "|")(1)
                            valor = XDoc.SelectNodes(regiao).item(i).SelectNodes(campo).item(0).Text
                            If valor <> "" Then pDados.add pPK & "_" & idItem & "|" & Replace(item, "|", ";") & "|" & valor
        
                        '' IMPOSTO
                        Case 2
                            regiao = Split((item), "|")(0)
                            subRegiao = Split((item), "|")(1)
                            campo = Split((item), "|")(2)
                            valor = XDoc.SelectNodes(regiao).item(i).SelectNodes(subRegiao).item(0).getElementsByTagName(campo).item(0).Text
                            If valor <> "" Then pDados.add pPK & "_" & idItem & "|" & Replace(item, "|", ";") & "|" & valor
        
                        Case Else
                    End Select
        
                    regiao = ""
                    subRegiao = ""
                    campo = ""
                    valor = ""
                    DoEvents
                Next item
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

Dim DadosGerais As New clsConexaoNfeCte

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
'        Debug.Print rstOrigem.Fields("NomeCampo").value
        rstDestino.Fields(rstOrigem.Fields("NomeCampo").value).value = rstOrigem.Fields("Valor").value
        
        '' #CAMPO_CHAVE_RELACIONAMENO
        If (pDestino = "tblDadosConexaoNFeCTe") Then rstDestino.Fields("Chave").value = rstOrigem.Fields("pk").value
    
        rstOrigem.MoveNext
        DoEvents
    Loop
        
    rstDestino.Update
    t = 0
    
    '' #COMPRAS
    If (pDestino = "tblCompraNF") Then Application.CurrentDb.Execute Replace(DadosGerais.UpdateProcessamentoConcluido, "strChave", rstPendentes.Fields("pk").value)

    '' #DADOS_GERAIS
    '' qryUpdateRegistroValido - Valor padrao ( 0 - INVALIDO )
    If (pDestino = "tblDadosConexaoNFeCTe") Then Application.CurrentDb.Execute "Update tblDadosConexaoNFeCTe SET registroValido = 0 where registroValido is null"
    
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

=======
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProcessamentoDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'' #####################################
'' 01 CARREGAR DE DADOS
'' #####################################

Option Compare Database

Public Enum enumTipoProcessamento
    opDadosGerais = 0
    opCompras = 1
End Enum

'' #tblProcessamento
Private Const qrySelecaoDeCampos As String = _
    "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((tblOrigemDestino.tabela)='strParametro') AND ((Len([Tag]))>0) AND ((tblOrigemDestino.TagOrigem)=1)) ORDER BY tblOrigemDestino.Tag, tblOrigemDestino.tabela;"

'' #tblProcessamento.Delete
Private Const qryDeleteProcessamento As String = _
    "DELETE * FROM tblProcessamento;"

'' #tblProcessamento.chave
Private Const qryUpdateProcessamento_Chave As String = _
    "UPDATE tblProcessamento SET tblProcessamento.chave = Replace([tblProcessamento].[chave],';','|');"

'' #tblProcessamento.NomeTabela
Private Const qryUpdateProcessamento_NomeTabela As String = _
    "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = ""tblRepositorio"" where tblProcessamento.NomeTabela is null;"

'' #tblProcessamento.NomeTabela - Limpar Itens Marcados Errados
Private Const qryUpdateProcessamento_LimparItensMarcadosErrados As String = _
    "UPDATE tblProcessamento SET tblProcessamento.NomeTabela = Null WHERE (((classificacao([tblProcessamento].[pk]))=1));"

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
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "######## - DeleteProcessamento()"
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryDeleteProcessamento
    Application.CurrentDb.Execute qryDeleteProcessamento
                                                    
End Sub
        
Public Sub UpdateProcessamentoLimparItensMarcadosErrados()
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "######## - UpdateProcessamentoLimparItensMarcadosErrados()"
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), qryUpdateProcessamento_LimparItensMarcadosErrados
    Application.CurrentDb.Execute qryUpdateProcessamento_LimparItensMarcadosErrados
                                                    
End Sub
        
        
Public Sub UpdateProcessamentoFormatarDados()
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "######## - UpdateProcessamentoFormatarDados()"
    Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                        qryUpdateProcessamento_opData, _
                                                        qryUpdateProcessamento_opTime): executarComandos qryProcessos
End Sub

Public Sub UpdateProcessamentoIdentificarCampos(pRepositorio As String)
    If DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='processamentoLog'") Then TextFile_Append CurrentProject.path & "\" & strLog(), "######## - UpdateProcessamentoIdentificarCampos(pRepositorio As String)"
    Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                    Replace(qryUpdateProcessamento_NomeTabela, "tblRepositorio", pRepositorio), _
                                                    Replace(qryUpdateProcessamento_NomeCampo, "tblRepositorio", pRepositorio), _
                                                    Replace(qryUpdateProcessamento_Formatacao, "tblRepositorio", pRepositorio), _
                                                    qryUpdateProcessamento_RelacaoCamposDeTabelas_Item_CompraNFItem, _
                                                    qryUpdateProcessamento_RelacaoCamposDeTabelas_ChvAcesso_CompraNF, _
                                                    qryUpdateProcessamento_RelacaoCamposDeTabelas_tblCompraNFItem_ChvAcesso_CompraNF): executarComandos qryProcessos
                                                
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

'Dim cont As Integer: cont = XDoc.getElementsByTagName("infNFe/det").Length
Dim item As Variant

Dim pDados As New Collection
Dim idItem As String: idItem = ""
Dim regiao As String: regiao = ""
Dim subRegiao As String: subRegiao = ""
Dim campo As String: campo = ""
Dim valor As String: valor = ""

'' DADOS GERAIS
Dim repositorioDadosGerais As String: repositorioDadosGerais = "tblDadosConexaoNFeCTe"

'' COMPRAS
Dim repositorioCompras As String: repositorioCompras = "tblCompraNF"
Dim repositorioComprasItens As String: repositorioComprasItens = "tblCompraNFItem"
    
    Select Case pTipoProcessamento
        
        '' DADOS GERAIS
        Case opDadosGerais
            
            '' IDENTIFICAÇÃO DO ARQUIVO
            pPK = getFileName(pPathFile)
            pDados.add pPK & "|" & "Chave" & "|" & pPK
            pDados.add pPK & "|" & "CaminhoDoArquivo" & "|" & pPathFile
            
            For Each item In carregarParametros(qrySelecaoDeCampos, repositorioDadosGerais)
                Select Case UBound(Split((item), "|"))
                    
                    Case 0
'                        Debug.Print CStr(item)
                        
                        valor = XDoc.getElementsByTagName(Split((item), "|")(0)).item(0).Text
                        If valor <> "" Then pDados.add pPK & "|" & Replace(item, "|", ";") & "|" & valor
                        
                    Case 1
'                        Debug.Print CStr(item)
                        
                        regiao = Split((item), "|")(0)
                        campo = Split((item), "|")(1)
                        valor = XDoc.SelectNodes(regiao).item(0).SelectNodes(campo).item(0).Text
                        If valor <> "" Then pDados.add pPK & "|" & Replace(item, "|", ";") & "|" & valor
                        
                    Case Else
                End Select
                
                regiao = ""
                campo = ""
                valor = ""
                DoEvents
            Next item
        
        '' COMPRAS
        Case opCompras
        
            pPK = DLookup("[Chave]", "[tblDadosConexaoNFeCTe]", "[CaminhoDoArquivo]='" & pPathFile & "'")
            pChvAcesso = DLookup("[ChvAcesso]", "[tblDadosConexaoNFeCTe]", "[CaminhoDoArquivo]='" & pPathFile & "'")

            '' IDENTIFICAÇÃO DO ARQUIVO
            pDados.add pPK & "|" & "CaminhoDoArquivo" & "|" & pPathFile
        
            '' CHAVE DE ACESSO
            pDados.add pPK & "|" & "ChvAcesso_CompraNF" & "|" & pChvAcesso
            
            '' CABEÇALHO DA COMPRA
            For Each item In carregarParametros(qrySelecaoDeCampos, repositorioCompras)
                Select Case UBound(Split((item), "|"))
                    Case 1
                        regiao = Split((item), "|")(0)
                        campo = Split((item), "|")(1)
                        valor = XDoc.SelectNodes(regiao).item(0).SelectNodes(campo).item(0).Text
                        If valor <> "" Then pDados.add pPK & "|" & Replace(item, "|", ";") & "|" & valor
                    Case Else
                End Select
                
                regiao = ""
                campo = ""
                valor = ""
                DoEvents
            Next item


            '' ITENS DA COMPRA
            For i = 0 To XDoc.getElementsByTagName("infNFe/det").Length - 1
                '' ID
                idItem = CStr(XDoc.getElementsByTagName("nfeProc/NFe/infNFe/det").item(i).Attributes(0).value)
                pDados.add pPK & "_" & idItem & "|" & "Item_CompraNFItem" & "|" & idItem
        
                pDados.add pPK & "_" & idItem & "|" & "tblCompraNFItem.ChvAcesso_CompraNF.opTexto" & "|" & pChvAcesso
        
                For Each item In carregarParametros(qrySelecaoDeCampos, repositorioComprasItens)
        
                    Select Case UBound(Split((item), "|"))
        
                        '' ITEM DE COMPRA
                        Case 1
                            regiao = Split((item), "|")(0)
                            campo = Split((item), "|")(1)
                            valor = XDoc.SelectNodes(regiao).item(i).SelectNodes(campo).item(0).Text
                            If valor <> "" Then pDados.add pPK & "_" & idItem & "|" & Replace(item, "|", ";") & "|" & valor
        
                        '' IMPOSTO
                        Case 2
                            regiao = Split((item), "|")(0)
                            subRegiao = Split((item), "|")(1)
                            campo = Split((item), "|")(2)
                            valor = XDoc.SelectNodes(regiao).item(i).SelectNodes(subRegiao).item(0).getElementsByTagName(campo).item(0).Text
                            If valor <> "" Then pDados.add pPK & "_" & idItem & "|" & Replace(item, "|", ";") & "|" & valor
        
                        Case Else
                    End Select
        
                    regiao = ""
                    subRegiao = ""
                    campo = ""
                    valor = ""
                    DoEvents
                Next item
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

Dim DadosGerais As New clsConexaoNfeCte

'' CONTROLE DE CADASTRO DE REGISTRO
Dim T As Integer
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
        If T = 0 Then rstDestino.AddNew: T = 1
'        Debug.Print rstOrigem.Fields("NomeCampo").value
        rstDestino.Fields(rstOrigem.Fields("NomeCampo").value).value = rstOrigem.Fields("Valor").value
        
        '' #CAMPO_CHAVE_RELACIONAMENO
        If (pDestino = "tblDadosConexaoNFeCTe") Then rstDestino.Fields("Chave").value = rstOrigem.Fields("pk").value
    
        rstOrigem.MoveNext
        DoEvents
    Loop
        
    rstDestino.Update
    T = 0
    
    '' #COMPRAS
    If (pDestino = "tblCompraNF") Then Application.CurrentDb.Execute Replace(DadosGerais.UpdateProcessamentoConcluido, "strChave", rstPendentes.Fields("pk").value)

    '' #DADOS_GERAIS
    '' qryUpdateRegistroValido - Valor padrao ( 0 - INVALIDO )
    If (pDestino = "tblDadosConexaoNFeCTe") Then Application.CurrentDb.Execute "Update tblDadosConexaoNFeCTe SET registroValido = 0 where registroValido is null"
    
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

>>>>>>> f4084cb29d769387d25e7b837853d2119e0da429
