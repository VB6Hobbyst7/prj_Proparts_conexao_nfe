Attribute VB_Name = "azs_Info"


'' #20211128_MoverArquivosProcessados


'' Done | Subistitui��o deste item
'' ----------------------------------------------------------------------------------
'' #20220119_dhEmiCopia
'' #20220119_GerarArquivosDeLancamentos_e_Manifestos
'' #20220112_InsertLogProcessados
'' #20220112_qryComprasItem_Update_STs_CTe_ST_CompraNFItem
'' #20220111_update_Almox_CompraNFItem
'' #20220111_update_FlagEst_CompraNFItem
'' #20220110_qryComprasItem_Update_AjustesCampos
'' #20220110_qryComprasItem_Update_STs
'' #20220106_update_IdProd_CompraNFItem
'' #20211202_qryComprasItem_Update_STs
'' #20211202_qryComprasItens_Update_Almox_CompraNFItem | #20211202_update_Almox_CompraNFItem
'' #20211128_MoverArquivosProcessados
'' #20211128_LimparRepositorios
'' #20211127_qryDadosGerais_Update_FiltroFil_DestinatarioNaoProparts
'' #20211127_qryDadosGerais_Update_FiltroFil_DestinatarioProparts
'' #20211127_qryDadosGerais_Update_FiltroFil_DestinatarioProparts_55
'' #20211127_qryDadosGerais_Update_FiltroFil_RemetenteProparts
'' #20211122_AjusteDeCampos_CTe
'' #20210823_qryUpdateNumPed_CompraNF
'' #20210823_qryDadosGerais_Update_Sit_CompraNF
'' #20210823_qryDadosGerais_Update_IdFornCompraNF
'' #20210823_qryDadosGerais_Update_ID_NatOp_CompraNF__FiltroCFOP
'' #20210823_qryDadosGerais_Update_IDVD
'' #20210823_qryCompras_Update_Dados_NFe
'' #20210823_qryCompras_Update_Dados
'' #20210823_qryComprasItens_Update_CFOP_CompraNF
'' #20210823_VTotProd_CompraNF
'' #20210823_ID_Prod_CompraNFItem
'' #20210823_FornecedoresValidos






'' ----------------------------------------------------------------------------------


''' #20211128_MoverArquivosProcessados
'Sub MoverArquivosProcessados()
'On Error GoTo adm_Err
'Dim db As DAO.Database: Set db = CurrentDb
'Dim rst As DAO.Recordset
'
'''#registroProcessado
''' 0|1 - caminhoDeColeta
''' 2 - caminhoDeColeta
''' 3 - caminhoDeColetaProcessados
''' 4 - caminhoDeColetaExpurgo
''' 5 - caminhoDeColetaAcoes
''' 8 - MOVER ARQUIVOS            -- CONTROLE INTERNO
''' 9 - FINAL                     -- CONTROLE INTERNO
'
''Dim sql_Update_tblDadosConexaoNFeCTe_registroProcessado_Reclassificar As String: sql_Update_tblDadosConexaoNFeCTe_registroProcessado_Reclassificar = _
''    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 1;"
''    Application.CurrentDb.Execute Replace(sql_Update_tblDadosConexaoNFeCTe_registroProcessado_Reclassificar, 1, 5)
'
'
''' Processados - registroProcessado(3)
''' 1. Classificar como "registroProcessado(3) - Processados" onde ...
''' 1.1 Onde temos arquivo para processar - "CaminhoDoArquivo" e
''' 1.2 Registro � valido "registroValido(1) - Registro V�lido" e
''' 1.3 Registro processado � "registroProcessado (2) - CadastroDeComprasEmServidor()".
'Dim sql_Update_registroProcessado_Processados As String: sql_Update_registroProcessado_Processados = _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 3 WHERE ((Not (tblDadosConexaoNFeCTe.CaminhoDoArquivo) Is Null) AND ((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=2));"
'    Application.CurrentDb.Execute sql_Update_registroProcessado_Processados
'
''' Expurgo - registroProcessado(4)
''' 1. Classificar como "registroProcessado(4) - Expurgo" onde ...
''' 1.1 Registro � valido "registroValido(1) - Registro V�lido" e n�o foi processado "registroProcessado(0) - N�o foi processado"
''' 2 Registro � invalido "registroValido(0) - Registro Inv�lido"
'Dim sql_Update_registroProcessado_Expurgo() As Variant: sql_Update_registroProcessado_Expurgo = Array( _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 4 WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0));", _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 4 WHERE (((tblDadosConexaoNFeCTe.registroValido)=0) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0));")
'    executarComandos sql_Update_registroProcessado_Expurgo
'
''' Finalizados - registroProcessado(9)
'Dim sql_Update_CopyFile_Final As String: sql_Update_CopyFile_Final = _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 9, tblDadosConexaoNFeCTe.CaminhoDoArquivo = [tblDadosConexaoNFeCTe].[CaminhoDestino], tblDadosConexaoNFeCTe.CaminhoDestino = Null WHERE (((tblDadosConexaoNFeCTe.[registroProcessado])=8))"
'
''' Atualiza��o do caminho de destino
'Dim sql_Update_CaminhoDestino As String: sql_Update_CaminhoDestino = _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.CaminhoDestino = strCaminhoDestino([tblDadosConexaoNFeCTe].[CaminhoDoArquivo]), tblDadosConexaoNFeCTe.registroProcessado = 8 WHERE (((tblDadosConexaoNFeCTe.registroProcessado)<8));"
'    Application.CurrentDb.Execute sql_Update_CaminhoDestino
'
''' Sele��o de arquivos para movimenta��o de pastas
'Dim sql_Select_CaminhoDestino As String: sql_Select_CaminhoDestino = _
'    "SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo, tblDadosConexaoNFeCTe.CaminhoDestino FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroProcessado)=8));"
'
'    Debug.Print "############################"
'
'    '' MOVER ARQUIVOS
'    Set rst = db.OpenRecordset(sql_Select_CaminhoDestino)
'    Do While Not rst.EOF
'        If (Dir(rst.Fields("CaminhoDoArquivo").value) <> "") Then
'            FileCopy rst.Fields("CaminhoDoArquivo").value, rst.Fields("CaminhoDestino").value
'            Kill rst.Fields("CaminhoDoArquivo").value
'        End If
'        rst.MoveNext
'    Loop
'
'    '' ATUALIZA��O - registroProcessado ( FINAL )
'    Application.CurrentDb.Execute sql_Update_CopyFile_Final
'
'db.Close
'
'
'adm_Exit:
'    Set db = Nothing
'    Set rst = Nothing
'
'    Exit Sub
'
'adm_Err:
'    Debug.Print "MoverArquivosProcessados() - " & Err.Description
'    TextFile_Append CurrentProject.path & "\" & strLog(), Err.Description
'    Resume adm_Exit
'
'End Sub
'
'


'Sub FSOMoveFile()
'    Dim FSO As New FileSystemObject
'    Set FSO = CreateObject("Scripting.FileSystemObject")
'
'    FSO.MoveFile "C:\Src\TestFile.txt", "C:\Dst\ModTestFile.txt"
'
'End Sub


'Sub azs_teste_gerarArquivosJson()
'
'Dim strCaminhoAcoes As String: strCaminhoAcoes = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColetaAcoes'")
'
''' #20220119_GerarArquivosDeLancamentos_e_Manifestos
''' LAN�AMENTO
'gerarArquivosJson opFlagLancadaERP, , strCaminhoAcoes
'
''' MANIFESTO
'gerarArquivosJson opManifesto, , strCaminhoAcoes
'
'
''Debug.Print DateDiff("s", "1/1/1970", "2021-03-03 08:20:00") * 1000
'
'
'End Sub
'
'
'Sub azs_teste_carregarDadosGerais()
'Dim pCaminho As String: pCaminho = _
'    "C:\ConexaoNFe\XML\68.365.5010003-77 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\recebimento\43210331275238000153550010000011371772829986-nfeproc.xml"
'
'    carregarDadosGerais CStr(pCaminho)
'
'End Sub


'Sub azs_teste_ProcessamentoDeArquivo_opCompras()
'Dim pCaminho As String
'
'Dim s As New clsProcessamentoDados
'Dim Chave As Variant
'
'
'    pCaminho = _
'        "C:\ConexaoNFe\XML\68.365.5010003-77 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\recebimento\43210331275238000153550010000011371772829986-nfeproc.xml"
'
'    s.DeleteProcessamento
'    s.ProcessamentoDeArquivo pCaminho, opCompras
'
'    '' IDENTIFICAR CAMPOS
'    s.UpdateProcessamentoIdentificarCampos "tblCompraNF"
'
'    '' CORRE��O DE DADOS MARCADOS ERRADOS EM ITENS DE COMPRAS
'    s.UpdateProcessamentoLimparItensMarcadosErrados
'
'    '' IDENTIFICAR CAMPOS DE ITENS DE COMPRAS
'    s.UpdateProcessamentoIdentificarCampos "tblCompraNFItem"
'
'    '' FORMATAR DADOS
'    s.UpdateProcessamentoFormatarDados
'
'
'Set s = Nothing
'
'End Sub



'Sub azs_teste_qryComprasCTe_Update_AjustesCampos()
'
''' BANCO_DESTINO
'Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
'Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
'Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
'Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
'Dim dbDestino As New Banco
'
''' BANCO_LOCAL
'Dim Scripts As New clsConexaoNfeCte
'Dim qryComprasCTe_Update_AjustesCampos As String: qryComprasCTe_Update_AjustesCampos = "UPDATE tblCompraNF SET tblCompraNF.HoraEntd_CompraNF = NULL ,tblCompraNF.IDVD_CompraNF = NULL WHERE (((tblCompraNF.ChvAcesso_CompraNF) IN (pLista_ChvAcesso_CompraNF)));"
'
'
'    '' BANCO_DESTINO
'    dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
'
'    dbDestino.SqlExecute Replace(qryComprasCTe_Update_AjustesCampos, "pLista_ChvAcesso_CompraNF", carregarComprasCTe)
'
'
'End Sub
'
'
'Sub azs_teste_xml_selectSingleNode()
'
'Dim pCaminho As String: pCaminho = _
'    "C:\xmls\68.365.5010003-77 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210348740351012767570000021186731952977908-cteproc.xml"
'
'
'    Dim objXML As Object, node As Object
'
'    Set objXML = CreateObject("MSXML2.DOMDocument")
'    objXML.async = False: objXML.validateOnParse = False
'
'    If Not objXML.Load(pCaminho) Then  'strXML is the string with XML'
'        Err.Raise objXML.parseError.errorCode, , objXML.parseError.reason
'
'    Else
'        Set node = objXML.selectSingleNode("cteProc")
'        Stop
'
'    End If
'End Sub
'
'
'Sub azs_teste_xml_nodeName()
'
'Dim pCaminho As String: pCaminho = _
'    "C:\xmls\68.365.5010003-77 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210348740351012767570000021186731952977908-cteproc.xml"
'
'Dim objXML As MSXML2.DOMDocument60: Set objXML = New MSXML2.DOMDocument60
'
''Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
'objXML.async = False: objXML.validateOnParse = False
'objXML.Load pCaminho
'
''Dim Nodes As IXMLDOMNodeList: Set Nodes = objXML.childNodes
'
'Dim objNode As IXMLDOMNode
'
'   For Each objNode In objXML.childNodes
'      If objNode.NodeType = NODE_TEXT Then
'          If objNode.ParentNode.nodeName = "pICMS" Then
'            Debug.Print CStr(objNode.NodeValue)
'          End If
'      End If
'   Next objNode
'End Sub
'
'
'Private Sub azs_teste_ValidacaoDeCampos()
'Dim db As DAO.Database: Set db = CurrentDb
'Dim rstRegistros As DAO.Recordset
'Dim rstItens As DAO.Recordset
'Dim sqlRegistros As String: sqlRegistros = "Select * from tblCompraNF where ChvAcesso_CompraNF = "
'Dim sqlItens As String: sqlItens = "Select * from tblCompraNFItem where ChvAcesso_CompraNF = "
'Dim arquivos As New Collection
'
'Dim item As Variant
'Dim TMP As String
'''42210300634453001303570010001139451001171544
' arquivos.add "42210300634453001303570010001139451001171544" '' 57
'
''arquivos.add "32210368365501000296550000000639051001364146"
'
''arquivos.add "32210304884082000569570000040073831040073834"
''arquivos.add "42210220147617000494570010009539201999046070"
''arquivos.add "32210368365501000296550000000638811001361356"
''arquivos.add "42210212680452000302550020000886301507884230"
'
''arquivos.Add "32210368365501000296550000000638841001361501"
'
'
'For Each item In arquivos
'
'
'    TMP = sqlRegistros & "'" & CStr(item) & "'"
'    Set rstRegistros = db.OpenRecordset(TMP)
'
'    Do While Not rstRegistros.EOF
'
'        TMP = ""
'        For i = 0 To rstRegistros.Fields.count - 1
'            TMP = rstRegistros.Fields(i).Name & vbTab & rstRegistros.Fields(i).value
'            TextFile_Append CurrentProject.path & "\" & CStr(item) & ".txt", TMP
'        Next i
'
'        TextFile_Append CurrentProject.path & "\" & CStr(item) & ".txt", vbNewLine & "#############################" & vbNewLine
'
'        TMP = ""
'        TMP = sqlItens & "'" & CStr(item) & "'"
'        Debug.Print TMP
'
'        Set rstItens = db.OpenRecordset(TMP)
'        Do While Not rstItens.EOF
'            For i = 0 To rstItens.Fields.count - 1
'                TMP = rstItens.Fields(i).Name & vbTab & rstItens.Fields(i).value
'                TextFile_Append CurrentProject.path & "\" & CStr(item) & ".txt", TMP
'            Next i
'
'            TextFile_Append CurrentProject.path & "\" & CStr(item) & ".txt", vbNewLine & "#############################" & vbNewLine
'
'            rstItens.MoveNext
'            DoEvents
'        Loop
'
'        Debug.Print "Concluido! - " & CStr(item) & ".txt"
'        rstRegistros.MoveNext
'        DoEvents
'        TMP = ""
'    Loop
'
'    rstRegistros.Close
'    rstItens.Close
'Next
'
'Debug.Print "Concluido!"
'Set rstRegistros = Nothing
'Set rstItens = Nothing
'
'End Sub
'
'
'
''Private Sub azs_teste_criarConsultasParaTestes()
''Dim db As DAO.Database: Set db = CurrentDb
''Dim rstOrigem As DAO.Recordset
''Dim strSQL As String
''Dim qrySelectTabelas As String: qrySelectTabelas = "Select Distinct tabela from tblOrigemDestino order by tabela"
''Dim tabela As Variant
''
'''' CRIAR CONSULTA PARA VALIDAR DADOS PROCESSADOS
''For Each tabela In carregarParametros(qrySelectTabelas)
''    strSQL = "Select "
''    Set rstOrigem = db.OpenRecordset("Select distinct Destino from tblOrigemDestino where tabela = '" & tabela & "'")
''    Do While Not rstOrigem.EOF
''        strSQL = strSQL & strSplit(rstOrigem.Fields("Destino").value, ".", 1) & ","
''        rstOrigem.MoveNext
''    Loop
''
''    strSQL = left(strSQL, Len(strSQL) - 1) & " from " & tabela
''    qryDeleteExists "qry_" & tabela
''    qryCreate "qry_" & tabela, strSQL
''Next tabela
''
''db.Close: Set db = Nothing
''
''End Sub
'
'Sub azs_teste_qryDadosGerais_Update_IdFornCompraNF()
'
'Dim qryDadosGerais_Update_IdFornCompraNF As String: qryDadosGerais_Update_IdFornCompraNF = _
'        "UPDATE (SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF, tmpClientes.C�DIGOClientes FROM tmpClientes) AS qryClientesFornecedor " & _
'        "INNER JOIN tblDadosConexaoNFeCTe ON tblDadosConexaoNFeCTe.CNPJ_emit = qryClientesFornecedor.strCNPJ_CPF " & _
'        "SET tblDadosConexaoNFeCTe.ID_Forn_CompraNF = qryClientesFornecedor.C�DIGOClientes;"
'
'        Application.CurrentDb.Execute qryDadosGerais_Update_IdFornCompraNF, dbSeeChanges
'
'
'End Sub
'
'
'Sub azs_teste_qryDadosGerais_Update_TransferenciaSisparts_CFOP()
'
'Dim qryDadosGerais_Update_TransferenciaSisparts_CFOP As String: qryDadosGerais_Update_TransferenciaSisparts_CFOP = _
'        "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.CFOP = ""6152"" WHERE (((tblDadosConexaoNFeCTe.ID_Tipo)=DLookUp(""id"",""tblTipos"",""Descricao='6 - NF-e Transfer�ncia com c�digo Sisparts'"")) AND ((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1));"
'
'        Application.CurrentDb.Execute qryDadosGerais_Update_TransferenciaSisparts_CFOP, dbSeeChanges
'
'
'End Sub
'
'Sub azs_teste_ReprocessamentoDeArquivos()
'
'
''' Corre��o de "CaminhoDoArquivo"
''' 1. Copiar "CaminhoDoArquivo_bkp" para "CaminhoDoArquivo"
'Dim sql_Update_tblDadosConexaoNFeCTe_CaminhoDoArquivo As String: sql_Update_tblDadosConexaoNFeCTe_CaminhoDoArquivo = _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.CaminhoDoArquivo = [tblDadosConexaoNFeCTe].[CaminhoDoArquivo_bkp];"
'    Application.CurrentDb.Execute sql_Update_tblDadosConexaoNFeCTe_CaminhoDoArquivo
'
'
''' Classifica��o de Processados
''' 2. Reclassificar como "registroProcessado(2)" onde "registroValido(1)"
'Dim sql_Update_tblDadosConexaoNFeCTe_registroProcessado As String: sql_Update_tblDadosConexaoNFeCTe_registroProcessado = _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 2 WHERE (((tblDadosConexaoNFeCTe.registroValido)=1));"
'    Application.CurrentDb.Execute sql_Update_tblDadosConexaoNFeCTe_registroProcessado
'
'
''' Classifica��o de Expurgo
''' 3. Reclassificar como "registroProcessado(4) - Expurgo" onde "registroValido(0) - N�o � valido"
''' 4. Reclassificar como "registroProcessado(4) - Expurgo" onde "ChvAcesso(null) - Sem chave de acesso"
'Dim sql_Update_tblDadosConexaoNFeCTe_registroProcessado_Expurgo() As Variant: sql_Update_tblDadosConexaoNFeCTe_registroProcessado_Expurgo = Array( _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 4 WHERE (((tblDadosConexaoNFeCTe.registroValido)=0));", _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 0 WHERE (((tblDadosConexaoNFeCTe.ChvAcesso) Is Null));")
'    executarComandos sql_Update_tblDadosConexaoNFeCTe_registroProcessado_Expurgo
'
'End Sub
'
'
''Sub teste_choose()
''
''    Debug.Print Choose(5, "caminhoDeColeta", "caminhoDeColeta", "caminhoDeColetaProcessados", "caminhoDeColetaExpurgo", "caminhoDeColetaAcoes")
''
''End Sub
'
'''' CARREGAR PARAMETROS UNICOS
''Public Function pegarValorDoParametro(pConsulta As String, pTipoDeParametro As String, Optional pCampo As String) As String
''Dim db As DAO.Database: Set db = CurrentDb
''Dim strTmp As String: strTmp = Replace(pConsulta, "strParametro", pTipoDeParametro)
''Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(strTmp)
''
''    pegarValorDoParametro = rst.Fields(IIf(pCampo <> "", pCampo, "ValorDoParametro")).value
''
''db.Close
''End Function
'
'
'
'''''''''#####################################
'''''''''#####################################
'''''''''#####################################
'
''Function carregarCamposValores(pRepositorio As String, pChvAcesso As String) As String
''
'''Dim pRepositorio As String: pRepositorio = "tblCompraNFItem"
'''Dim pChvAcesso As String: pChvAcesso = "32210368365501000296550000000638791001361285"
''
''Dim Scripts As New clsConexaoNfeCte
''Dim db As DAO.Database: Set db = CurrentDb
''Dim rstCampos As DAO.Recordset: Set rstCampos = db.OpenRecordset(Replace(Scripts.SelectCamposNomes, "pRepositorio", pRepositorio))
''Dim rstOrigem As DAO.Recordset
''
''Dim tmpScript As String
''Dim tmpValidarCampo As String: tmpValidarCampo = right(pRepositorio, Len(pRepositorio) - 3)
''
''Dim sqlOrigem As String: sqlOrigem = _
''    "Select * from (" & Replace(Scripts.SelectRegistroValidoPorcessado, "pRepositorio", pRepositorio) & ") as tmpRepositorio where tmpRepositorio.ChvAcesso_CompraNF = '" & pChvAcesso & "'"
''
''    Set rstOrigem = db.OpenRecordset(sqlOrigem)
''
''    rstOrigem.MoveLast
''    rstOrigem.MoveFirst
''    Do While Not rstOrigem.EOF
''        tmpScript = tmpScript & "("
''
''        '' LISTAGEM DE CAMPOS
''        rstCampos.MoveFirst
''        Do While Not rstCampos.EOF
''
''            '' CRIAR SCRIPT DE INCLUSAO DE DADOS NA TABELA DESTINO
''            '' 2. campos x formatacao
''            If InStr(rstCampos.Fields("campo").value, tmpValidarCampo) Then
''
''                If InStr(rstCampos.Fields("campo").value, "NumPed_CompraNF") Then tmpScript = tmpScript & "strNumPed_CompraNF,": GoTo pulo
''
''                If rstCampos.Fields("formatacao").value = "opTexto" Then
''                    tmpScript = tmpScript & "'" & rstOrigem.Fields(rstCampos.Fields("campo").value).value & "',"
''
''                ElseIf rstCampos.Fields("formatacao").value = "opNumero" Or rstCampos.Fields("formatacao").value = "opMoeda" Then
''                    tmpScript = tmpScript & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", rstOrigem.Fields(rstCampos.Fields("campo").value).value, rstCampos.Fields("valorPadrao").value) & ","
''
''                ElseIf rstCampos.Fields("formatacao").value = "opTime" Then
''                    tmpScript = tmpScript & "'" & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", Format(rstOrigem.Fields(rstCampos.Fields("campo").value).value, DATE_TIME_FORMAT), rstCampos.Fields("valorPadrao").value) & "',"
''
''                ElseIf rstCampos.Fields("formatacao").value = "opData" Then
''                    tmpScript = tmpScript & "'" & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", Format(rstOrigem.Fields(rstCampos.Fields("campo").value).value, DATE_FORMAT), rstCampos.Fields("valorPadrao").value) & "',"
''
''                End If
''
''            End If
''
''pulo:
''            rstCampos.MoveNext
''            DoEvents
''        Loop
''
''        tmpScript = left(tmpScript, Len(tmpScript) - 1) & "),"
''        rstOrigem.MoveNext
''        DoEvents
''    Loop
''
''    Set Scripts = Nothing
''    rstCampos.Close
''    rstOrigem.Close
''    db.Close
''
''    carregarCamposValores = left(tmpScript, Len(tmpScript) - 1)
''
''End Function
'
''Sub teste__carregarScript_Insert()
''
'''' 01
'''Debug.Print carregarScript_Insert("tblCompraNF", "32210368365501000296550000000638811001361356")
''
'''' 02
''Debug.Print carregarScript_Insert("tblCompraNFItem", "32210368365501000296550000000638791001361285")
''
'''' 23
'''Debug.Print carregarScript_Insert("tblCompraNFItem", "32210368365501000296550000000638811001361356")
''
''End Sub
''
''Function carregarScript_Insert(pRepositorio As String, pChvAcesso As String) As String
''
''Dim strCamposNomes As String: _
''    strCamposNomes = carregarCamposNomes(pRepositorio)
''
''Dim strCamposNomesTmp As String: _
''    strCamposNomesTmp = Replace(strCamposNomes, "_" & right(pRepositorio, Len(pRepositorio) - 3), "")
''
'''Dim strCamposValores As Collection: _
'''    strCamposValores = carregarCamposValores(pRepositorio, pChvAcesso)
''
''Dim item As Variant
''
''    For Each item In carregarCamposValores(pRepositorio, pChvAcesso)
''        Debug.Print CStr(i)
''    Next item
''
''
''Dim tmpScript As String: _
''    tmpScript = "INSERT INTO " & pRepositorio & " ( " & strCamposNomes & " ) SELECT " & strCamposNomesTmp & " FROM ( VALUES " & strCamposValores & " ) AS TMP ( " & strCamposNomesTmp & " ) LEFT JOIN " & pRepositorio & " ON " & pRepositorio & ".ChvAcesso_CompraNF = tmp.ChvAcesso WHERE " & pRepositorio & ".ChvAcesso_CompraNF IS NULL;"
''    '"INSERT INTO " & pRepositorio & " ( " & strCamposNomes & " ) SELECT " & strCamposNomesTmp & " FROM ( VALUES ( " & strCamposValores & " ) ) AS TMP ( " & strCamposNomesTmp & " ) LEFT JOIN " & pRepositorio & " ON " & pRepositorio & ".ChvAcesso_CompraNF = tmp.ChvAcesso WHERE " & pRepositorio & ".ChvAcesso_CompraNF IS NULL;"
''
''    carregarScript_Insert = tmpScript
''
''End Function
'
'
'Sub azs_teste_compras_atualizarCampos()
'Dim DadosGerais As New clsConexaoNfeCte
'
'    DadosGerais.compras_atualizarCampos
'
'Set DadosGerais = Nothing
'End Sub

'Sub azs_teste_update_Almox_CompraNFItem()
'
''' BANCO_DESTINO
'Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
'Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
'Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
'Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
'Dim dbDestino As New Banco: dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
'
'Dim sql_comprasItens_update_Almox_CompraNFItem As String:
'    sql_comprasItens_update_Almox_CompraNFItem = "UPDATE tblCompraNFItem " & _
'                                                    "SET tblCompraNFItem.Almox_CompraNFItem = tmpEstoqueAlmox.Codigo_Almox " & _
'                                                    "FROM tmpEstoqueAlmox RIGHT JOIN tblCompraNF ON tmpEstoqueAlmox.CodUnid_Almox = tblCompraNF.Fil_CompraNF " & _
'                                                    "INNER JOIN tblCompraNFItem ON tblCompraNF.ID_CompraNF = tblCompraNFItem.ID_CompraNF_CompraNFItem " & _
'                                                    "WHERE tmpEstoqueAlmox.Codigo_Almox IN (12,1,6) ; " ''AND tblCompraNFItem.Almox_CompraNFItem IS NULL
'
'
''' #ANALISE_DE_PROCESSAMENTO
'Dim DT_PROCESSO As Date: DT_PROCESSO = Now()
'
'    '' #20211202_update_Almox_CompraNFItem
'    'dbDestino.SqlExecute sql_comprasItens_update_Almox_CompraNFItem
'    Application.CurrentDb.Execute sql_comprasItens_update_Almox_CompraNFItem
'
'
''' #ANALISE_DE_PROCESSAMENTO
'statusFinal DT_PROCESSO, "azs_teste_update_Almox_CompraNFItem()"
'
'End Sub


'Sub teste_qryComprasItem_Update_STs()
'Dim scripts As New clsConexaoNfeCte
'
'Dim qryProcessos() As Variant: qryProcessos = Array( _
'                                                    scripts.UpdateItens_AjustesCampos, _
'                                                    scripts.UpdateItens_STs, _
'                                                    scripts.UpdateItens_STs_CTe_ST_CompraNFItem _
'                                                    ): executarComandos qryProcessos
'Set scripts = Nothing
'End Sub

''' #20220111_update_Almox_CompraNFItem
'Private Const qryComprasItem_Update_Almox_CompraNFItem_55_1907 As String = _
'            "UPDATE tblCompraNF " & _
'            "INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF " & _
'            "SET tblCompraNFItem.Almox_CompraNFItem = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro] = 'Almox_CompraNFItem|55|1.907'"") " & _
'            "WHERE (((tblCompraNF.CFOP_CompraNF) = ""1.907"") AND ((tblCompraNF.ModeloDoc_CompraNF) = ""55""));"
'
''' #20220111_update_Almox_CompraNFItem
'Private Const qryComprasItem_Update_Almox_CompraNFItem_55_2152_PSC As String = _
'            "UPDATE tblCompraNF " & _
'            "INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF " & _
'            "SET tblCompraNFItem.Almox_CompraNFItem = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro] = 'Almox_CompraNFItem|55|2.152|PSC'"") " & _
'            "WHERE (((tblCompraNF.CFOP_CompraNF) = ""2.152"") AND ((tblCompraNF.Fil_CompraNF) = ""PSC"")    AND ((tblCompraNF.ModeloDoc_CompraNF) = ""55"") );"
'
''' #20220111_update_Almox_CompraNFItem
'Private Const qryComprasItem_Update_Almox_CompraNFItem_55_2152_PSP As String = _
'            "UPDATE tblCompraNF " & _
'            "INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF " & _
'            "SET tblCompraNFItem.Almox_CompraNFItem = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro] = 'Almox_CompraNFItem|55|2.152|PSP'"") " & _
'            "WHERE (((tblCompraNF.CFOP_CompraNF) = ""2.152"") AND ((tblCompraNF.Fil_CompraNF) = ""PSP"")    AND ((tblCompraNF.ModeloDoc_CompraNF) = ""55"") );"
'
''' #20220111_update_Almox_CompraNFItem
'Private Const qryComprasItens_Update_Almox_CompraNFItem As String = _
'            "UPDATE (tmpEstoqueAlmox RIGHT JOIN tblCompraNF ON tmpEstoqueAlmox.CodUnid_Almox = tblCompraNF.Fil_CompraNF) INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF SET tblCompraNFItem.Almox_CompraNFItem = [tmpEstoqueAlmox].[Codigo_Almox] WHERE (((tmpEstoqueAlmox.Codigo_Almox) In (""12"",""1"",""6"")));"
'
'Sub teste_update()
'
'Dim qryProcessos() As Variant: qryProcessos = Array( _
'                                                    qryComprasItens_Update_Almox_CompraNFItem, _
'                                                    qryComprasItem_Update_Almox_CompraNFItem_55_1907, _
'                                                    qryComprasItem_Update_Almox_CompraNFItem_55_2152_PSC, _
'                                                    qryComprasItem_Update_Almox_CompraNFItem_55_2152_PSP _
'                                                    ): executarComandos qryProcessos
'
'
'End Sub



'Sub azs_teste_sql_comprasItens_update_IdProd()
'
''' BANCO_DESTINO
'Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
'Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
'Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
'Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
'Dim dbDestino As New Banco
'
''' BANCO_LOCAL
'Dim Scripts As New clsConexaoNfeCte
'
'Dim sql_comprasItens_update_IdProd As String:
'    sql_comprasItens_update_IdProd = "UPDATE tblCompraNFItem SET ID_Prod_CompraNFItem = tbProdutos.[c�digo] " & _
'                                        "FROM tblCompraNFItem AS tbItens " & _
'                                        "INNER JOIN tblCompraNF as tbCompras ON tbCompras.ID_CompraNF = tbItens.ID_CompraNF_CompraNFItem " & _
'                                        "INNER join [Cadastro de Produtos] as tbProdutos on tbProdutos.modelo = 'Transporte'  " & _
'                                        "WHERE tbCompras.Sit_CompraNF = 6 and tbItens.ID_Prod_CompraNFItem=0;"
'
'
'    '' BANCO_DESTINO
'    dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
'
'    dbDestino.SqlExecute sql_comprasItens_update_IdProd
'
'
'End Sub
'


'Sub azs_teste_FuncionamentoGeralDeProcessamentoDeArquivos()
'Dim strCaminhoAcoes As String: strCaminhoAcoes = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColetaAcoes'")
'Dim pArquivos As Collection: Set pArquivos = New Collection
'
'    ''==================================================
'    '' LIMPAR REPOSITORIOS
'    ''==================================================
'
''    '' Limpar toda a base de dados
''    Application.CurrentDb.Execute "Delete from tblDadosConexaoNFeCTe"
''
''    '' Limpar repositorio de itens de compras
''    Application.CurrentDb.Execute _
''            "Delete from tblCompraNFItem"
''
''    '' Limpar repositorio de compras
''    Application.CurrentDb.Execute _
''            "Delete from tblCompraNF"
'
'    ''==================================================
'    '' PROCESSAR DADOS
'    ''==================================================
'
'    '' Carregar todos os arquivos para processamento.
'    processarDadosGerais pArquivos
'
'    '' Processamento de arquivos pendentes da pasta de coleta.
'    processarArquivosPendentes
'
'    '' EXPORTA��O DE DADOS
'    CadastroDeComprasEmServidor
'
'    ''==================================================
'    '' PROCESSAR ARQUIVOS
'    ''==================================================
'
'    '' MOVER ARQUIVOS
''    MoverArquivosProcessados
'
'    '' LAN�AMENTO
''    gerarArquivosJson opFlagLancadaERP, , strCaminhoAcoes
'
'    '' MANIFESTO
''    gerarArquivosJson opManifesto, , strCaminhoAcoes
'
'
'Debug.Print "### Concluido! - testeDeFuncionamentoGeral"
'TextFile_Append CurrentProject.path & "\" & strLog(), "Concluido! - testeDeFuncionamentoGeral"
'
'End Sub
'
'
'Sub azs_teste_processarDadosGerais()
'Dim pArquivos As Collection: Set pArquivos = New Collection
'Dim arquivos As Collection: Set arquivos = New Collection
'Dim item As Variant
'
'    pArquivos.add "C:\xmls\processados\68.365.5010003-77 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\42210312680452000302550020000895841453583169-nfeproc.xml"
'    pArquivos.add "C:\xmls\processados\68.365.5010003-77 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\42210312680452000302550020000902571508970265-nfeproc.xml"
'    pArquivos.add "C:\xmls\processados\68.365.5010003-77 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\32210368365501000296550000000638791001361285-nfeproc.xml"
'    pArquivos.add "C:\xmls\processados\68.365.5010001-05 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\42210368365501000377550000000066481001365721-nfeproc.xml"
'
'    pArquivos.add "C:\xmls\processados\68.365.5010002-96 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\32210320147617002608570010019864161998013580-cteproc_x.xml"
'    pArquivos.add "C:\xmls\processados\68.365.5010003-77 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\42210320147617000494570010009617261999038272-cteproc.xml"
'    pArquivos.add "C:\xmls\processados\68.365.5010003-77 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\42210348740351012767570000020998441987851951-cteproc.xml"
'    pArquivos.add "C:\xmls\processados\68.365.5010003-77 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\42210300634453001303570010001139451001171544-cteproc.xml"
'
'    pArquivos.add "C:\xmls\expurgo\68.365.5010003-77 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\42210307872326000158550040001546741011035210-nfeproc.xml"
'
'
'    For Each item In pArquivos
'        Debug.Print getFileName(CStr(item))
'        Debug.Print DLookup("ID", "tblDadosConexaoNFeCTe", "Chave='" & getFileName(CStr(item)) & "'")
'        If (IsNull(DLookup("ID", "tblDadosConexaoNFeCTe", "Chave='" & getFileName(CStr(item)) & "'"))) Then arquivos.add CStr(item)
'    Next
'
'    For Each item In arquivos
'        Debug.Print "#"
'        Debug.Print CStr(item)
'    Next
'
''processarDadosGerais pArquivos
'
'End Sub



'Sub azs_teste_ProcessamentoDeArquivo_opCompras()
'Dim pCaminho As String
'
'Dim s As New clsProcessamentoDados
'Dim ChavesDeAcesso As Collection: Set ChavesDeAcesso = New Collection
'Dim Chave As Variant
'
''' NFe
''ChavesDeAcesso.add "42210312680452000302550020000895841453583169" ' - ARQUIVO X PLANILHA
''ChavesDeAcesso.add "42210312680452000302550020000902571508970265"
''ChavesDeAcesso.add "32210368365501000296550000000638791001361285"
''ChavesDeAcesso.add "42210368365501000377550000000066481001365721"
'
''' CTe
''ChavesDeAcesso.add "32210320147617002608570010019864161998013580"
''ChavesDeAcesso.add "42210320147617000494570010009617261999038272"
''ChavesDeAcesso.add "42210348740351012767570000020998441987851951"
''ChavesDeAcesso.add "42210300634453001303570010001139451001171544"
'
''' ENTENDIMENTO
''ChavesDeAcesso.add "42210307872326000158550040001546741011035210"
'
'
'ChavesDeAcesso.add "32210368365501000296550000000638961001363203"
'
'    For Each Chave In ChavesDeAcesso
'
''        pCaminho = _
''            DLookup("CaminhoDoArquivo", "tblDadosConexaoNFeCTe", "ChvAcesso='" & CStr(Chave) & "'")
'
'
'        ' pCaminho = "C:\ConexaoNFe\XML\68.365.5010003-77 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\32210368365501000296550000000638961001363203-nfeproc.xml"
'
'        pCaminho = _
'            "C:\ConexaoNFe\XML\68.365.5010003-77 - Proparts Com�rcio de Artigos Esportivos e Tecnologia Ltda\recebimento\43210331275238000153550010000011371772829986-nfeproc.xml"
'
'
'        Debug.Print pCaminho
'
'
'        s.DeleteProcessamento
'        s.ProcessamentoDeArquivo pCaminho, opCompras
'
'        '' IDENTIFICAR CAMPOS
'        s.UpdateProcessamentoIdentificarCampos "tblCompraNF"
'
'        '' CORRE��O DE DADOS MARCADOS ERRADOS EM ITENS DE COMPRAS
'        s.UpdateProcessamentoLimparItensMarcadosErrados
'
'        '' IDENTIFICAR CAMPOS DE ITENS DE COMPRAS
'        s.UpdateProcessamentoIdentificarCampos "tblCompraNFItem"
'
'        '' FORMATAR DADOS
'        s.UpdateProcessamentoFormatarDados
'
'    Next
'
'Set s = Nothing
'
'End Sub


