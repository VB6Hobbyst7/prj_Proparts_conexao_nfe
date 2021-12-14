Attribute VB_Name = "azs_modMSSQL"
'Public Const DATE_TIME_FORMAT               As String = "yyyy/mm/dd hh:mm:ss"
'Public Const DATE_FORMAT                    As String = "yyyy/mm/dd"
'
''' #20210823_CadastroDeComprasEmServidor
'Sub CadastroDeComprasEmServidor()
'
''' #ANALISE_DE_PROCESSAMENTO
'Dim DT_PROCESSO As Date: DT_PROCESSO = Now()
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
'Dim db As DAO.Database: Set db = CurrentDb
'Dim rstChvAcesso As DAO.Recordset: Set rstChvAcesso = db.OpenRecordset(Replace(Scripts.SelectRegistroValidoPorcessado, "pRepositorio", "tblCompraNF"))
'
'''' #AILTON - TESTES
''Dim TMP2 As String: TMP2 = "select * from (" & Replace(Scripts.SelectRegistroValidoPorcessado, "pRepositorio", "tblCompraNF") & ") as tmp where tmp.ChvAcesso_CompraNF = '42210300634453001303570010001139451001171544'"
''Dim rstChvAcesso As DAO.Recordset: Set rstChvAcesso = db.OpenRecordset(TMP2)
'
'Dim qryCompras_Insert_Compras As String
'Dim qryComprasItens_Update_IDCompraNF As String
'
''' CONTROLE DE "NumPed_CompraNF" RELACIONADO COM REGISTROS DO SERVIDOR
'Dim contador As Long: contador = 1
'
''' CONTROLE DE REPOSITORIOS x CHAVES DE ACESSO
'Dim item As Variant
'Dim pChvAcesso As String
'Dim pRepositorio As String
'
''' SCRIPT DE INCLUS�O DE DADOS NO SERVIDOR
'Dim tmpCamposNomes As String
'Dim strCamposNomes As String
'Dim strCamposNomesTmp As String
'Dim strRepositorio As String
'Dim tmpScript As String: _
'    tmpScript = "INSERT INTO pRepositorio ( strCamposNomes ) SELECT strCamposTmp FROM ( VALUES strCamposValores ) AS TMP ( strCamposTmp ) LEFT JOIN pRepositorio ON pRepositorio.ChvAcesso_CompraNF = tmp.ChvAcesso WHERE pRepositorio.ChvAcesso_CompraNF IS NULL;"
'
'Dim tmpScriptItens As String: _
'    tmpScriptItens = "INSERT INTO pRepositorio ( strCamposNomes ) SELECT strCamposTmp FROM ( VALUES strCamposValores ) AS TMP ( strCamposTmp );"
'
'Dim qryComprasCTe_Update_AjustesCampos As String: _
'    qryComprasCTe_Update_AjustesCampos = "UPDATE tblCompraNF SET tblCompraNF.HoraEntd_CompraNF = NULL ,tblCompraNF.IDVD_CompraNF = NULL WHERE (((tblCompraNF.ChvAcesso_CompraNF) IN (pLista_ChvAcesso_CompraNF)));"
'
'
''' #CONTADOR
'Dim contadorDeRegistros As Long: contadorDeRegistros = 1
'Dim totalDeRegistros As Long
'
'
''' VALIDAR CONCILIA��O
'Dim TMP As String
'
'    '' BANCO_DESTINO
'    dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
'
'    '' #BARRA_PROGRESSO
'    totalDeRegistros = rstChvAcesso.RecordCount
'    SysCmd acSysCmdInitMeter, "Exporta��o...", totalDeRegistros
'
'    '' CADASTRO
'    Do While Not rstChvAcesso.EOF
'        pRepositorio = "tblCompraNF"
'        pChvAcesso = rstChvAcesso.Fields("ChvAcesso_CompraNF").value
'
'        tmpCamposNomes = carregarCamposNomes(pRepositorio)
'        strCamposNomes = Replace(tmpScript, "strCamposNomes", tmpCamposNomes)
'        strCamposNomesTmp = Replace(Replace(tmpScript, "strCamposNomes", tmpCamposNomes), "strCamposTmp", Replace(tmpCamposNomes, "_" & right(pRepositorio, Len(pRepositorio) - 3), ""))
'        strRepositorio = Replace(strCamposNomesTmp, "pRepositorio", pRepositorio)
'
'        '' CONTADOR
'        dbDestino.SqlSelect "SELECT max(NumPed_CompraNF)+1 as contador from tblCompraNF"
'        contador = IIf(IsNull(dbDestino.rs.Fields("contador").value), 1, dbDestino.rs.Fields("contador").value)
'
'        '' CADASTRO DE COMPRAS
'        For Each item In carregarCamposValores(pRepositorio, pChvAcesso)
'            TMP = Replace(Replace(strRepositorio, "strCamposValores", item), "strNumPed_CompraNF", contador)
'            Debug.Print TMP
'            dbDestino.SqlExecute TMP
'        Next item
'
'        '' RELACIONAR ITENS DE COMPRAS COM COMPRAS J� CADASTRADAS NO SERVIDOR
'        dbDestino.SqlSelect "SELECT ChvAcesso_CompraNF,ID_CompraNF FROM tblCompraNF where ChvAcesso_CompraNF = '" & pChvAcesso & "';"
'        qryComprasItens_Update_IDCompraNF = Replace(Replace(Scripts.UpdateComprasItens_IDCompraNF, "strChave", pChvAcesso), "strID_Compra", dbDestino.rs.Fields("ID_CompraNF").value)
'        If Not dbDestino.rs.EOF Then
'            pRepositorio = "tblCompraNFItem"
'
'            tmpCamposNomes = carregarCamposNomes(pRepositorio)
'            strCamposNomes = Replace(tmpScriptItens, "strCamposNomes", tmpCamposNomes)
'            strCamposNomesTmp = Replace(Replace(tmpScriptItens, "strCamposNomes", tmpCamposNomes), "strCamposTmp", Replace(tmpCamposNomes, "_" & right(pRepositorio, Len(pRepositorio) - 3), ""))
'            strRepositorio = Replace(strCamposNomesTmp, "pRepositorio", pRepositorio)
'
'            '' #20210823_qryUpdateNumPed_CompraNF
'            Application.CurrentDb.Execute qryComprasItens_Update_IDCompraNF
'
'            '' CADASTRO DE ITENS DE COMPRAS
'            For Each item In carregarCamposValores(pRepositorio, pChvAcesso)
'                TMP = Replace(Replace(strRepositorio, "strCamposValores", item), "strNumPed_CompraNF", contador)
'                Debug.Print TMP
'                dbDestino.SqlExecute TMP
'            Next item
'
'        End If
'
'        '' MUDAR STATUS DO REGISTRO
'        Application.CurrentDb.Execute Replace(Scripts.compras_atualizarEnviadoParaServidor, "strChave", rstChvAcesso.Fields("chvAcesso_CompraNF").value)
'
'        rstChvAcesso.MoveNext
'        contador = contador + 1
'        contadorDeRegistros = contadorDeRegistros + 1
'
'        '' #BARRA_PROGRESSO
'        SysCmd acSysCmdUpdateMeter, contadorDeRegistros
'
'        DoEvents
'    Loop
'
'    '' #20211122_AjusteDeCampos_CTe
'    dbDestino.SqlExecute Replace(qryComprasCTe_Update_AjustesCampos, "pLista_ChvAcesso_CompraNF", carregarComprasCTe)
'    dbDestino.SqlExecute "UPDATE tblCompraNF SET HoraEntd_CompraNF = NULL, IDVD_CompraNF = NULL WHERE tblCompraNF.IDVD_CompraNF=0;"
'
'    '' #20211128_LimparRepositorios
'    '' Limpar repositorio de itens de compras
'    Application.CurrentDb.Execute _
'            "Delete from tblCompraNFItem"
'
'    '' Limpar repositorio de compras
'    Application.CurrentDb.Execute _
'            "Delete from tblCompraNF"
'
'
'rstChvAcesso.Close
'dbDestino.CloseConnection
'db.Close
'
'Set Scripts = Nothing
'Set rstChvAcesso = Nothing
'Set db = Nothing
'
''' #ANALISE_DE_PROCESSAMENTO
'statusFinal DT_PROCESSO, "CadastroDeComprasEmServidor - Exportar compras ( Quantidade de registros: " & contador & " )"
'
''' #BARRA_PROGRESSO
'SysCmd acSysCmdRemoveMeter
'
'Debug.Print "Concluido!"
'
'End Sub
'
'Function carregarCamposValores(pRepositorio As String, pChvAcesso As String) As Collection
'Set carregarCamposValores = New Collection
'Dim Scripts As New clsConexaoNfeCte
'
''Dim pRepositorio As String: pRepositorio = "tblCompraNFItem"
''Dim pChvAcesso As String: pChvAcesso = "32210368365501000296550000000638791001361285"
'
''' BANCO_LOCAL
'Dim db As DAO.Database: Set db = CurrentDb
'
'Dim rstCampos As DAO.Recordset: Set rstCampos = db.OpenRecordset(Replace(Scripts.SelectCamposNomes, "pRepositorio", pRepositorio))
'Dim rstOrigem As DAO.Recordset
'
''' VALIDAR CONCILIA��O
'Dim tmpScript As String
'Dim tmpValidarCampo As String: tmpValidarCampo = right(pRepositorio, Len(pRepositorio) - 3)
'
'Dim sqlOrigem As String: sqlOrigem = _
'    "Select * from (" & Replace(Scripts.SelectRegistroValidoPorcessado, "pRepositorio", pRepositorio) & ") as tmpRepositorio where tmpRepositorio.ChvAcesso_CompraNF = '" & pChvAcesso & "'"
'
'    Set rstOrigem = db.OpenRecordset(sqlOrigem)
'
'    rstOrigem.MoveLast
'    rstOrigem.MoveFirst
'    Do While Not rstOrigem.EOF
'        tmpScript = "("
'
'        '' LISTAGEM DE CAMPOS
'        rstCampos.MoveFirst
'        Do While Not rstCampos.EOF
'
'            '' CRIAR SCRIPT DE INCLUSAO DE DADOS NA TABELA DESTINO
'            '' 2. campos x formatacao
'            If InStr(rstCampos.Fields("campo").value, tmpValidarCampo) Then
'
'                If InStr(rstCampos.Fields("campo").value, "NumPed_CompraNF") Then tmpScript = tmpScript & "strNumPed_CompraNF,": GoTo pulo
'
'                If rstCampos.Fields("formatacao").value = "opTexto" Then
'                    tmpScript = tmpScript & "'" & rstOrigem.Fields(rstCampos.Fields("campo").value).value & "',"
'
'                ElseIf rstCampos.Fields("formatacao").value = "opNumero" Or rstCampos.Fields("formatacao").value = "opMoeda" Then
'                    tmpScript = tmpScript & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", rstOrigem.Fields(rstCampos.Fields("campo").value).value, rstCampos.Fields("valorPadrao").value) & ","
'
'                ElseIf rstCampos.Fields("formatacao").value = "opTime" Then
'                    tmpScript = tmpScript & "'" & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", Format(rstOrigem.Fields(rstCampos.Fields("campo").value).value, DATE_TIME_FORMAT), rstCampos.Fields("valorPadrao").value) & "',"
'
'                ElseIf rstCampos.Fields("formatacao").value = "opData" Then
'                    tmpScript = tmpScript & "'" & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", Format(rstOrigem.Fields(rstCampos.Fields("campo").value).value, DATE_FORMAT), rstCampos.Fields("valorPadrao").value) & "',"
'
'                End If
'
'            End If
'
'pulo:
'            rstCampos.MoveNext
'            DoEvents
'        Loop
'
'        carregarCamposValores.add left(tmpScript, Len(tmpScript) - 1) & ")"
'        rstOrigem.MoveNext
'        DoEvents
'    Loop
'
'    Set Scripts = Nothing
'    rstCampos.Close
'    rstOrigem.Close
'    db.Close
'
'End Function
'
'Function carregarCamposNomes(pRepositorio As String) As String
'Dim Scripts As New clsConexaoNfeCte
'
''' BANCO_LOCAL
'Dim db As DAO.Database: Set db = CurrentDb
'Dim rstCampos As DAO.Recordset
'
'Dim tmpValidarCampo As String: tmpValidarCampo = right(pRepositorio, Len(pRepositorio) - 3)
'
''' VALIDAR CONCILIA��O
'Dim tmpScript As String
'
'    '' MONTAR STRING DE NOME DE COLUNAS
'    Set rstCampos = db.OpenRecordset(Replace(Scripts.SelectCamposNomes, "pRepositorio", pRepositorio))
'    Do While Not rstCampos.EOF
'        If InStr(rstCampos.Fields("campo").value, tmpValidarCampo) Then
'            tmpScript = tmpScript & rstCampos.Fields("campo").value & ","
'        End If
'        rstCampos.MoveNext
'        DoEvents
'    Loop
'
'    Set Scripts = Nothing
'    rstCampos.Close
'    db.Close
'
'    carregarCamposNomes = left(tmpScript, Len(tmpScript) - 1)
'
'End Function
'
''' #20211122_AjusteDeCampos_CTe
'Function carregarComprasCTe() As String
'Dim Scripts As New clsConexaoNfeCte
'
''' BANCO_LOCAL
'Dim db As DAO.Database: Set db = CurrentDb
'Dim rstCampos As DAO.Recordset
'
''' VALIDAR CONCILIA��O
'Dim tmpScript As String
'
'    '' MONTAR STRING DE NOME DE COLUNAS
'    Set rstCampos = db.OpenRecordset("qryCompras_CTe_Select_AjustesCampos")
'    Do While Not rstCampos.EOF
'        tmpScript = tmpScript & "'" & rstCampos.Fields("ChvAcesso_CompraNF").value & "',"
'
'        rstCampos.MoveNext
'        DoEvents
'    Loop
'
'    Set Scripts = Nothing
'    rstCampos.Close
'    db.Close
'
'    carregarComprasCTe = left(tmpScript, Len(tmpScript) - 1)
'
'End Function
'