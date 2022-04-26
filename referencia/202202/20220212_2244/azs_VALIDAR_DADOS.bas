Attribute VB_Name = "azs_VALIDAR_DADOS"
Option Compare Database


''' #20220202_Controle_InibirReprocessamento
'Public Sub Controle_InibirReprocessamento()
'On Error GoTo adm_Err
'
''' BANCO_DESTINO
'Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
'Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
'Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
'Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
'Dim dbDestino As New Banco: dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
'
''' BANCO_LOCAL
'Dim db As DAO.Database: Set db = CurrentDb
'Dim sql_Select_ChvAcesso_Pendentes As String: sql_Select_ChvAcesso_Pendentes = _
'    "SELECT tblDadosConexaoNFeCTe.ChvAcesso FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0)) ORDER BY tblDadosConexaoNFeCTe.ID;"
'Dim rstChvAcesso As DAO.Recordset: Set rstChvAcesso = db.OpenRecordset(sql_Select_ChvAcesso_Pendentes)
'
'
'Dim sql_Update_registroProcessado_Expurgo() As Variant: sql_Update_registroProcessado_Expurgo = Array( _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 4 WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.ID_TIPO)=0));", _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 4 WHERE (((tblDadosConexaoNFeCTe.registroValido)=0) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0));", _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 4 WHERE (((tblDadosConexaoNFeCTe.codMod)=55) AND ((tblDadosConexaoNFeCTe.CFOP)=5927));", _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 4 WHERE (((tblDadosConexaoNFeCTe.codMod)=55) AND ((tblDadosConexaoNFeCTe.CFOP)=1949));")
'    executarComandos sql_Update_registroProcessado_Expurgo
'
'
'
''' CRIAR REPOSITORIO DE EXPURGO PARA REGISTROS COM TENTATIVA DE REPROCESSAMENTO
'CreateDir CStr(DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColetaTentativaDeReprocessamento'"))
'
'
'    Do While Not rstChvAcesso.EOF
'        dbDestino.SqlSelect "SELECT COUNT(*) as RegistroExistente FROM tblCompraNF where ChvAcesso_CompraNF = '" & rstChvAcesso.Fields("ChvAcesso").value & "'"
'        If (dbDestino.rs.Fields("RegistroExistente").value = 1) Then
'            Debug.Print rstChvAcesso.Fields("ChvAcesso").value
'            mover_ArquivosComTentativaDeReprocessamento rstChvAcesso.Fields("ChvAcesso").value
'
'
'            Application.CurrentDb.Execute _
'                    "Delete from tblCompraNFItem where ChvAcesso_CompraNF = '" & rstChvAcesso.Fields("ChvAcesso").value & "'"
'
'            Application.CurrentDb.Execute _
'                    "Delete from tblCompraNF where ChvAcesso_CompraNF = '" & rstChvAcesso.Fields("ChvAcesso").value & "'"
'
'
'            Application.CurrentDb.Execute _
'                    "Delete from tblDadosConexaoNFeCTe where ChvAcesso = '" & rstChvAcesso.Fields("ChvAcesso").value & "'"
'
'        End If
'
'        rstChvAcesso.MoveNext
'        DoEvents
'    Loop
'
'    rstChvAcesso.Close
'    dbDestino.CloseConnection
'    db.Close
'
'adm_Exit:
'    Set dbDestino = Nothing
'    Set rstChvAcesso = Nothing
'    Set db = Nothing
'    Exit Sub
'
'adm_Err:
'    Debug.Print "expurgoDeRegistroCadastradoEmServidor() - " & Err.Description
'    TextFile_Append CurrentProject.path & "\" & strLog(), Err.Description
'    Resume adm_Exit
'End Sub
'
''' #20220202_Controle_InibirReprocessamento
'Function mover_ArquivosComTentativaDeReprocessamento(pChvAcesso As String)
'
'Dim pSource As String: _
'    pSource = DLookup("[CaminhoDoArquivo]", "[tblDadosConexaoNFeCTe]", "[ChvAcesso]='" & pChvAcesso & "'")
'
'Dim pFileName As String: _
'    pFileName = getNomeDoArquivo(pSource)
'
'Dim pDestination As String: _
'    pDestination = CStr(DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColetaTentativaDeReprocessamento'"))
'
'Dim tFileLog As String: _
'    tFileLog = left(Split(strLog, ".")(0), 6) & ".csv"
'
'    '' mover arquivo
'    pSource = getPath(pSource)
'    FileCopy pSource & pFileName, pDestination & pFileName
'    Kill pSource & pFileName
'
'    '' log
'    If (Dir(pDestination & tFileLog) = "") Then TextFile_Append pDestination & tFileLog, "arquivo;data_tentativa"
'    TextFile_Append pDestination & tFileLog, pFileName & ";" & strControle
'
'End Function
'
'
