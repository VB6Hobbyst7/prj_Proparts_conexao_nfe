Attribute VB_Name = "azs_VALIDAR_DADOS"
Option Compare Database



Sub teste_moveFile()

Dim pChvAcesso As String: pChvAcesso = ""

Dim pSource As String: _
    pSource = DLookup("[CaminhoDoArquivo]", "[tblDadosConexaoNFeCTe]", "[ChvAcesso]='" & pPathFile & "'")
    
Dim pFileName As String: _
    pFileName = getNomeDoArquivo(pSource)

Dim pDestination As String: _
    pDestination = CStr(DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColetaTentativaDeReprocessamento'"))
    
Dim tFileLog As String: _
    tFileLog = left(Split(strLog, ".")(0), 6) & ".csv"


    '' mover arquivo
    FileCopy pSource & pFileName, pDestination & pFileName
    Kill pSource & pFileName
    
    '' log
    If (Dir(pDestination & tFileLog) = "") Then TextFile_Append pDestination & tFileLog, "arquivo;data_tentativa"
    TextFile_Append pDestination & tFileLog, pFileName & ";" & strControle

End Sub


Sub Controle_InibirReprocessamento()
On Error GoTo adm_Err

'' BANCO_DESTINO
Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
Dim dbDestino As New Banco: dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer

'' BANCO_LOCAL
Dim db As DAO.Database: Set db = CurrentDb
Dim sql_Select_ChvAcesso_Pendentes As String: sql_Select_ChvAcesso_Pendentes = _
    "SELECT tblDadosConexaoNFeCTe.ChvAcesso FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0)) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0) ORDER BY tblDadosConexaoNFeCTe.ID;"
Dim rstChvAcesso As DAO.Recordset: Set rstChvAcesso = db.OpenRecordset(sql_Select_ChvAcesso_Pendentes)

'' CRIAR REPOSITORIO DE EXPURGO PARA REGISTROS COM TENTATIVA DE REPROCESSAMENTO
CreateDir CStr(DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColetaTentativaDeReprocessamento'"))

'' Expurgo - registroProcessado(4)
Dim sql_Update_registroProcessado_Expurgo As String: sql_Update_registroProcessado_Expurgo = _
    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 4 WHERE ((tblDadosConexaoNFeCTe.ChvAcesso)='strChvAcesso');"
    
    Do While Not rstChvAcesso.EOF
        dbDestino.SqlSelect "SELECT COUNT(*) as RegistroExistente FROM tblCompraNF where ChvAcesso_CompraNF = '" & rstChvAcesso.Fields("ChvAcesso").value & "'"
        If (dbDestino.rs.Fields("RegistroExistente").value = 1) Then
            Debug.Print rstChvAcesso.Fields("ChvAcesso").value
            Application.CurrentDb.Execute Replace(sql_Update_registroProcessado_Expurgo, "strChvAcesso", rstChvAcesso.Fields("ChvAcesso").value)
        End If
        
        
'        '' mover arquivo
'        If (Dir(rst.Fields("CaminhoDestino").value) <> "") Then Kill rst.Fields("CaminhoDestino").value
'        FileCopy rst.Fields("CaminhoDoArquivo").value, rst.Fields("CaminhoDestino").value
'        Kill rst.Fields("CaminhoDoArquivo").value
        
        
        rstChvAcesso.MoveNext
    Loop

    rstChvAcesso.Close
    dbDestino.CloseConnection
    db.Close

adm_Exit:
    Set dbDestino = Nothing
    Set rstChvAcesso = Nothing
    Set db = Nothing
    Exit Sub

adm_Err:
    Debug.Print "expurgoDeRegistroCadastradoEmServidor() - " & Err.Description
    TextFile_Append CurrentProject.path & "\" & strLog(), Err.Description
    Resume adm_Exit
End Sub


