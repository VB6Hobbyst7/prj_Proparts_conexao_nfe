Attribute VB_Name = "02_modMSSQL"
Public Const DATE_TIME_FORMAT               As String = "yyyy/mm/dd hh:mm:ss"
Public Const DATE_FORMAT                    As String = "yyyy/mm/dd"

Sub relacionarComprasItens()
'' BANCO_DESTINO
Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
Dim dbDestino As New Banco

'' BANCO_ORIGEM
Dim Scripts As New clsConexaoNfeCte
Dim db As DAO.Database: Set db = CurrentDb
Dim rstChvAcesso As DAO.Recordset: Set rstChvAcesso = db.OpenRecordset(Scripts.SelectRegistroValidoPorcessado)

Dim sqlTmp As String: sqlTmp = "UPDATE tblCompraNFItem SET tblCompraNFItem.ID_CompraNF_CompraNFItem = strID_Compra WHERE  tblCompraNFItem.ChvAcesso_CompraNF = 'strChave'"
Dim Tmp As String

    '' BANCO_DESTINO
    dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
    
    Do While Not rstChvAcesso.EOF
        dbDestino.SqlSelect "SELECT ChvAcesso_CompraNF,ID_CompraNF FROM tblCompraNF where ChvAcesso_CompraNF = '" & rstChvAcesso.Fields("ChvAcesso_CompraNF").value & "';"
        
        Tmp = Replace(Replace(sqlTmp, "strChave", dbDestino.rs.Fields("ChvAcesso_CompraNF").value), "strID_Compra", dbDestino.rs.Fields("ID_CompraNF").value)
        
        Debug.Print Tmp
        If Not dbDestino.rs.EOF Then Application.CurrentDb.Execute Tmp
        
        rstChvAcesso.MoveNext
        DoEvents
    Loop

rstChvAcesso.Close
dbDestino.CloseConnection
db.Close

End Sub


Sub CadastroDeComprasEmServidor()

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

'' BANCO_DESTINO
Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
Dim dbDestino As New Banco

'' BANCO_ORIGEM
Dim Scripts As New clsConexaoNfeCte
Dim db As DAO.Database: Set db = CurrentDb
Dim rstChvAcesso As DAO.Recordset: Set rstChvAcesso = db.OpenRecordset(Scripts.SelectRegistroValidoPorcessado)
Dim tmpScript As String

Dim contador As Long: contador = 1

    '' BANCO_DESTINO
    dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer

    '' CADASTRO
    Do While Not rstChvAcesso.EOF
        
        '' CONTADOR
        dbDestino.SqlSelect "SELECT max(NumPed_CompraNF)+1 as contador from tblCompraNF"
        contador = IIf(IsNull(dbDestino.rs.Fields("contador").value), 1, dbDestino.rs.Fields("contador").value)
        
        '' CADASTRO DE COMPRAS
        tmpScript = Replace(carregarScript_Insert("tblCompraNF", rstChvAcesso.Fields("ChvAcesso_CompraNF").value), "strNumPed_CompraNF", contador)
        Debug.Print tmpScript
        dbDestino.SqlExecute tmpScript
        
        '' RELACIONAR ITENS DE COMPRAS COM COMPRAS JÁ CADASTRADAS NO SERVIDOR
        relacionarComprasItens
        
        
        '' CADASTRO DE ITENS DE COMPRAS
        
        
        
        
        rstChvAcesso.MoveNext
        contador = contador + 1
        DoEvents
    Loop

rstChvAcesso.Close
dbDestino.CloseConnection
db.Close


Set Scripts = Nothing
Set rstChvAcesso = Nothing
Set db = Nothing

'' #ANALISE_DE_PROCESSAMENTO
statusFinal DT_PROCESSO, "CadastroDeComprasEmServidor - Exportar compras ( Quantidade de registros: " & contador & " )"


Debug.Print "Concluido!"

End Sub

Function carregarScript_Insert(pRepositorio As String, pChvAcesso As String) As String

Dim strCamposNomes As String: _
    strCamposNomes = carregarCamposNomes(pRepositorio)

Dim strCamposNomesTmp As String: _
    strCamposNomesTmp = Replace(strCamposNomes, "_CompraNF", "")

Dim strCamposValores As String: _
    strCamposValores = carregarCamposValores(pRepositorio, pChvAcesso)

Dim tmpScript As String: _
    tmpScript = "INSERT INTO tblCompraNF ( " & strCamposNomes & " ) SELECT " & strCamposNomesTmp & " FROM ( VALUES ( " & strCamposValores & " ) ) AS TMP ( " & strCamposNomesTmp & " ) LEFT JOIN tblCompraNF ON tblCompraNF.ChvAcesso_CompraNF = tmp.ChvAcesso WHERE tblCompraNF.ChvAcesso_CompraNF IS NULL;"
    
    carregarScript_Insert = tmpScript

End Function

Function carregarCamposValores(pRepositorio As String, pChvAcesso As String) As String
Dim Scripts As New clsConexaoNfeCte
Dim db As DAO.Database: Set db = CurrentDb
Dim rstCampos As DAO.Recordset: Set rstCampos = db.OpenRecordset(Replace(Scripts.SelectCamposNomes, "pRepositorio", pRepositorio))
Dim rstOrigem As DAO.Recordset

Dim tmpScript As String
Dim Tmp As String: Tmp = right(pRepositorio, Len(pRepositorio) - 3)

    Set rstOrigem = db.OpenRecordset("Select * from (" & Replace(Scripts.SelectRegistroValidoPorcessado, "pRepositorio", pRepositorio) & ") as tmpRepositorio where tmpRepositorio.ChvAcesso_CompraNF = '" & pChvAcesso & "'")
    
    Do While Not rstOrigem.EOF
        tmpScript = ""
    
        '' LISTAGEM DE CAMPOS
        rstCampos.MoveFirst
        Do While Not rstCampos.EOF
    
            '' CRIAR SCRIPT DE INCLUSAO DE DADOS NA TABELA DESTINO
            '' 2. campos x formatacao
            If InStr(rstCampos.Fields("campo").value, Tmp) Then
    
                If InStr(rstCampos.Fields("campo").value, "NumPed_CompraNF") Then tmpScript = tmpScript & "strNumPed_CompraNF,": GoTo pulo
                    
    
                If rstCampos.Fields("formatacao").value = "opTexto" Then
                    tmpScript = tmpScript & "'" & rstOrigem.Fields(rstCampos.Fields("campo").value).value & "',"
    
                ElseIf rstCampos.Fields("formatacao").value = "opNumero" Or rstCampos.Fields("formatacao").value = "opMoeda" Then
                    tmpScript = tmpScript & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", rstOrigem.Fields(rstCampos.Fields("campo").value).value, rstCampos.Fields("valorPadrao").value) & ","
    
                ElseIf rstCampos.Fields("formatacao").value = "opTime" Then
                    tmpScript = tmpScript & "'" & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", Format(rstOrigem.Fields(rstCampos.Fields("campo").value).value, DATE_TIME_FORMAT), rstCampos.Fields("valorPadrao").value) & "',"
    
                ElseIf rstCampos.Fields("formatacao").value = "opData" Then
                    tmpScript = tmpScript & "'" & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", Format(rstOrigem.Fields(rstCampos.Fields("campo").value).value, DATE_FORMAT), rstCampos.Fields("valorPadrao").value) & "',"
    
                End If
    
            End If
            
pulo:
            rstCampos.MoveNext
            DoEvents
        Loop
        
        rstOrigem.MoveNext
        DoEvents
    Loop

    Set Scripts = Nothing
    rstCampos.Close
    rstOrigem.Close
    db.Close

    carregarCamposValores = left(tmpScript, Len(tmpScript) - 1)

End Function

Function carregarCamposNomes(pRepositorio As String) As String
Dim Scripts As New clsConexaoNfeCte
Dim db As DAO.Database: Set db = CurrentDb
Dim rstCampos As DAO.Recordset
Dim tmpScript As String
Dim Tmp As String

    '' 1. cabecalho
    Set rstCampos = db.OpenRecordset(Replace(Scripts.SelectCamposNomes, "pRepositorio", pRepositorio))
    Do While Not rstCampos.EOF
        tmpScript = tmpScript & rstCampos.Fields("campo").value & ","
        rstCampos.MoveNext
        DoEvents
    Loop

    Set Scripts = Nothing
    rstCampos.Close
    db.Close

    carregarCamposNomes = left(tmpScript, Len(tmpScript) - 1)

End Function

''' #20210823_qryUpdateNumPed_CompraNF
'Sub CadastroDeComprasComControleDeNumeroDePedidos()
'
'Dim qryDadosGerais_Select_ChvAcesso_RegistrosValidosProcessados As String: qryDadosGerais_Select_ChvAcesso_RegistrosValidosProcessados = "Select Distinct ChvAcesso from tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido) = 1) AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1) AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0)) AND tblDadosConexaoNFeCTe.ChvAcesso IS NOT NULL;"
'
''' BANCO_ORIGEM
'Dim contador As Long
'Dim sqlItens As String
'Dim sqlScript As String
'Dim sqlScriptTemplate As String: sqlScriptTemplate = _
'        "INSERT INTO tblCompraNF (ChvAcesso_CompraNF,NumPed_CompraNF) " & _
'        "SELECT ChvAcesso, NumPed FROM (VALUES ('sqlItens',contador)) AS tmp(ChvAcesso, NumPed) " & _
'        "LEFT JOIN tblCompraNF ON tblCompraNF.ChvAcesso_CompraNF = tmp.ChvAcesso WHERE tblCompraNF.ChvAcesso_CompraNF IS NULL;"
'
'Dim db As DAO.Database: Set db = CurrentDb
'Dim rstChvAcesso As DAO.Recordset: Set rstChvAcesso = db.OpenRecordset(qryDadosGerais_Select_ChvAcesso_RegistrosValidosProcessados)
'
''' BANCO_DESTINO
'Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
'Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
'Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
'Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
'Dim dbDestino As New Banco
'
'dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
'
'    Do While Not rstChvAcesso.EOF
'
'        '' CONTADOR
'        dbDestino.SqlSelect "SELECT max(NumPed_CompraNF)+1 as contador from tblCompraNF"
'        contador = IIf(IsNull(dbDestino.rs.Fields("contador").value), 1, dbDestino.rs.Fields("contador").value)
'
'        '' CADASTRO
'        sqlScript = Replace(Replace(sqlScriptTemplate, "sqlItens", rstChvAcesso.Fields("ChvAcesso").value), "contador", Format(contador, "000000"))
'        Debug.Print rstChvAcesso.Fields("ChvAcesso").value
'        Debug.Print sqlScript
'
'        dbDestino.SqlExecute sqlScript
'
'        rstChvAcesso.MoveNext
'        'dbDestino.rs.Close
'        contador = contador + 1
'        sqlItens = ""
'        DoEvents
'    Loop
'
'rstChvAcesso.Close
'db.Close
'
'dbDestino.CloseConnection
'
'Debug.Print "Concluido!"
'
'End Sub
