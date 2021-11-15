Attribute VB_Name = "02_modMSSQL"
Public Const DATE_TIME_FORMAT               As String = "yyyy/mm/dd hh:mm:ss"
Public Const DATE_FORMAT                    As String = "yyyy/mm/dd"

'' #20210823_CadastroDeComprasEmServidor
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
Dim qryCompras_Insert_Compras As String
Dim qryComprasItens_Update_IDCompraNF As String
Dim tmp As String

Dim contador As Long: contador = 1

    '' BANCO_DESTINO
    dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer

    '' CADASTRO
    Do While Not rstChvAcesso.EOF
        
        '' CONTADOR
        dbDestino.SqlSelect "SELECT max(NumPed_CompraNF)+1 as contador from tblCompraNF"
        contador = IIf(IsNull(dbDestino.rs.Fields("contador").value), 1, dbDestino.rs.Fields("contador").value)
        
        '' CADASTRO DE COMPRAS
        qryCompras_Insert_Compras = Replace(carregarScript_Insert("tblCompraNF", rstChvAcesso.Fields("ChvAcesso_CompraNF").value), "strNumPed_CompraNF", contador)
        dbDestino.SqlExecute qryCompras_Insert_Compras
        
        '' RELACIONAR ITENS DE COMPRAS COM COMPRAS J� CADASTRADAS NO SERVIDOR
        dbDestino.SqlSelect "SELECT ChvAcesso_CompraNF,ID_CompraNF FROM tblCompraNF where ChvAcesso_CompraNF = '" & rstChvAcesso.Fields("ChvAcesso_CompraNF").value & "';"
        qryComprasItens_Update_IDCompraNF = Replace(Replace(Scripts.UpdateComprasItens_IDCompraNF, "strChave", rstChvAcesso.Fields("ChvAcesso_CompraNF").value), "strID_Compra", dbDestino.rs.Fields("ID_CompraNF").value)
        If Not dbDestino.rs.EOF Then
            
            '' #20210823_qryUpdateNumPed_CompraNF
            Application.CurrentDb.Execute qryComprasItens_Update_IDCompraNF
            
            '' CADASTRO DE ITENS DE COMPRAS
            
            '' .................
        End If
                
                
        '' MUDAR STATUS DO REGISTRO
        Application.CurrentDb.Execute Replace(Scripts.compras_atualizarEnviadoParaServidor, "strChave", rstChvAcesso.Fields("chvAcesso_CompraNF").value)

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


Sub teste__carregarScript_Insert()

'' 01
'Debug.Print carregarScript_Insert("tblCompraNF", "32210368365501000296550000000638811001361356")

'' 02
Debug.Print carregarScript_Insert("tblCompraNFItem", "32210368365501000296550000000638791001361285")

'' 23
'Debug.Print carregarScript_Insert("tblCompraNFItem", "32210368365501000296550000000638811001361356")

End Sub

Function carregarScript_Insert(pRepositorio As String, pChvAcesso As String) As String

Dim strCamposNomes As String: _
    strCamposNomes = carregarCamposNomes(pRepositorio)

Dim strCamposNomesTmp As String: _
    strCamposNomesTmp = Replace(strCamposNomes, "_" & right(pRepositorio, Len(pRepositorio) - 3), "")

Dim strCamposValores As String: _
    strCamposValores = carregarCamposValores(pRepositorio, pChvAcesso)

Dim tmpScript As String: _
    tmpScript = "INSERT INTO " & pRepositorio & " ( " & strCamposNomes & " ) SELECT " & strCamposNomesTmp & " FROM ( VALUES " & strCamposValores & " ) AS TMP ( " & strCamposNomesTmp & " ) LEFT JOIN " & pRepositorio & " ON " & pRepositorio & ".ChvAcesso_CompraNF = tmp.ChvAcesso WHERE " & pRepositorio & ".ChvAcesso_CompraNF IS NULL;"
    '"INSERT INTO " & pRepositorio & " ( " & strCamposNomes & " ) SELECT " & strCamposNomesTmp & " FROM ( VALUES ( " & strCamposValores & " ) ) AS TMP ( " & strCamposNomesTmp & " ) LEFT JOIN " & pRepositorio & " ON " & pRepositorio & ".ChvAcesso_CompraNF = tmp.ChvAcesso WHERE " & pRepositorio & ".ChvAcesso_CompraNF IS NULL;"
    
    carregarScript_Insert = tmpScript

End Function

Function carregarCamposValores(pRepositorio As String, pChvAcesso As String) As String

'Dim pRepositorio As String: pRepositorio = "tblCompraNFItem"
'Dim pChvAcesso As String: pChvAcesso = "32210368365501000296550000000638791001361285"

Dim Scripts As New clsConexaoNfeCte
Dim db As DAO.Database: Set db = CurrentDb
Dim rstCampos As DAO.Recordset: Set rstCampos = db.OpenRecordset(Replace(Scripts.SelectCamposNomes, "pRepositorio", pRepositorio))
Dim rstOrigem As DAO.Recordset

Dim tmpScript As String
Dim tmpValidarCampo As String: tmpValidarCampo = right(pRepositorio, Len(pRepositorio) - 3)

Dim sqlOrigem As String: sqlOrigem = _
    "Select * from (" & Replace(Scripts.SelectRegistroValidoPorcessado, "pRepositorio", pRepositorio) & ") as tmpRepositorio where tmpRepositorio.ChvAcesso_CompraNF = '" & pChvAcesso & "'"
    
    Set rstOrigem = db.OpenRecordset(sqlOrigem)
    
    rstOrigem.MoveLast
    rstOrigem.MoveFirst
    Do While Not rstOrigem.EOF
        tmpScript = tmpScript & "("
        
        '' LISTAGEM DE CAMPOS
        rstCampos.MoveFirst
        Do While Not rstCampos.EOF
    
            '' CRIAR SCRIPT DE INCLUSAO DE DADOS NA TABELA DESTINO
            '' 2. campos x formatacao
            If InStr(rstCampos.Fields("campo").value, tmpValidarCampo) Then
    
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
        
        tmpScript = left(tmpScript, Len(tmpScript) - 1) & "),"
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
Dim tmp As String

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
