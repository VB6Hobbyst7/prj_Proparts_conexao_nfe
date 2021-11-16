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

'' BANCO_LOCAL
Dim Scripts As New clsConexaoNfeCte
Dim db As DAO.Database: Set db = CurrentDb
Dim rstChvAcesso As DAO.Recordset: Set rstChvAcesso = db.OpenRecordset(Replace(Scripts.SelectRegistroValidoPorcessado, "pRepositorio", "tblCompraNF"))

''' #AILTON - TESTES
'Dim TMP2 As String: TMP2 = "select * from (" & Replace(Scripts.SelectRegistroValidoPorcessado, "pRepositorio", "tblCompraNF") & ") as tmp where tmp.ChvAcesso_CompraNF = '42210300634453001303570010001139451001171544'"
'Dim rstChvAcesso As DAO.Recordset: Set rstChvAcesso = db.OpenRecordset(TMP2)

Dim qryCompras_Insert_Compras As String
Dim qryComprasItens_Update_IDCompraNF As String

'' CONTROLE DE "NumPed_CompraNF" RELACIONADO COM REGISTROS DO SERVIDOR
Dim contador As Long: contador = 1

'' CONTROLE DE REPOSITORIOS x CHAVES DE ACESSO
Dim item As Variant
Dim pChvAcesso As String
Dim pRepositorio As String

'' SCRIPT DE INCLUSÃO DE DADOS NO SERVIDOR
Dim tmpCamposNomes As String
Dim strCamposNomes As String
Dim strCamposNomesTmp As String
Dim strRepositorio As String
Dim tmpScript As String: _
    tmpScript = "INSERT INTO pRepositorio ( strCamposNomes ) SELECT strCamposTmp FROM ( VALUES strCamposValores ) AS TMP ( strCamposTmp ) LEFT JOIN pRepositorio ON pRepositorio.ChvAcesso_CompraNF = tmp.ChvAcesso WHERE pRepositorio.ChvAcesso_CompraNF IS NULL;"

Dim tmpScriptItens As String: _
    tmpScriptItens = "INSERT INTO pRepositorio ( strCamposNomes ) SELECT strCamposTmp FROM ( VALUES strCamposValores ) AS TMP ( strCamposTmp );"


'' VALIDAR CONCILIAÇÃO
Dim TMP As String

    '' BANCO_DESTINO
    dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdInitMeter, "Exportação...", rstChvAcesso.RecordCount

    '' CADASTRO
    Do While Not rstChvAcesso.EOF
        pRepositorio = "tblCompraNF"
        pChvAcesso = rstChvAcesso.Fields("ChvAcesso_CompraNF").value

        tmpCamposNomes = carregarCamposNomes(pRepositorio)
        strCamposNomes = Replace(tmpScript, "strCamposNomes", tmpCamposNomes)
        strCamposNomesTmp = Replace(Replace(tmpScript, "strCamposNomes", tmpCamposNomes), "strCamposTmp", Replace(tmpCamposNomes, "_" & right(pRepositorio, Len(pRepositorio) - 3), ""))
        strRepositorio = Replace(strCamposNomesTmp, "pRepositorio", pRepositorio)

        '' CONTADOR
        dbDestino.SqlSelect "SELECT max(NumPed_CompraNF)+1 as contador from tblCompraNF"
        contador = IIf(IsNull(dbDestino.rs.Fields("contador").value), 1, dbDestino.rs.Fields("contador").value)

        '' CADASTRO DE COMPRAS
        For Each item In carregarCamposValores(pRepositorio, pChvAcesso)
            TMP = Replace(Replace(strRepositorio, "strCamposValores", item), "strNumPed_CompraNF", contador)
            Debug.Print TMP
            dbDestino.SqlExecute TMP
        Next item

        '' RELACIONAR ITENS DE COMPRAS COM COMPRAS JÁ CADASTRADAS NO SERVIDOR
        dbDestino.SqlSelect "SELECT ChvAcesso_CompraNF,ID_CompraNF FROM tblCompraNF where ChvAcesso_CompraNF = '" & pChvAcesso & "';"
        qryComprasItens_Update_IDCompraNF = Replace(Replace(Scripts.UpdateComprasItens_IDCompraNF, "strChave", pChvAcesso), "strID_Compra", dbDestino.rs.Fields("ID_CompraNF").value)
        If Not dbDestino.rs.EOF Then
            pRepositorio = "tblCompraNFItem"

            tmpCamposNomes = carregarCamposNomes(pRepositorio)
            strCamposNomes = Replace(tmpScriptItens, "strCamposNomes", tmpCamposNomes)
            strCamposNomesTmp = Replace(Replace(tmpScriptItens, "strCamposNomes", tmpCamposNomes), "strCamposTmp", Replace(tmpCamposNomes, "_" & right(pRepositorio, Len(pRepositorio) - 3), ""))
            strRepositorio = Replace(strCamposNomesTmp, "pRepositorio", pRepositorio)

            '' #20210823_qryUpdateNumPed_CompraNF
            Application.CurrentDb.Execute qryComprasItens_Update_IDCompraNF

            '' CADASTRO DE ITENS DE COMPRAS
            For Each item In carregarCamposValores(pRepositorio, pChvAcesso)
                TMP = Replace(Replace(strRepositorio, "strCamposValores", item), "strNumPed_CompraNF", contador)
                Debug.Print TMP
                dbDestino.SqlExecute TMP
            Next item

        End If

        '' MUDAR STATUS DO REGISTRO
        Application.CurrentDb.Execute Replace(Scripts.compras_atualizarEnviadoParaServidor, "strChave", rstChvAcesso.Fields("chvAcesso_CompraNF").value)

        rstChvAcesso.MoveNext
        contador = contador + 1

        '' #BARRA_PROGRESSO
        SysCmd acSysCmdUpdateMeter, contador

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

'' #BARRA_PROGRESSO
SysCmd acSysCmdRemoveMeter

Debug.Print "Concluido!"

End Sub

Function carregarCamposValores(pRepositorio As String, pChvAcesso As String) As Collection
Set carregarCamposValores = New Collection
Dim Scripts As New clsConexaoNfeCte

'Dim pRepositorio As String: pRepositorio = "tblCompraNFItem"
'Dim pChvAcesso As String: pChvAcesso = "32210368365501000296550000000638791001361285"

'' BANCO_LOCAL
Dim db As DAO.Database: Set db = CurrentDb

Dim rstCampos As DAO.Recordset: Set rstCampos = db.OpenRecordset(Replace(Scripts.SelectCamposNomes, "pRepositorio", pRepositorio))
Dim rstOrigem As DAO.Recordset

'' VALIDAR CONCILIAÇÃO
Dim tmpScript As String
Dim tmpValidarCampo As String: tmpValidarCampo = right(pRepositorio, Len(pRepositorio) - 3)

Dim sqlOrigem As String: sqlOrigem = _
    "Select * from (" & Replace(Scripts.SelectRegistroValidoPorcessado, "pRepositorio", pRepositorio) & ") as tmpRepositorio where tmpRepositorio.ChvAcesso_CompraNF = '" & pChvAcesso & "'"
    
    Set rstOrigem = db.OpenRecordset(sqlOrigem)
    
    rstOrigem.MoveLast
    rstOrigem.MoveFirst
    Do While Not rstOrigem.EOF
        tmpScript = "("
        
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
        
        carregarCamposValores.add left(tmpScript, Len(tmpScript) - 1) & ")"
        rstOrigem.MoveNext
        DoEvents
    Loop

    Set Scripts = Nothing
    rstCampos.Close
    rstOrigem.Close
    db.Close

End Function

Function carregarCamposNomes(pRepositorio As String) As String
Dim Scripts As New clsConexaoNfeCte

'' BANCO_LOCAL
Dim db As DAO.Database: Set db = CurrentDb
Dim rstCampos As DAO.Recordset

'' VALIDAR CONCILIAÇÃO
Dim tmpScript As String

    '' MONTAR STRING DE NOME DE COLUNAS
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



'' #20211110_0730
'' qryProcessamento_Select_CompraComItens
Function carregarCamposNomesProcessamento() As String

Dim colCadastros As New Collection

Dim Scripts As New clsConexaoNfeCte

'' BANCO_LOCAL
Dim db As DAO.Database: Set db = CurrentDb
Dim rstDados As DAO.Recordset
Dim rstCampos As DAO.Recordset

'' VALIDAR CONCILIAÇÃO
Dim tmpCampoNome As String
Dim tmpCampoValor As String

    Set rstDados = db.OpenRecordset("SELECT DISTINCT pk, NomeTabela FROM qryProcessamento_Select_CompraComItens")
    Do While Not rstDados.EOF

        '' MONTAR STRING DE NOME DE COLUNAS
        Set rstCampos = db.OpenRecordset(Replace("Select * from qryProcessamento_Select_CompraComItens where NomeTabela='pRepositorio';", "pRepositorio", rstDados.Fields("NomeTabela").value))
        Do While Not rstCampos.EOF
            
            If rstDados.Fields("NomeTabela").value = "tblCompraNFItem" And rstCampos.Fields("NomeCampo").value = "ChvAcesso_CompraNF" Then
                tmpCampoNome = tmpCampoNome & "ID_CompraNF_CompraNFItem,"
                GoTo puloCampoNome
            End If
            
            tmpCampoNome = tmpCampoNome & rstCampos.Fields("NomeCampo").value & ","
            
puloCampoNome:
            rstCampos.MoveNext
            DoEvents
        Loop
    
        tmpCampoNome = "INSERT INTO " & rstDados.Fields("NomeTabela").value & " (" & left(tmpCampoNome, Len(tmpCampoNome) - 1) & ") "
    
        rstCampos.MoveFirst
        Do While Not rstCampos.EOF
                  
            If rstDados.Fields("NomeTabela").value = "tblCompraNFItem" And rstCampos.Fields("NomeCampo").value = "ChvAcesso_CompraNF" Then
                tmpCampoValor = tmpCampoValor & "(SELECT ID_CompraNF FROM tblCompraNF where ChvAcesso_CompraNF = '" & rstCampos.Fields("valor").value & "'),"
                GoTo puloCampoValor
            End If
                  
            If rstCampos.Fields("formatacao").value = "opTexto" Then
                tmpCampoValor = tmpCampoValor & "'" & rstCampos.Fields("Valor").value & "',"
    
            ElseIf rstCampos.Fields("formatacao").value = "opNumero" Or rstCampos.Fields("formatacao").value = "opMoeda" Then
                tmpCampoValor = tmpCampoValor & IIf((rstCampos.Fields("Valor").value) <> "", rstCampos.Fields("Valor").value, 0) & ","
    
            ElseIf rstCampos.Fields("formatacao").value = "opTime" Then
                tmpCampoValor = tmpCampoValor & "'" & IIf((rstCampos.Fields("Valor").value) <> "", Format(rstCampos.Fields("Valor").value, DATE_TIME_FORMAT), "00:00:00") & "',"
    
            ElseIf rstCampos.Fields("formatacao").value = "opData" Then
                tmpCampoValor = tmpCampoValor & "'" & IIf((rstCampos.Fields("Valor").value) <> "", Format(rstCampos.Fields("Valor").value, DATE_FORMAT), Null) & "',"
    
            End If
            
puloCampoValor:
            rstCampos.MoveNext
            DoEvents
        Loop
    
        tmpCampoValor = "Select " & left(tmpCampoValor, Len(tmpCampoValor) - 1) & ";"
        
        colCadastros.add tmpCampoNome & tmpCampoValor
        tmpCampoNome = ""
        tmpCampoValor = ""

        rstDados.MoveNext
        DoEvents
    Loop

    CadastroDeCompra colCadastros

    Set Scripts = Nothing
    rstCampos.Close
    rstDados.Close
    db.Close

End Function



Sub CadastroDeCompra(colCadastros As Collection)

'' BANCO_DESTINO
Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
Dim dbDestino As New Banco
Dim i As Variant

'' Dim t As String: t = "INSERT INTO tblCompraNFItem (Item_CompraNFItem,ID_Prod_CompraNFItem,CFOP_CompraNFItem,BaseCalcICMSSubsTrib_CompraNFItem,BaseCalcIPI_CompraNFItem,DebICMS_CompraNFItem,DebIPI_CompraNFItem,ICMS_CompraNFItem,IPI_CompraNFItem,QtdFat_CompraNFItem,VUnt_CompraNFItem,VTot_CompraNFItem,VTotBaseCalcICMS_CompraNFItem,ID_CompraNF_CompraNFItem) SELECT 1,(select CodigoProd_Grade from tabGradeProdutos where CodigoForn_Grade='00.1918.117.006') as tmpID_Prod_CompraNFItem,'6152',00,4527.48,181.10,452.75,4.00,10.00,1.0000,4527.4818,4527.48,4527.48,(Select ID_CompraNF from tblCompraNF where ChvAcesso_CompraNF='32210268365501000296550000000637741001351624') as tmpPK;"

    '' BANCO_DESTINO
    dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer

    For Each i In colCadastros
        dbDestino.SqlExecute CStr(i)
    Next

End Sub




