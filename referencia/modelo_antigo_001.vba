'' #ADMINISTRACAO - RESPONSAVEL POR TRAZER OS DADOS DO SERVIDOR PARA AUXILIO NO PROCESSAMENTO. QUANDO NECESSARIO
Public Function ADM_carregarDadosDoServidor()
On Error GoTo adm_Err
Dim retVal As Variant: retVal = MsgBox("Deseja carregar dados do servidor?", vbQuestion + vbYesNo, "ADM_carregarDadosDoServidor")

    If retVal = vbYes Then
    
        '' NATUREZA DE OPERAÇÃO
        Application.CurrentDb.Execute "Delete from tmpNatOp"
        ImportarDados "tblNatOp", "tmpNatOp"
        
        '' CADASTRO DE EMPRESA
        Application.CurrentDb.Execute "Delete from tmpEmpresa"
        ImportarDados "tblEmpresa", "tmpEmpresa"
        
'        '' CADASTRO DE CLIENTES
'        Application.CurrentDb.Execute "Delete from tmpClientes"
'        ImportarDados "Clientes", "tmpClientes"
        
        MsgBox "Fim!", vbOKOnly + vbExclamation, "ADM_carregarDadosDoServidor"
    
    End If

adm_Exit:
    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit
    
End Function

Private Sub ImportarDados(pOrigem As String, pDestino As String)

'' #BANCO_ORIGEM
Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
Dim tmpOrigem As String: tmpOrigem = "Select * from " & pOrigem

Dim dboOrigem As New Banco: dboOrigem.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
dboOrigem.SqlSelect tmpOrigem

'' #BANCO_LOCAL
Dim db As DAO.Database: Set db = CurrentDb
Dim tmpDestino As String: tmpDestino = "Select * from " & pDestino
Dim rstDestino As DAO.Recordset: Set rstDestino = db.OpenRecordset(tmpDestino)
Dim rstOrigem As DAO.Recordset

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()
Dim fld As Variant, t As Variant

'' #BARRA_PROGRESSO
Dim contadorDeRegistros As Long: contadorDeRegistros = 1
SysCmd acSysCmdInitMeter, "Transferindo " & pOrigem & " ...", dboOrigem.rs.RecordCount

Do While Not dboOrigem.rs.EOF

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdUpdateMeter, contadorDeRegistros

    '' listar campos da tabela
    For Each fld In rstDestino.Fields
        If t = 0 Then rstDestino.AddNew: t = 1
        rstDestino.Fields(fld.Name).value = dboOrigem.rs.Fields(fld.Name).value
        DoEvents
    Next
    rstDestino.Update
    t = 0
    
    '' #BARRA_PROGRESSO
    contadorDeRegistros = contadorDeRegistros + 1
    dboOrigem.rs.MoveNext
    DoEvents
Loop

'' #BARRA_PROGRESSO
SysCmd acSysCmdRemoveMeter

'dbDestino.CloseConnection
db.Close: Set db = Nothing

'' #ANALISE_DE_PROCESSAMENTO
statusFinal DT_PROCESSO, "Processamento - ImportarDados ( " & pDestino & " )"

End Sub