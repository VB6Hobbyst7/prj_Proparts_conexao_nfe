Attribute VB_Name = "modConexaoNfeCte"
Option Compare Database


'' 02.01.CARREGAR COMPRAS ANTES DE ENVIAR PARA O SERVIDOR
Function proc_01()
On Error GoTo adm_Err
Dim s As New clsConexaoNfeCte

'' #ANALISE_DE_PROCESSAMENTO

    s.carregar_DadosGerais
    
    
        
    
            
    MsgBox "Fim!", vbOKOnly + vbExclamation, "carregarComprasItens"
    

adm_Exit:
    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit

End Function


''' 03.ENVIAR DADOS PARA SERVIDOR
'Function enviarComprasParaServidor()
'On Error GoTo adm_Err
'
''' VARIAVEL DE PARAMETRO
'Dim pDestino As String: pDestino = "tblCompraNF"
'
''' ---------------------
''' VARIAVEIS GERAIS
''' ---------------------
'
''' LISTAGEM DE CAMPOS DA TABELA ORIGEM/DESTINO
'Dim strCampo As String
'
''' SCRIPT
'Dim tmpScript As String
'Dim tmp As String
'
''' ---------------------
''' BANCO LOCAL
''' ---------------------
'Dim db As dao.Database: Set db = CurrentDb
'Dim rstOrigem As dao.Recordset
'
''' ---------------------
''' BANCO DESTINO
''' ---------------------
'
'Dim strUsuarioNome As String
'Dim strUsuarioSenha As String
'Dim strOrigem As String
'Dim strBanco As String
'
'Dim dbDestino As New Banco
'
'Dim retVal As Variant: retVal = MsgBox("Deseja enviar compras para o servidor?", vbQuestion + vbYesNo, "ADM_enviarComprasParaServidor")
'
'    If retVal = vbYes Then
'        strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
'        strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
'        strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
'        strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
'
'        dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
'        dbDestino.SqlSelect "SELECT * FROM " & pDestino
'
'        tmpScript = "Insert into " & pDestino & " ("
'
'        '' 1. cabeçalho
'        Dim rstCampos As dao.Recordset
'        Set rstCampos = db.OpenRecordset("Select distinct campo,formatacao,valorPadrao from tblOrigemDestino where tblOrigemDestino.tabela = '" & pDestino & "'")
'        Do While Not rstCampos.EOF
'            tmpScript = tmpScript & rstCampos.Fields("campo").value & ","
'            rstCampos.MoveNext
'            DoEvents
'        Loop
'        tmpScript = left(tmpScript, Len(tmpScript) - 1) & ") values ("
'
'
'        '' BANCO LOCAL
'        Set rstOrigem = db.OpenRecordset("Select * from " & pDestino)
'        Do While Not rstOrigem.EOF
'            tmp = ""
'
'            '' LISTAGEM DE CAMPOS
'            rstCampos.MoveFirst
'            Do While Not rstCampos.EOF
'
'                '' CRIAR SCRIPT DE INCLUSÃO DE DADOS NA TABELA DESTINO
'                '' 2. campos x formatação
'
'                If rstCampos.Fields("formatacao").value = "opTexto" Then
'                    tmp = tmp & "'" & rstOrigem.Fields(rstCampos.Fields("campo").value).value & "',"
'
'                ElseIf rstCampos.Fields("formatacao").value = "opNumero" Or rstCampos.Fields("formatacao").value = "opMoeda" Then
'                    tmp = tmp & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", rstOrigem.Fields(rstCampos.Fields("campo").value).value, rstCampos.Fields("valorPadrao").value) & ","
'
'                ElseIf rstCampos.Fields("formatacao").value = "opTime" Or rstCampos.Fields("formatacao").value = "opData" Then
'                    tmp = tmp & "'" & IIf((rstOrigem.Fields(rstCampos.Fields("campo").value).value) <> "", rstOrigem.Fields(rstCampos.Fields("campo").value).value, rstCampos.Fields("valorPadrao").value) & "',"
'
'                End If
'
'                rstCampos.MoveNext
'                DoEvents
'            Loop
'
'            '' BANCO DESTINO
'            tmp = left(tmp, Len(tmp) - 1) & ")"
'
'            'Debug.Print tmpScript & tmp
'
'            rstOrigem.MoveNext
'
'            dbDestino.SqlExecute tmpScript & tmp
'
'            DoEvents
'        Loop
'
'        dbDestino.CloseConnection
'        db.Close: Set db = Nothing
'
'
'        MsgBox "Fim!", vbOKOnly + vbExclamation, "enviarComprasParaServidor"
'
'    End If
'
'adm_Exit:
'    Exit Function
'
'adm_Err:
'    MsgBox Error$
'    Resume adm_Exit
'
'
'End Function


''' 02.01.CARREGAR COMPRAS ANTES DE ENVIAR PARA O SERVIDOR
'Function carregarComprasItens()
'On Error GoTo adm_Err
'Dim strProcessamento As String: strProcessamento = "tblCompraNFItem"
'Dim s As New clsConexaoNfeCte
'Dim t As Variant
'
''' #ANALISE_DE_PROCESSAMENTO
'Dim DT_PROCESSO As Date: DT_PROCESSO = Now()
'Dim contadorDeRegistros As Long: contadorDeRegistros = 1
'Dim count As Long
'
'Dim retVal As Variant: retVal = MsgBox("Deseja carregar as compras com base no processamento de dados gerais ?", vbQuestion + vbYesNo, "carregarComprasItens")
'
'    If retVal = vbYes Then
'
'        count = DCount("*", "tblDadosConexaoNFeCTe", "(((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0))")
'
'        '' #BARRA_PROGRESSO
'        SysCmd acSysCmdInitMeter, "Processando Itens...", count
'
'        '' #CARREGAR DADOS
'        For Each t In carregarParametros(qrySelectProcessamentoItensCompras)
'
'            '' #BARRA_PROGRESSO
'            SysCmd acSysCmdUpdateMeter, contadorDeRegistros
'
'            '' 01.PROCESSAMENTO DE DADOS VINDOS DO XML
'            s.processar_ComprasItens CStr(t)
'
'            '' 02.TRANSFERIR DADOS PROCESSADOS PARA A TABELA DESTINO TEMPORARIA - COMPRAS ITENS
'            s.TransferirDadosProcessados strProcessamento
'
'            '' #BARRA_PROGRESSO
'            contadorDeRegistros = contadorDeRegistros + 1
'            DoEvents
'        Next
'
'
'        '' STATUS DE CONCLUSAO
''        s.compras_atualizarItensCompras
'
'        '' #BARRA_PROGRESSO
'        SysCmd acSysCmdRemoveMeter
'
'        '' #ANALISE_DE_PROCESSAMENTO
'        statusFinal DT_PROCESSO, "Processamento - carregarComprasItens"
'
'        MsgBox "Fim!", vbOKOnly + vbExclamation, "carregarComprasItens"
'
'    End If
'
'adm_Exit:
'    Exit Function
'
'adm_Err:
'    MsgBox Error$
'    Resume adm_Exit
'
'End Function


''' 02.CARREGAR COMPRAS ANTES DE ENVIAR PARA O SERVIDOR
'Function carregarCompras()
'Dim strProcessamento As String: strProcessamento = "tblCompraNF"
'Dim s As New clsConexaoNfeCte
'Dim t As Variant
'
'On Error GoTo adm_Err
'Dim retVal As Variant: retVal = MsgBox("Deseja carregar as compras com base no processamento de dados gerais ?", vbQuestion + vbYesNo, "carregarCompras")
'
'    If retVal = vbYes Then
'
'        '' #CARREGAR DADOS
'        For Each t In Array(strProcessamento)
'
'            '' PROCESSAR APENAS ARQUIVOS VALIDOS
'            s.ProcessarArquivosXml CStr(t), carregarParametros(qrySelectProcessamentoPendente)
'
'            '' FORMATAR CAMPOS
'            s.FormatarCampos
'
'            '' #TRATAMENTO
'            s.TratamentoDeCompras
'
'            '' #TRANSFERIR DADOS PROCESSADOS - COMPRAS
'            s.TransferirDadosProcessados strProcessamento
'
'            '' ATUALIZAR CAMPOS DE COMPRAS
'            s.compras_atualizarCampos
'
'        Next
'
'        '' #VALIDAR_DADOS
'        criarConsultasParaTestes
'
'        MsgBox "Fim!", vbOKOnly + vbExclamation, "carregarCompras"
'
'    End If
'
'adm_Exit:
'    Exit Function
'
'adm_Err:
'    MsgBox Error$
'    Resume adm_Exit
'
'End Function

''' 01.CARREGAR DADOS GERAIS - CONCLUIDO
'Function carregarDadosGerais()
'On Error GoTo adm_Err
'
'Dim strProcessamento As String: strProcessamento = "tblDadosConexaoNFeCTe"
'Dim s As New clsConexaoNfeCte
'Dim t As Variant
'
'Dim retVal As Variant: retVal = MsgBox("Deseja carregar os dados gerais dos arquivos XML ?", vbQuestion + vbYesNo, "carregarDadosGerais")
'
'    If retVal = vbYes Then
'
'        '' #LIMPAR_BASE_DE_TESTES
'        Application.CurrentDb.Execute "DELETE FROM tblDadosConexaoNFeCTe"
'        Application.CurrentDb.Execute "DELETE FROM tblCompraNF"
'        Application.CurrentDb.Execute "DELETE FROM tblCompraNFItem"
'
'        '' #CARREGAR DADOS
'        For Each t In Array(strProcessamento)
'
'            '' #PROCESSAMENTO DE ARQUIVO - ENVIO DE DADOS PARA tblProcessamento
'            s.ProcessarArquivosXml CStr(t), GetFilesInSubFolders(pegarValorDoParametro(qryParametros, strCaminhoDeColeta))
'
'            '' FORMATAR CAMPOS
'            s.FormatarCampos
'
'            '' #TRANSFERIR DADOS PROCESSADOS - DADOS GERAIS - ENVIO DE DADOS PARA tblDadosConexaoNFeCTe
'            s.TransferirDadosProcessados strProcessamento
'
'            '' #TRATAMENTO
'            s.TratamentoDeDadosGerais
'
'            '' #ARQUIVOS - GERAR ARQUIVOS
'            s.CriarTipoDeArquivo opFlagLancadaERP
'            s.CriarTipoDeArquivo opManifesto
'
'        Next
'
'        MsgBox "Fim!", vbOKOnly + vbExclamation, "carregarDadosGerais"
'
'    End If
'
'adm_Exit:
'    Exit Function
'
'adm_Err:
'    MsgBox Error$
'    Resume adm_Exit
'
'End Function

''' #ADMINISTRACAO
'Sub ADM_criarTabelas()
'
'    ''tblCompraNF
'    Application.CurrentDb.Execute createTable("tblCompraNF")
'    Application.CurrentDb.Execute createTable("tblCompraNFItem")
'
'End Sub

'' #ADMINISTRACAO - RESPONSAVEL POR TRAZER OS DADOS DO SERVIDOR PARA AUXILIO NO PROCESSAMENTO. QUANDO NECESSARIO
Function ADM_carregarDadosDoServidor()
On Error GoTo adm_Err
Dim retVal As Variant: retVal = MsgBox("Deseja carregar dados do servidor?", vbQuestion + vbYesNo, "ADM_carregarDadosDoServidor")

    If retVal = vbYes Then
    
        '' NATUREZA DE OPERAÇÃO
        Application.CurrentDb.Execute "Delete from tmpNatOp"
        ImportarDados "tblNatOp", "tmpNatOp"
        
        '' CADASTRO DE EMPRESA
        Application.CurrentDb.Execute "Delete from tmpEmpresa"
        ImportarDados "tblEmpresa", "tmpEmpresa"
        
        '' CADASTRO DE CLIENTES
        Application.CurrentDb.Execute "Delete from tmpClientes"
        ImportarDados "Clientes", "tmpClientes"
        
        MsgBox "Fim!", vbOKOnly + vbExclamation, "ADM_carregarDadosDoServidor"
    
    End If

adm_Exit:
    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit
    
End Function


