Attribute VB_Name = "00_Info"
Option Compare Database


'tmpClientes
'tmpEmpresa
'tmpGradeProdutos
'tmpNatOp
'tmpProdutos




'' LIMPAR TODA A BASE DE DADOS
Public Const dataBaseClear As Boolean = True

'' REPROCESSAR ARQUIVOS PENDENTES
Public Const dataBaseReplay As Boolean = False

'' EXPORTAR DADOS PARA SERVIDOR
Public Const dataBaseExportarDados As Boolean = False

'' PROCESSAMENTO DE ARQUIVOS
Public Const dataBaseTratamentoDeArquivos As Boolean = False
Public Const dataBaseGerarLancamentoManifesto As Boolean = False


Sub teste_FuncionamentoGeralDeProcessamentoDeArquivos()
Dim strCaminhoAcoes As String: strCaminhoAcoes = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColetaAcoes'")
    
    ''==================================================
    '' REPOSITORIO GERAL
    ''==================================================

    '' LIMPAR TODA A BASE DE DADOS
    If dataBaseClear Then
    
        '' Limpar toda a base de dados
        Application.CurrentDb.Execute "Delete from tblDadosConexaoNFeCTe"

        '' Limpar repositorio de itens de compras
        Application.CurrentDb.Execute _
                "Delete from tblCompraNFItem"
    
        '' Limpar repositorio de compras
        Application.CurrentDb.Execute _
                "Delete from tblCompraNF"

        '' Carregar todos os arquivos para processamento.
        ProcessarDadosGerais
        
    Else
        
        '' Carregar todos os arquivos para processamento.
        ProcessarDadosGerais
    
    
    End If

    ''==================================================
    '' REPOSITORIOS DE COMPRAS
    ''==================================================
    
    '' REPROCESSAR ARQUIVOS VALIDOS
    If dataBaseReplay Then
    
        '' Ajustar marcação de registro
        Application.CurrentDb.Execute _
            "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado=0 WHERE tblDadosConexaoNFeCTe.registroValido=1 AND tblDadosConexaoNFeCTe.ID_Tipo>0"
        
        '' Limpar repositorio de itens de compras
        Application.CurrentDb.Execute _
                "Delete from tblCompraNFItem"
    
        '' Limpar repositorio de compras
        Application.CurrentDb.Execute _
                "Delete from tblCompraNF"

        '' Processamento de arquivos pendentes da pasta de coleta.
        processarArquivosPendentes
            
    Else
    
        '' Processamento de arquivos pendentes da pasta de coleta.
        processarArquivosPendentes
    
    End If


    ''==================================================
    '' EXPORTAR DADOS PARA O SERVIDOR
    ''==================================================

    '' EXPORTAÇÃO DE DADOS
    If dataBaseExportarDados Then _
            CadastroDeComprasEmServidor

    ''==================================================
    '' PROCESSAMENTO DE ARQUIVOS
    ''==================================================

    '' #### TRANSFERENCIAS DE ARQUIVOS
    If dataBaseTratamentoDeArquivos Then _

        '' Transferir Arquivos Validos para pasta de processados
        tratamentoDeArquivosValidos
    
        '' Transferir Arquivos Invalidos para pasta de Expurgo
        tratamentoDeArquivosInvalidos

    End If

    '' #### GERAR ARQUIVOS DE LANÇAMENTO E MANIFESTO
    If dataBaseGerarLancamentoManifesto Then
    
        '' LANÇAMENTO
        gerarArquivosJson opFlagLancadaERP, , strCaminhoAcoes
    
        '' MANIFESTO
        gerarArquivosJson opManifesto, , strCaminhoAcoes
        
    End If
    
Debug.Print "### Concluido! - testeDeFuncionamentoGeral"
TextFile_Append CurrentProject.path & "\" & strLog(), "Concluido! - testeDeFuncionamentoGeral"

End Sub



Sub TransferirProcessamentoParaRepositorios(pChave As String)
    
Dim qryProcessamento_Select_CompraComItens As String: _
    qryProcessamento_Select_CompraComItens = _
    "SELECT DISTINCT tblProcessamento.NomeTabela " & _
        "   ,tblProcessamento.pk " & _
        "   ,tblProcessamento.NomeCampo " & _
        "   ,tblProcessamento.valor " & _
        "   ,tblProcessamento.formatacao " & _
        "FROM tblParametros " & _
        "INNER JOIN tblProcessamento ON (tblParametros.ValorDoParametro = tblProcessamento.NomeCampo) AND (tblParametros.TipoDeParametro = tblProcessamento.NomeTabela) " & _
        "WHERE (((tblProcessamento.pk) LIKE 'pChave*') AND ((tblProcessamento.NomeCampo) IS NOT NULL)) " & _
        "ORDER BY tblProcessamento.NomeTabela,tblProcessamento.pk"

Dim qryCompras_Insert_Processamento As String: _
    qryCompras_Insert_Processamento = "INSERT INTO pRepositorio (strCamposNomes) SELECT strCamposValores  "

Dim DadosGerais As New clsConexaoNfeCte
Dim db As DAO.Database: Set db = CurrentDb
Dim rstProcessamento As DAO.Recordset: Set rstProcessamento = db.OpenRecordset(Replace(qryProcessamento_Select_CompraComItens, "pChave", pChave))
Dim rstRegistros As DAO.Recordset: Set rstRegistros = db.OpenRecordset("select distinct tmp.pk from (" & Replace(qryProcessamento_Select_CompraComItens, "pChave", pChave) & ") as tmp order by tmp.pk;")
Dim rstFiltered As DAO.Recordset

Dim item As Variant

Dim pRepositorio As String
Dim strChave As String
Dim strCamposNomes As String
Dim strCamposValores As String
Dim strCompras_Insert_Processamento As String
    
    '' REGISTROS
    Do While Not rstRegistros.EOF
        
        '' VARIAVEIS
        strCamposNomes = ""
        strCamposValores = ""
        strCompras_Insert_Processamento = qryCompras_Insert_Processamento
        
        '' SELEÇÃO DE ITEM
        strChave = rstRegistros.Fields("pk").value
        rstProcessamento.Filter = "pk = '" & strChave & "'"
            
        '' PROCESSAMENTO DO ITEM
        Set rstFiltered = rstProcessamento.OpenRecordset
        Do While Not rstFiltered.EOF
        
            pRepositorio = rstFiltered.Fields("NomeTabela").value

            '' NOME DAS COLUNAS
            strCamposNomes = strCamposNomes & rstFiltered.Fields("NomeCampo").value & ","
            
            '' VALOR DAS COLUNAS
            If rstFiltered.Fields("formatacao").value = "opTexto" Then
                strCamposValores = strCamposValores & "'" & rstFiltered.Fields("Valor").value & "',"
                
            ElseIf rstFiltered.Fields("formatacao").value = "opNumero" Or rstFiltered.Fields("formatacao").value = "opMoeda" Then
                strCamposValores = strCamposValores & rstFiltered.Fields("Valor").value & ","
            
            ElseIf rstFiltered.Fields("formatacao").value = "opTime" Then
                strCamposValores = strCamposValores & "'" & Format(rstFiltered.Fields("Valor").value, DATE_TIME_FORMAT) & "',"
            
            ElseIf rstFiltered.Fields("formatacao").value = "opData" Then
                strCamposValores = strCamposValores & "'" & Format(rstFiltered.Fields("Valor").value, DATE_FORMAT) & "',"
            
            End If
       
            rstFiltered.MoveNext
            DoEvents
        Loop
        
        '' EXIT SCRIPT
        strCamposNomes = left(strCamposNomes, Len(strCamposNomes) - 1)
        strCamposValores = left(strCamposValores, Len(strCamposValores) - 1)
        
        ''Debug.Print
        Application.CurrentDb.Execute Replace(Replace(Replace(strCompras_Insert_Processamento, "strCamposNomes", strCamposNomes), "strCamposValores", strCamposValores), "pRepositorio", pRepositorio)
        
        rstRegistros.MoveNext
        DoEvents
    Loop
        
    '' registroProcessado
    Application.CurrentDb.Execute Replace(DadosGerais.UpdateProcessamentoConcluido, "strChave", strChave)
        
        
'' Cleanup
rstFiltered.Close
rstProcessamento.Close
rstRegistros.Close

Set rstFiltered = Nothing
Set rstProcessamento = Nothing
Set rstRegistros = Nothing

End Sub

