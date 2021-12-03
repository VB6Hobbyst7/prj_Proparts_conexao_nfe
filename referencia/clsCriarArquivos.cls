VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCriarArquivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Enum enumTipoArquivo
    opFlagLancadaERP = 0
    opManifesto = 1
End Enum

'' #00. CRIAÇÃO DE ARQUIVOS JSON
Public Sub criarArquivoJson(pArquivo As enumTipoArquivo, qryDados As String, Optional caminhoDeProcessados As String)
Dim db As DAO.Database: Set db = CurrentDb
Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(qryDados)
Dim dados As New clsConexaoNfeCteDados ''Set dados = New clsConexaoNfeCteDados

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()
Dim contadorDeRegistros As Long: contadorDeRegistros = 1
Dim nomeArquivo As String
    
    '' #BARRA_PROGRESSO
    SysCmd acSysCmdInitMeter, "Processando Json...", rst.RecordCount

    '' CRIAÇÃO DE ARQUIVOS
    Do While Not rst.EOF
    
        '' #BARRA_PROGRESSO
        SysCmd acSysCmdUpdateMeter, contadorDeRegistros
    
        '' #CARREGAR_DADOS_ARQUIVO
        dados.ChvAcesso = rst.Fields("ChvAcesso").value
        dados.dhEmi = rst.Fields("dhEmi").value
        dados.codUsuarioErp = DLookup("ValorDoParametro", "tblParametros", "TipoDeParametro='UsuarioErpCod'")
        dados.nomeUsuarioErp = DLookup("ValorDoParametro", "tblParametros", "TipoDeParametro='UsuarioErpNome'")
        
        '' CAMINHO DE SAIDA DO ARQUIVO
        dados.caminhoDeProcessados = CreateDir(DLookup("ValorDoParametro", "tblParametros", "TipoDeParametro='caminhoDeColetaProcessados'"))
        If caminhoDeProcessados <> "" Then dados.caminhoDeProcessados = caminhoDeProcessados
                        
        Select Case pArquivo
            Case opFlagLancadaERP
                dados.Comando = DLookup("ValorDoParametro", "tblParametros", "TipoDeParametro='Comando'")
                dados.criarERP
                nomeArquivo = "FlagLancadaERP"
            Case opManifesto
                dados.codTipoEvento = DLookup("ValorDoParametro", "tblParametros", "TipoDeParametro='codTipoEvento'")
                dados.criarManifesto
                nomeArquivo = "Manifesto"
            Case Else
        End Select
        
        rst.MoveNext
        DoEvents
        
        '' #BARRA_PROGRESSO
        contadorDeRegistros = contadorDeRegistros + 1
    Loop
    
    '' #BARRA_PROGRESSO
    SysCmd acSysCmdRemoveMeter
    
    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "Processamento - CriarTipoDeArquivo (" & nomeArquivo & ")"
    
Cleanup:
    db.Close: Set db = Nothing
    
End Sub

