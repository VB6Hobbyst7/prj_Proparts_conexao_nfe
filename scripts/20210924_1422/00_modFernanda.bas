Attribute VB_Name = "00_modFernanda"
Option Compare Database

''----------------------------
'' ### EXEMPLOS DE FUNÇÕES
''
'' 01. testeUnitario_carregarDadosGerais
'' 02. testeUnitario_carregarArquivosPendentes
'' 03. enviar_compras_para_servidor
'' 04. exemplos_criacao_arquivos_json
'' 99. FUNÇÃO_AUXILIAR: carregarDadosGerais(strArquivo As String)
'' 99. FUNÇÃO_AUXILIAR: carregarArquivosPendentes(strArquivo As String)
''----------------------------


'' ### TO-DO ###
''

'' #20210823_XML_CONTROLE
'' - Quando importar cada XML, precisa recortar o arquivo da pasta da empresa e colar dentro de uma pasta chama “Processados”, porém dentro de cada pasta de cada empresa, pois não podemos misturar os XML´s de cada empresa.
'' #20210823_XML_FORMULARIO
'' - Não encontrei um formulário com os XML´s que não foram processados e o motivo.
'' #20210823_VTotProd_CompraNF
'' ID_Prod_CompraNFItem


'' ### DONE ###
''
'' Consultas
'' #20210823_EXPORTACAO_LIMITE
'' #20210823_qryUpdateCFOP_PSC_PES -- FiltroCFOP
'' #20210823_qryUpdate_IDVD
'' #20210823_qryUpdateID_NatOp_CompraNF
'' #20210823_qryUpdateCFOP_FilCompra
'' #20210823_qryUpdate_ModeloDoc_CFOP
'' #20210823_qryUpdateFilCompraNF
'' #20210823_qryUpdateIdFornCompraNF
'' #20210823_qryUpdateNumPed_CompraNF
'' #20210823_qryUpdateSit_CompraNF


'' #20210823_XML_CONTROLE
Sub tratamentoDeArquivos()

Dim objFSO As Object: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim db As DAO.Database: Set db = Application.CurrentDb
Dim rstArquivos As DAO.Recordset
Dim strFileName As String
Dim strFilePath As String: strFilePath = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeProcessados'")
Dim qryArquivos As String: qryArquivos = _
    "SELECT CaminhoDoArquivo FROM  tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=2));"


CreateDir strFilePath


Set rstArquivos = db.OpenRecordset(sqlArquivos)
Do While Not rstArquivos.EOF
       
    strFileName = rstArquivos.Fields("CaminhoDoArquivo").value
    
    objFSO.CopyFile strFileName, strFilePath & getFileNameAndExt(rstArquivos.Fields("CaminhoDoArquivo").value)
    'If (Dir(strFileName) <> "") Then Kill strFilePath & strFileName
    
    rstArquivos.MoveNext
    DoEvents
Loop


db.Close: Set db = Nothing

End Sub



Sub testeGeral()

    Application.CurrentDb.Execute "Delete from tblCompraNFItem"
    Application.CurrentDb.Execute "Delete from tblCompraNF"
    Application.CurrentDb.Execute "Delete from tblDadosConexaoNFeCTe"
    
    testeUnitario_carregarDadosGerais
    testeUnitario_carregarArquivosPendentes
    
    Debug.Print "Concluido! - testeGeral"

End Sub


Sub enviar_compras_para_servidor()

    '' CADASTRO DE CABEÇALHO DE COMPRAS
    enviar_ComprasParaServidor "tblCompraNF"

    '' RELACIONAMENTO DE ID_COMPRAS COM CHAVES DE ACESSO CADASTRADAS DO SERVIDOR
    criarTabelaTemporariaParaRelacionarIdCompraComChvAcesso
    relacionarIdCompraComChvAcesso
    
    '' CADASTRO DE ITENS DE COMPRAS
    enviar_ComprasParaServidor "tblCompraNFItem"

    Debug.Print "Concluido! - testeGeral"

End Sub


Sub exemplos_criacao_arquivos_json()
Dim s As New clsCriarArquivos
Dim qrySelectRegistroValido As String: qrySelectRegistroValido = _
            "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi]))>0) AND ((tblDadosConexaoNFeCTe.registroValido)=1));"

Dim strCaminhoDeSaida As String: strCaminhoDeSaida = "C:\temp\" & strControle
CreateDir strCaminhoDeSaida
    
    '' NO PROCESSAMENTO DO ARQUIVO DE XML
    s.criarArquivoJson opFlagLancadaERP, qrySelectRegistroValido, strCaminhoDeSaida

    '' SELEÇÃO PELO USUARIO
    s.criarArquivoJson opManifesto, qrySelectRegistroValido, strCaminhoDeSaida


    MsgBox "Concluido!", vbOKOnly + vbInformation, "teste_arquivos_json"

Cleanup:

    Set s = Nothing

End Sub

Sub testeUnitario_carregarDadosGerais()
Dim Item As Variant

''#######################################################################################
''### BASE DE TESTES
''#######################################################################################

'' REPOSITORIO
Dim arquivos As Collection: Set arquivos = New Collection

''' RETORNO SIMBÓLICO DE MERCADORIA DEPOSITADA EM DEPÓSITO FECHA
arquivos.add "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210212680452000302550020000886301507884230-nfeproc.xml"

''' TRANSF. DE MERCADORIAS
arquivos.add "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210368365501000296550000000638811001361356-nfeproc.xml"

''' #TIPO 01 - CTE - TRANSPORTE RODOVIARIO
arquivos.add "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210220147617000494570010009539201999046070-cteproc.xml"

''' PREST. SERV. TRANSPORTE A ESTABELECIMENTO COMERCIAL
arquivos.add "C:\temp\Coleta\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210304884082000569570000040073831040073834-cteproc.xml"

''#######################################################################################
''### PROCESSAMENTO
''#######################################################################################

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

'' #CONTADOR
Dim contadorDeRegistros As Long: contadorDeRegistros = 1


    '' #BARRA_PROGRESSO
    SysCmd acSysCmdInitMeter, "Pendentes ...", arquivos.count

    For Each Item In arquivos
    
        carregarDadosGerais CStr(Item)

        '' #BARRA_PROGRESSO
        contadorDeRegistros = contadorDeRegistros + 1
        SysCmd acSysCmdUpdateMeter, contadorDeRegistros

        DoEvents
    Next Item


    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "Processamento - Importar Dados Gerais ( Quantidade de registros: " & contadorDeRegistros & " )"
    
    '' #BARRA_PROGRESSO
    SysCmd acSysCmdRemoveMeter

'    MsgBox "Concluido!", vbOKOnly + vbInformation, "testeUnitario_carregarDadosGerais"

End Sub

Sub testeUnitario_carregarArquivosPendentes()
Dim Item As Variant

''#######################################################################################
''### BASE DE TESTES
''#######################################################################################
Dim DadosGerais As New clsConexaoNfeCte

'' REPOSITORIO
Dim arquivos As Collection: Set arquivos = New Collection

''' DADOS
For Each Item In carregarParametros(DadosGerais.SelectArquivosPendentes)
    arquivos.add CStr(Item)
Next

''#######################################################################################
''### PROCESSAMENTO
''#######################################################################################

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

'' #CONTADOR
Dim contadorDeRegistros As Long: contadorDeRegistros = 1


    '' #BARRA_PROGRESSO
    SysCmd acSysCmdInitMeter, "Pendentes ...", arquivos.count

    For Each Item In arquivos
    
        carregarArquivosPendentes CStr(Item)

        '' #BARRA_PROGRESSO
        contadorDeRegistros = contadorDeRegistros + 1
        SysCmd acSysCmdUpdateMeter, contadorDeRegistros

        DoEvents
    Next Item


    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "Processamento - Importar Dados Gerais ( Quantidade de registros: " & contadorDeRegistros & " )"
    
    '' #BARRA_PROGRESSO
    SysCmd acSysCmdRemoveMeter

'    MsgBox "Concluido!", vbOKOnly + vbInformation, "testeUnitario_carregarArquivosPendentes"

End Sub

Function carregarDadosGerais(strArquivo As String)
On Error GoTo adm_Err
Dim s As New clsProcessamentoDados
Dim DadosGerais As New clsConexaoNfeCte
Dim Item As Variant
Dim strRepositorio As String

    ''#######################################################################################
    ''### ENVIAR DADOS DE ARQUIVOS PARA TABELA DE PROCESSAMENTO
    ''#######################################################################################
    
    '' REPOSITORIO
    strRepositorio = "tblDadosConexaoNFeCTe"

    '' LIMPAR TABELA DE PROCESSAMENTOS
    s.DeleteProcessamento

    '' PROCESSAMENTO
    s.ProcessamentoDeArquivo strArquivo, opDadosGerais

    '' IDENTIFICAR CAMPOS
    s.UpdateProcessamentoIdentificarCampos strRepositorio
    
    '' FORMATAR DADOS
    s.UpdateProcessamentoFormatarDados
        
        
    ''#######################################################################################
    ''### TRANSFERIR DADOS PROCESSADOS PARA REPOSITORIO
    ''#######################################################################################
        
    '' TRANSFERENCIA DE DADOS
    s.ProcessamentoTransferir strRepositorio
    
    ''#######################################################################################
    ''### TRATAMENTO DE DADOS IMPORTADOS
    ''#######################################################################################
    
    '' CLASSIFICAR DADOS GERAIS
    DadosGerais.TratamentoDeDadosGerais


adm_Exit:
    Set s = Nothing
    Set DadosGerais = Nothing
    
    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit

End Function

Function carregarArquivosPendentes(strArquivo As String)
On Error GoTo adm_Err

Dim s As New clsProcessamentoDados
Dim DadosGerais As New clsConexaoNfeCte
Dim strRepositorio As String
    
    ''#######################################################################################
    ''### ENVIAR DADOS DE ARQUIVOS PARA TABELA DE PROCESSAMENTO
    ''#######################################################################################

    '' REPOSITORIO
    strRepositorio = "tblCompraNF"

    '' LIMPAR TABELA DE PROCESSAMENTOS
    s.DeleteProcessamento

    '' PROCESSAMENTO
    s.ProcessamentoDeArquivo strArquivo, opCompras

    '' IDENTIFICAR CAMPOS
    s.UpdateProcessamentoIdentificarCampos strRepositorio
    
    '' CORREÇÃO DE DADOS MARCADOS ERRADOS EM ITENS DE COMPRAS
    s.UpdateProcessamentoLimparItensMarcadosErrados
    
    '' IDENTIFICAR CAMPOS DE ITENS DE COMPRAS
    s.UpdateProcessamentoIdentificarCampos strRepositorio & "Item"
    
    '' FORMATAR DADOS
    s.UpdateProcessamentoFormatarDados
    
    ''#######################################################################################
    ''### TRANSFERIR DADOS PROCESSADOS PARA REPOSITORIO
    ''#######################################################################################

    '' TRANSFERIR DADOS PROCESSADOS
    s.ProcessamentoTransferir strRepositorio
    s.ProcessamentoTransferir strRepositorio & "Item"

    ''#######################################################################################
    ''### FORMATAR DADOS PROCESSADOS
    ''#######################################################################################

    '' COMPRAS ATUALIAR CAMPOS
    DadosGerais.compras_atualizarCampos
    
    '' COMPRAS ITENS CTE
    DadosGerais.compras_carregarItensCTe

    '' FORMATAR ITENS DE COMPRA
    DadosGerais.FormatarItensDeCompras
    
    '' CADASTRO DE NUMERO DE PEDIDOS
    DadosGerais.UpdateNumPed_CompraNF

adm_Exit:
    Set s = Nothing
    Set DadosGerais = Nothing

    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit

End Function






