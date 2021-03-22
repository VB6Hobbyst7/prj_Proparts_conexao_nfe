Attribute VB_Name = "modConexaoNfeCte_02_Operacional"
Option Compare Database

Private FileCollention As New Collection
Private con As ADODB.Connection

'' PROCESSAMENTOS - LIMPAR TABELA
Private Const qryDeleteProcessamento As String = "DELETE * FROM tblProcessamento;"

'' PROCESSAMENTOS - PENDENTES
Private Const sqyProcessamentosPendentes As String = "SELECT DISTINCT pk from tblProcessamento;"

'' CONSULTA PARA CRIAÇÃO DE ARQUIVOS JSON
Private Const sqyDadosJson As String = "SELECT chave, Comando, dhEmi, CaminhoDoArquivo,codTipoEvento FROM tblDadosConexaoNFeCTe WHERE LEN(Chave)>0;"

'' CARREGAR TAGs DE VINDAS DO XML
Private Const qryTags As String = "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((Len([Tag]))>0) AND ((tblOrigemDestino.tabela) = 'strParametro') AND ((tblOrigemDestino.TagOrigem)=1));"

'' ATUALIZAR CAMPOS ( Nometabela e NomeCampo ) PARA USO DA TABELA DE PROCESSAMENTO
Private Const qryUpdateProcessamento As String = "UPDATE (SELECT tblOrigemDestino.Destino, tblOrigemDestino.Tag, strSplit([Destino],'.',0) AS NomeTabela, strSplit([Destino],'.',1) AS NomeCampo FROM tblOrigemDestino WHERE tblOrigemDestino.Tabela = 'strParametro' ) as qryOrigemDestino INNER JOIN tblProcessamento ON qryOrigemDestino.Tag = tblProcessamento.chave SET tblProcessamento.NomeTabela = [qryOrigemDestino].[NomeTabela], tblProcessamento.NomeCampo = [qryOrigemDestino].[NomeCampo];"

'' CONSULTA DA TABELA PARAMETRO
Private Const qryParametros As String = "SELECT tblParametros.ValorDoParametro FROM tblParametros WHERE (((tblParametros.TipoDeParametro) = 'strParametro'))"
Private Const strCaminhoDeColeta As String = "caminhoDeColeta"
Private Const strUsuarioErpCod As String = "UsuarioErpCod"
Private Const strUsuarioErpNome As String = "UsuarioErpNome"
Private Const strCodTipoEvento As String = "codTipoEvento"
Private Const strComando As String = "Comando"

''###########################################################################


'' #TESTES - SELEÇÃO DE FORNECEDORES VALIDOS
Private Const qryUpdate_FornecedoresValidos As String = "UPDATE (SELECT Replace(Replace(Replace([CNPJ_CPF],""."",""""),""-"",""""),""/"","""") AS strCNPJ_CPF, Clientes.CÓDIGOClientes AS ID_Cad FROM Clientes WHERE (((Replace(Replace(Replace([CNPJ_CPF],""."",""""),""-"",""""),""/"",""""))<>""00000000000000"" And (Replace(Replace(Replace([CNPJ_CPF],""."",""""),""-"",""""),""/"",""""))<>""99999999999""))) AS qryFornecedoresValidos     INNER JOIN tblDadosConexaoNFeCTe ON qryFornecedoresValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit SET tblDadosConexaoNFeCTe.registroValido = 1;"



Sub teste_FornecedoresValidos()
Dim qry() As Variant: qry = Array(qryUpdate_FornecedoresValidos)

    executarComandos qry

End Sub


Sub TESTES_20210314_2111()
Dim t As Variant

    For Each t In Array("tblDadosConexaoNFeCTe")
        ProcessarArquivosXml CStr(t)
        TransferirDadosConexaoNFeCTe
        ''qryUpdateDadosConexaoNFeCTe_IdTipo
    Next


'    For Each t In Array("tblCompraNF")  ''Array("tblDadosConexaoNFeCTe", "tblCompraNF", "tblCompraNFItem")
'        ProcessarArquivosXml CStr(t)
'        TransferirCompras
'    Next


End Sub



'' #########################################################################################
'' ### #Proparts - Módulo principal para processamento de dados
'' #########################################################################################

Sub teste()
Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
Dim qry() As Variant: qry = Array("chCTe")

'' cte
'XDoc.Load "C:\temp\Coleta\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210204884082000569570000039548351039548356-cteproc.xml"

'' nfe
'XDoc.Load "C:\temp\Coleta\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\29210220961864000187550010000001891138200000-nfeproc.xml"

For Each Item In qry
    Set lists = XDoc.SelectNodes("//" & Item)
    For Each fieldnode In lists
        If (fieldnode.HasChildNodes) Then
            For Each childNode In fieldnode.ChildNodes
                Debug.Print fieldnode.text
            Next childNode
        End If
    Next fieldnode
Next Item

Set XDoc = Nothing

End Sub


Private Sub ProcessarArquivosXml(pTabelaDestino As String)
    Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
    Dim cadastro As clsProcessamento
    Dim col As New Collection
    Dim strPk As String
    Dim i As Variant
    
    '' LIMPAR TABELA DE PROCESSAMENTOS
    Application.CurrentDb.Execute qryDeleteProcessamento
        
    '' 01.Leitura e identificação do arquivo
    For Each fileName In GetFilesInSubFolders(pegarValorDoParametro(qryParametros, strCaminhoDeColeta))
        XDoc.Load fileName
        
        '' 01.CRIAR CHAVE UNICA DE REGISTRO PARA CONTROLE DE DADOS
        strPk = Controle & getFileName(CStr(fileName))
        col.add strPk & "|" & "CaminhoDoArquivo" & "|" & fileName
        
        '' 02.CARREGAR CAMPOS DE ORIGEM X DESTINO DO REGISTRO
        For Each Item In carregarParametros(qryTags, pTabelaDestino)
            Set lists = XDoc.SelectNodes("//" & Item)
            For Each fieldnode In lists
                If (fieldnode.HasChildNodes) Then
                    For Each childNode In fieldnode.ChildNodes
                        col.add strPk & "|" & Item & "|" & fieldnode.text
                    Next childNode
                End If
            Next fieldnode
        Next Item

        '' 03. REALIZAR CADASTRO DE TODOS OS ITENS COLETADOS NA TABELA DE PROCESSAMENTO
        If (col.count > 2) Then
            
            '' CADASTRAR REGISTRO
            Set cadastro = New clsProcessamento
            For Each i In col
                With cadastro
                    .pk = Split(i, "|")(0)
                    .Chave = Split(i, "|")(1)
                    .valor = Split(i, "|")(2)
                    .cadastrar
                End With
            Next i
            
            '' ATUALIZAR CAMPOS DE RELACIONAMENTOS
            Application.CurrentDb.Execute Replace(qryUpdateProcessamento, "strParametro", pTabelaDestino)
            
        End If
        
        '' LIMPAR COLEÇÃO
        ClearCollection col

    Next fileName

    Set XDoc = Nothing
    
    MsgBox "Concluido!", vbOKOnly + vbInformation, "ProcessarArquivosXml"
    
End Sub

Sub TransferirDadosConexaoNFeCTe()
Dim dados As New clsDadosConexaoNFeCTe
Dim db As DAO.Database: Set db = CurrentDb
Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(sqyProcessamentosPendentes)

    Do While Not rst.EOF
        dados.Comando = pegarValorDoParametro(qryParametros, strComando)
        dados.codTipoEvento = pegarValorDoParametro(qryParametros, strCodTipoEvento)
        dados.carregar_dados rst.Fields("pk").Value
        dados.cadastrar
        
        rst.MoveNext
    Loop

    db.Close
    
    Set db = Nothing
    
    MsgBox "Concluido!", vbOKOnly + vbInformation, "TransferirDadosConexaoNFeCTe"

End Sub

Sub TransferirCompras()
Dim dados As New clsCompraNF
Dim db As DAO.Database: Set db = CurrentDb
Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(sqyProcessamentosPendentes)

    Do While Not rst.EOF
        
        dados.carregar_dados rst.Fields("pk").Value
        dados.cadastrar
        
        rst.MoveNext
    Loop

    db.Close
    
    Set db = Nothing
    
    MsgBox "Concluido!", vbOKOnly + vbInformation, "TransferirDadosConexaoNFeCTe"

End Sub


Sub CriarArquivoFlagLancadaERP()
Dim dados As New clsDadosConexaoNFeCTe
Dim db As DAO.Database: Set db = CurrentDb
Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(sqyDadosJson)

    Do While Not rst.EOF
        
        '' CRIAÇÃO DE Json
        dados.Chave = rst.Fields("chave").Value
        dados.Comando = rst.Fields("Comando").Value
        dados.dhEmi = rst.Fields("dhEmi").Value
        dados.codUsuarioErp = pegarValorDoParametro(qryParametros, strUsuarioErpCod)
        dados.nomeUsuarioErp = pegarValorDoParametro(qryParametros, strUsuarioErpNome)
        dados.CaminhoDoArquivo = getPath(rst.Fields("CaminhoDoArquivo").Value)
        dados.criarERP
                
        rst.MoveNext
    Loop

    db.Close
    
    Set db = Nothing
    
    MsgBox "Concluido!", vbOKOnly + vbInformation, "CriarArquivoFlagLancadaERP"

End Sub

Sub CriarArquivoDeManifesto()
Dim dados As New clsDadosConexaoNFeCTe
Dim db As DAO.Database: Set db = CurrentDb
Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(sqyDadosJson)

    Do While Not rst.EOF
        
        '' CRIAÇÃO DE Json
        dados.Chave = rst.Fields("chave").Value
        dados.codTipoEvento = rst.Fields("codTipoEvento").Value
        dados.dhEmi = rst.Fields("dhEmi").Value
        dados.codUsuarioErp = pegarValorDoParametro(qryParametros, strUsuarioErpCod)
        dados.nomeUsuarioErp = pegarValorDoParametro(qryParametros, strUsuarioErpNome)
        dados.CaminhoDoArquivo = rst.Fields("CaminhoDoArquivo").Value
        
        '' CRIAÇÃO DE MANIFESTO
        dados.criarManifesto
        
        rst.MoveNext
    Loop

    db.Close
    
    Set db = Nothing
    
    MsgBox "Concluido!", vbOKOnly + vbInformation, "CriarArquivoDeManifesto"

End Sub


''#######################################################################################
''### LIMBO
''#######################################################################################

'' #DESCONTINUADO - CONTROLE DE FORNECEDORES VALIDOS
''Private Const qryFornecedoresValidos As String = "SELECT Replace(Replace(Replace([CNPJ_CPF],""."",""""),""-"",""""),""/"","""") AS Expr1, Clientes.CÓDIGOClientes AS ID_Cad FROM Clientes WHERE (((Replace(Replace(Replace([CNPJ_CPF],""."",""""),""-"",""""),""/"",""""))<>""00000000000000"" And (Replace(Replace(Replace([CNPJ_CPF],""."",""""),""-"",""""),""/"",""""))<>""99999999999""));"

'' #DESCONTINUADO - CARREGAR TAGs DE VINDAS DA TABELA
'Private Const qryTagsLocais As String = "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((Len([Tag]))>0) AND ((tblOrigemDestino.Destino) Like 'strParametro*') AND ((tblOrigemDestino.TagOrigem)=2));"

'Sub CadastrarCompras()
'Dim dados As New clsDadosConexaoNFeCTe
'Dim sqyProcessamentosPendentes As String: sqyProcessamentosPendentes = "SELECT DISTINCT pk from tblProcessamento;"
'
'Dim db As DAO.Database: Set db = CurrentDb
'Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(sqyProcessamentosPendentes)
'
'    Do While Not rst.EOF
'        dados.carregar_dados rst.Fields("pk").Value
'        dados.cadastrar
'        rst.MoveNext
'    Loop
'
'    db.Close
'
'    Set db = Nothing
'
'    MsgBox "Concluido!", vbOKOnly + vbInformation, "CadastroCompras"
'
'End Sub

'Public Sub CadastroDeProcessamento(obj As Collection)
'Dim cadastro As clsProcessamento: Set cadastro = New clsProcessamento
'Dim I As Variant
'
'For Each I In obj
'    With cadastro
'        .pk = Split(I, "|")(0)
'        .Chave = Split(I, "|")(1)
'        .valor = Split(I, "|")(2)
'        .cadastrar
'    End With
'Next I
'
'End Sub
