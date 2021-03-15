Attribute VB_Name = "modConexaoNfeCte_02_Operacional"
Option Compare Database

Private FileCollention As New Collection
Private con As ADODB.Connection

'' DESATIVADO
'Private Const script_tblDadosConexaoNFeCTe As String = "INSERT INTO tblDadosConexaoNFeCTe (codMod,dhEmi,CNPJ_emit,Razao_emit,CNPJ_Rem,CPNJ_Dest,CaminhoDoArquivo) VALUES(strOrigem,'strCaminhoDoArquivo')"

'' PROCESSAMENTOS - LIMPAR TABELA
Private Const qryDeleteProcessamento As String = "DELETE * FROM tblProcessamento;"

'' PROCESSAMENTOS - PENDENTES
Private Const sqyProcessamentosPendentes As String = "SELECT DISTINCT pk from tblProcessamento;"

'' CONSULTA PARA CRIAÇÃO DE ARQUIVOS JSON
Private Const sqyDadosJson As String = "SELECT chave, Comando, dhEmi, CaminhoDoArquivo,codTipoEvento FROM tblDadosConexaoNFeCTe WHERE LEN(Chave)>0;"

'' CARREGAR TAGs DE VINDAS DO XML
'Private Const qryTags As String = "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((Len([Tag]))>0) AND ((tblOrigemDestino.Destino) Like 'strParametro*') AND ((tblOrigemDestino.TagOrigem)=1));"
Private Const qryTags As String = "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((Len([Tag]))>0) AND ((tblOrigemDestino.tabela) = 'strParametro') AND ((tblOrigemDestino.TagOrigem)=1));"

'' CARREGAR TAGs DE VINDAS DA TABELA
Private Const qryTagsLocais As String = "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((Len([Tag]))>0) AND ((tblOrigemDestino.Destino) Like 'strParametro*') AND ((tblOrigemDestino.TagOrigem)=2));"

'' ATUALIZAR CAMPOS ( Nometabela e NomeCampo ) PARA USO DA TABELA DE PROCESSAMENTO
Private Const qryUpdateProcessamento As String = "UPDATE (SELECT tblOrigemDestino.Destino, tblOrigemDestino.Tag, strSplit([Destino],'.',0) AS NomeTabela, strSplit([Destino],'.',1) AS NomeCampo FROM tblOrigemDestino WHERE tblOrigemDestino.Tabela = 'strParametro' ) as qryOrigemDestino INNER JOIN tblProcessamento ON qryOrigemDestino.Tag = tblProcessamento.chave SET tblProcessamento.NomeTabela = [qryOrigemDestino].[NomeTabela], tblProcessamento.NomeCampo = [qryOrigemDestino].[NomeCampo];"

'' CONSULTA DA TABELA PARAMETRO
Private Const qryParametros As String = "SELECT tblParametros.ValorDoParametro FROM tblParametros WHERE (((tblParametros.TipoDeParametro) = 'strParametro'))"
Private Const strCaminhoDeColeta As String = "caminhoDeColeta"
Private Const strUsuarioErpCod As String = "UsuarioErpCod"
Private Const strUsuarioErpNome As String = "UsuarioErpNome"
Private Const strCodTipoEvento As String = "codTipoEvento"
Private Const strComando As String = "Comando"
Private Const strTagOrigemPrincipal As String = "tagOrigemPrincipal"


Sub TESTES_20210314_2111()
Dim arr As Variant
Dim t As Variant

    For Each t In Array("tblCompraNF")  ''Array("tblDadosConexaoNFeCTe", "tblCompraNF", "tblCompraNFItem")
        ProcessarArquivosXml CStr(t)
    Next

End Sub



'' #########################################################################################
'' ### #Proparts - Módulo principal para processamento de dados
'' #########################################################################################

Private Sub ProcessarArquivosXml(pTabelaDestino As String)
    Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
    Dim cadastro As clsProcessamento
    Dim col As New Collection
    Dim strPk As String
    Dim I As Variant
    
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
            For Each I In col
                With cadastro
                    .pk = Split(I, "|")(0)
                    .Chave = Split(I, "|")(1)
                    .valor = Split(I, "|")(2)
                    .cadastrar
                End With
            Next I
            
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


Sub CadastrarCompras()
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

End Sub




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
