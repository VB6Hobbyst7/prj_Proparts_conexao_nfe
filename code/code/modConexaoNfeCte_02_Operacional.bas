Attribute VB_Name = "modConexaoNfeCte_02_Operacional"
Option Compare Database

Private FileCollention As New Collection
Private con As ADODB.Connection

Private Const script_tblDadosConexaoNFeCTe As String = "INSERT INTO tblDadosConexaoNFeCTe (codMod,dhEmi,CNPJ_emit,Razao_emit,CNPJ_Rem,CPNJ_Dest,CaminhoDoArquivo) VALUES(strOrigem,'strCaminhoDoArquivo')"
Private Const qryParametro As String = "SELECT tblParametros.ValorDoParametro FROM tblParametros WHERE (((tblParametros.TipoDeParametro) = 'strParametro'))"
Private Const qryTags As String = "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((Len([Tag]))>0) AND ((tblOrigemDestino.Destino) Like 'strParametro*') AND ((tblOrigemDestino.TagOrigem)=1));"
Private Const qryTagsLocais As String = "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((Len([Tag]))>0) AND ((tblOrigemDestino.Destino) Like 'strParametro*') AND ((tblOrigemDestino.TagOrigem)=2));"
Private Const qryUpdateProcessamento As String = "UPDATE (SELECT tblOrigemDestino.Destino, tblOrigemDestino.Tag, strSplit([Destino],'.',0) AS NomeTabela, strSplit([Destino],'.',1) AS NomeCampo FROM tblOrigemDestino) as qryOrigemDestino INNER JOIN tblProcessamento ON qryOrigemDestino.Tag = tblProcessamento.chave SET tblProcessamento.NomeTabela = [qryOrigemDestino].[NomeTabela], tblProcessamento.NomeCampo = [qryOrigemDestino].[NomeCampo] ;" 'WHERE (((Len([NomeTabela]))<0))

Private Const strCaminhoDeColeta As String = "caminhoDeColeta"


'' #Libs
'' carregarDados
'' carregarERP
'' carregarManifesto
'' main ---->   EM DESENVOLVIMENTO

''#CarregarValorDeParametro
''#ConsultarArquivosEmPastas
''#ExtrairConteudoDeArquivo
''#ExtrairCaminhoDoArquivo
''#ExtrairNomeDoArquivo
''#CriarPastasDestino
''#FormatarTimeStamp
''#ExecutarConsultas
''#CriarArquivosJson


'' #########################################################################################
'' ### #Proparts - Módulo principal para processamento de dados
'' #########################################################################################

Private Sub LeituraDeArquivos() ''#ExtrairConteudoDeArquivo - Armazenar em tabela para tratamento de dados
    Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
    Dim cadastro As clsProcessamento
    Dim col As New Collection
    Dim strPk As String
    
    Dim qryAtulizacoes() As Variant: arr = Array(qryUpdateProcessamento)
    
        
    '' 01. Leitura e identificação do arquivo
    For Each fileName In GetFilesInSubFolders(pegarValorDoParametro(strCaminhoDeColeta, qryParametro))
        XDoc.Load fileName
        
        Set cadastro = New clsProcessamento
        
        '' IDENTIFICAÇÃO DO ARQUIVO (PK)
        strPk = Controle & getFileName(CStr(fileName))
        
        col.add strPk & "|" & "CaminhoDoArquivo" & "|" & fileName
        
                
        '' 02. Separação por tags e coleta de dados
        For Each Item In carregarParametros("tblDadosConexaoNFeCTe", qryTags)
            Set lists = XDoc.SelectNodes("//" & Item)
            
            For Each fieldnode In lists
                If (fieldnode.HasChildNodes) Then
                    For Each childNode In fieldnode.ChildNodes
                       
                        col.add strPk & "|" & Item & "|" & fieldnode.text
                        
                    Next childNode
                End If
            
            Next fieldnode
        
        Next Item

        '' REALIZAR CADASTRO
        If (col.count > 2) Then
            
            '' CADASTRAR REGISTRO
            CadastroDeProcessamento col
            
            '' ATUALIZAR CAMPOS DE RELACIONAMENTOS
            Application.CurrentDb.Execute qryUpdateProcessamento
            
        End If
        
        '' LIMPAR COLEÇÃO
        ClearCollection col

    Next fileName

    Set XDoc = Nothing
    
    MsgBox "Concluido!", vbOKOnly + vbInformation, "LeituraDeArquivos"
    
End Sub


Sub DadosConexaoNFeCTe()
Dim dados As New clsDadosConexaoNFeCTe
Dim sqyProcessamentosPendentes As String: sqyProcessamentosPendentes = "SELECT DISTINCT pk from tblProcessamento;"

Dim db As DAO.Database: Set db = CurrentDb
Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(sqyProcessamentosPendentes)

    Do While Not rst.EOF
        dados.carregar_dados rst.Fields("pk").Value
        dados.cadastrar
        rst.MoveNext
    Loop

    db.Close
    
    Set db = Nothing
    
    MsgBox "Concluido!", vbOKOnly + vbInformation, "DadosConexaoNFeCTe"

End Sub
