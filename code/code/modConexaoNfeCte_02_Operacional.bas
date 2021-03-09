Attribute VB_Name = "modConexaoNfeCte_02_Operacional"
Option Compare Database

Private FileCollention As New Collection
Private con As ADODB.Connection

Private Const script_tblDadosConexaoNFeCTe As String = "INSERT INTO tblDadosConexaoNFeCTe (codMod,dhEmi,CNPJ_emit,Razao_emit,CNPJ_Rem,CPNJ_Dest,CaminhoDoArquivo) VALUES(strOrigem,'strCaminhoDoArquivo')"
Private Const qryParametro As String = "SELECT tblParametros.ValorDoParametro FROM tblParametros WHERE (((tblParametros.TipoDeParametro) = 'strParametro'))"
Private Const qryTags As String = "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((Len([Tag]))>0) AND ((tblOrigemDestino.Destino) Like 'strParametro*'));"

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
        
    '' 01. Leitura e identificação do arquivo
    For Each fileName In GetFilesInSubFolders(pegarValorDoParametro(strCaminhoDeColeta, qryParametro))
        XDoc.Load fileName
        
        Set cadastro = New clsProcessamento
        
        '' IDENTIFICAÇÃO DO ARQUIVO (PK)
        strPk = Controle & getFileName(CStr(fileName))
        
'        Debug.Print strPk & "|" & "CaminhoDoArquivo" & "|" & fileName
        col.add strPk & "|" & "CaminhoDoArquivo" & "|" & fileName
                
        '' 02. Separação por tags e coleta de dados
        For Each Item In carregarParametros("tblDadosConexaoNFeCTe", qryTags)
            Set lists = XDoc.SelectNodes("//" & Item)
            
            For Each fieldnode In lists
                If (fieldnode.HasChildNodes) Then
                    For Each childNode In fieldnode.ChildNodes
                       
'                        Debug.Print strPk & "|" & Item & "|" & fieldnode.text
                        col.add strPk & "|" & Item & "|" & fieldnode.text
                        
                    Next childNode
                End If
            
            Next fieldnode
        
        Next Item

        '' REALIZAR CADASTRO
        If (col.count > 2) Then
            CadastroDeProcessamento col
        End If
        
        '' LIMPAR COLEÇÃO
        ClearCollection col

    Next fileName

    Set XDoc = Nothing
End Sub



Sub CadastroDeProcessamento(obj As Collection)
Dim cadastro As clsProcessamento: Set cadastro = New clsProcessamento
Dim I As Variant

For Each I In obj
    With cadastro
        .pk = Split(I, "|")(0)
        .chave = Split(I, "|")(1)
        .valor = Split(I, "|")(2)
        .cadastrar
    End With
Next I

End Sub
