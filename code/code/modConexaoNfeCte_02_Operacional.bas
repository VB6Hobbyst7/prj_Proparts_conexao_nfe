Attribute VB_Name = "modConexaoNfeCte_02_Operacional"
Option Compare Database

Private FileCollention As New Collection
Private con As ADODB.Connection


Private Const script_tblDadosConexaoNFeCTe As String = "INSERT INTO tblDadosConexaoNFeCTe (codMod,dhEmi,CNPJ_emit,Razao_emit,CNPJ_Rem,CPNJ_Dest,CaminhoDoArquivo) VALUES(strOrigem,'strCaminhoDoArquivo')"

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


'' #####################################################################
'' ### #Ailton - EM TESTES
'' #####################################################################


Sub teste_20210308_1418()
Dim cadastro As clsFactory: Set cadastro = New clsFactory
Dim T As New Collection
Dim pPath As String: pPath = "c:\temp\20210308_1418\"
CreateDir pPath

'' CRIAR CLASSES e CONSULTAS
With cadastro
    .strFilePath = pPath
    .letNameTable carregarParametros("tabelaAuxiliar")
    .criarScriptClasse
    .criarScriptConsulta
End With

End Sub


'Sub teste_carregarDados()
'Dim tmp As clsDadosConexaoNFeCTe: Set tmp = New clsDadosConexaoNFeCTe
'
'    tmp.cadastroDados pegarValorDoParametro(strCaminhoDeColeta)
'
'End Sub


'' #####################################################################
'' ### #PRINCIPAL
'' #####################################################################

Private Sub LeituraDeArquivos() ''#ExtrairConteudoDeArquivo - Armazenar em tabela para tratamento de dados
    Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
    
    '' 01. Leitura e identificação do arquivo
    For Each fileName In GetFilesInSubFolders(pegarValorDoParametro(strCaminhoDeColeta))

        XDoc.Load fileName
        
        Debug.Print vbNewLine
        Debug.Print fileName
        
        '' 02. Separação por tags e coleta de dados
        For Each Item In Array("ide/mod", "ide/dhEmi", "emit/CNPJ", "emit/xNome", "rem/CNPJ", "dest/CNPJ", "infCTeNorm/infDoc/infNFe/chave")
        
            Set lists = XDoc.SelectNodes("//" & Item)
            
            For Each fieldNode In lists
            
                If (fieldNode.HasChildNodes) Then
                    For Each childNode In fieldNode.ChildNodes
                        Debug.Print "[" & fieldNode.BaseName & "] = [" & fieldNode.text & "]"
                    Next childNode
                End If
            Next fieldNode
        Next

    Next

    Set XDoc = Nothing
End Sub
