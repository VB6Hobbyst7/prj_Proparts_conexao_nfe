'' #CARREGAR DADOS DO ARQUIVO COM BASE NA TABELA DE ORIGEM DESTINO
Private Sub ProcessarArquivosXml(pTabelaDestino As String, pArquivos As Collection)
Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
Dim cadastro As clsProcessamento
Dim cRegistros As New Collection
Dim colCampos As New Collection
Dim strPk As String
Dim i As Variant
Dim fileName As Variant
Dim Item As Variant
Dim lists As Variant
Dim fieldnode As Variant
Dim childNode As Variant
    
'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()
Dim contadorDeArquivos As Long: contadorDeArquivos = 1

    statusFinal DT_PROCESSO, pTabelaDestino & " - Quantidade de arquivos: " & pArquivos.count

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdInitMeter, "Processando arquivos...", pArquivos.count
        
    '' LIMPAR TABELA DE PROCESSAMENTOS
    Application.CurrentDb.Execute qryDeleteProcessamento
        
        
    '' 01.Leitura e identificação do arquivo
    For Each fileName In pArquivos
        XDoc.Load fileName
        
        '' #BARRA_PROGRESSO
        SysCmd acSysCmdUpdateMeter, contadorDeArquivos
        
        '' 01.CRIAR CHAVE UNICA DE REGISTRO PARA CONTROLE DE DADOS
        strPk = IIf(IsNull(DLookup("[chave]", "[tblDadosConexaoNFeCTe]", "[CaminhoDoArquivo]='" & fileName & "'")), getFileName(CStr(fileName)), DLookup("[chave]", "[tblDadosConexaoNFeCTe]", "[CaminhoDoArquivo]='" & fileName & "'"))
        cRegistros.add strPk & "|" & "CaminhoDoArquivo" & "|" & fileName
        
        '' 02.CARREGAR CAMPOS DE ORIGEM X DESTINO DO REGISTRO
        For Each Item In carregarParametros(qryTags, pTabelaDestino)
            Set lists = XDoc.SelectNodes("//" & Item)
            For Each fieldnode In lists
                If (fieldnode.HasChildNodes) Then
                    For Each childNode In fieldnode.ChildNodes
                        cRegistros.add strPk & "|" & Item & "|" & fieldnode.text
                    Next childNode
                End If
            Next fieldnode
            
            DoEvents
            
        Next Item

        '' 03. REALIZAR CADASTRO DE TODOS OS ITENS COLETADOS NA TABELA DE PROCESSAMENTO
        If (cRegistros.count > 2) Then
            
            '' CADASTRAR REGISTRO
             cadastroProcessamento cRegistros
        
        End If
        
        '' LIMPAR COLEÇÃO
        ClearCollection cRegistros

        '' #BARRA_PROGRESSO
        contadorDeArquivos = contadorDeArquivos + 1
        
    Next fileName
                
                
'    '' ATUALIZAR CAMPOS DE RELACIONAMENTOS
'    Application.CurrentDb.Execute qryUpdateProcessamento
    

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdRemoveMeter

    Set XDoc = Nothing

    '' #ANALISE_DE_PROCESSAMENTO
    statusFinal DT_PROCESSO, "Processamento - ProcessarArquivosXml"
        
End Sub