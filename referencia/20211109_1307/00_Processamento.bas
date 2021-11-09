Attribute VB_Name = "00_Processamento"
Option Compare Database

'' #PENDENTE
'' * proxima atualização
''
'' #TESTES
'' * 32210268365501000296550000000637741001351624 | 55
'' * 32210304884082000569570000040073831040073834 | 57




'' ### ao montar analisa "contains" '' remover duplicados
'' ### ListagemDeArquivosValidos() '' listar apenas arquivos validos para cadastro
'' ### export() '' chvCadastro



'' #20211109
Sub teste_ListagemDeArquivosValidos()

'' #ENTRADA - LISTAGEM DE TODOS OS ARQUIVOS COLETADOS
Dim pColArquivos As Collection: Set pColArquivos = New Collection
    pColArquivos.add "C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210348740351012767570000021186701559009401-cteproc.xml", "C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210348740351012767570000021186701559009401-cteproc.xml"
    pColArquivos.add "C:\xmls\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\32210348740351015359570000000443691303812742-cteproc.xml", "C:\xmls\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\32210348740351015359570000000443691303812742-cteproc.xml"
    pColArquivos.add "C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210307872326000158550040001550831011035318-nfeproc.xml", "C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210307872326000158550040001550831011035318-nfeproc.xml"
    pColArquivos.add "C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210312680452000302550020000902331810472980-nfeproc.xml", "C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210312680452000302550020000902331810472980-nfeproc.xml"
    pColArquivos.add "C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210320147617000494570010009658691999034138-cteproc.xml", "C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210320147617000494570010009658691999034138-cteproc.xml"
    pColArquivos.add "C:\xmls\68.365.5010001-05 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\32210268365501000296550000000637741001351624-nfeproc.xml", "C:\xmls\68.365.5010001-05 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\32210268365501000296550000000637741001351624-nfeproc.xml"

Dim item As Variant

'' Importação
For Each item In ListagemDeArquivosValidosParaCadastros(pColArquivos)
    Debug.Print CStr(item)
Next

End Sub

'' #20211109
Function ListagemDeArquivosValidosParaCadastros(pColArquivos As Collection) As Collection
Dim colArquivo As Variant

    '' remover duplicados - #PENDENTE
    '' remover cadastrados
    For Each colArquivo In pColArquivos
        If DLookup("[ID]", "[tblDadosConexaoNFeCTe]", "[Chave]='" & getFileName(CStr(colArquivo)) & "'") <> "" Then pColArquivos.remove CStr(colArquivo)
    Next

'' return as collection
Set ListagemDeArquivosValidosParaCadastros = pColArquivos

End Function


Sub teste_TransferirProcessamentoParaRepositorios()
Dim Processamento As New clsProcessamentoDados
Dim DadosGerais As New clsConexaoNfeCte

Dim strChave As String: strChave = "32210268365501000296550000000637741001351624" ' 55
'Dim strChave As String: strChave = "32210304884082000569570000040073831040073834" ' 57
Dim strRepositorio As String: strRepositorio = "tblCompraNF"

Dim strArquivo As String: _
    strArquivo = "C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210348740351012767570000021186701559009401-cteproc.xml"
    '' strArquivo = "C:\xmls\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\32210348740351015359570000000443691303812742-cteproc.xml"
    '' strArquivo = "C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210307872326000158550040001550831011035318-nfeproc.xml"
    '' strArquivo = "C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210312680452000302550020000902331810472980-nfeproc.xml"
    '' strArquivo = "C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210320147617000494570010009658691999034138-cteproc.xml"
    '' strArquivo = "C:\xmls\68.365.5010001-05 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\32210268365501000296550000000637741001351624-nfeproc.xml"
    ''strArquivo = DLookup("[CaminhoDoArquivo]", "[tblDadosConexaoNFeCTe]", "[ChvAcesso]='" & strChave & "'")


    '' #LIMPAR REPOSITORIO DE COLETA
'    Processamento.DeleteProcessamento
    
    '' #IMPORTACAO
    Processamento.ProcessamentoDeArquivo strArquivo, opDadosGerais
        
    '' #IDENTIFICAÇÃO - CAMPOS
'    Processamento.UpdateProcessamentoIdentificarCampos "tblDadosConexaoNFeCTe"
        
    '' #FORMATAÇÃO
'    Processamento.UpdateProcessamentoFormatarDados
        
    '' #ARQUIVO
'    DadosGerais.ProcessamentoTransferir "tblDadosConexaoNFeCTe"
    
    
    '' #CLASSIFICAÇÃO
'    DadosGerais.TratamentoDeDadosGerais
    
    
''' ################################################
'
'
'    '' #IMPORTACAO - CADASTRO DE DADOS DO ARQUIVO NA TABELA DE PROCESSAMENTO -
'    Processamento.ProcessamentoDeArquivo strArquivo, "tblCompraNF"
'
'
'    '' #IDENTIFICAÇÃO - CAMPOS
'    Processamento.UpdateProcessamentoIdentificarCampos "tblCompraNF"
'
'    '' #CORREÇÃO DE DADOS MARCADOS ERRADOS EM ITENS DE COMPRAS
'    Processamento.UpdateProcessamentoLimparItensMarcadosErrados
'
'    '' #IDENTIFICAÇÃO - CAMPOS
'    Processamento.UpdateProcessamentoIdentificarCampos "tblCompraNFItem"
'
'    '' FORMATAR DADOS
'    Processamento.UpdateProcessamentoFormatarDados
'
'    '' #CLASSIFICAÇÃO
'    DadosGerais.compras_atualizarCampos
'
'
'    '' TRANSFERIR
''    TransferirProcessamentoParaRepositorios strChave


Set Processamento = Nothing
Set DadosGerais = Nothing

End Sub



'' 01. Seleção
Sub teste_CarregarArquivosDeColeta()
Dim arquivo As Variant
Dim Processamento As New clsProcessamentoDados
Dim DadosGerais As New clsConexaoNfeCte

Dim strRepositorio As String: strRepositorio = "tblDadosConexaoNFeCTe"

    For Each arquivo In CarregarArquivosDeColeta
        Debug.Print arquivo
        TextFile_Append CurrentProject.path & "\CarregarArquivosDeColeta.log", CStr(arquivo)
                
        '' LIMPAR TABELA DE PROCESSAMENTOS
        Processamento.DeleteProcessamento
        
        '' PROCESSAMENTO
        Processamento.ProcessamentoDeArquivo CStr(arquivo), opDadosGerais
        
        '' IDENTIFICAR CAMPOS
        Processamento.UpdateProcessamentoIdentificarCampos strRepositorio
                
        '' FORMATAR DADOS
        Processamento.UpdateProcessamentoFormatarDados
        
        '' TRANSFERENCIA DE DADOS
        Processamento.ProcessamentoTransferir strRepositorio
    
        '' CLASSIFICAR DADOS GERAIS
        DadosGerais.TratamentoDeDadosGerais
    
    
        DoEvents
    Next

End Sub

'' 02. Classificação
Sub teste_CarregarArquivosParaProcessamento()
Dim arquivo As Variant

    For Each arquivo In CarregarArquivosParaProcessamento
        Debug.Print arquivo
        TextFile_Append CurrentProject.path & "\CarregarArquivosParaProcessamento.log", CStr(arquivo)
        
        
        
        
        
        
        DoEvents
    Next

End Sub



'' #############################################
'' ### LIB'S
'' #############################################


Function CarregarArquivosParaProcessamento() As Collection: Set CarregarArquivosParaProcessamento = New Collection
Dim DadosGerais As New clsConexaoNfeCte

    For Each item In carregarParametros(DadosGerais.SelectArquivosPendentes)
        CarregarArquivosParaProcessamento.add CStr(item)
    Next

Set DadosGerais = Nothing
End Function


Function CarregarArquivosDeColeta() As Collection: Set CarregarArquivosDeColeta = New Collection
Dim DadosGerais As New clsConexaoNfeCte

'' REPOSITORIOS
For Each caminhoAntigo In Array(DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColeta'"))
    For Each caminhoNovo In carregarParametros(DadosGerais.SelectColetaEmpresa)
        For Each item In GetFilesInSubFolders(CStr(Replace(Replace(caminhoAntigo, "empresa", caminhoNovo), "recebimento\", "")))
            CarregarArquivosDeColeta.add CStr(item)
        Next
    Next
Next

End Function


