Attribute VB_Name = "azsCadastroRegistros"
Option Compare Database

Sub createTable()

'azs_createTable "tblCompraNF"
azs_createTable "tblCompraNFItem"
'azs_createTable "tblDadosConexaoNFeCTe"

End Sub


Sub azs_createTable(pTabelaNome As String)
'On Error Resume Next

Dim pTabelaCampos As String: pTabelaCampos = "SELECT DISTINCT campo FROM tblOrigemDestino WHERE tabela = '" & pTabelaNome & "' AND len(campo)>0 ;" '' ORDER BY id
Dim qryProcessos() As Variant
Dim script As String

    script = "SELECT "
    For Each Item In carregarParametros(pTabelaCampos)
        script = script & "'' as " & Item & " ,"
    Next Item
    script = left(script, Len(script) - 1) & " INTO " & pTabelaNome & ";"
    
    '' EXCLUIR CASO EXISTA
    If Not IsNull(DLookup("Name", "MSysObjects", "type in(1,6) and Name='" & pTabelaNome & "'")) Then Application.CurrentDb.Execute "DROP TABLE " & pTabelaNome
    
    '' CRIAR TABELA
    qryProcessos = Array(script, "DELETE FROM " & pTabelaNome)
    executarComandos qryProcessos
    
    
Set con = Nothing

Cleanup:
'        Set pTabelaNome = Nothing
'        Set pTabelaCampos = Nothing

End Sub

Sub start_cadastro()
    CadastroDeRegistros registro
End Sub

Private Sub CadastroDeRegistros(Itens As Collection)
Dim con As ADODB.Connection: Set con = CurrentProject.Connection
Dim i As Variant

    For Each i In Itens
        con.Execute i
    Next i

Set con = Nothing

End Sub


Private Function registro() As Collection
Set registro = New Collection

    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS00/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS20/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS30/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS40/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS41/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS50/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS51/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS60/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMSSN101/CSOSN',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMSSN102/CSOSN',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMSSN500/CSOSN',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS00/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS20/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS30/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS40/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS41/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS50/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS51/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS60/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMSSN101/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMSSN102/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMSSN500/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS00/vICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/vICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS20/vICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS51/vICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/vICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/vICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMSSN101/vCredICMSSN',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS00/modBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/modBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS20/modBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS51/modBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/modBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/modBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/modBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS30/modBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/modBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/modBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS00/pICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/pICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS20/pICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS51/pICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/pICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/pICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMSSN101/pCredSN',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/pICMSST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS30/pICMSST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/pICMSST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/pICMSST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/pMVAST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS30/pMVAST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/pMVAST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/pMVAST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS20/pRedBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS51/pRedBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/pRedBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/pRedBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/pRedBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS30/pRedBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/pRedBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/pRedBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS00/vBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/vBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS20/vBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS51/vBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/vBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/vBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/vBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS30/vBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/vBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/vBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/vICMSST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS30/vICMSST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/vICMSST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/vICMSST',1)"
    
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/IPI|cEnq',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/IPI/IPITrib|CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/IPI/IPITrib|vBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/IPI/IPITrib|CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/IPI/IPITrib|pIPI',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/IPI/IPITrib|vIPI',1)"
End Function


'Sub TESTE_CARREGAR_REGISTRO_UNICO()
'
'    '' LIMPAR TABELA DE PROCESSAMENTOS
'    Application.CurrentDb.Execute qryDeleteProcessamento
'
'
'    '' VENDA MERCADORIAS ADQUIRIDAS E/OU RECEB TERCEIROS
'    '' NUMERO_NF: 629140
'    '' NUMERO_ITENS: 01
'    carregar_Compras "C:\temp\Coleta\68.365.5010001-05 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\26210324073694000155550010006291401018935070-nfeproc.xml"
'
'    '' Retorno simbólico de mercadoria depositada em depósito fecha
''    carregar_Compras "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210312680452000302550020000896201269925336-nfeproc.xml"
'
'
'    '' TRANSPORTE RODOVIARIO
''    carregar_Compras "C:\temp\Coleta\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210248740351015359570000000309211914301218-cteproc.xml"
'
'
'    '' ARQUIVO - CERTIFICADO
''    carregar_Compras "C:\temp\Coleta\68.365.5010001-05 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\35210365833410000169550000006711211238251650-nfeproc.xml"
'
'
'End Sub


'Sub TESTE_CARREGAR_REGISTRO_LISTA()
'Dim Item As Variant
'
''' #BARRA_PROGRESSO
'Dim DT_PROCESSO As Date: DT_PROCESSO = Now()
'Dim contadorDeRegistros As Long: contadorDeRegistros = 1
'
''' LIMPAR TABELA DE PROCESSAMENTOS
'Application.CurrentDb.Execute qryDeleteProcessamento
'
'    '' #BARRA_PROGRESSO
'    SysCmd acSysCmdInitMeter, "Processando ...", DCount("*", "tblDadosConexaoNFeCTe", "(((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0))")
'
'    '' CARREGAR_DADOS
'    For Each Item In carregarParametros(qrySelecaoDeArquivosPendentes)
'        '' #BARRA_PROGRESSO
'        SysCmd acSysCmdUpdateMeter, contadorDeRegistros
'
'        carregar_Compras CStr(Item)
'
'        '' #BARRA_PROGRESSO
'        contadorDeRegistros = contadorDeRegistros + 1
'
'        DoEvents
'    Next Item
'
'    '' #BARRA_PROGRESSO
'    SysCmd acSysCmdRemoveMeter
'    statusFinal DT_PROCESSO, "Processamento - teste_listarArquivos"
'
'End Sub


'Sub TESTE_CARREGAR_REGISTRO_UNICO__20210520_1600()
'Dim s As New clsConexaoNfeCte
'
'    '' LIMPAR TABELA DE PROCESSAMENTOS
'    Application.CurrentDb.Execute qryDeleteProcessamento
'
'
'    '' VENDA MERCADORIAS ADQUIRIDAS E/OU RECEB TERCEIROS
'    '' NUMERO_NF: 629140
'    '' NUMERO_ITENS: 01
'    s.processamentoDeCompras "C:\temp\Coleta\68.365.5010001-05 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\26210324073694000155550010006291401018935070-nfeproc.xml"
'
'    '' SAIDA DE VENDA
'    '' NUMERO_NF: 8745405
'    '' NUMERO_ITENS: 10
'    s.processamentoDeCompras "C:\temp\Coleta\68.365.5010001-05 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\35210343283811001202550010087454051410067364-nfeproc.xml"
'
'    '' Retorno simbólico de mercadoria depositada em depósito fecha
''    processamentoDeCompras "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210312680452000302550020000896201269925336-nfeproc.xml"
'
'
'    '' TRANSPORTE RODOVIARIO
''    processamentoDeCompras "C:\temp\Coleta\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210248740351015359570000000309211914301218-cteproc.xml"
'
'
'    '' ARQUIVO - CERTIFICADO
''    processamentoDeCompras "C:\temp\Coleta\68.365.5010001-05 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\35210365833410000169550000006711211238251650-nfeproc.xml"
'
'
'End Sub

