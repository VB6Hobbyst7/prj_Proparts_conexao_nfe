Attribute VB_Name = "modConexaoNfeCte"
Option Compare Database

Private Const qryDeleteProcessamento As String = _
        "DELETE * FROM tblProcessamento;"

Private Const qrySelecaoDeCampos As String = _
        "SELECT tblOrigemDestino.Tag FROM tblOrigemDestino WHERE (((tblOrigemDestino.tabela)='strParametro') AND ((Len([Tag]))>0) AND ((tblOrigemDestino.TagOrigem)=1)) ORDER BY tblOrigemDestino.Tag, tblOrigemDestino.tabela;"
       
Private Const qrySelecaoDeArquivosPendentes As String = _
        "SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1)) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0) ORDER BY tblDadosConexaoNFeCTe.CaminhoDoArquivo;"

Sub teste_listarArquivos()
Dim Item As Variant

'' #BARRA_PROGRESSO
Dim contadorDeRegistros As Long: contadorDeRegistros = 1

'' LIMPAR TABELA DE PROCESSAMENTOS
Application.CurrentDb.Execute qryDeleteProcessamento

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdInitMeter, "Processando ...", DCount("*", "tblDadosConexaoNFeCTe", "(((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0))")

    '' CARREGAR_DADOS
    For Each Item In carregarParametros(qrySelecaoDeArquivosPendentes)
        '' #BARRA_PROGRESSO
        SysCmd acSysCmdUpdateMeter, contadorDeRegistros
    
        carregar_Compras CStr(Item)
        
        '' #BARRA_PROGRESSO
        contadorDeRegistros = contadorDeRegistros + 1
        
        DoEvents
    Next Item

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdRemoveMeter

End Sub
        
Sub TESTE_CARREGAR()

    '' LIMPAR TABELA DE PROCESSAMENTOS
    Application.CurrentDb.Execute qryDeleteProcessamento


    '' VENDA MERCADORIAS ADQUIRIDAS E/OU RECEB TERCEIROS
    '' NUMERO_NF: 629140
    '' NUMERO_ITENS: 01
    carregar_Compras "C:\temp\Coleta\68.365.5010001-05 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\26210324073694000155550010006291401018935070-nfeproc.xml"


    '' Retorno simbólico de mercadoria depositada em depósito fecha
'    carregar_Compras "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210312680452000302550020000896201269925336-nfeproc.xml"


    '' TRANSPORTE RODOVIARIO
'    carregar_Compras "C:\temp\Coleta\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\32210248740351015359570000000309211914301218-cteproc.xml"


    '' ARQUIVO - CERTIFICADO
'    carregar_Compras "C:\temp\Coleta\68.365.5010001-05 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\35210365833410000169550000006711211238251650-nfeproc.xml"


End Sub


Public Function carregar_Compras(pPathFile As String)
On Error Resume Next

Dim s As New clsConexaoNfeCte

'' CHAVES DE CONTROLE
Dim pPK As String: pPK = DLookup("[Chave]", "[tblDadosConexaoNFeCTe]", "[CaminhoDoArquivo]='" & pPathFile & "'")
Dim pChvAcesso As String: pChvAcesso = DLookup("[ChvAcesso]", "[tblDadosConexaoNFeCTe]", "[CaminhoDoArquivo]='" & pPathFile & "'")

'' CARREGAR ARQUIVO
Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
XDoc.Load pPathFile

Dim cont As Integer: cont = XDoc.getElementsByTagName("infNFe/det").Length
Dim Item As Variant

Dim pDados As New Collection

Dim idItem As String: idItem = ""

    '' IDENTIFICAÇÃO DO ARQUIVO
    pDados.add pPK & "|" & "CaminhoDoArquivo" & "|" & pPathFile

    '' CHAVE DE ACESSO
    pDados.add pPK & "|" & "ChvAcesso_CompraNF" & "|" & pChvAcesso

    '' CABEÇALHO DA COMPRA
    For Each Item In carregarParametros(qrySelecaoDeCampos, "tblCompraNF")
        Select Case UBound(Split((Item), "|"))
            
            '' ITEM DE COMPRA
            Case 1
                regiao = Split((Item), "|")(0)
                campo = Split((Item), "|")(1)
                valor = XDoc.SelectNodes(regiao).Item(0).SelectNodes(campo).Item(0).text
                If valor <> "" Then pDados.add pPK & "|" & campo & "|" & valor
                
            Case Else
        End Select
        
        regiao = ""
        campo = ""
        valor = ""
        
        DoEvents
    
    Next Item


    '' ITENS DA COMPRA
    For i = 0 To cont - 1
        '' ID
        idItem = CStr(XDoc.getElementsByTagName("nfeProc/NFe/infNFe/det").Item(i).Attributes(0).value)
        pDados.add pPK & "_" & idItem & "|" & "Item_CompraNFItem" & "|" & idItem
        pDados.add pPK & "_" & idItem & "|" & "ChvAcesso_CompraNF" & "|" & pChvAcesso

        For Each Item In carregarParametros(qrySelecaoDeCampos, "tblCompraNFItem")
            Select Case UBound(Split((Item), "|"))

                '' ITEM DE COMPRA
                Case 1
                    regiao = Split((Item), "|")(0)
                    campo = Split((Item), "|")(1)
                    valor = XDoc.SelectNodes(regiao).Item(i).SelectNodes(campo).Item(0).text
                    If valor <> "" Then pDados.add pPK & "_" & idItem & "|" & campo & "|" & valor

                '' IMPOSTO
                Case 2
                    regiao = Split((Item), "|")(0)
                    subRegiao = Split((Item), "|")(1)
                    campo = Split((Item), "|")(2)
                    valor = XDoc.SelectNodes(regiao).Item(i).SelectNodes(subRegiao).Item(0).getElementsByTagName(campo).Item(0).text
                    If valor <> "" Then pDados.add pPK & "_" & idItem & "|" & campo & "|" & valor
                    
                Case Else
            End Select

            regiao = ""
            subRegiao = ""
            campo = ""
            valor = ""

            DoEvents

        Next Item

        DoEvents

    Next i

    '' CADASTRAR DADOS
    s.cadastroProcessamento pDados
    
    '' LIMPAR COLEÇÃO
    ClearCollection pDados

    '' ATUALIZAR CAMPOS DE RELACIONAMENTOS
    Application.CurrentDb.Execute Replace(qryUpdateProcessamento, "strParametro", "tblCompraNFItem")
    
    '' TRANSFERIR PROCESSADOS
'    s.TransferirDadosProcessados "tblCompraNFItem"

Set XDoc = Nothing

End Function


'' ############################################################################################################################


'Function proc_01()
'On Error GoTo adm_Err
'Dim s As New clsConexaoNfeCte
'
''' #ANALISE_DE_PROCESSAMENTO
'
''' #ADMINISTRACAO - RESPONSAVEL POR TRAZER OS DADOS DO SERVIDOR PARA AUXILIO NO PROCESSAMENTO. QUANDO NECESSARIO
''    s.ADM_carregarDadosDoServidor
'
''' 01.CARREGAR DADOS GERAIS - CONCLUIDO
''    s.carregar_DadosGerais
'
''' 02.CARREGAR COMPRAS ANTES DE ENVIAR PARA O SERVIDOR
'    s.carregar_Compras
'
''' 03.ENVIAR DADOS PARA SERVIDOR
''    s.enviar_ComprasParaServidor
'
'    MsgBox "Fim!", vbOKOnly + vbExclamation, "proc_01"
'
'
'adm_Exit:
'    Exit Function
'
'adm_Err:
'    MsgBox Error$
'    Resume adm_Exit
'
'End Function

