Attribute VB_Name = "modConexaoNfeCte"
Option Compare Database




Function proc_01()
On Error GoTo adm_Err
Dim s As New clsConexaoNfeCte

'' #ANALISE_DE_PROCESSAMENTO

'' #ADMINISTRACAO - RESPONSAVEL POR TRAZER OS DADOS DO SERVIDOR PARA AUXILIO NO PROCESSAMENTO. QUANDO NECESSARIO
'    s.ADM_carregarDadosDoServidor

'' 01.CARREGAR DADOS GERAIS - CONCLUIDO
'    s.carregar_DadosGerais

'' 02.CARREGAR COMPRAS ANTES DE ENVIAR PARA O SERVIDOR
    s.carregar_Compras

'' 03.ENVIAR DADOS PARA SERVIDOR
'    s.enviar_ComprasParaServidor
    
    MsgBox "Fim!", vbOKOnly + vbExclamation, "proc_01"
    

adm_Exit:
    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit

End Function


Function proc_TESTE()
On Error GoTo adm_Err
Dim s As New clsConexaoNfeCte


    
    '' REGISTRO OK
    s.carregar_ComprasItens "C:\temp\Coleta\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\42210312680452000302550020000895841453583169-nfeproc.xml"
    
    
'    s.carregar_ComprasItensPorConsulta
    
'    s.TransferirDadosProcessados "tblCompraNFItem"

'    s.TransferirDadosProcessados_v02 "tblCompraNFItem"
    
    MsgBox "Fim!", vbOKOnly + vbExclamation, "proc_01"
    

adm_Exit:
    Exit Function

adm_Err:
    MsgBox Error$
    Resume adm_Exit

End Function
