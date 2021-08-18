Sub TESTE_CARREGAR_REGISTRO_LISTA()
Dim s As New clsConexaoNfeCte
Dim Item As Variant
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()
Dim DT_PROCESSO_ITEM As Date
Dim contadorDeRegistros As Long: contadorDeRegistros = 1
SysCmd acSysCmdInitMeter, "Processando...", DCount("*", "tblDadosConexaoNFeCTe", "(((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=1) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0))")

    '' LIMPAR TABELA DE PROCESSAMENTOS
    Application.CurrentDb.Execute qryDeleteProcessamento

    '' CARREGAR_DADOS
    For Each Item In carregarParametros(qrySelecaoDeArquivosPendentes)
        DT_PROCESSO_ITEM = Now()

        SysCmd acSysCmdUpdateMeter, contadorDeRegistros
        s.ProcessamentoDeArquivo CStr(Item), opCompras
        contadorDeRegistros = contadorDeRegistros + 1

        statusFinal DT_PROCESSO_ITEM, "Processamento - " & CStr(Item)
        DoEvents
    Next Item

    '' FORMATAR CAMPOS
'    s.FormatarCamposEmProcessamento

    '' #TRANSFERIR DADOS PROCESSADOS - COMPRAS
    s.ProcessamentoTransferir "tblCompraNF"

    '' #TRANSFERIR DADOS PROCESSADOS - COMPRAS ITENS
    s.ProcessamentoTransferir "tblCompraNFItem"

    '' #TRATAMENTO
'    s.TratamentoDeCompras
'    s.compras_atualizarCampos


    SysCmd acSysCmdRemoveMeter
    statusFinal DT_PROCESSO, "Processamento - teste_listarArquivos"

End Sub