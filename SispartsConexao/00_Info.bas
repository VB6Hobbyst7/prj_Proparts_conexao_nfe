Attribute VB_Name = "00_Info"
Option Compare Database
'' ### TO-DO ###
''
'' #20210823_BaseCalcICMS_CompraNF
'' #20210823_VTotICMS_CompraNF

'' #20210823_CadastroDeComprasEmServidor
'' #20210823_qryUpdateNumPed_CompraNF
'' #20210823_FornecedoresValidos

'' ### DONE ###
''
'' Consultas
'' #20210823_EXPORTACAO_LIMITE
'' #20210823_qryDadosGerais_Update_ID_NatOp_CompraNF__FiltroCFOP -- FiltroCFOP
'' #20210823_qryDadosGerais_Update_IDVD
'' #20210823_qryUpdateID_NatOp_CompraNF
'' #20210823_qryUpdateCFOP_FilCompra
'' #20210823_qryUpdate_ModeloDoc_CFOP
'' #20210823_qryUpdateFilCompraNF
'' #20210823_qryDadosGerais_Update_IdFornCompraNF
'' #20210823_qryDadosGerais_Update_Sit_CompraNF
'' #20210823_XML_CONTROLE | Quando importar cada XML, precisa recortar o arquivo da pasta da empresa e colar dentro de uma pasta chama “Processados”, porém dentro de cada pasta de cada empresa, pois não podemos misturar os XML´s de cada empresa.
'' #20210823_XML_FORMULARIO | Não encontrei um formulário com os XML´s que não foram processados e o motivo. | <<< ATENÇÃO - NÃO DEFINIMOS COMO CLASSIFICAREMOS OS MOTIMOS DE NÃO PROCESSAMENTO DE ARQUIVOS >>>
'' #20210823_VTotProd_CompraNF
'' #20210823_ID_Prod_CompraNFItem


'' LIMPAR TODA A BASE DE DADOS
Public Const dataBaseClear As Boolean = True

'' REPROCESSAR ARQUIVOS PENDENTES
Public Const dataBaseReplay As Boolean = False

'' EXPORTAR DADOS PARA SERVIDOR
Public Const dataBaseExportarDados As Boolean = False

'' PROCESSAMENTO DE ARQUIVOS
Public Const dataBaseTratamentoDeArquivos As Boolean = False
Public Const dataBaseGerarLancamentoManifesto As Boolean = False

Sub teste_FuncionamentoGeralDeProcessamentoDeArquivos()
Dim strCaminhoAcoes As String: strCaminhoAcoes = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColetaAcoes'")
    
    ''==================================================
    '' REPOSITORIO GERAL
    ''==================================================

    '' LIMPAR TODA A BASE DE DADOS
    If dataBaseClear Then
    
        '' Limpar toda a base de dados
        Application.CurrentDb.Execute "Delete from tblDadosConexaoNFeCTe"

        '' Limpar repositorio de itens de compras
        Application.CurrentDb.Execute _
                "Delete from tblCompraNFItem"
    
        '' Limpar repositorio de compras
        Application.CurrentDb.Execute _
                "Delete from tblCompraNF"

        '' Carregar todos os arquivos para processamento.
        processarDadosGerais
        
    Else
        
        '' Carregar todos os arquivos para processamento.
        processarDadosGerais
    
    
    End If

    ''==================================================
    '' REPOSITORIOS DE COMPRAS
    ''==================================================
    
    '' REPROCESSAR ARQUIVOS VALIDOS
    If dataBaseReplay Then
    
        '' Ajustar marcação de registro
        Application.CurrentDb.Execute _
            "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado=0 WHERE tblDadosConexaoNFeCTe.registroValido=1 AND tblDadosConexaoNFeCTe.ID_Tipo>0"
        
        '' Limpar repositorio de itens de compras
        Application.CurrentDb.Execute _
                "Delete from tblCompraNFItem"
    
        '' Limpar repositorio de compras
        Application.CurrentDb.Execute _
                "Delete from tblCompraNF"

        '' Processamento de arquivos pendentes da pasta de coleta.
        processarArquivosPendentes
            
    Else
    
        '' Processamento de arquivos pendentes da pasta de coleta.
        processarArquivosPendentes
    
    End If


    ''==================================================
    '' EXPORTAR DADOS PARA O SERVIDOR
    ''==================================================

    '' EXPORTAÇÃO DE DADOS
    If dataBaseExportarDados Then _
            CadastroDeComprasEmServidor

    ''==================================================
    '' PROCESSAMENTO DE ARQUIVOS
    ''==================================================

    '' #### TRANSFERENCIAS DE ARQUIVOS
    If dataBaseTratamentoDeArquivos Then _

        '' Transferir Arquivos Validos para pasta de processados
        tratamentoDeArquivosValidos
    
        '' Transferir Arquivos Invalidos para pasta de Expurgo
        tratamentoDeArquivosInvalidos

    End If

    '' #### GERAR ARQUIVOS DE LANÇAMENTO E MANIFESTO
    If dataBaseGerarLancamentoManifesto Then
    
        '' LANÇAMENTO
        gerarArquivosJson opFlagLancadaERP, , strCaminhoAcoes
    
        '' MANIFESTO
        gerarArquivosJson opManifesto, , strCaminhoAcoes
        
    End If
    
Debug.Print "### Concluido! - testeDeFuncionamentoGeral"
TextFile_Append CurrentProject.path & "\" & strLog(), "Concluido! - testeDeFuncionamentoGeral"

End Sub

