# Relação de funções concluidas da aplicação:

* (Concluído) 1. Não reprocessar o mesmo arquivo - ( Controle por nome de arquivo )
	- (Concluído) 1.1 Endentimento de principais pastas:
	- (Concluído) 1.1.001 Empresa/Recebimento - Responsável pelo recebimento de arquivos para inicio do processamento.
	- (Concluído) 1.1.002 Processados - Todos os arquivos processados sem erros
	- (Concluído) 1.1.003 Expurgo 	- Todos os arquivos com alguma pendencia de relacionamento(s) de dados ou invalido
	- (Concluído) 1.1.004 Ações		- Todos os arquivos Json's ( Pasta de ações (Lançada ERP e Manifesto de Confirmação) ) conforme documentação do site

* (Concluído) 2. Controle de registro processado com fluxo de movimentação de arquivos por pasta
	- (Concluído) 2.1 Endentimento de fluxo (registroProcessado):
	- (Concluído) 2.1.1 [0] - Pendente de processamento
	- (Concluído) 2.1.2 [1] - Registro OK
	- (Concluído) 2.1.3 [2] - Enviado para servidor
	- (Concluído) 2.1.4 [3] - Mover para pasta de processados
	- (Concluído) 2.1.5 [4] - Mover para pasta de expurgo
	- (Concluído) 2.1.6 [9] - Processamento Finalizado

* (Concluído) 3. Controle de dados gerais
	- (Concluído) 3.1.001 FiltroFil_DestinatarioNaoProparts
	- (Concluído) 3.1.002 FiltroFil_DestinatarioProparts
	- (Concluído) 3.1.003 FiltroFil_DestinatarioProparts_55
	- (Concluído) 3.1.004 FiltroFil_RemetenteProparts
	- (Concluído) 3.1.005 NumPed_CompraNF
	- (Concluído) 3.1.006 Update_Sit_CompraNF
	- (Concluído) 3.1.007 Update_IdFornCompraNF
	- (Concluído) 3.1.008 Update_ID_NatOp_CompraNF__FiltroCFOP
	- (Concluído) 3.1.009 Update_IDVD
	- (Concluído) 3.1.010 FornecedoresValidos
	- (Concluído) 3.1.011 Update_IdTipo
	- (Concluído) 3.1.012 Update_IdTipo_CTe
	- (Concluído) 3.1.013 Update_IdTipo_RetornoArmazem
	- (Concluído) 3.1.014 Update_IdTipo_RetornoArmazem_CFOP
	- (Concluído) 3.1.015 Update_IdTipo_TransferenciaSisparts
	- (Concluído) 3.1.016 Update_IdTipo_TransferenciaSisparts_CFOP
	- (Concluído) 3.1.017 Update_Sit_CompraNF
	- (Concluído) 3.1.018 Update_IdFornCompraNF
	- (Concluído) 3.1.019 Update_ProcessamentoConcluído
	- (Concluído) 3.1.020 Update_ProcessamentoConcluído_CTE
	- (Concluído) 3.1.021 Update_ProcessamentoConcluído_Servidor
	- (Concluído) 3.1.022 Update_FornecedoresValidos
	- (Concluído) 3.1.023 Update_RegistrosValidos
	- (Concluído) 3.1.024 Select_ArquivosPendentes
	- (Concluído) 3.1.025 Delete_RegistrosInvalidos - [Pausado]

* (Concluído) 4. Controle Notas Fiscais e Itens
	- (Concluído) 4.1.001 RegistroValidoPorcessado
	- (Concluído) 4.1.002 RegistroConcluído
	- (Concluído) 4.1.003 Update_NumPed_Contador
	- (Concluído) 4.1.004 Update_IDCompraNF
	- (Concluído) 3.1.005 Update_AjustesCampos_LOCAL
	- (Concluído) 4.1.006 Update_Transferencia_DadosGerais_Para_Compras
	- (Concluído) 4.1.007 Update_CFOP_CompraNF
	- (Concluído) 4.1.008 Update_Dados_ID_Prod_CompraNFItem
	- (Concluído) 4.1.009 Insert_Dados_CTeItens
	- (Concluído) 4.1.010 AjusteDeCampos_CTe
	- (Concluído) 4.1.011 qryComprasItens_Update_CFOP_CompraNF
	- (Concluído) 4.1.012 ID_Prod_CompraNFItem

* (Concluído) 5. Controle De Parametros do sistema
	- (Concluído) 5.1.001 Relacionamento do campo FiltroFil
	- (Concluído) 5.1.002 registroProcessado
	- (Concluído) 5.1.003 registroValido
	- (Concluído) 5.1.004 Login e senha do Banco(Servidor) - Exportação de dados
	- (Concluído) 5.1.005 Caminhos de Coleta de arquivos
	- (Concluído) 5.1.006 Identificação de Robô para Gerar Json's de Lançada ERP e Manifesto de Confirmação
	- (Concluído) 5.1.007 Identificação Tipos de evento para uso da função "Gerar Json's de Lançada ERP e Manifesto de Confirmação"
	- (Concluído) 5.1.008 Link's de Documentação de sistema "conexaonfe"
	- (Concluído) 5.1.009 Controle de Campos para exportação de dados para o Banco(Servidor)
	- (Concluído) 5.1.010 Controle de Processamento de registros
	- (Concluído) 5.1.011 Controle de Origem de dados coletados ( Exemplo: Xml, Table, Query e Nulo )
	- (Concluído) 5.1.012 Relação de tipos de cadastro
	- (Concluído) 5.1.013 Controle de log
	- (Concluído) 5.1.014 Controle interno do sistema (system)

* (Concluído) 6. Identificação de tabelas ordenadas por grau de importancia para o sistema
	- (Concluído) 6.1.001 tblParametros		|	Controle de parametros internos 
	- (Concluído) 6.1.001 tblOrigemDestino	|	Relacionamento de campos do arquivo com as tabelas finais
	- (Concluído) 6.1.001 tblTipos			| 	Controle de tipos de Notas Fiscais
	- (Concluído) 6.1.001 tblProcessamento	|	Repositorio de transito de dados coletados para tratamento e futura transição para tabelas finais
	- (Concluído) 6.1.001 tblDadosConexaoNFeCTe	| Repositorio de dados gerais para controle de arquivos processados de suas paras originais
	- (Concluído) 6.1.001 tblConexao			|	Repositorio para auxilio de vinculo de tabelas/Consultas ligadas ao Banco(Servidor)
	- (Concluído) 6.1.001 tblCompraNF			| Repositorio de transito de dados coletados para tratamento e futura transição para tabelas Compras
	- (Concluído) 6.1.001 tblCompraNFItem		| Repositorio de transito de dados coletados para tratamento e futura transição para tabelas Compras
