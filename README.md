~~~~
#Previsão
#Replanejado
#EmPausa
#EmDesenvolvimento
#EmTestes
#Concluido
~~~~
___

# Proparts ( Conexão NF-e )

Leitura de arquivos do tipo xml (NF-e ou CT-e) para importação em banco de dados e criação de dois arquivos do tipo json dos seguintes modelos: Atualizada no ERP (lançada) e Envio do manifesto pelo ERP.


## Alguns Links
[Arquivos Integração com ERP](http://docs.conexaonfe.com.br/arquivos-integracao/#envio-do-manifesto-pelo-erp)

[Aplicativo da Manifestação do Destinatário](http://www.mdehom.fazenda.sp.gov.br/docs/manual.pdf)

https://www.nfe.fazenda.gov.br/portal/listaConteudo.aspx?tipoConteudo=33ol5hhSYZk=



[MSXML 4.0 Service Pack 3 (Microsoft XML Core Services)](https://www.microsoft.com/en-us/download/details.aspx?id=15697)

[The Markdown (Basic Syntax)](https://www.markdownguide.org/basic-syntax/)
[Basic writing and formatting syntax](https://docs.github.com/pt/github/writing-on-github/basic-writing-and-formatting-syntax)





# To-Do

###### 01 - 23/02 - Coleta - 23/02 - EmTestes - Previsão: 02/03 - Replanejado: [09/03]

	* Regras de negócio

		- [TESTES] Dados Gerais
		- [AJUSTES] Compras e Itens da compra
		- [PENDENTE] Origem x Destino - ( Mapeamento de campos do arquivo com os campos da tabela )
		- [X] Parametros - ( Variaveis de ambiente )
		- [X] Tipos - ( Controle de registros )

	* Cadatro de Variaveis de ambiente

		* Processamento de Arquivos
			- [X] Pendente
			- [X] Processamento
			- [X] Transferir
			- [X] Rejeitado

		* Controle de usuário
			- [X] UsuarioNome
			- [X] UsuarioCodigo

###### 02 - 24/02 - Arquivo - 24/02 - EmDesenvolvimento - Previsão: 02/03 - Replanejado: [09/03]

	* Criar Tabelas Auxiliares

		- [X] Dados Gerais
		- [X] Compras e Itens da compra
		- [X] Origem x Destino
		- [X] Parametros
		- [X] Tipos

	* Criar Consultas para cadastro ( Usadas pelos módulos )

		- [X] Dados Gerais
		- [X] Compras e Itens da compra
		- [X] Origem x Destino
		- [X] Parametros
		- [X] Tipos

	* Módulos ( Regras de negócios X Consultas )

		- [X] Dados Gerais
		- [ ] Compras e Itens da compra
		- [X] Origem x Destino
		- [X] Parametros
		- [X] Tipos


###### 03 - 25/03 - Extração - 25/03 - EmDesenvolvimento - 01/03 - Previsão: 02/03 - Replanejado: [09/03]

	* Atualizada no ERP (Lançada) - Concluido - 04/03;
	* Envio do manifesto pelo ERP - EmTestes - 05/03;

###### 04 - 26/02 - Envio - 26/02 - EmDesenvolvimento - 02/03 - Previsão: 02/03 - Replanejado: [09/03]

	* Atualizada no ERP (Lançada) - Concluido - 04/03;
	* Envio do manifesto pelo ERP - EmTestes - 05/03;


###### 05 - 01/03 - Implantação e testes - 01/03 - EmPausa - Previsão: 01/03 - Replanejar

	* Implantação e Testes com massa de dados ;
	* Envio do manifesto pelo ERP.

###### 06 - 01/03 - Validação e Entrega - 01/03 - EmPausa - Previsão: 01/03 - Replanejar

	* Validação do dados processados atraves de evidencias do fluxo (Inicio - Fim) ;
	* Entrega do projeto com documentação enexo ao códigos para futuras pesquisas e melhorias caso necessário.

