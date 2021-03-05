# prj_Proparts_conexa_nfe

Leitura de arquivos do tipo xml (NF-e ou CT-e) para importação em banco de dados e criação de dois arquivos do tipo json dos seguintes modelos: Atualizada no ERP (lançada) e Envio do manifesto pelo ERP.


## Alguns Links
[MSXML 4.0 Service Pack 3 (Microsoft XML Core Services)](https://www.microsoft.com/en-us/download/details.aspx?id=15697)

[Aplicativo da Manifestação do Destinatário](http://www.mdehom.fazenda.sp.gov.br/docs/manual.pdf)

https://www.nfe.fazenda.gov.br/portal/listaConteudo.aspx?tipoConteudo=33ol5hhSYZk=

[The Markdown (Basic Syntax)](https://www.markdownguide.org/basic-syntax/)


## Tags

#### #Previsão
#### #Replanejado
#### #EmDesenvolvimento
#### #EmTestes
#### #Concluido




# To-Do

## 01 - 23/02 - Coleta - 23/02 - EmTestes

- [ ] Identificação e relacionamento de campos do arquivo com os campos da tabela
- [ ] Origem/Destino - Processamento e movimentação de arquivos para pastas (Processados e/ou Sem Identificação).


## 02 - 24/02 - Arquivo - Em desenvolvimento - Previsão: 02/03 - Replanejado: [09/03]

- [ ] Criação de tabelas auxiliares ( tblDadosConexaoNFeCTe e tblTipos )
	- [ ] tblDadosConexaoNFeCTe  | Dados coletados dos arquivos tipo xml;
	- [ ] tblTipos | Responsável pela identificação de arquivos processados.
	      
- [ ] Envio de dados coletados para as tabelas ( tblDadosConexaoNFeCTe, tblCompraNF e tblCompraNFItem ).


## 03 - 25/03 - Extração - 01/03 - Previsão: 02/03 - Replanejado: [09/03]

- [X] Atualizada no ERP (Lançada) - Concluido - 04/03;
- [ ] Envio do manifesto pelo ERP - EmTestes - 05/03;
	- [ ] 


## 04 - 26/02 - Envio - 02/03 - Previsão: 02/03 - Replanejado: [09/03]

- [ ] Atualizada no ERP (Lançada);
- [ ] Envio do manifesto pelo ERP.


## 05 - 01/03 - Implantação e testes - Previsão: 01/03 - Replanejar

- [ ] Implantação e Testes com massa de dados ;
- [ ] Envio do manifesto pelo ERP.


## 06 - 01/03 - Validação e Entrega - Previsão: 01/03 - Replanejar

- [ ] Validação do dados processados atraves de evidencias do fluxo (Inicio - Fim) ;
- [ ] Entrega do projeto com documentação enexo ao códigos para futuras pesquisas e melhorias caso necessário.

