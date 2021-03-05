# prj_Proparts_conexa_nfe

Leitura de arquivos do tipo xml (NF-e ou CT-e) para importação em banco de dados e criação de dois arquivos do tipo json dos seguintes modelos: Atualizada no ERP (lançada) e Envio do manifesto pelo ERP.


#### #Previsão
#### #Replanejado
#### #Concluido
#### #EmTestes



# To-Do

## 01 - 23/02 - Coleta - 23/02 - ** CONCLUIDO **

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

