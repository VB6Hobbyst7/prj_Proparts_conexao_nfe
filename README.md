# prj_Proparts_conexa_nfe

Leitura de arquivos do tipo xml (NF-e ou CT-e) para importação em banco de dados e criação de dois arquivos do tipo json dos seguintes modelos: Atualizada no ERP (lançada) e Envio do manifesto pelo ERP.


To-Do

# 01 - 23/02 - Coleta - OK

- [ ] Identificação e relacionamento de campos do arquivo com os campos da tabela
- [ ] Origem/Destino - Processamento e movimentação de arquivos para pastas (Processados e/ou Sem Identificação).


# 02 - 24/02 - Arquivo - Em desenvolvimento - previsão: 02/03

- [ ] Criação de tabelas auxiliares ( tblDadosConexaoNFeCTe e tblTipos )
	- [ ] tblDadosConexaoNFeCTe  | Dados coletados dos arquivos tipo xml;
	- [ ] tblTipos | Responsável pela identificação de arquivos processados.
	      
- [ ] Envio de dados coletados para as tabelas ( tblDadosConexaoNFeCTe, tblCompraNF e tblCompraNFItem ).


# 03 - 25/03 - Extração - 01/03 - previsão: 02/03

- [ ] Atualizada no ERP (Lançada);
- [ ] Envio do manifesto pelo ERP.


# 04 - 26/02 - Envio - 02/03 - previsão: 02/03

- [ ] Atualizada no ERP (Lançada);
- [ ] Envio do manifesto pelo ERP.


# 05 - 01/03 - Implantação e testes - previsão: 01/03

- [ ] Implantação e Testes com massa de dados ;
- [ ] Envio do manifesto pelo ERP.


# 06 - 01/03 - Validação e Entrega - previsão: 01/03

- [ ] Validação do dados processados atraves de evidencias do fluxo (Inicio - Fim) ;
- [ ] Entrega do projeto com documentação enexo ao códigos para futuras pesquisas e melhorias caso necessário.

