# prj_Proparts_conexa_nfe

Leitura de arquivos do tipo xml (NF-e ou CT-e) para importação em banco de dados e criação de dois arquivos do tipo json dos seguintes modelos: Atualizada no ERP (lançada) e Envio do manifesto pelo ERP.


To-Do

#01 - Coleta - Coleta de dados em arquivos do tipo xml

- [ ] Identificação e relacionamento de campos do arquivo com os campos da tabela
- [ ] Origem/Destino - Processamento e movimentação de arquivos para pastas 
      (Processados e/ou Sem Identificação).
      
#02 - Arquivo - Cadastro de dados coletados em tabelas indicadas

- [ ] Criação de tabelas auxiliares ( tblDadosConexaoNFeCTe e tblTipos )
	- [ ] tblDadosConexaoNFeCTe  | Dados coletados dos arquivos tipo xml;
	- [ ] tblTipos | Responsável pela identificação de arquivos processados.
	      
- [ ] Envio de dados coletados para as tabelas ( tblDadosConexaoNFeCTe, tblCompraNF e tblCompraNFItem ).
      
#03 - Extração - Criação de arquivos tipo Json com dados vindos da (01 - Coleta)

- [ ] Atualizada no ERP (Lançada);
- [ ] Envio do manifesto pelo ERP.
      
#04 - Envio - Transferencia de arquivos vindos da (03 - Extração)

- [ ] Atualizada no ERP (Lançada);
- [ ] Envio do manifesto pelo ERP.

#05 - Implantação e Testes

- [ ] Implantação e Testes com massa de dados ;
- [ ] Envio do manifesto pelo ERP.

#06 - Validação e Entrega

- [ ] Validação do dados processados atraves de evidencias do fluxo (Inicio - Fim) ;
- [ ] Entrega do projeto com documentação enexo ao códigos para futuras pesquisas e melhorias caso necessário.

