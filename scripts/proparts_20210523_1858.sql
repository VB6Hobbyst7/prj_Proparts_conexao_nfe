

-- qrySelectComprasItens
SELECT tblProcessamento.valor, *
FROM tblProcessamento
WHERE (((tblProcessamento.[NomeTabela])='tblCompraNFItem') AND ((tblProcessamento.valor)="35210343283811001202550010087454051410067364"))
ORDER BY tblProcessamento.ID;



-- compras

SELECT tblProcessamento.NomeCampo, tblProcessamento.valor
FROM tblProcessamento
WHERE (((tblProcessamento.NomeCampo)="ChvAcesso_CompraNF") AND ((tblProcessamento.NomeTabela)='tblCompraNF'))
ORDER BY tblProcessamento.ID;


-- itens compra

tblProcessamento.valor, *
FROM tblProcessamento
WHERE (((tblProcessamento.valor) In (SELECT tblProcessamento.valor as chave
FROM tblProcessamento
WHERE (((tblProcessamento.NomeCampo)="ChvAcesso_CompraNF") AND ((tblProcessamento.NomeTabela)='tblCompraNF'))
ORDER BY tblProcessamento.ID)) AND ((tblProcessamento.NomeTabela)='tblCompraNFItem'))
ORDER BY tblProcessamento.ID;



SELECT tblProcessamento.pk FROM tblProcessamento WHERE (((tblProcessamento.valor) In (SELECT tblProcessamento.valor as chave FROM tblProcessamento WHERE (((tblProcessamento.NomeCampo)="ChvAcesso_CompraNF") AND ((tblProcessamento.NomeTabela)='tblCompraNF')) ORDER BY tblProcessamento.ID)) AND ((tblProcessamento.NomeTabela)='tblCompraNFItem')) ORDER BY tblProcessamento.ID;
