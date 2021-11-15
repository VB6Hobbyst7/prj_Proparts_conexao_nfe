/*

	01 CARREGAR DE DADOS

*/
-- qrySelecaoDeCampos
SELECT tblOrigemDestino.Tag
FROM tblOrigemDestino
WHERE (
		((tblOrigemDestino.tabela) = 'strParametro')
		AND ((Len([Tag])) > 0)
		AND ((tblOrigemDestino.TagOrigem) = 1)
		)
ORDER BY tblOrigemDestino.Tag
	,tblOrigemDestino.tabela;

-- qryDeleteProcessamento
DELETE *
FROM tblProcessamento;

-- qryUpdateProcessamento_Chave
UPDATE tblProcessamento
SET tblProcessamento.chave = Replace([tblProcessamento].[chave], ';', '|');

-- qryUpdateProcessamento_NomeTabela
UPDATE tblProcessamento
SET tblProcessamento.NomeTabela = "" tblRepositorio ""
WHERE tblProcessamento.NomeTabela IS NULL;

-- qryUpdateProcessamento_LimparItensMarcadosErrados
UPDATE tblProcessamento
SET tblProcessamento.NomeTabela = NULL
WHERE (((classificacao([tblProcessamento].[pk])) = 1));

-- qryUpdateProcessamento_NomeCampo
UPDATE tblProcessamento
SET tblProcessamento.NomeCampo = DLookUp("" campo "", "" tblOrigemDestino "", "" tag = '"" & [tblProcessamento].[chave] & ""'
		AND Tabela = 'tblRepositorio' "")
WHERE (((tblProcessamento.NomeTabela) = "" tblRepositorio ""));

-- qryUpdateProcessamento_Formatacao
UPDATE tblProcessamento
SET tblProcessamento.formatacao = DLookUp("" formatacao "", "" tblOrigemDestino "", "" tag = '"" & [tblProcessamento].[chave] & ""'
		AND Tabela = 'tblRepositorio' "")
WHERE (((tblProcessamento.NomeTabela) = "" tblRepositorio ""));

-- qryUpdateProcessamento_RelacaoCamposDeTabelas_Item_CompraNFItem
UPDATE tblProcessamento
SET tblProcessamento.NomeTabela = "" tblCompraNFItem ""
	,tblProcessamento.NomeCampo = [tblProcessamento].[chave]
	,tblProcessamento.formatacao = DLookUp("" formatacao "", "" tblOrigemDestino "", "" campo = 'Item_CompraNFItem' "")
WHERE (((tblProcessamento.chave) = "" Item_CompraNFItem ""));

-- qryUpdateProcessamento_RelacaoCamposDeTabelas_ChvAcesso_CompraNF
UPDATE tblProcessamento
SET tblProcessamento.NomeTabela = "" tblCompraNF ""
	,tblProcessamento.NomeCampo = [tblProcessamento].[chave]
	,tblProcessamento.formatacao = "" opTexto ""
WHERE (((tblProcessamento.chave) = "" ChvAcesso_CompraNF ""));

-- qryUpdateProcessamento_RelacaoCamposDeTabelas_tblCompraNFItem_ChvAcesso_CompraNF
UPDATE tblProcessamento
SET tblProcessamento.NomeTabela = strSplit([tblProcessamento].[chave], '.', 0)
	,tblProcessamento.NomeCampo = strSplit([tblProcessamento].[chave], '.', 1)
	,tblProcessamento.formatacao = strSplit([tblProcessamento].[chave], '.', 2)
WHERE (((tblProcessamento.chave) = "" tblCompraNFItem.ChvAcesso_CompraNF.opTexto ""));

-- qryUpdateProcessamento_opData
UPDATE tblProcessamento
SET tblProcessamento.valor = Mid([tblProcessamento].[valor], 1, 10)
WHERE formatacao = 'opData';

-- qryUpdateProcessamento_opTime
UPDATE tblProcessamento
SET tblProcessamento.valor = Mid([tblProcessamento].[valor], 12, 8)
WHERE formatacao = 'opTime';
