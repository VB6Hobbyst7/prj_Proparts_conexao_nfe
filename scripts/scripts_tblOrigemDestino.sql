-- '' ########################################
-- '' #tblOrigemDestino
-- '' ########################################

-- '' #tblOrigemDestino
-- azsUpdateOrigemDestino_tabela_campo

UPDATE tblOrigemDestino
SET tblOrigemDestino.tabela = strSplit([Destino], '.', 0)
	,tblOrigemDestino.campo = strSplit([Destino], '.', 1);

-- #######################################################################################################################

-- '' -- CARREGAR TAGs DE VINDAS DO XML
-- '' #tblOrigemDestino.tabela
-- qryTags

SELECT tblOrigemDestino.Tag
FROM tblOrigemDestino
WHERE (
		((Len([Tag])) > 0)
		AND ((tblOrigemDestino.tabela) = 'strParametro')
		AND ((tblOrigemDestino.TagOrigem) = 1)
		);

-- #######################################################################################################################
