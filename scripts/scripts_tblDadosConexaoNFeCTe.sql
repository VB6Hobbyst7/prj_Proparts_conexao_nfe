-- '' ########################################
-- '' #tblDadosConexaoNFeCTe - #DADOS_GERAIS
-- '' ########################################
-- '' tblDadosConexaoNFeCTe.Pendentes
-- qrySelecaoDeArquivosPendentes
-- '' tblDadosConexaoNFeCTe.Pendentes
-- 

SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo
FROM tblDadosConexaoNFeCTe
WHERE (
		((tblDadosConexaoNFeCTe.registroValido) = 1)
		AND ((tblDadosConexaoNFeCTe.registroProcessado) = 0)
		)
	AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0)
ORDER BY tblDadosConexaoNFeCTe.CaminhoDoArquivo;

-- 
-- 
-- #######################################################################################################################


