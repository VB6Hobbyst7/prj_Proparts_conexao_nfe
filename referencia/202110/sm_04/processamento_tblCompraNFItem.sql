-- qryUpdateItens_CFOP_CompraNFItem
UPDATE (
		tblDadosConexaoNFeCTe INNER JOIN tblCompraNF ON tblDadosConexaoNFeCTe.ChvAcesso = tblCompraNF.ChvAcesso_CompraNF
		)
INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF

SET tblCompraNFItem.CFOP_CompraNFItem = [tblCompraNF].[CFOP_CompraNF]
WHERE (
		((tblDadosConexaoNFeCTe.registroValido) = 1)
		AND ((tblDadosConexaoNFeCTe.registroProcessado) = 0)
		AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0)
		);

-- qryUpdateItens_ID_Prod_CompraNFItem
UPDATE tblCompraNFItem
SET tblCompraNFItem.ID_Prod_CompraNFItem = DLookUp("CodigoProd_Grade", "dbo_tabGradeProdutos", "CodigoForn_Grade='" & [tblCompraNFItem].[ID_Prod_CompraNFItem] & "'");
