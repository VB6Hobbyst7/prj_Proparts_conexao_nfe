-- tblCompraNFItem.ID_Prod_CompraNFItem

UPDATE tblDadosConexaoNFeCTe
INNER JOIN (
	tblCompraNF INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF
	) ON tblDadosConexaoNFeCTe.ChvAcesso = tblCompraNF.ChvAcesso_CompraNF

SET tblCompraNFItem.ID_Prod_CompraNFItem = DLookUp("CodigoProd_Grade", "tmpGradeProdutos", "CodigoForn_Grade='" & [tblCompraNFItem].[ID_Prod_CompraNFItem] & "'")
WHERE (
		((tblDadosConexaoNFeCTe.registroValido) = 1)
		AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1)
		);


-- tblCompraNFItem.CFOP_CompraNFItem
UPDATE (
	tblDadosConexaoNFeCTe INNER JOIN tblCompraNF ON tblDadosConexaoNFeCTe.ChvAcesso = tblCompraNF.ChvAcesso_CompraNF
		)
INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF

SET tblCompraNFItem.CFOP_CompraNFItem = [tblCompraNF].[CFOP_CompraNF]
WHERE (
		((tblDadosConexaoNFeCTe.registroValido) = 1)
		AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1)
		AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0)
		);


-- tblCompraNF.ID_Forn_CompraNF
UPDATE (
	SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF
			,tmpClientes.CÓDIGOClientes
		FROM tmpClientes
		) AS qryClientesFornecedor
INNER JOIN tblCompraNF ON tblCompraNF.CNPJ_CPF_CompraNF = qryClientesFornecedor.strCNPJ_CPF

SET tblCompraNF.ID_Forn_CompraNF = qryClientesFornecedor.CÓDIGOClientes;


-- tmpIDVD_CompraNF.IDVD_CompraNF
UPDATE (
	SELECT tblCompraNF.IDVD_CompraNF
			,((Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6))) AS strIDVD_CompraNF
			,[tblCompraNF].[Obs_CompraNF]
		FROM tblCompraNF
		WHERE (
				((Left([Obs_CompraNF], 6)) = 'PEDIDO ')
				AND ((tblCompraNF.CNPJ_CPF_CompraNF) = '12680452000302')
				)
		) AS tmpIDVD_CompraNF
SET tmpIDVD_CompraNF.IDVD_CompraNF = tmpIDVD_CompraNF.strIDVD_CompraNF;




UPDATE (
		SELECT ((Left(Trim(Replace(Replace(tblDadosConexaoNFeCTe.IDVD_CompraNF, 'PEDIDO: ', ''), 'PEDIDO ', '')), 6))) AS strIDVD_CompraNF, tblDadosConexaoNFeCTe.IDVD_CompraNF
		FROM tblDadosConexaoNFeCTe
		WHERE (((Left(tblDadosConexaoNFeCTe.IDVD_CompraNF, 6)) = 'PEDIDO ') AND ((tblDadosConexaoNFeCTe.CNPJ_emit) = '12680452000302'))
		) AS tmpIDVD_CompraNF
SET tmpIDVD_CompraNF.IDVD_CompraNF = tmpIDVD_CompraNF.strIDVD_CompraNF;
