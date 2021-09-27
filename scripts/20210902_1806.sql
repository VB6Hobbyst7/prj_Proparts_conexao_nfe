-- #qryUpdateCFOP_FilCompra

UPDATE tblCompraNF 
	SET 
	tblCompraNF.CFOP_CompraNF = DLookUp("[FiltroCFOP]","[tblDadosConexaoNFeCTe]","[ChvAcesso]='" & [tblCompraNF].[ChvAcesso_CompraNF] & "'")
	, tblCompraNF.Fil_CompraNF = DLookUp("[ID_EMPRESA]","[tblDadosConexaoNFeCTe]","[ChvAcesso]='" & [tblCompraNF].[ChvAcesso_CompraNF] & "'");


-- #tblCompraNF.ID_Forn_CompraNF

UPDATE (
		SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF
			,tmpClientes.CÓDIGOClientes
		FROM tmpClientes
		) AS qryClientesFornecedor
INNER JOIN tblCompraNF ON tblCompraNF.CNPJ_CPF_CompraNF = qryClientesFornecedor.strCNPJ_CPF

SET tblCompraNF.ID_Forn_CompraNF = qryClientesFornecedor.CÓDIGOClientes;


