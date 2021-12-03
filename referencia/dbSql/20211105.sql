-- tmpClientes
-- tmpEmpresa
-- tmpGradeProdutos
-- tmpNatOp
-- tmpProdutos



-- qryDadosGerais_Update_FornecedoresValidos
UPDATE (
		SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF
		FROM tmpClientes
		) AS qryFornecedoresValidos
INNER JOIN tblDadosConexaoNFeCTe ON qryFornecedoresValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit

SET tblDadosConexaoNFeCTe.registroValido = 1
WHERE ((tblDadosConexaoNFeCTe.registroProcessado) = 0);

-- qryDadosGerais_Update_RegistrosValidos
UPDATE (
		SELECT STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF
		FROM tmpEmpresa
		) AS qryRegistrosValidos
INNER JOIN tblDadosConexaoNFeCTe ON qryRegistrosValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit

SET tblDadosConexaoNFeCTe.registroValido = 1;

-- qryDadosGerais_Update_IdTipo
UPDATE tblDadosConexaoNFeCTe
SET tblDadosConexaoNFeCTe.ID_Tipo = 0
WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) IS NULL));

-- qryDadosGerais_Update_IdTipo_RetornoArmazem
UPDATE tblDadosConexaoNFeCTe
SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp("" [ValorDoParametro] "", "" [tblParametros] "", "" [TipoDeParametro] = 'RetornoArmazem' "")
WHERE (
		((tblDadosConexaoNFeCTe.ID_Tipo) = 0)
		AND ((tblDadosConexaoNFeCTe.codMod) = CInt('55'))
		AND ((tblDadosConexaoNFeCTe.CNPJ_emit) IN ('12680452000302'))
		);

-- qryDadosGerais_Update_IdTipo_TransferenciaSisparts
UPDATE tblDadosConexaoNFeCTe
SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp("" [ValorDoParametro] "", "" [tblParametros] "", "" [TipoDeParametro] = 'TransferenciaSisparts' "")
WHERE (
		((tblDadosConexaoNFeCTe.ID_Tipo) = 0)
		AND ((tblDadosConexaoNFeCTe.codMod) = CInt('55'))
		AND (
			(tblDadosConexaoNFeCTe.CNPJ_emit) IN (
				SELECT CNPJ_Empresa
				FROM [tmpEmpresa]
				)
			)
		);

-- qryDadosGerais_Update_IdTipo_CTe
UPDATE tblDadosConexaoNFeCTe
SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp("" [ValorDoParametro] "", "" [tblParametros] "", "" [TipoDeParametro] = 'CTe' "")
WHERE (
		((tblDadosConexaoNFeCTe.ID_Tipo) = 0)
		AND ((tblDadosConexaoNFeCTe.codMod) = CInt('57'))
		);

-- qryDadosGerais_Update_Sit_CompraNF
UPDATE (
		SELECT strSplit(ValorDoParametro, '|', 0) AS strFinalidade
			,strSplit(ValorDoParametro, '|', 1) AS strSit_CompraNF
		FROM tblParametros
		WHERE TipoDeParametro = 'Sit_CompraNF'
		) AS qrySit_CompraNF
INNER JOIN (
	SELECT *
	FROM tblDadosConexaoNFeCTe
	WHERE tblDadosConexaoNFeCTe.registroValido IN (
			SELECT TOP 1 cint(tblParametros.ValorDoParametro)
			FROM [tblParametros]
			WHERE TipoDeParametro = 'registroValido'
			)
	) AS tmpDadosConexaoNFeCTe ON (cint(qrySit_CompraNF.strFinalidade) = tmpDadosConexaoNFeCTe.ID_TIPO)

SET tmpDadosConexaoNFeCTe.Sit_CompraNF = [qrySit_CompraNF].[strSit_CompraNF];

-- qryDadosGerais_Update_IdEmpresa
UPDATE (
	SELECT tmpEmpresa.ID_Empresa
			,STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF
		FROM tmpEmpresa
		) AS qryEmpresas
INNER JOIN tblDadosConexaoNFeCTe ON qryEmpresas.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_Rem

SET tblDadosConexaoNFeCTe.ID_Empresa = qryEmpresas.ID_Empresa;

-- qryDadosGerais_Update_IdEmpresa_TransferenciaEntreFiliais
UPDATE (
		SELECT tmpEmpresa.ID_Empresa
			,STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF
		FROM tmpEmpresa
		) AS qryEmpresas
INNER JOIN tblDadosConexaoNFeCTe ON qryEmpresas.strCNPJ_CPF = tblDadosConexaoNFeCTe.CPNJ_Dest

SET tblDadosConexaoNFeCTe.ID_Empresa = qryEmpresas.ID_Empresa
WHERE (((tblDadosConexaoNFeCTe.CFOP) = '6152'));

-- qryDadosGerais_Update_IDVD
UPDATE tblDadosConexaoNFeCTe
SET tblDadosConexaoNFeCTe.IDVD_CompraNF = val((Left(Trim(Replace(Replace(tblDadosConexaoNFeCTe.Obs_CompraNF, 'PEDIDO: ', ''), 'PEDIDO ', '')), 6)))
WHERE (
		((Left(tblDadosConexaoNFeCTe.Obs_CompraNF, 6)) = 'PEDIDO ')
		AND ((tblDadosConexaoNFeCTe.CNPJ_emit) = '12680452000302')
		AND (
			((tblDadosConexaoNFeCTe.registroValido) = 1)
			AND ((tblDadosConexaoNFeCTe.registroProcessado) = 0)
			)
		AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0)
		);

-- qryDadosGerais_Update_IdFornCompraNF
UPDATE (
		SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF
			,tmpClientes.CÓDIGOClientes
		FROM tmpClientes
		) AS qryClientesFornecedor
INNER JOIN tblDadosConexaoNFeCTe ON tblDadosConexaoNFeCTe.CNPJ_emit = qryClientesFornecedor.strCNPJ_CPF

SET tblDadosConexaoNFeCTe.ID_Forn_CompraNF = qryClientesFornecedor.CÓDIGOClientes;

-- qryDadosGerais_Update_ID_NatOp_CompraNF__FiltroCFOP
UPDATE (
		SELECT tmpNatOp.ID_NatOper
			,tmpNatOp.Fil_NatOper
			,tmpNatOp.CFOP_NatOper
			,qryPscPes.strXMLCFOP
			,qryPscPes.strEstado
		FROM (
			SELECT strSplit(ValorDoParametro, '|', 0) AS strFil_NatOper
				,strSplit(ValorDoParametro, '|', 1) AS strEstado
				,strSplit(ValorDoParametro, '|', 2) AS strXMLCFOP
				,strSplit(ValorDoParametro, '|', 3) AS strCFOP_NatOper
			FROM tblParametros
			WHERE TipoDeParametro = 'FiltroFil'
				AND strSplit(ValorDoParametro, '|', 0) IN (
					'PSC'
					,'PES'
					)
			) AS qryPscPes
		INNER JOIN tmpNatOp ON (qryPscPes.strCFOP_NatOper = tmpNatOp.CFOP_NatOper)
			AND (qryPscPes.strFil_NatOper = tmpNatOp.Fil_NatOper)
		) AS tmpPscPes
INNER JOIN (
	SELECT *
	FROM tblDadosConexaoNFeCTe
	WHERE tblDadosConexaoNFeCTe.registroValido IN (
			SELECT TOP 1 cint(tblParametros.ValorDoParametro)
			FROM [tblParametros]
			WHERE TipoDeParametro = 'registroValido'
			)
		AND tblDadosConexaoNFeCTe.ID_NatOp_CompraNF IS NULL
	) AS tmpDadosConexaoNFeCTe ON (tmpPscPes.strXMLCFOP = tmpDadosConexaoNFeCTe.CFOP)
	AND (tmpPscPes.Fil_NatOper = tmpDadosConexaoNFeCTe.ID_Empresa)

SET tmpDadosConexaoNFeCTe.ID_NatOp_CompraNF = [tmpPscPes].[ID_NatOper]
	,tmpDadosConexaoNFeCTe.FiltroCFOP = [tmpPscPes].[CFOP_NatOper];
