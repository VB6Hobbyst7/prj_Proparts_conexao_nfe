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
					,'PSP'
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
