--''=====================================================================================================================
--'' registroValido - REGISTROS VALIDOS
--''=====================================================================================================================

-- qryUpdateFornecedoresValidos, _
UPDATE (
		SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF
		FROM tmpClientes
		) AS qryFornecedoresValidos
INNER JOIN tblDadosConexaoNFeCTe ON qryFornecedoresValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit
SET tblDadosConexaoNFeCTe.registroValido = 1;


-- qryUpdateRegistrosValidos, _
UPDATE (
		SELECT STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF
		FROM tmpEmpresa
		) AS qryRegistrosValidos
INNER JOIN tblDadosConexaoNFeCTe ON qryRegistrosValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit
SET tblDadosConexaoNFeCTe.registroValido = 1;



--''=====================================================================================================================
--'' ID_TIPO - APENAS TIPOS COM ID DE VALOR ZERO(0) SERÃO ATUALIZADOS PARA NÃO COMPROMETER OS REGISTROS JÁ PROCESSADOS
--''=====================================================================================================================

-- '' INICIAR TODOS OS REGISTROS COM "0" (ZERO) ONDE O VALOR DO CAMPO "ID_TIPO" ATUALMENTE É "NULL" PARA INICIO DAS IDENTIFICAÇÕES DE TIPOS 
-- qryUpdateIdTipo, _
UPDATE tblDadosConexaoNFeCTe
SET tblDadosConexaoNFeCTe.ID_Tipo = 0
WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) IS NULL));


-- '' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 4 - NF-e Retorno Armazém
-- qryUpdateRetornoArmazem, _
UPDATE tblDadosConexaoNFeCTe
SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='RetornoArmazem'")
WHERE (
		((tblDadosConexaoNFeCTe.ID_Tipo) = 0)
		AND ((tblDadosConexaoNFeCTe.codMod) = CInt('55'))
		AND ((tblDadosConexaoNFeCTe.CNPJ_emit) IN ('12680452000302'))
		);


-- '' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 6 - NF-e Transferência com código Sisparts
-- qryUpdateTransferenciaSisparts, _
UPDATE tblDadosConexaoNFeCTe
SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='TransferenciaSisparts'")
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


-- '' RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 0 - CT-e
-- qryUpdateCTe, _
UPDATE tblDadosConexaoNFeCTe
SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='CTe'")
WHERE (
		((tblDadosConexaoNFeCTe.ID_Tipo) = 0)
		AND ((tblDadosConexaoNFeCTe.codMod) = CInt('57'))
		);


--''=====================================================================================================================
--'' ID_Empresa - CLASSIFICAÇÃO DE EMPRESAS
--''=====================================================================================================================


-- qryUpdateIdEmpresa, _
UPDATE (
		SELECT tmpEmpresa.ID_Empresa
			,STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF
		FROM tmpEmpresa
		) AS qryEmpresas
INNER JOIN tblDadosConexaoNFeCTe ON qryEmpresas.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_Rem
SET tblDadosConexaoNFeCTe.ID_Empresa = qryEmpresas.ID_Empresa;


-- qryUpdateIdEmpresa_TransferenciaEntreFiliais, _
UPDATE (
		SELECT tmpEmpresa.ID_Empresa
			,STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF
		FROM tmpEmpresa
		) AS qryEmpresas
INNER JOIN tblDadosConexaoNFeCTe ON qryEmpresas.strCNPJ_CPF = tblDadosConexaoNFeCTe.CPNJ_Dest
SET tblDadosConexaoNFeCTe.ID_Empresa = qryEmpresas.ID_Empresa
WHERE (((tblDadosConexaoNFeCTe.CFOP) = '6152'));


-- '' RELACIONAMENTO DE CFOP PARA DENTRO E FORA DO ESTADO
-- qryUpdateCFOP_PSC_PES, _
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
	
	
-- '' CONTROLE DE FINALIDADES
-- qryUpdateSit_CompraNF
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
