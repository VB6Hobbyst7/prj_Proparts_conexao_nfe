/****** 

	MSSQL 

******/
SELECT *
-- SELECT IDVD_CompraNF 
-- delete
FROM [SispartsConexao].[dbo].[tblCompraNF];
-- where ChvAcesso_CompraNF = '32210204884082000569570000039547081039547080';

SELECT *
-- delete
FROM [SispartsConexao].[dbo].[tblCompraNFItem];
-- inner join [SispartsConexao].[dbo].[tblCompraNF] ON [SispartsConexao].[dbo].[tblCompraNFItem].ID_CompraNF_CompraNFItem = [SispartsConexao].[dbo].[tblCompraNF].ID_CompraNF
-- where [SispartsConexao].[dbo].[tblCompraNF].ChvAcesso_CompraNF = '32210204884082000569570000039547081039547080';

/*

	02 PROCESSAMENTO DE DADOS

*/
/****************************************************************************************************************************************
	ID_TIPO - APENAS TIPOS COM ID DE VALOR ZERO(0) SERÃƒO ATUALIZADOS PARA NÃƒO COMPROMETER OS REGISTROS JÃ� PROCESSADOS
****************************************************************************************************************************************/
-- TIPOS DE CADASTRO ( NÃƒO PROCESSADO )
-- qryUpdateIdTipo
UPDATE tblDadosConexaoNFeCTe
SET tblDadosConexaoNFeCTe.ID_Tipo = 0
WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) IS NULL));

-- RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 4 - NF-e Retorno ArmazÃ©m
-- qryUpdateRetornoArmazem
UPDATE tblDadosConexaoNFeCTe
SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp("" [ValorDoParametro] "", "" [tblParametros] "", "" [TipoDeParametro] = 'RetornoArmazem' "")
WHERE (
		((tblDadosConexaoNFeCTe.ID_Tipo) = 0)
		AND ((tblDadosConexaoNFeCTe.codMod) = CInt('55'))
		AND ((tblDadosConexaoNFeCTe.CNPJ_emit) IN ('12680452000302'))
		);

-- RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 6 - NF-e TransferÃªncia com cÃ³digo Sisparts
-- qryUpdateTransferenciaSisparts
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

-- RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 0 - CT-e
-- qryUpdateCTe
UPDATE tblDadosConexaoNFeCTe
SET tblDadosConexaoNFeCTe.ID_Tipo = DLookUp("" [ValorDoParametro] "", "" [tblParametros] "", "" [TipoDeParametro] = 'CTe' "")
WHERE (
		((tblDadosConexaoNFeCTe.ID_Tipo) = 0)
		AND ((tblDadosConexaoNFeCTe.codMod) = CInt('57'))
		);

/********************************************************************
	DADOS GERAIS
	
	registroValido
		0 			- KO
		1			- OK		
		
	registroProcessado
		0 			- KO
		1			- OK	
********************************************************************/
-- #####################################
--	REGISTROS VALIDOS
-- #####################################
-- qryUpdateRegistrosValidos
UPDATE (
		SELECT STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF
		FROM tmpEmpresa
		) AS qryRegistrosValidos
INNER JOIN tblDadosConexaoNFeCTe ON qryRegistrosValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit

SET tblDadosConexaoNFeCTe.registroValido = 1;

-- qryUpdateFornecedoresValidos
UPDATE (
		SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF
		FROM tmpClientes
		) AS qryFornecedoresValidos
INNER JOIN tblDadosConexaoNFeCTe ON qryFornecedoresValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit

SET tblDadosConexaoNFeCTe.registroValido = 1;

-- #####################################
--	REGISTROS PROCESSADOS
-- #####################################
-- qryUpdateProcessamentoConcluido
UPDATE tblDadosConexaoNFeCTe
SET tblDadosConexaoNFeCTe.registroProcessado = 1
WHERE (
		((tblDadosConexaoNFeCTe.registroValido) = 1)
		AND ((tblDadosConexaoNFeCTe.registroProcessado) = 0)
		AND ((tblDadosConexaoNFeCTe.Chave) = 'strChave')
		);

-- qryUpdateProcessamentoConcluido_CTE
UPDATE tblDadosConexaoNFeCTe
SET tblDadosConexaoNFeCTe.registroProcessado = 1
WHERE (
		((tblDadosConexaoNFeCTe.registroValido) = 1)
		AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1)
		AND ((tblDadosConexaoNFeCTe.ID_Tipo) = DLookUp("" [ValorDoParametro] "", "" [tblParametros] "", "" [TipoDeParametro] = 'Cte' ""))
		);

-- qryUpdateProcessamentoConcluido_ItensCompras
UPDATE tblDadosConexaoNFeCTe
SET tblDadosConexaoNFeCTe.registroProcessado = 1
FROM tblDadosConexaoNFeCTe
WHERE (
		((tblDadosConexaoNFeCTe.registroValido) = 1)
		AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1)
		AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0)
		);

-- #####################################
--	SELEÃ‡ÃƒO DE DADOS
-- #####################################
-- qrySelecaoDeArquivosPendentes
SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo
FROM tblDadosConexaoNFeCTe
WHERE (
		((tblDadosConexaoNFeCTe.registroValido) = 1)
		AND ((tblDadosConexaoNFeCTe.registroProcessado) = 0)
		)
	AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0)
ORDER BY tblDadosConexaoNFeCTe.CaminhoDoArquivo;

/********************************************************************
	FiltroFil
********************************************************************/
-- qryUpdateIdEmpresa (CTE)
UPDATE (
		SELECT tmpEmpresa.ID_Empresa
			,STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF
		FROM tmpEmpresa
		) AS qryEmpresas
INNER JOIN tblDadosConexaoNFeCTe ON qryEmpresas.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_Rem

SET tblDadosConexaoNFeCTe.ID_Empresa = qryEmpresas.ID_Empresa;

-- qryUpdateIdEmpresa_TransferenciaEntreFiliais
UPDATE (
		SELECT tmpEmpresa.ID_Empresa
			,STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF
		FROM tmpEmpresa
		) AS qryEmpresas
INNER JOIN tblDadosConexaoNFeCTe ON qryEmpresas.strCNPJ_CPF = tblDadosConexaoNFeCTe.CPNJ_Dest

SET tblDadosConexaoNFeCTe.ID_Empresa = qryEmpresas.ID_Empresa
WHERE (((tblDadosConexaoNFeCTe.CFOP) = '6152'));

/********************************************************************
	FiltroCFOP
********************************************************************/
-- qryUpdateCFOP_PSC_PES
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
		AND tblDadosConexaoNFeCTe.FiltroCFOP = 0
	) AS tmpDadosConexaoNFeCTe ON (tmpPscPes.strXMLCFOP = tmpDadosConexaoNFeCTe.CFOP)
	AND (tmpPscPes.Fil_NatOper = tmpDadosConexaoNFeCTe.ID_Empresa)

SET tmpDadosConexaoNFeCTe.FiltroCFOP = [tmpPscPes].[ID_NatOper];

/********************************************************************
	CONSULTA PARA CRIAÃ‡ÃƒO DE ARQUIVOS JSON
********************************************************************/
-- sqyDadosJson
SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso
	,tblDadosConexaoNFeCTe.dhEmi
FROM tblDadosConexaoNFeCTe
WHERE (
		((Len([ChvAcesso])) > 0)
		AND ((Len([dhEmi])) > 0)
		);

/********************************************************************
	COMPRAS
********************************************************************/
-- qryUpdateBaseCalcICMS
UPDATE tblCompraNF
SET tblCompraNF.BaseCalcICMS_CompraNF = replace([tblCompraNF].[BaseCalcICMS_CompraNF] / 100, "", "", ""."")
WHERE (((tblCompraNF.BaseCalcICMS_CompraNF) > "" 0 ""))
	OR (((tblCompraNF.BaseCalcICMS_CompraNF) IS NOT NULL));

-- qryUpdate_IDVD
UPDATE tblCompraNF
SET tblCompraNF.IDVD_CompraNF = Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6)
WHERE (
		((Left([Obs_CompraNF], 6)) = 'PEDIDO ')
		AND ((tblCompraNF.CNPJ_CPF_CompraNF) = '12680452000302')
		AND ((Val(Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6))) > 0)
		);

-- qryUpdateCFOP_FilCompra
UPDATE tblCompraNF
SET tblCompraNF.CFOP_CompraNF = DLookUp("" [FiltroCFOP] "", "" [tblDadosConexaoNFeCTe] "", "" [ChvAcesso] = '"" & [tblCompraNF].[ChvAcesso_CompraNF] & ""' "")
	,tblCompraNF.Fil_CompraNF = DLookUp("" [ID_EMPRESA] "", "" [tblDadosConexaoNFeCTe] "", "" [ChvAcesso] = '"" & [tblCompraNF].[ChvAcesso_CompraNF] & ""' "");

-- qryUpdateFilCompraNF
UPDATE (
		SELECT tmpEmpresa.ID_Empresa
			,STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF
			,tmpEmpresa.CNPJ_Empresa
		FROM tmpEmpresa
		WHERE (((tmpEmpresa.CNPJ_Empresa) IS NOT NULL))
		) AS qryEmpresas
INNER JOIN tblCompraNF ON qryEmpresas.strCNPJ_CPF = tblCompraNF.CNPJ_CPF_CompraNF

SET tblCompraNF.Fil_CompraNF = qryEmpresas.ID_Empresa;

-- qryUpdateIdFornCompraNF
UPDATE (
		SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF
			,tmpClientes.CÓDIGOClientes
		FROM tmpClientes
		) AS qryClientesFornecedor
INNER JOIN tblCompraNF ON tblCompraNF.CNPJ_CPF_CompraNF = qryClientesFornecedor.strCNPJ_CPF

SET tblCompraNF.ID_Forn_CompraNF = qryClientesFornecedor.CÓDIGOClientes;

-- qryUpdate_ModeloDoc_CFOP
UPDATE tblCompraNF
INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso

SET tblCompraNF.ModeloDoc_CompraNF = [tblDadosConexaoNFeCTe].[codMod]
	,tblCompraNF.CFOP_CompraNF = [tblDadosConexaoNFeCTe].[FiltroCFOP]
WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) > 0));

/********************************************************************
	COMPRAS ITENS
********************************************************************/
-- qryUpdateItens_ID_Prod_CompraNFItem
UPDATE tblCompraNFItem
SET tblCompraNFItem.ID_Prod_CompraNFItem = DLookUp("" CodigoProd_Grade "", "" dbo_tabGradeProdutos "", "" CodigoForn_Grade = '"" & [tblCompraNFItem].[ID_Prod_CompraNFItem] & ""' "");

-- qryInsertItensCTe
INSERT INTO tblCompraNFItem (
	ChvAcesso_CompraNF
	,VUnt_CompraNFItem
	,Num_CompraNFItem
	,VTot_CompraNFItem
	,DebICMS_CompraNFItem
	,VTotBaseCalcICMS_CompraNFItem
	,ID_NatOp_CompraNFItem
	,Item_CompraNFItem
	,ID_Grade_CompraNFItem
	,QtdFat_CompraNFItem
	,IPI_CompraNFItem
	,FlagEst_CompraNFItem
	,BaseCalcICMS_CompraNFItem
	)
SELECT tblCompraNF.ChvAcesso_CompraNF
	,tblCompraNF.VTotNF_CompraNF AS VUnt_CompraNFItem
	,tblCompraNF.NumNF_CompraNF
	,tblCompraNF.VTotNF_CompraNF
	,IIf([VTotICMS_CompraNF] <> "", Replace(Nz([tblCompraNF].[BaseCalcICMS_CompraNF], 0) / 100, ",", "."), 0) AS strVTotICMS
	,tblCompraNF.BaseCalcICMS_CompraNF
	,tblCompraNF.ID_NatOp_CompraNF
	,1 AS strItem
	,1 AS strIDGrade
	,1 AS strQtdFat
	,0 AS strIPI
	,0 AS strFlag
	,100 AS strBaseCalcICMS
FROM tblCompraNF
INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso
WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = DLookUp("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='Cte'")));
