-- CTE
--> RELACIONAR COM REMETENTE.
UPDATE (
		SELECT tmpEmpresa.ID_Empresa
			,STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF
		FROM tmpEmpresa
		) AS qryEmpresas
INNER JOIN tblDadosConexaoNFeCTe ON qryEmpresas.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_Rem
SET tblDadosConexaoNFeCTe.ID_Empresa = qryEmpresas.ID_Empresa;


-- TRANSFERENCIA ENTRE FILIAIS
UPDATE (
		SELECT tmpEmpresa.ID_Empresa
			,STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF
		FROM tmpEmpresa
		) AS qryEmpresas
INNER JOIN tblDadosConexaoNFeCTe ON qryEmpresas.strCNPJ_CPF = tblDadosConexaoNFeCTe.CPNJ_Dest
SET tblDadosConexaoNFeCTe.ID_Empresa = qryEmpresas.ID_Empresa;
WHERE (((tblDadosConexaoNFeCTe.CFOP) = '6152'));


-- RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 4 - NF-e Retorno Armazém
UPDATE (
		SELECT ValorDoParametro
			,TipoDeParametro
		FROM tblParametros
		WHERE TipoDeParametro = 'RetornoArmazem'
		) AS tmpRetornoArmazem
		INNER JOIN (
		(SELECT TOP 1 cInt('55') AS strMod
			,'12680452000302' AS strCNPJ_CPF
			,'RetornoArmazem' AS strTipoDeParametro
		FROM tblParametros
		) AS qryRetornoArmazem INNER JOIN tblDadosConexaoNFeCTe ON 
		(qryRetornoArmazem.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit)
		AND (qryRetornoArmazem.strMod = tblDadosConexaoNFeCTe.codMod)
	) ON (tmpRetornoArmazem.TipoDeParametro = qryRetornoArmazem.strTipoDeParametro)
	AND (tmpRetornoArmazem.TipoDeParametro = qryRetornoArmazem.strTipoDeParametro)
SET tblDadosConexaoNFeCTe.ID_Tipo = [tmpRetornoArmazem].[ValorDoParametro]
WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = 0)) AND (((tblDadosConexaoNFeCTe.CFOP) = '5907'));


-- RELACIONAR COM ID DE TIPOS DE CADASTROS (tblTipos) - 0 - CT-e
UPDATE (
		SELECT ValorDoParametro
			,TipoDeParametro
		FROM tblParametros
		WHERE TipoDeParametro = 'CTe'
		) AS tmpCTe
INNER JOIN (
	(
		SELECT TOP 1 cInt('57') AS strMod
			,'CTe' AS strTipoDeParametro
		FROM tblParametros
		) AS qryCTe INNER JOIN tblDadosConexaoNFeCTe ON (qryCTe.strMod = tblDadosConexaoNFeCTe.codMod)
	) ON (tmpCTe.TipoDeParametro = qryCTe.strTipoDeParametro)
	AND (tmpCTe.TipoDeParametro = qryCTe.strTipoDeParametro)

SET tblDadosConexaoNFeCTe.ID_Tipo = [tmpCTe].[ValorDoParametro]
WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = 0));


-- ** NOVO ** -- CONCLUIDO
-- pendente
-- qryUpdateIdFornCompraNF
ID_Forn_CompraNF

UPDATE (SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF, tmpClientes.CÓDIGOClientes
		FROM tmpClientes
		) AS qryClientesFornecedor
INNER JOIN tblCompraNF ON tblCompraNF.CNPJ_CPF_CompraNF = qryClientesFornecedor.strCNPJ_CPF
SET tblCompraNF.ID_Forn_CompraNF = qryClientesFornecedor.CÓDIGOClientes;



-- ** NOVO **
-- '' #AILTON - qryExpurgoDeComprasJaCadastradas - ( ChvAcesso )
-- '' #DUVIDA - QUAL O OBJETIVO ?
-- '' #ENTENDIMENTO - NÃO GERAR DUPLICIDADE NO CADASTRO DE COMPRA VERIFICANDO A EXISTENCIA DA MESMA PELOS SEGUINTES CAMPOS: XMLCNPJEmi E XMLNumNF
SELECT Clientes.CNPJ_CPF, tblCompraNF.ID_Forn_CompraNF, tblCompraNF.NumNF_CompraNF, tblCompraNF.DTEntd_CompraNF FROM tblCompraNF INNER JOIN Clientes ON tblCompraNF.ID_Forn_CompraNF = Clientes.CÓDIGOClientes WHERE  (Clientes.CNPJ_CPF='" & Format(XMLCNPJEmi, "00\.000\.000/0000\-00") & "') AND (tblCompraNF.NumNF_CompraNF=" & XMLNumNF & ");

-- ** NOVO **
-- '' #AILTON - qryExpurgoDeComprasJaCadastradas - ( ChvAcesso )
-- '' NÃO PROCESSAR REGISTROS JÁ CADASTRADOS ( RELACIONAMENTO POR ChvAcesso )


-- chave ( itensCompra )
NumNF_CompraNF + ID_Forn_CompraNF


-- ** NOVO ** -- COMPRAS
tmpCompraNF.ModeloDoc_CompraNF = tblDadosConexaoNFeCTe.codMod
tmpCompraNF.CFOP_CompraNF = tmpDadosConexaoNFeCTe.FiltroCFOP



-- ** UPDATE ** -- qryUpdate_IDVD
UPDATE tblCompraNF
SET tblCompraNF.IDVD_CompraNF = Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6)
WHERE (
		((Left([Obs_CompraNF], 6)) = 'PEDIDO ')
		AND ((tblCompraNF.CNPJ_CPF_CompraNF) = '12680452000302')
		AND ((Val(Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6))) > 0));



-- ** OK ** -- qryUpdateCFOP_PSC_PES

UPDATE  ( SELECT 
           tmpNatOp.ID_NatOper, tmpNatOp.Fil_NatOper, tmpNatOp.CFOP_NatOper, qryPscPes.strXMLCFOP, qryPscPes.strEstado  
       FROM (SELECT  
               strSplit(ValorDoParametro,'|',0) AS strFil_NatOper,  strSplit(ValorDoParametro,'|',1) AS strEstado,  strSplit(ValorDoParametro,'|',2) AS strXMLCFOP,  strSplit(ValorDoParametro,'|',3) AS strCFOP_NatOper  
             FROM  
               tblParametros  
             WHERE  
               TipoDeParametro='FiltroFil' And strSplit(ValorDoParametro,'|',0) In ('PSC','PES'))  AS qryPscPes  
       INNER JOIN tmpNatOp ON (qryPscPes.strCFOP_NatOper = tmpNatOp.CFOP_NatOper) AND (qryPscPes.strFil_NatOper = tmpNatOp.Fil_NatOper) )  AS tmpPscPes  
INNER JOIN  
   (   SELECT  *  
       FROM  tblDadosConexaoNFeCTe  
       WHERE tblDadosConexaoNFeCTe.registroValido IN (SELECT TOP 1 cint(tblParametros.ValorDoParametro) FROM [tblParametros] WHERE TipoDeParametro = 'registroValido')  
       AND tblDadosConexaoNFeCTe.FiltroCFOP = 0 )  AS tmpDadosConexaoNFeCTe 
ON (tmpPscPes.strXMLCFOP = tmpDadosConexaoNFeCTe.CFOP) AND (tmpPscPes.Fil_NatOper = tmpDadosConexaoNFeCTe.ID_Empresa)
SET  tmpDadosConexaoNFeCTe.FiltroCFOP = [tmpPscPes].[ID_NatOper];



-- 