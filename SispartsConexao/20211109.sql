INSERT INTO tblProcessamento (
	PK
	,NomeCampo
	,NomeTabela
	,formatacao
	,Valor
	)
SELECT '32210268365501000296550000000637741001351624'
	,'CFOP_CompraNF'
	,'tblCompraNF'
	,'opTexto'
	,tDados.FiltroCFOP 
FROM tblDadosConexaoNFeCTe AS tDados
WHERE ChvAcesso = '32210268365501000296550000000637741001351624';


SELECT tblDadosConexaoNFeCTe.CFOP
	,tblDadosConexaoNFeCTe.ID_NatOp_CompraNF
	,tblDadosConexaoNFeCTe.ID_Forn_CompraNF
	,tblDadosConexaoNFeCTe.codMod
	,tblDadosConexaoNFeCTe.ID_Empresa
	,tblDadosConexaoNFeCTe.ChvAcesso
-- FROM tblDadosConexaoNFeCTe WHERE tblDadosConexaoNFeCTe.ChvAcesso = '42210320147617000494570010009658691999034138'; -- 32210268365501000296550000000637741001351624
FROM (
	SELECT DISTINCT tblProcessamento.NomeTabela
		,tblProcessamento.pk
		,tblProcessamento.NomeCampo
		,tblProcessamento.valor
		,tblProcessamento.formatacao
		,tblParametros.ID
	FROM tblParametros
	INNER JOIN tblProcessamento ON (tblParametros.ValorDoParametro = tblProcessamento.NomeCampo)
		AND (tblParametros.TipoDeParametro = tblProcessamento.NomeTabela)
	WHERE (
			((tblProcessamento.NomeTabela) = "tblCompraNF")
			AND (
				(tblProcessamento.NomeCampo) IS NOT NULL
				AND (tblProcessamento.NomeCampo) = "ChvAcesso_CompraNF"
				)
			)
	ORDER BY tblProcessamento.NomeTabela
		,tblProcessamento.pk
		,tblParametros.ID
	) AS tmpPk
INNER JOIN tblDadosConexaoNFeCTe ON trim(tmpPk.valor) = trim(tblDadosConexaoNFeCTe.ChvAcesso);





SELECT DISTINCT pk, NomeTabela FROM qryProcessamento_Select_CompraComItens;



INSERT INTO tblProcessamento (PK,NomeCampo,NomeTabela,formatacao,Valor) 
SELECT Distinct PK,NomeCampo,'tblCompraNF','opTexto',val((Left(Trim(Replace(Replace(Valor, 'PEDIDO: ', ''), 'PEDIDO ', '')), 6)))  as IDVD_CompraNF
FROM tblProcessamento
WHERE NomeTabela = 'tblCompraNF'
	AND len(formatacao) > 0
	AND len(NomeCampo) > 0
	AND Left(Valor, 6) = 'PEDIDO '
	AND NomeCampo = 'Obs_CompraNF';






SELECT Distinct NomeCampo,Valor
FROM tblProcessamento
WHERE NomeTabela = 'tblCompraNF'
	AND len(formatacao) > 0
	AND len(NomeCampo) > 0
	AND NomeCampo IN ('CNPJ_CPF_CompraNF' , 'Obs_CompraNF')
 	AND Left(Valor, 6) = 'PEDIDO '	
	AND Valor = '12680452000302'








SELECT * 
FROM 
(
SELECT Distinct NomeCampo,Valor
FROM tblProcessamento
WHERE NomeTabela = 'tblCompraNF'
	AND len(formatacao) > 0
	AND len(NomeCampo) > 0
	AND Valor = '12680452000302'
-- 	AND Left(Valor, 6) = 'PEDIDO '
	AND NomeCampo IN ('CNPJ_CPF_CompraNF' , 'Obs_CompraNF')
) AS tmp 
WHERE NomeCampo IN ('CNPJ_CPF_CompraNF' , 'Obs_CompraNF');

-- Left(tblDadosConexaoNFeCTe.Obs_CompraNF, 6)) = 'PEDIDO ')
-- tblDadosConexaoNFeCTe.CNPJ_CPF_CompraNF) = '12680452000302'