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