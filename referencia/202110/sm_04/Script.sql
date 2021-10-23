








SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0)) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0) ORDER BY tblDadosConexaoNFeCTe.CaminhoDoArquivo;

-- sqlArquivosValidos
SELECT ID AS ChvAcesso
	,CaminhoDoArquivo
FROM tblDadosConexaoNFeCTe
WHERE (
		(NOT (tblDadosConexaoNFeCTe.CaminhoDoArquivo) IS NULL)
		AND ((tblDadosConexaoNFeCTe.registroValido) = 1)
		AND ((tblDadosConexaoNFeCTe.registroProcessado) = 2)
		);

-- sqlArquivosInvalidosNaoProcessado
SELECT ID AS ChvAcesso
	,CaminhoDoArquivo
FROM tblDadosConexaoNFeCTe
WHERE (
		(NOT (tblDadosConexaoNFeCTe.CaminhoDoArquivo) IS NULL)
		AND ((tblDadosConexaoNFeCTe.registroValido) = 0)
		AND ((tblDadosConexaoNFeCTe.registroProcessado) = 0)
		);

-- LISTAR REGISTROS VALIDOS PARA AÇÕES DE Lançada ERP e Manifesto de Confirmação.
SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso
	,tblDadosConexaoNFeCTe.dhEmi
FROM tblDadosConexaoNFeCTe
WHERE (
		((Len([ChvAcesso])) > 0)
		AND ((Len([dhEmi])) > 0)
		AND ((tblDadosConexaoNFeCTe.registroValido) = 1)
		);




-- qrySelecaoDeCaminhoDeColeta
SELECT tblParametros.ValorDoParametro FROM tblParametros WHERE tblParametros.TipoDeParametro = 'CaminhoDeColeta';

UPDATE tblDadosConexaoNFeCTe
SET tblDadosConexaoNFeCTe.CaminhoDoArquivo = getPath(Replace([tblDadosConexaoNFeCTe].[CaminhoDoArquivo], DLookUp("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColeta'"), IIf([tblDadosConexaoNFeCTe].[registroProcessado] = 3, DLookUp("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeProcessados'"), IIf([tblDadosConexaoNFeCTe].[registroProcessado] = 4, DLookUp("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeExpurgo'"), getPath([tblDadosConexaoNFeCTe].[CaminhoDoArquivo]))))) & getFileNameAndExt([tblDadosConexaoNFeCTe].[CaminhoDoArquivo])
WHERE tblDadosConexaoNFeCTe.ChvAcesso = "32210268365501000296550000000637741001351624";


SELECT *
FROM 
	tblDadosConexaoNFeCTe
where tblDadosConexaoNFeCTe.ID = 1043;


SELECT * FROM (
	SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso
		,tblDadosConexaoNFeCTe.dhEmi
	FROM tblDadosConexaoNFeCTe
	WHERE (
			((Len([ChvAcesso])) > 0)
			AND ((Len([dhEmi])) > 0)
			AND ((tblDadosConexaoNFeCTe.registroValido) = 1)
			)
) AS tmpSelecao WHERE tmpSelecao.ChvAcesso =  "26210324073694000155550010006291401018935070";



-- ATUALIZAR REGISTRO PARA "PROCESSADO"
UPDATE tblDadosConexaoNFeCTe
SET tblDadosConexaoNFeCTe.registroProcessado = 2
WHERE (
		((tblDadosConexaoNFeCTe.ChvAcesso) = "32210268365501000296550000000637741001351624")
		AND ((tblDadosConexaoNFeCTe.registroValido) = 1)
		AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1)
		);

-- CONSULTAR REGISTROS PENDENTES DE PROCESSAMENTO
SELECT *
FROM tblCompraNFItem
WHERE ChvAcesso_CompraNF IN (
		SELECT ChvAcesso
		FROM tblDadosConexaoNFeCTe
		WHERE (
				((tblDadosConexaoNFeCTe.registroValido) = 1)
				AND (tblDadosConexaoNFeCTe.registroProcessado) = 1
				)
		);


/**** VALIDAR: tblCompraNF

	SELECT 
		ChvAcesso_CompraNF
		,CNPJ_CPF_CompraNF
		,NomeCompleto_CompraNF
		,NumPed_CompraNF
		,NumNF_CompraNF
		,DTEmi_CompraNF
		,DTEntd_CompraNF
		,HoraEntd_CompraNF
		,ID_Forn_CompraNF
		,ModeloDoc_CompraNF
		,Obs_CompraNF
		,Serie_CompraNF
		,Sit_CompraNF
		,TPNF_CompraNF
		,BaseCalcICMS_CompraNF
		,VTotICMS_CompraNF
		,VTotNF_CompraNF
		,VTotProd_CompraNF
		,IDVD_CompraNF
		-- DELETE
		-- Select VTotServ_CompraNF,VTotSeguro_CompraNF
		,VTotOutDesp_CompraNF
		,VTotIPI_CompraNF
		,VTotISS_CompraNF
		,TxDesc_CompraNF
		,VTotDesc_CompraNF
		,VTotISS_CompraNF
		,VTotISS_CompraNF
		-- select * 
	FROM tblCompraNF;


	SELECT
		Item_CompraNFItem
		,ID_CompraNF_CompraNFItem
		,ID_Prod_CompraNFItem
		,CFOP_CompraNFItem
		,BaseCalcICMS_CompraNFItem
		,BaseCalcICMSSubsTrib_CompraNFItem
		,BaseCalcIPI_CompraNFItem
		,DebICMS_CompraNFItem
		,DebIPI_CompraNFItem
		,ICMS_CompraNFItem
		,IPI_CompraNFItem
		,QtdFat_CompraNFItem
		,VUnt_CompraNFItem
		,TxMLSubsTrib_CompraNFItem
		,VTot_CompraNFItem
		,VTotBaseCalcICMS_CompraNFItem
		,VTotDesc_CompraNFItem
		,VTotFrete_CompraNFItem
		,VTotICMSSubsTrib_CompraNFItem
		,VTotOutDesp_CompraNFItem
		-- DELETE
		-- Select *
	from tblCompraNFItem;



 * */


UPDATE (SELECT tblCompraNF.IDVD_CompraNF, ((Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6))) AS strIDVD_CompraNF, [tblCompraNF].[Obs_CompraNF] from tblCompraNF WHERE (((Left([Obs_CompraNF], 6)) = 'PEDIDO ') AND ((tblCompraNF.CNPJ_CPF_CompraNF) = '12680452000302'))) AS tmpIDVD_CompraNF SET tmpIDVD_CompraNF.IDVD_CompraNF = tmpIDVD_CompraNF.strIDVD_CompraNF;


UPDATE (SELECT tblCompraNF.IDVD_CompraNF, ((Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6))) AS strIDVD_CompraNF,
[tblCompraNF].[Obs_CompraNF] from tblCompraNF 
WHERE (((Left([Obs_CompraNF], 6)) = 'PEDIDO ') AND ((tblCompraNF.CNPJ_CPF_CompraNF) = '12680452000302'))) AS tmpIDVD_CompraNF
SET tmpIDVD_CompraNF.IDVD_CompraNF = tmpIDVD_CompraNF.strIDVD_CompraNF
WHERE val(tmpIDVD_CompraNF.strIDVD_CompraNF) > 0 ;



(SELECT tblCompraNF.IDVD_CompraNF, ((Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6))) AS strIDVD_CompraNF,
[tblCompraNF].[Obs_CompraNF] from tblCompraNF 
WHERE (((Left([Obs_CompraNF], 6)) = 'PEDIDO ') AND ((tblCompraNF.CNPJ_CPF_CompraNF) = '12680452000302'))) AS tmpIDVD_CompraNF


SELECT 
((Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6))) AS strIDVD_CompraNF,
[tblCompraNF].[Obs_CompraNF] from tblCompraNF 
WHERE (((Left([Obs_CompraNF], 6)) = 'PEDIDO ') AND ((tblCompraNF.CNPJ_CPF_CompraNF) = '12680452000302'));



SELECT 
(Val(Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6))) AS tmp,
[tblCompraNF].[Obs_CompraNF] from tblCompraNF 
WHERE (((Left([Obs_CompraNF], 6)) = 'PEDIDO ') AND ((tblCompraNF.CNPJ_CPF_CompraNF) = '12680452000302'));




-- qryUpdate_IDVD
UPDATE tblCompraNF
SET tblCompraNF.IDVD_CompraNF = Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6)
-- SELECT [tblCompraNF].[Obs_CompraNF] from tblCompraNF 
WHERE (((Left([Obs_CompraNF], 6)) = 'PEDIDO ') AND ((tblCompraNF.CNPJ_CPF_CompraNF) = '12680452000302') AND ((Val(Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6))) > 0));


-- VALIDAR: IDVD_CompraNF
SELECT [tblCompraNF].[Obs_CompraNF] from tblCompraNF 
WHERE 
(Left([Obs_CompraNF], 6) = 'PEDIDO ') AND (tblCompraNF.CNPJ_CPF_CompraNF = '12680452000302');



-- PENDENTE: ARQUIVOS/REGISTROS SEM IDENTIFICAÇÃO DE TIPOS
SELECT 
	tblDadosConexaoNFeCTe.CaminhoDoArquivo 
FROM 
	tblDadosConexaoNFeCTe 
WHERE 
	(((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0)) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0) 
ORDER BY 
	tblDadosConexaoNFeCTe.CaminhoDoArquivo;


-- CLASSIFICOES ( registroValido e registroProcessado) : REGISTROS VALIDOS E ...
-- [registroProcessado]
-- 0 - Pendente de processamento
-- 1 - Registro OK
-- 2 - Enviado para servidor
-- 3 - Arquivo movido para pasta de processados
-- 4 - Arquivo movido para pasta de expurgo

SELECT ChvAcesso, CaminhoDoArquivo, registroProcessado 
-- SELECT COUNT(*)
FROM  tblDadosConexaoNFeCTe 
-- WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=2));
-- WHERE tblDadosConexaoNFeCTe.ChvAcesso ="42210312680452000302550020000897601860571690";
-- UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 0
WHERE tblDadosConexaoNFeCTe.registroValido=1 AND tblDadosConexaoNFeCTe.registroProcessado=3;



-- TAG: EXPURGO
UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 4 WHERE tblDadosConexaoNFeCTe.ID =276;



-- SELECAO: DADOS GERAIS




-- SELECAO: ITENS DE NOTAS
SELECT *
-- SELECT count(*)
-- delete 
from tblCompraNFItem
-- where ChvAcesso_CompraNF = '32210304884082000569570000040073831040073834'
where ID_CompraNF_CompraNFItem = 2196
;






-- SELECAO: NOTAS
SELECT VTotProd_CompraNF,*
-- delete 
from tblCompraNF
where ChvAcesso_CompraNF = '32210304884082000569570000040073831040073834'
; 


SELECT ChvAcesso FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND (tblDadosConexaoNFeCTe.registroProcessado)=1)
;


Select * from tblCompraNF Where ChvAcesso_CompraNF IN (SELECT ChvAcesso FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND (tblDadosConexaoNFeCTe.registroProcessado)=1))
;


SELECT CaminhoDoArquivo FROM  tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=2))
;
