

UPDATE tmpCompraNFItem
SET ID_Prod_CompraNFItem = tmpProdutos.código
FROM tmpCompraNFItem
INNER JOIN tmpCompraNF ON tmpCompraNF.ID_CompraNF = tmpCompraNFItem.ID_CompraNF_CompraNFItem
INNER JOIN tmpProdutos ON tmpProdutos.modelo = "Transporte"
WHERE tmpCompraNF.Sit_CompraNF = 6;




-- 42210312680452000302550020000895841453583169  1.907
-- ST ICMS 050 o correto e 150

-- 42210348740351012767570000021186731952977908  2.353
-- ID_Prod com 0 correto e 34125 (Transporte)

-- 32210368365501000296550000000638961001363203  2.152
-- VTotIPI = R$ 95,24
-- VTotProd = 1321,38 <> 1416,62



-- #20220106_update_Modelo_CadatroDeProdutos
	-- UPDATE [Cadastro de Produtos] SET [modelo]='XXX' where [código]=529;
	SELECT [código],[modelo] FROM [Cadastro de Produtos] where [código] IN (34125,529);


-- #20220106_update_IdProd_CompraNFItem
	/* RELACIONAR OS ITENS DE COMPRA COMO TRANSPORTE. ONDE A TABELA DE PRODUTOS É IGUAL A "TRANSPORTE" E O CAMPO "ID_Prod_CompraNFItem" É IGUAL A ZERO(0) E O CAMPO "Sit_CompraNF" VINDO DA TABELA COMPRAS É IGUAL A 6 */
	UPDATE tblCompraNFItem SET ID_Prod_CompraNFItem = tbProdutos.[código]
	FROM tblCompraNFItem AS tbItens
	INNER JOIN tblCompraNF as tbCompras ON tbCompras.ID_CompraNF = tbItens.ID_CompraNF_CompraNFItem
	INNER join [Cadastro de Produtos] as tbProdutos on tbProdutos.modelo = 'Transporte' 
	WHERE tbCompras.Sit_CompraNF = 6 and tbItens.ID_Prod_CompraNFItem=0;
	
	
	/* ZERAR O CAMPO ID_Prod_CompraNFItem PARA TIPOS (6) CTe*/
	UPDATE tblCompraNFItem SET ID_Prod_CompraNFItem = 0
	-- SELECT COUNT(*) -- SELECT ID_Prod_CompraNFItem
	FROM tblCompraNFItem AS tbItens
	INNER JOIN tblCompraNF as tbCompras ON tbCompras.ID_CompraNF = tbItens.ID_CompraNF_CompraNFItem
	WHERE tbCompras.Sit_CompraNF = 6;


SELECT ID_NatOper
	,Fil_NatOper
	,CFOP_NatOper
	,Descr_NatOper
	,STICMS_NatOper
	,STIPI_NatOper
	,STPC_NatOper
-- SELECT COUNT(*)	-- SELECT * -- DELETE
FROM tblNatOp;



SELECT ChvAcesso_CompraNF
	,VTotIPI_CompraNF 
	,VTotProd_CompraNF 
	,NumPed_CompraNF
	,ID_NatOp_CompraNF
	,tblNatOp.descr_natoper
	,CNPJ_CPF_CompraNF
	,NomeCompleto_CompraNF
	,NumNF_CompraNF
	,DTEmi_CompraNF
	,HoraEntd_CompraNF
	,ModeloDoc_CompraNF
	,Obs_CompraNF
	,Serie_CompraNF
	,TPNF_CompraNF
	,BaseCalcICMS_CompraNF
	,VTotICMS_CompraNF
	,VTotIPI_CompraNF
	,VTotNF_CompraNF
	,IDVD_CompraNF
-- SELECT COUNT(*)	-- SELECT * -- DELETE
-- SELECT Sit_CompraNF
FROM tblCompraNF
INNER JOIN tblNatOp ON tblCompraNF.ID_NatOp_CompraNF = tblNatOp.ID_NatOper
WHERE tblCompraNF.ChvAcesso_CompraNF = '42210348740351012767570000021186731952977908';
-- WHERE ChvAcesso_CompraNF = '42210312680452000302550020000895841453583169';

SELECT Item_CompraNFItem
	,ChvAcesso_CompraNF
	,ID_Prod_CompraNFItem
	,tbCompras.Sit_CompraNF
	,CFOP_CompraNFItem
	,tbCompras.CFOP_CompraNF
	,ST_CompraNFItem
	,ICMS_CompraNFItem
	,STIPI_CompraNFItem
	,STPIS_CompraNFItem
	,STCOFINS_CompraNFItem
	,BaseCalcICMS_CompraNFItem
	,BaseCalcICMSSubsTrib_CompraNFItem
	,ID_Grade_CompraNFItem
	,FlagEst_CompraNFItem
	,BaseCalcICMS_CompraNFItem
	,BaseCalcICMSSubsTrib_CompraNFItem
	,ID_NatOp_CompraNFItem
	,tbCompras.ID_NatOp_CompraNF
	,BaseCalcIPI_CompraNFItem
	,DebICMS_CompraNFItem
	,DebIPI_CompraNFItem
	,ICMS_CompraNFItem
	,IPI_CompraNFItem
	,QtdFat_CompraNFItem
	,VUnt_CompraNFItem
	,VTot_CompraNFItem
	,VTotBaseCalcICMS_CompraNFItem
	,ID_CompraNF_CompraNFItem
-- SELECT COUNT(*)	-- SELECT * -- DELETE
FROM tblCompraNFItem as tbItens
INNER JOIN tblCompraNF as tbCompras ON tbCompras.ID_CompraNF = tbItens.ID_CompraNF_CompraNFItem
WHERE tbCompras.ChvAcesso_CompraNF = '42210348740351012767570000021186731952977908';
-- WHERE tblCompraNF.ChvAcesso_CompraNF = '42210312680452000302550020000895841453583169';
