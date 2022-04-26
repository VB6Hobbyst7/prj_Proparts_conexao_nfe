
----------------------------------

SELECT COUNT(*) as RegistroExistente 
FROM tblCompraNF where ChvAcesso_CompraNF = '32210268365501000296550000000637741001351624';


SELECT COUNT(*) as RegistroExistente,ChvAcesso_CompraNF FROM tblCompraNF where ChvAcesso_CompraNF = '32210268365501000296550000000637741001351624' GROUP BY ChvAcesso_CompraNF 





SELECT COUNT(*)	-- SELECT * -- DELETE
FROM tblCompraNF;


SELECT COUNT(*)	-- SELECT * -- DELETE
FROM tblCompraNFItem


SELECT ChvAcesso_CompraNF
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
	,CodMunOrigem_CompraNF
	,CodMunDestino_CompraNF
-- SELECT ChvAcesso_CompraNF,NumPed_CompraNF,ID_Forn_CompraNF
-- SELECT COUNT(*)	-- SELECT * -- DELETE
FROM tblCompraNF
INNER JOIN tblNatOp ON tblCompraNF.ID_NatOp_CompraNF = tblNatOp.ID_NatOper
WHERE ChvAcesso_CompraNF = '42210312680452000302550020000902571508970265'

--WHERE ModeloDoc_CompraNF = 55
--	AND ID_NatOp_CompraNF = 122;
	

SELECT Item_CompraNFItem
	,ID_Grade_CompraNFItem
	,FlagEst_CompraNFItem
	,BaseCalcICMS_CompraNFItem
	,BaseCalcICMSSubsTrib_CompraNFItem
	,ID_NatOp_CompraNFItem
	,tblCompraNF.ID_NatOp_CompraNF
	,CFOP_CompraNFItem
	,tblCompraNF.CFOP_CompraNF
	,ID_Prod_CompraNFItem
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
-- SELECT ID_CompraNFItem
-- SELECT DISTINCT(ID_Prod_CompraNFItem)
-- SELECT DISTINCT ModeloDoc_CompraNF, CFOP_CompraNF, Fil_CompraNF, Almox_CompraNFItem,FlagEst_CompraNFItem
-- SELECT COUNT(*)	-- SELECT * -- DELETE
FROM tblCompraNFItem
INNER JOIN tblCompraNF ON tblCompraNF.ID_CompraNF = tblCompraNFItem.ID_CompraNF_CompraNFItem
WHERE tblCompraNF.ModeloDoc_CompraNF = 55
	AND tblCompraNF.ChvAcesso_CompraNF = '42210312680452000302550020000902571508970265'
	
	
	
	/*** CADASTRO DE NATUREZA DE OPERAÇÃO


	SispartsConexao
	WINAP6LLZINDIHW
	41L70n@#
	sa



	select 
		ID_NatOper
		,CFOP_NatOper
		,Fil_NatOper 
		,descr_natoper
	-- SELECT COUNT(*)	-- SELECT * -- DELETE
	from tblNatOp
	where CFOP_NatOper='2.152'
	where CFOP_NatOper='1.907'
	where ID_NatOper=419
 
 */
	/*** CADASTRO DE CLIENTES
	
	SET IDENTITY_INSERT Clientes ON

	INSERT INTO Clientes (CÃ“DIGOClientes,NomeCompleto,CNPJ_CPF,Estado,envia,opMTB,opAv,opTri,opEstr,opDH,opCami,opOut,OptSimples_Cad,vdRS,vdAV,vdSP,vdSC,vdSR,vdTR,vd661,FlagSel_Cad,FlagAtivo,FlagSemRest,FlagBloq,FlagMsgCob,FlagSitTransp_Cad) 
	SELECT * FROM ( VALUES
	(23306,'Proparts Com Art Esp e Tec - Filial Es','68.365.501/0002-96','ES',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
	,(46036,'Proparts Com Art Esp e Tec - Filial Sc','68.365.501/0003-77','SC',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
	,(1025,'Proparts Com Art Esp e Tec - Matriz Sp','68.365.501/0001-05','SP',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
	) AS tmp(str_CODIGOClientes,str_NomeCompleto,str_CNPJ_CPF,str_Estado,envia,opMTB,opAv,opTri,opEstr,opDH,opCami,opOut,OptSimples_Cad,vdRS,vdAV,vdSP,vdSC,vdSR,vdTR,vd661,FlagSel_Cad,FlagAtivo,FlagSemRest,FlagBloq,FlagMsgCob,FlagSitTransp_Cad);

	SET IDENTITY_INSERT Clientes OFF
	
	
	*/
	/*** CADASTRO DE COMPRAS ( SEM ITENS )
	
	INSERT INTO tblCompraNF (
		ChvAcesso_CompraNF
		,CNPJ_CPF_CompraNF
		,NomeCompleto_CompraNF
		,NumPed_CompraNF
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
		)
	SELECT '32210268365501000296550000000637741001351624'
		,'68365501000296'
		,'PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
		,(
			SELECT max(IsNull(NumPed_CompraNF, 0)) + 1 AS contador
			FROM tblCompraNF
			)
		,63774
		,'2021/02/15'
		,'1899/12/30 09:44:00'
		,'55'
		,'PEDIDO 322295    TICKET TRANSF;'
		,'0'
		,'1'
		,4527.48
		,181.10
		,0
		,4980.23
		,'322295';	
	
*/

	
-- qryComprasItens_Select_DISTINCT_Almox_CompraNFItem
SELECT DISTINCT tblCompraNF.ModeloDoc_CompraNF, tblCompraNF.CFOP_CompraNF, tblCompraNF.Fil_CompraNF, tblCompraNFItem.Almox_CompraNFItem
FROM tblCompraNF INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF
ORDER BY tblCompraNF.ModeloDoc_CompraNF, tblCompraNF.CFOP_CompraNF;


-- qryComprasItens_Select_DISTINCT_ST_CompraNFItem
SELECT DISTINCT tblCompraNF.ModeloDoc_CompraNF, tblCompraNFItem.ST_CompraNFItem
FROM tblCompraNF INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF
ORDER BY tblCompraNF.ModeloDoc_CompraNF;



SELECT DISTINCT ModeloDoc_CompraNF, CFOP_CompraNF, Fil_CompraNF, Almox_CompraNFItem,FlagEst_CompraNFItem
FROM tblCompraNFItem
INNER JOIN tblCompraNF ON tblCompraNF.ID_CompraNF = tblCompraNFItem.ID_CompraNF_CompraNFItem
ORDER BY ModeloDoc_CompraNF, CFOP_CompraNF;


-- #20211202_update_Almox_CompraNFItem
UPDATE tblCompraNFItem
SET tblCompraNFItem.Almox_CompraNFItem = tabEstoqueAlmox.Codigo_Almox
FROM tabEstoqueAlmox RIGHT JOIN tblCompraNF ON tabEstoqueAlmox.CodUnid_Almox = tblCompraNF.Fil_CompraNF
INNER JOIN tblCompraNFItem ON tblCompraNF.ID_CompraNF = tblCompraNFItem.ID_CompraNF_CompraNFItem
WHERE tabEstoqueAlmox.Codigo_Almox IN (12,1,6) AND tblCompraNFItem.Almox_CompraNFItem IS NULL; 

-- VALIDAR DADOS
-- #20211202_update_Almox_CompraNFItem
SELECT ID_CompraNF,Fil_CompraNF, Almox_CompraNFItem,ID_CompraNF_CompraNFItem,Codigo_Almox,Desc_Almox,CodUnid_Almox
FROM tabEstoqueAlmox RIGHT JOIN tblCompraNF ON tabEstoqueAlmox.CodUnid_Almox = tblCompraNF.Fil_CompraNF
INNER JOIN tblCompraNFItem ON tblCompraNF.ID_CompraNF = tblCompraNFItem.ID_CompraNF_CompraNFItem
WHERE tabEstoqueAlmox.Codigo_Almox IN (12,1,6) AND tblCompraNFItem.Almox_CompraNFItem IS NULL;

SELECT Almox_CompraNFItem,ID_CompraNF_CompraNFItem
-- SELECT DISTINCT Almox_CompraNFItem
from 
	tblCompraNFItem
	

SELECT 
	Codigo_Almox,Desc_Almox,CodUnid_Almox
FROM tabEstoqueAlmox
	
	
	
--ALTER TABLE tblCompraNF ADD CodMunOrigem_CompraNF varchar(255);
--
--ALTER TABLE tblCompraNF ADD CodMunDestino_CompraNF varchar(255);
--
--ALTER TABLE tblCompraNF DROP COLUMN CodMunOrigem;
--
--ALTER TABLE tblCompraNF DROP COLUMN CodMunDestino;
	