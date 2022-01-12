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
-- SELECT ChvAcesso_CompraNF,NumPed_CompraNF,ID_Forn_CompraNF
-- SELECT COUNT(*)	-- SELECT * -- DELETE
FROM tblCompraNF
INNER JOIN tblNatOp ON tblCompraNF.ID_NatOp_CompraNF = tblNatOp.ID_NatOper
WHERE ChvAcesso_CompraNF = '42210312680452000302550020000902571508970265'


WHERE ModeloDoc_CompraNF = 55
	AND ID_NatOp_CompraNF = 122;
	



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
-- SELECT COUNT(*)	-- SELECT * -- DELETE
FROM tblCompraNFItem
INNER JOIN tblCompraNF ON tblCompraNF.ID_CompraNF = tblCompraNFItem.ID_CompraNF_CompraNFItem
WHERE tblCompraNF.ModeloDoc_CompraNF = 55
	AND tblCompraNF.ChvAcesso_CompraNF = '42210312680452000302550020000902571508970265'