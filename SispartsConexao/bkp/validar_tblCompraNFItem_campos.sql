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
FROM tblCompraNFItem
INNER JOIN tblCompraNF ON tblCompraNF.ID_CompraNF = tblCompraNFItem.ID_CompraNF_CompraNFItem
WHERE tblCompraNF.ModeloDoc_CompraNF = 55
	AND tblCompraNF.ChvAcesso_CompraNF = '42210212680452000302550020000886301507884230'