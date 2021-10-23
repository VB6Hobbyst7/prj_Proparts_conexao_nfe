SELECT ID_CompraNFItem
	,IDOLD_CompraNFItem
	,ID_CompraNF_CompraNFItem
	,ID_CompraNFOLD_CompraNFItem
	,Item_CompraNFItem
	,ID_Prod_CompraNFItem
	,ID_ProdOld_CompraNFItem
	,ID_Grade_CompraNFItem
	,Almox_CompraNFItem
	,QtdFat_CompraNFItem
	,TxDesc_CompraNFItem
	,VUntDesc_CompraNFItem
	,ICMS_CompraNFItem
	,ISS_CompraNFItem
	,IPI_CompraNFItem
	,ID_NatOp_CompraNFItem
	,ID_NatOpOLD_CompraNFItem
	,CFOP_CompraNFItem
	,ST_CompraNFItem
	,FlagEst_CompraNFItem
	,EstDe_CompraNFItem
	,EstPara_CompraNFItem
	,DTEmi_CompraNFItem
	,Esp_CompraNFItem
	,Série_CompraNFItem
	,Num_CompraNFItem
	,Dia_CompraNFItem
	,UF_CompraNFItem
	,VCntb_CompraNFItem
	,BaseCalcICMS_CompraNFItem
	,VTotBaseCalcICMS_CompraNFItem
	,DebICMS_CompraNFItem
	,IseICMS_CompraNFItem
	,OutICMS_CompraNFItem
	,BaseCalcIPI_CompraNFItem
	,DebIPI_CompraNFItem
	,IseIPI_CompraNFItem
	,OutIPI_CompraNFItem
	,Obs_CompraNFItem
	,TxMLSubsTrib_CompraNFItem
	,TxIntSubsTrib_CompraNFItem
	,TxExtSubsTrib_CompraNFItem
	,BaseCalcICMSSubsTrib_CompraNFItem
	,VTotICMSSubsTrib_compranfitem
	,VTotDesc_CompraNFItem
	,VTotFrete_CompraNFItem
	,VTotSeg_CompraNFItem
	,STIPI_CompraNFItem
	,STPIS_CompraNFItem
	,STCOFINS_CompraNFItem
	,nID_CompraNFItem
	,PIS_CompraNFItem
	,COFINS_CompraNFItem
	,VTotBaseCalcPIS_CompraNFItem
	,VTotBaseCalcCOFINS_CompraNFItem
	,VTotPIS_CompraNFItem
	,VTotCOFINS_CompraNFItem
	,VTotOutDesp_CompraNFItem
	,VUntCustoSI_CompraNFItem
	,VTotDebISSRet_CompraNFItem
	,VTotIseICMS_CompraNFItem
	,VTotOutICMS_CompraNFItem
	,SNCredICMS_CompraNFItem
	,VTotSNCredICMS_CompraNFItem
	,VUnt_CompraNFItem
	,VTot_CompraNFItem	
	-- DELETE
	-- SELECT *	-- SELECT COUNT(*)	-- SELECT DISTINCT ID_Prod_CompraNFItem
FROM tblCompraNFItem
INNER JOIN tblCompraNF ON tblCompraNF.ID_CompraNF = tblCompraNFItem.ID_CompraNF_CompraNFItem
where tblCompraNF.ChvAcesso_CompraNF = '32210304884082000569570000040073831040073834';

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
	,TxMLSubsTrib_CompraNFItem
	,VTotBaseCalcICMS_CompraNFItem
	,VTotDesc_CompraNFItem
	,VTotFrete_CompraNFItem
	,VTotICMSSubsTrib_CompraNFItem
	,VTotOutDesp_CompraNFItem
	,VUnt_CompraNFItem
	,VTot_CompraNFItem
FROM tblCompraNFItem
INNER JOIN tblCompraNF ON tblCompraNF.ID_CompraNF = tblCompraNFItem.ID_CompraNF_CompraNFItem
where tblCompraNF.ChvAcesso_CompraNF = '32210368365501000296550000000638841001361501';
