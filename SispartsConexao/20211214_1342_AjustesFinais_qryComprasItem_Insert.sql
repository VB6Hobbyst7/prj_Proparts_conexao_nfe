INSERT INTO tblCompraNFItem (
	ChvAcesso_CompraNF
	,Num_CompraNFItem
	,DebICMS_CompraNFItem
	,VTotBaseCalcICMS_CompraNFItem
	,ID_NatOp_CompraNFItem
	,Item_CompraNFItem
	,ID_Grade_CompraNFItem
	,QtdFat_CompraNFItem
	,IPI_CompraNFItem
	,FlagEst_CompraNFItem
	,BaseCalcICMS_CompraNFItem
	,IseICMS_CompraNFItem
	,BaseCalcIPI_CompraNFItem
	,DebIPI_CompraNFItem
	,IseIPI_CompraNFItem
	,TxMLSubsTrib_CompraNFItem
	,BaseCalcICMSSubsTrib_CompraNFItem
	,VTotICMSSubsTrib_compranfitem
	,VTotDesc_CompraNFItem
	,VTotFrete_CompraNFItem
	,VTotPIS_CompraNFItem
	,VTotBaseCalcPIS_CompraNFItem
	,VTotCOFINS_CompraNFItem
	,VTotIseICMS_CompraNFItem
	,VTotBaseCalcCOFINS_CompraNFItem
	,VTotSNCredICMS_CompraNFItem
	,VTotSeg_CompraNFItem
	,VTotOutDesp_CompraNFItem
	,VUnt_CompraNFItem
	,VTot_CompraNFItem
	,ID_Prod_CompraNFItem
	,Almox_CompraNFItem
	,ICMS_CompraNFItem
	,ST_CompraNFItem
	,STCOFINS_CompraNFItem
	,STPIS_CompraNFItem
	,STIPI_CompraNFItem
	)
SELECT tblCompraNF.ChvAcesso_CompraNF
	,tblCompraNF.NumNF_CompraNF
	,IIf([VTotICMS_CompraNF] <> "", Replace(Nz(tblCompraNF.VTotICMS_CompraNF, 0) / 100, ",", "."), 0) AS strVTotICMS
	,tblCompraNF.BaseCalcICMS_CompraNF
	,tblDadosConexaoNFeCTe.ID_NatOp_CompraNF
	,1 AS str_Item_CompraNFItem
	,1 AS str_ID_Grade_CompraNFItem
	,1 AS str_QtdFat_CompraNFItem
	,0 AS str_IPI_CompraNFItem
	,0 AS str_FlagEst_CompraNFItem
	,100 AS str_BaseCalcICMS_CompraNFItem
	,0 AS str_IseICMS_CompraNFItem
	,0 AS str_BaseCalcIPI_CompraNFItem
	,0 AS str_DebIPI_CompraNFItem
	,0 AS str_IseIPI_CompraNFItem
	,0 AS str_TxMLSubsTrib_CompraNFItem
	,0 AS str_BaseCalcICMSSubsTrib_CompraNFItem
	,0 AS str_VTotICMSSubsTrib_compranfitem
	,0 AS str_VTotDesc_CompraNFItem
	,0 AS str_VTotFrete_CompraNFItem
	,0 AS str_VTotPIS_CompraNFItem
	,0 AS str_VTotBaseCalcPIS_CompraNFItem
	,0 AS str_VTotCOFINS_CompraNFItem
	,0 AS str_VTotIseICMS_CompraNFItem
	,0 AS str_VTotBaseCalcCOFINS_CompraNFItem
	,0 AS str_VTotSNCredICMS_CompraNFItem
	,0 AS str_VTotSeg_CompraNFItem
	,0 AS str_VTotOutDesp_CompraNFItem
	,tblCompraNF.VTotNF_CompraNF AS str_VUnt_CompraNFItem
	,tblCompraNF.VTotNF_CompraNF AS str_VTot_CompraNFItem
	,DLookUp("[Código]", "[tmpProdutos]", "[Modelo] = 'TRANSPORTE'") AS str_ID_Prod_CompraNFItem
	,IIf(tblDadosConexaoNFeCTe.codMod = 57, IIf(tblDadosConexaoNFeCTe.ID_EMPRESA = "PSC", 12, IIf(tblDadosConexaoNFeCTe.ID_EMPRESA = "PSP", "6", IIf(tblDadosConexaoNFeCTe.ID_EMPRESA = "PES", "1", "")))) AS str_Almox_CompraNFItem
	,tblDadosConexaoNFeCTe.ICMS_CompraNFItem
	,"0" & DLookUp("[STICMS_NatOper]", "[tmpNatOp]", "[ID_NatOper] = " & tblDadosConexaoNFeCTe.ID_NatOp_CompraNF & "") AS str_ST_CompraNFItem
	,DLookUp("[STPC_NatOper]", "[tmpNatOp]", "[ID_NatOper] = " & tblDadosConexaoNFeCTe.ID_NatOp_CompraNF & "") AS str_STCOFINS_CompraNFItem
	,DLookUp("[STPC_NatOper]", "[tmpNatOp]", "[ID_NatOper] = " & tblDadosConexaoNFeCTe.ID_NatOp_CompraNF & "") AS str_STPIS_CompraNFItem
	,DLookUp("[STIPI_NatOper]", "[tmpNatOp]", "[ID_NatOper] = " & tblDadosConexaoNFeCTe.ID_NatOp_CompraNF & "") AS str_STIPI_CompraNFItem
FROM tblCompraNF
INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso
WHERE (
		((tblDadosConexaoNFeCTe.ID_Tipo) = DLookUp("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='Cte'"))
		AND ((tblDadosConexaoNFeCTe.registroValido) = 1)
		AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1)
		);
