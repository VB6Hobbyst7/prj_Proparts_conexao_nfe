/********************************************************************

-- qryInsertItensCTe ( ANTIGO )

INSERT INTO tblCompraNFItem (
	ChvAcesso_CompraNF
	,VUnt_CompraNFItem
	,DTEmi_CompraNFItem
	,Num_CompraNFItem
	,VTot_CompraNFItem
	,DebICMS_CompraNFItem
	,VTotBaseCalcICMS_CompraNFItem
	,ID_NatOp_CompraNFItem
	,Item_CompraNFItem
	,ID_Grade_CompraNFItem
	,QtdFat_CompraNFItem
	,ISS_CompraNFItem
	,IPI_CompraNFItem
	,FlagEst_CompraNFItem
	,BaseCalcICMS_CompraNFItem
	,IseICMS_CompraNFItem
	,OutICMS_CompraNFItem
	,BaseCalcIPI_CompraNFItem
	,DebIPI_CompraNFItem
	,IseIPI_CompraNFItem
	,OutIPI_CompraNFItem
	,TxMLSubsTrib_CompraNFItem
	,TxIntSubsTrib_CompraNFItem
	,TxExtSubsTrib_CompraNFItem
	,BaseCalcICMSSubsTrib_CompraNFItem
	,VTotICMSSubsTrib_compranfitem
	,VTotFrete_CompraNFItem
	,VTotDesc_CompraNFItem
	,VTotPIS_CompraNFItem
	,VTotBaseCalcPIS_CompraNFItem
	,VTotCOFINS_CompraNFItem
	,VTotIseICMS_CompraNFItem
	,VTotBaseCalcCOFINS_CompraNFItem
	,SNCredICMS_CompraNFItem
	,VTotSNCredICMS_CompraNFItem
	,VTotSeg_CompraNFItem
	,VTotOutDesp_CompraNFItem
	)
SELECT tblCompraNF.ChvAcesso_CompraNF
	,tblCompraNF.VTotNF_CompraNF
	,tblCompraNF.DTEmi_CompraNF
	,tblCompraNF.NumNF_CompraNF
	,tblCompraNF.VTotNF_CompraNF
	,IIf([VTotICMS_CompraNF] <> """", replace(tblCompraNF.BaseCalcICMS_CompraNF / 100, "", "", "".""), 0) AS strVTotICMS
	,tblCompraNF.BaseCalcICMS_CompraNF
	,tblCompraNF.ID_NatOp_CompraNF
	,1 AS strItem
	,1 AS strIDGrade
	,1 AS strQtdFat
	,0 AS strISS
	,0 AS strIPI
	,0 AS strFlag
	,100 AS strBaseCalcICMS
	,0 AS strIseICMS
	,0 AS strOutICMS
	,0 AS strBaseCalcIPI
	,0 AS strDebIPI
	,0 AS strIseIPI
	,0 AS strOutIPI
	,0 AS strTxMLSubsTrib
	,0 AS strTxIntSubsTrib
	,0 AS strTxExtSubsTrib
	,0 AS strBaseCalcICMSSubsTrib
	,0 AS strVTotICMSSubsTrib
	,0 AS strVTotFrete
	,0 AS strVTotDesc
	,0 AS strVTotPIS
	,0 AS strVTotBaseCalcPIS
	,0 AS strVTotCOFINS
	,0 AS strVTotIseICMS
	,0 AS strVTotBaseCalcCOFINS
	,0 AS strSNCredICMS
	,0 AS strVTotSNCredICMS
	,0 AS strVTotSeg
	,0 AS strVTotOutDesp
FROM tblCompraNF
INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso
WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = DLookUp("" [ValorDoParametro] "", "" [tblParametros] "", "" [TipoDeParametro] = 'Cte' "")))
	AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1);


********************************************************************/