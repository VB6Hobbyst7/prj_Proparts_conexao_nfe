-- SELECT ChvAcesso_CompraNF,ID_CompraNF FROM tblCompraNF where ChvAcesso_CompraNF = '42210312680452000302550020000895841453583169';


SELECT Item
	,ID_CompraNF
	,ID_Prod
	,CFOP
	,BaseCalcICMS
	,BaseCalcICMSSubsTrib
	,BaseCalcIPI
	,DebICMS
	,DebIPI
	,ICMS
	,IPI
	,QtdFat
	,VUnt
	,TxMLSubsTrib
	,VTot
	,VTotBaseCalcICMS
	,VTotDesc
	,VTotFrete
	,VTotICMSSubsTrib
	,VTotOutDesp
FROM (
	VALUES (
		1
		,33980
		,19231
		,'0'
		,0
		,00
		,8827.47
		,353.10
		,882.75
		,4.00
		,10.00
		,78.0000
		,113.1727000000
		,0
		,8827.47
		,8827.47
		,0
		,0
		,0
		,0
		)
		,(
		2
		,33980
		,32088
		,'0'
		,0
		,00
		,37.50
		,1.50
		,3.75
		,4.00
		,10.00
		,1.0000
		,37.5000000000
		,0
		,37.50
		,37.50
		,0
		,0
		,0
		,0
		)
	) AS TMP(Item, ID_CompraNF, ID_Prod, CFOP, BaseCalcICMS, BaseCalcICMSSubsTrib, BaseCalcIPI, DebICMS, DebIPI, ICMS, IPI, QtdFat, VUnt, TxMLSubsTrib, VTot, VTotBaseCalcICMS, VTotDesc, VTotFrete, VTotICMSSubsTrib, VTotOutDesp)
