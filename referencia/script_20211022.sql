select *
-- delete
from tblCompraNF;

INSERT INTO tblCompraNF (
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
	,CFOP_CompraNF
	,Fil_CompraNF
	,ID_NatOp_CompraNF
	)
SELECT ChvAcesso
	,CNPJ_CPF
	,NomeCompleto
	,NumPed
	,NumNF
	,DTEmi
	,DTEntd
	,HoraEntd
	,ID_Forn
	,ModeloDoc
	,Obs
	,Serie
	,Sit
	,TPNF
	,BaseCalcICMS
	,VTotICMS
	,VTotNF
	,VTotProd
	,IDVD
	,CFOP
	,Fil
	,ID_NatOp
FROM (
	VALUES (
		'32210304884082000569570000040073831040073834'
		,'04884082000569'
		,'JADLOG LOGISTICA S.A.'
		,000004
		,4007383
		,''
		,'2021/10/20'
		,'00:00:00'
		,0
		,'57'
		,''
		,'0'
		,'0'
		,'1'
		,0
		,0
		,36.52
		,36.52
		,''
		,'2353'
		,'PES'
		,214
		)
	) AS TMP(ChvAcesso, CNPJ_CPF, NomeCompleto, NumPed, NumNF, DTEmi, DTEntd, HoraEntd, ID_Forn, ModeloDoc, Obs, Serie, Sit, TPNF, BaseCalcICMS, VTotICMS, VTotNF, VTotProd, IDVD, CFOP, Fil, ID_NatOp)
LEFT JOIN tblCompraNF ON tblCompraNF.ChvAcesso_CompraNF = tmp.ChvAcesso WHERE tblCompraNF.ChvAcesso_CompraNF IS NULL;
