SELECT tmpCompras_ID_CompraNF.ChvAcesso_CompraNF FROM tmpCompras_ID_CompraNF;




SELECT tblCompraNF.*
FROM tmpCompras_ID_CompraNF
RIGHT JOIN tblCompraNF ON tmpCompras_ID_CompraNF.NumPed_CompraNF = tblCompraNF.NumPed_CompraNF;


SELECT 
BaseCalcICMS_CompraNF
	,ChvAcesso_CompraNF
	,CNPJ_CPF_CompraNF
	,DTEmi_CompraNF
	,DTEntd_CompraNF
	,HoraEntd_CompraNF
	,IDVD_CompraNF
	,NomeCompleto_CompraNF
	,NumNF_CompraNF
	,NumPed_CompraNF
	,Obs_CompraNF
	,Serie_CompraNF
	,TPNF_CompraNF
	,VTotICMS_CompraNF
	--,VTotIPI_CompraNF
	,VTotNF_CompraNF
	,VTotProd_CompraNF
FROM 
tblCompraNF;


INSERT INTO tblCompraNF (
	BaseCalcICMS_CompraNF
	,ChvAcesso_CompraNF
	,CNPJ_CPF_CompraNF
	,DTEmi_CompraNF
	,DTEntd_CompraNF
	,HoraEntd_CompraNF
	,IDVD_CompraNF
	,NomeCompleto_CompraNF
	,NumNF_CompraNF
	,NumPed_CompraNF
	,Obs_CompraNF
	,Serie_CompraNF
	,TPNF_CompraNF
	,VTotICMS_CompraNF
	,VTotIPI_CompraNF
	,VTotNF_CompraNF
	,VTotProd_CompraNF
	)
VALUES (
	0
	,'32210304884082000569570000040073831040073834'
	,'04884082000569'
	,''
	,''
	,'00:00:00'
	,''
	,'JADLOG LOGISTICA S.A.'
	,4007383
	,000013
	,''
	,'0'
	,'0'
	,4.38
	,36.52
	,0
	);
