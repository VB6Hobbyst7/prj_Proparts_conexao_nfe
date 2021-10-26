
SELECT COUNT(*) 
	-- DELETE
	-- SELECT * 
	-- SELECT ChvAcesso_CompraNF,NumPed_CompraNF,IDVD_CompraNF,ModeloDoc_CompraNF
	-- SELECT ChvAcesso_CompraNF,ID_CompraNF
	-- SELECT max(NumPed_CompraNF)
FROM tblCompraNF 
where ChvAcesso_CompraNF = '32210268365501000296550000000637821001352053'
order by IDVD_CompraNF;




SELECT COUNT(*) 
	-- DELETE	
	-- SELECT *
	-- SELECT DISTINCT(ID_Prod_CompraNFItem)
FROM tblCompraNFItem;


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
		'32210268365501000296550000000637741001351624'
		,'68365501000296'
		,'PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
		,1
		,63774
		,'2021/02/15'
		,''
		,'00:00:00'
		,0
		,'55'
		,'PEDIDO 322295    TICKET TRANSF;'
		,'0'
		,'5'
		,'1'
		,4527.48
		,0
		,4980.23
		,4980.23
		,''
		,'0'
		,''
		,0
		)
	) AS TMP(ChvAcesso, CNPJ_CPF, NomeCompleto, NumPed, NumNF, DTEmi, DTEntd, HoraEntd, ID_Forn, ModeloDoc, Obs, Serie, Sit, TPNF, BaseCalcICMS, VTotICMS, VTotNF, VTotProd, IDVD, CFOP, Fil, ID_NatOp)
LEFT JOIN tblCompraNF ON tblCompraNF.ChvAcesso_CompraNF = tmp.ChvAcesso
WHERE tblCompraNF.ChvAcesso_CompraNF IS NULL;
