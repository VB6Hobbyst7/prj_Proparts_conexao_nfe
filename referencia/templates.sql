SELECT ChvAcesso_CompraNF,NumPed_CompraNF,CNPJ_CPF_CompraNF,NomeCompleto_CompraNF,NumNF_CompraNF,DTEmi_CompraNF,HoraEntd_CompraNF,ModeloDoc_CompraNF,Obs_CompraNF,Serie_CompraNF,TPNF_CompraNF,BaseCalcICMS_CompraNF,VTotICMS_CompraNF,VTotIPI_CompraNF,VTotNF_CompraNF,IDVD_CompraNF
	-- SELECT ChvAcesso_CompraNF,NumPed_CompraNF,ID_Forn_CompraNF
	-- SELECT COUNT(*)	-- DELETE
FROM tblCompraNF
-- where ChvAcesso_CompraNF = '32210204884082000569570000039547081039547081'

SELECT Item_CompraNFItem,ID_Prod_CompraNFItem,CFOP_CompraNFItem,BaseCalcICMSSubsTrib_CompraNFItem,BaseCalcIPI_CompraNFItem,DebICMS_CompraNFItem,DebIPI_CompraNFItem,ICMS_CompraNFItem,IPI_CompraNFItem,QtdFat_CompraNFItem,VUnt_CompraNFItem,VTot_CompraNFItem,VTotBaseCalcICMS_CompraNFItem,ID_CompraNF_CompraNFItem
	-- SELECT ID_CompraNFItem
	-- SELECT DISTINCT(ID_Prod_CompraNFItem)
	-- SELECT COUNT(*)	-- DELETE
FROM tblCompraNFItem;

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