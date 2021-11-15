select * 
-- delete
from tblCompraNF;


INSERT INTO tblCompraNF (
	ChvAcesso_CompraNF
	,CNPJ_CPF_CompraNF
	,NomeCompleto_CompraNF
	,NumNF_CompraNF
	,DTEmi_CompraNF
	,HoraEntd_CompraNF
	,Obs_CompraNF
	,Serie_CompraNF
	,TPNF_CompraNF
	,BaseCalcICMS_CompraNF
	,VTotICMS_CompraNF
	,VTotNF_CompraNF
	,IDVD_CompraNF
	)
SELECT '32210268365501000296550000000637741001351624'
	,'68365501000296'
	,'PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
	,63774
	,'2021/02/15'
	,'1899/12/30 09:44:00'
	,'PEDIDO 322295    TICKET TRANSF;'
	,'0'
	,'1'
	,4527.48
	,181.10
	,4980.23
	,'322295';
	
select 
	Item_CompraNFItem
	,ID_CompraNF_CompraNFItem
	,ID_Prod_CompraNFItem
	,CFOP_CompraNFItem
	,BaseCalcICMSSubsTrib_CompraNFItem
	,BaseCalcIPI_CompraNFItem
	,DebICMS_CompraNFItem
	,DebIPI_CompraNFItem
	,ICMS_CompraNFItem
	,IPI_CompraNFItem
	,QtdFat_CompraNFItem
	,VUnt_CompraNFItem
	,VTot_CompraNFItem
	,VTotBaseCalcICMS_CompraNFItem
-- select *
from tblCompraNFItem;



INSERT INTO tblCompraNFItem (
	Item_CompraNFItem
	,ID_Prod_CompraNFItem
	,CFOP_CompraNFItem
	,BaseCalcICMSSubsTrib_CompraNFItem
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
	)
SELECT 1
	,tabGradeProdutos.CodigoProd_Grade
	,tblCompraNF.CFOP_CompraNF
	,00
	,4527.48
	,181.10
	,452.75
	,4.00
	,10.00
	,1.0000
	,4527.4818000000
	,4527.48
	,4527.48
	,tblCompraNF.ID_CompraNF as tmpPK
-- select 
from 
	tblCompraNF,tabGradeProdutos
where 
	ChvAcesso_CompraNF='32210268365501000296550000000637741001351624'
	and tabGradeProdutos.CodigoForn_Grade='00.1918.117.006'