-- CadastroDeComprasEmServidor
-- CHAVE_COMPRA_ITEM
-- Item_CompraNFItem,ID_Prod_CompraNFItem,QtdFat_CompraNFItem,VTot_CompraNFItem

select 
Item_CompraNFItem & ID_Prod_CompraNFItem & QtdFat_CompraNFItem & VTot_CompraNFItem as pk
,Item_CompraNFItem
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
,VUnt_CompraNFItem
,TxMLSubsTrib_CompraNFItem
,VTot_CompraNFItem
,VTotBaseCalcICMS_CompraNFItem
,VTotDesc_CompraNFItem
,VTotFrete_CompraNFItem
,VTotICMSSubsTrib_CompraNFItem
,VTotOutDesp_CompraNFItem
from tblCompraNFItem


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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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
		),(
		'32210268365501000296550000000637741001351628'
		,'68365501000296'
		,'AZS - PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
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