/****** Script for SelectTopNRows command from SSMS  ******/
SELECT 
	ID_CompraNF
	,ChvAcesso_CompraNF
	,NumNF_CompraNF
	,Serie_CompraNF
	,NomeCompleto_CompraNF
	,CNPJ_CPF_CompraNF
	,DTEmi_CompraNF
	,DTEntd_CompraNF
	,HoraEntd_CompraNF
	,Obs_CompraNF
	,BaseCalcICMS_CompraNF
	,TPNF_CompraNF
	,VTotICMS_CompraNF
	,VTotNF_CompraNF
	,VTotProd_CompraNF
	-- select * 
	-- delete 
  FROM [SispartsConexao].[dbo].[tblCompraNF]
  -- where ChvAcesso_CompraNF = '42210312680452000302550020000901171433418381'
  WHERE ID_CompraNF = 618
  ORDER BY ID_CompraNF;



 SELECT 
	ID_CompraNF_CompraNFItem
	,Item_CompraNFItem
	,BaseCalcICMS_CompraNFItem
	,BaseCalcICMSSubsTrib_CompraNFItem
	,BaseCalcIPI_CompraNFItem
	,CFOP_CompraNFItem
	,DebICMS_CompraNFItem
	,DebIPI_CompraNFItem
	,ID_Prod_CompraNFItem
	,QtdFat_CompraNFItem
	,VTot_CompraNFItem
	,VTotDesc_CompraNFItem
	,VTotFrete_CompraNFItem
	,VTotOutDesp_CompraNFItem
	,VUnt_CompraNFItem
	-- select ID_CompraNF_CompraNFItem 
	-- select *
	-- delete 
 FROM [SispartsConexao].[dbo].[tblCompraNFItem]
 WHERE ID_CompraNF_CompraNFItem = 618
 ORDER BY id_CompraNFItem,ID_CompraNF_CompraNFItem;


 