SELECT COUNT(*) 
	-- DELETE
	-- SELECT * 
	-- SELECT ChvAcesso_CompraNF,NumPed_CompraNF,ID_Forn_CompraNF
	-- SELECT max(NumPed_CompraNF)
FROM tblCompraNF
-- where ChvAcesso_CompraNF = '32210204884082000569570000039547081039547081'

SELECT COUNT(*) 
	-- DELETE
	-- SELECT DISTINCT(ID_Prod_CompraNFItem)
FROM tblCompraNFItem;

---###################################################################################

-- CadastroDeComprasEmServidor
-- CHAVE_COMPRA_ITEM
-- Item_CompraNFItem,ID_Prod_CompraNFItem,QtdFat_CompraNFItem,VTot_CompraNFItem


---###################################################################################

insert into tblCompraNF (ChvAcesso_CompraNF) 
select ChvAcesso from ( values
('42210312680452000302550020000898451202677594'),
('42210312680452000302550020000898461903898004')) as tmp(ChvAcesso)
left join tblCompraNF on tblCompraNF.ChvAcesso_CompraNF = tmp.ChvAcesso  where tblCompraNF.ChvAcesso_CompraNF is null;

---###################################################################################