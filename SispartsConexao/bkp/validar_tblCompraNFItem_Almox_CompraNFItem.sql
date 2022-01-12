
-- VALIDAR DADOS
-- #20211202_update_Almox_CompraNFItem
SELECT ID_CompraNF
	,Fil_CompraNF
	,Almox_CompraNFItem
	,ID_CompraNF_CompraNFItem
	
	,ST_CompraNFItem
	,STIPI_CompraNFItem
	,STPIS_CompraNFItem
	,STCOFINS_CompraNFItem
FROM tblCompraNFItem
INNER JOIN tblCompraNF ON tblCompraNF.ID_CompraNF = tblCompraNFItem.ID_CompraNF_CompraNFItem;





SELECT 
	Codigo_Almox,Desc_Almox,CodUnid_Almox
FROM tabEstoqueAlmox
WHERE tabEstoqueAlmox.Codigo_Almox IN (12,1,6);


/* update_Almox_CompraNFItem
	
	-- #20211202_update_Almox_CompraNFItem
	UPDATE tblCompraNFItem
	SET tblCompraNFItem.Almox_CompraNFItem = tabEstoqueAlmox.Codigo_Almox
	FROM tabEstoqueAlmox
	RIGHT JOIN tblCompraNF ON tabEstoqueAlmox.CodUnid_Almox = tblCompraNF.Fil_CompraNF
	INNER JOIN tblCompraNFItem ON tblCompraNF.ID_CompraNF = tblCompraNFItem.ID_CompraNF_CompraNFItem
	WHERE tabEstoqueAlmox.Codigo_Almox IN (12,1,6)
		AND tblCompraNFItem.Almox_CompraNFItem IS NULL;

	
	*/
