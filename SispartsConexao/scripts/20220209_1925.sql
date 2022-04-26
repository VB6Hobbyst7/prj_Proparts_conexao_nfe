-- 20220209_1925
SELECT NumNF_CompraNF
	,VTotNF_CompraNF
	,VTotProd_CompraNF
-- delete
	FROM tblCompraNF
WHERE ChvAcesso_CompraNF IN (
		'42220104884082000305570000159044501159044500'
		)
GROUP BY NumNF_CompraNF
	,VTotNF_CompraNF
	,VTotProd_CompraNF;


SELECT NumNF_CompraNF, sum(VTot_CompraNFItem)
-- delete
-- select * 
FROM tblCompraNFItem
INNER JOIN tblCompraNF ON tblCompraNF.ID_CompraNF = tblCompraNFItem.ID_CompraNF_CompraNFItem
WHERE tblCompraNF.ChvAcesso_CompraNF IN (
		'42220104884082000305570000159044501159044500'
		)
GROUP BY NumNF_CompraNF;
