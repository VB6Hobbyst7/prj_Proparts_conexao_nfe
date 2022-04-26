-- 20220209_1923
SELECT NumNF_CompraNF
	,VTotNF_CompraNF
	,VTotProd_CompraNF
-- delete
	FROM tblCompraNF
WHERE ChvAcesso_CompraNF IN (
		'42220148740351012767570000028353921443763418'
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
		'42220148740351012767570000028353921443763418'
		)
GROUP BY NumNF_CompraNF;
