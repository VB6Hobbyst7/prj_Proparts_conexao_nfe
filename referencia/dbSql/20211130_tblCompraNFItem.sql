SELECT COUNT(*) as contador 
-- DELETE
-- Select *
-- SELECT COUNT(*) -- 1220 
FROM tblCompraNFItem where ID_CompraNF_CompraNFItem = (
	SELECT ID_CompraNF
		-- DELETE
		-- Select *
		-- SELECT COUNT(*) -- 562  
		FROM tblCompraNF where ChvAcesso_CompraNF = '42210312680452000302550020000896031976539618'
	);
