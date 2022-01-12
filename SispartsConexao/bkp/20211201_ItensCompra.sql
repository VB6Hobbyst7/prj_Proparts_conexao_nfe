INSERT INTO tblCompraNFItem (
	ID_Grade_CompraNFItem
	, Almox_CompraNFItem
	, ID_NatOp_CompraNFItem
	, CFOP_CompraNFItem
	, FlagEst_CompraNFItem
	, ST_CompraNFItem
	, STIPI_CompraNFItem
	, STPIS_CompraNFItem
	, STCOFINS_CompraNFItem 	
	, BaseCalcICMS_CompraNFItem
	, BaseCalcICMSSubsTrib_CompraNFItem
	)
SELECT 
	1 AS ID_Grade_CompraNFItem
	, Almox_CompraNFItem 		-- é o Almox de cada Proparst, eu passei para vc as regras, veio com zero. PSC = 12    PES = 1      PSP = 6
	, ID_NatOp_CompraNFItem 	-- veio com zero, o correto é o mesmo ID da tblCompraNF
	, CFOP_CompraNFItem 		-- o correto é o mesmo CFOP da tblCompraNF
	, 1 as FlagEst_CompraNFItem		-- engessar = 1
	, ST_CompraNFItem			-- vieram vazio, mas deve ser consequencia do ID_NatOp_CompraNFItem que veio com zero
	, STIPI_CompraNFItem		-- vieram vazio, mas deve ser consequencia do ID_NatOp_CompraNFItem que veio com zero
	, STPIS_CompraNFItem		-- vieram vazio, mas deve ser consequencia do ID_NatOp_CompraNFItem que veio com zero
	, STCOFINS_CompraNFItem 	-- vieram vazio, mas deve ser consequencia do ID_NatOp_CompraNFItem que veio com zero
	,100 as BaseCalcICMS_CompraNFItem	-- 100 engessado
	, 0 as BaseCalcICMSSubsTrib_CompraNFItem	-- 0 engessado




SELECT
	,1 AS ID_Grade_CompraNFItem
	FROM tblCompraNFItem
	WHERE ID_CompraNF_CompraNFItem = (
			SELECT ID_CompraNF
			-- DELETE
			-- Select *
			-- SELECT COUNT(*) -- 562  
			FROM tblCompraNF
			WHERE ChvAcesso_CompraNF = '42210312680452000302550020000896031976539618'
			);