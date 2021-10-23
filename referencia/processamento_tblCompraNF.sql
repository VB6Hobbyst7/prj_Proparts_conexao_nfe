

-- qryUpdate_tblCompraNF
UPDATE tblDadosConexaoNFeCTe
INNER JOIN tblCompraNF ON tblDadosConexaoNFeCTe.ChvAcesso = tblCompraNF.ChvAcesso_CompraNF

SET tblCompraNF.CFOP_CompraNF = [tblDadosConexaoNFeCTe].[FiltroCFOP]
	,tblCompraNF.ID_NatOp_CompraNF = [tblDadosConexaoNFeCTe].[ID_NatOp_CompraNF]
	,tblCompraNF.Sit_CompraNF = [tblDadosConexaoNFeCTe].[Sit_CompraNF]
	,tblCompraNF.ModeloDoc_CompraNF = [tblDadosConexaoNFeCTe].[codMod]
	,tblCompraNF.Fil_CompraNF = [tblDadosConexaoNFeCTe].[ID_Empresa]
	,tblCompraNF.BaseCalcICMS_CompraNF = replace(Nz([tblCompraNF].[BaseCalcICMS_CompraNF], 0) / 100, ",", ".")
	,tblCompraNF.TPNF_CompraNF = 1

WHERE (((tblDadosConexaoNFeCTe.registroValido) = 1) AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1)) AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0);

---######################################################################


-- qryUpdateID_NatOp_CompraNF, _
UPDATE tblDadosConexaoNFeCTe
INNER JOIN tblCompraNF ON tblDadosConexaoNFeCTe.ChvAcesso = tblCompraNF.ChvAcesso_CompraNF

SET tblCompraNF.CFOP_CompraNF = [tblDadosConexaoNFeCTe].[FiltroCFOP]
	,tblCompraNF.ID_NatOp_CompraNF = [tblDadosConexaoNFeCTe].[ID_NatOp_CompraNF]
	,tblCompraNF.Sit_CompraNF = [tblDadosConexaoNFeCTe].[Sit_CompraNF]
WHERE (((tblDadosConexaoNFeCTe.registroValido) = 1) AND ((tblDadosConexaoNFeCTe.registroProcessado) = 0)) AND ((tblDadosConexaoNFeCTe.ID_Tipo) > 0);


/*** DESCONTINUAR

-- qryUpdateFilCompraNF, _
UPDATE (
		SELECT tmpEmpresa.ID_Empresa
			,STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF
			,tmpEmpresa.CNPJ_Empresa
		FROM tmpEmpresa
		WHERE (((tmpEmpresa.CNPJ_Empresa) IS NOT NULL))
		) AS qryEmpresas
INNER JOIN tblCompraNF ON qryEmpresas.strCNPJ_CPF = tblCompraNF.CNPJ_CPF_CompraNF
SET tblCompraNF.Fil_CompraNF = qryEmpresas.ID_Empresa;


-- qryUpdateCFOP_FilCompra, _
UPDATE tblCompraNF SET tblCompraNF.CFOP_CompraNF = DLookUp("[FiltroCFOP]", "[tblDadosConexaoNFeCTe]", "[ChvAcesso] = '" & [tblCompraNF].[ChvAcesso_CompraNF] & '")	,tblCompraNF.Fil_CompraNF = DLookUp("[ID_EMPRESA]", "[tblDadosConexaoNFeCTe]", "[ChvAcesso] = '" & [tblCompraNF].[ChvAcesso_CompraNF] & '");


-- qryUpdateBaseCalcICMS, _
UPDATE tblCompraNF
SET tblCompraNF.BaseCalcICMS_CompraNF = replace(Nz([tblCompraNF].[BaseCalcICMS_CompraNF], 0) / 100, ",", ".")
WHERE (((tblCompraNF.BaseCalcICMS_CompraNF) > "0"))	OR (((tblCompraNF.BaseCalcICMS_CompraNF) IS NOT NULL));



-- qryUpdate_ModeloDoc_CFOP
UPDATE tblCompraNF
INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso
SET tblCompraNF.ModeloDoc_CompraNF = [tblDadosConexaoNFeCTe].[codMod]
	,tblCompraNF.CFOP_CompraNF = [tblDadosConexaoNFeCTe].[FiltroCFOP]
WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) > 0));


*/









-- qryUpdateIdFornCompraNF, _
UPDATE (
		SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF
			,tmpClientes.CÓDIGOClientes
		FROM tmpClientes
		) AS qryClientesFornecedor
INNER JOIN tblCompraNF ON tblCompraNF.CNPJ_CPF_CompraNF = qryClientesFornecedor.strCNPJ_CPF
SET tblCompraNF.ID_Forn_CompraNF = qryClientesFornecedor.CÓDIGOClientes;




-- qryUpdate_IDVD, _
UPDATE (
		SELECT tblCompraNF.IDVD_CompraNF
			,((Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6))) AS strIDVD_CompraNF
			,[tblCompraNF].[Obs_CompraNF]
		FROM tblCompraNF
		WHERE (
				((Left([Obs_CompraNF], 6)) = 'PEDIDO ')
				AND ((tblCompraNF.CNPJ_CPF_CompraNF) = '12680452000302')
				)
		) AS tmpIDVD_CompraNF
SET tmpIDVD_CompraNF.IDVD_CompraNF = tmpIDVD_CompraNF.strIDVD_CompraNF;



-- qryInsertItensCTe, _
INSERT INTO tblCompraNFItem (
	ChvAcesso_CompraNF
	,VUnt_CompraNFItem
	,Num_CompraNFItem
	,VTot_CompraNFItem
	,DebICMS_CompraNFItem
	,VTotBaseCalcICMS_CompraNFItem
	,ID_NatOp_CompraNFItem
	,Item_CompraNFItem
	,ID_Grade_CompraNFItem
	,QtdFat_CompraNFItem
	,IPI_CompraNFItem
	,FlagEst_CompraNFItem
	,BaseCalcICMS_CompraNFItem
	)
SELECT tblCompraNF.ChvAcesso_CompraNF
	,tblCompraNF.VTotNF_CompraNF AS VUnt_CompraNFItem
	,tblCompraNF.NumNF_CompraNF
	,tblCompraNF.VTotNF_CompraNF
	,IIf([VTotICMS_CompraNF] <> "", Replace(Nz([tblCompraNF].[BaseCalcICMS_CompraNF], 0) / 100, ",", "."), 0) AS strVTotICMS
	,tblCompraNF.BaseCalcICMS_CompraNF
	,tblCompraNF.ID_NatOp_CompraNF
	,1 AS strItem
	,1 AS strIDGrade
	,1 AS strQtdFat
	,0 AS strIPI
	,0 AS strFlag
	,100 AS strBaseCalcICMS
FROM tblCompraNF
INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso
WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) = DLookUp("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='Cte'")));


-- qryUpdateProcessamentoConcluido_CTE
UPDATE tblDadosConexaoNFeCTe
SET tblDadosConexaoNFeCTe.registroProcessado = 1
WHERE (
		((tblDadosConexaoNFeCTe.registroValido) = 1)
		AND ((tblDadosConexaoNFeCTe.registroProcessado) = 1)
		AND ((tblDadosConexaoNFeCTe.ID_Tipo) = DLookUp(" [ValorDoParametro] ", " [tblParametros] ", " [TipoDeParametro] = 'Cte' "))
		);


-- qryUpdateItens_ID_Prod_CompraNFItem
UPDATE tblCompraNFItem
SET tblCompraNFItem.ID_Prod_CompraNFItem = DLookUp("CodigoProd_Grade ", " dbo_tabGradeProdutos ", " CodigoForn_Grade = '" & [tblCompraNFItem].[ID_Prod_CompraNFItem] & "' ");


-- qryUpdateNumPed_CompraNF
UPDATE TblCompraNF
SET TblCompraNF.NumPed_CompraNF = Format(IIf(IsNull(DLookup(" [ValorDoParametro] ", " [tblParametros] ", " [TipoDeParametro] = 'NumPed_CompraNF' ")), '000001', DLookup(" [ValorDoParametro] ", " [tblParametros] ", " [TipoDeParametro] = 'NumPed_CompraNF' ") + 1), '000000')
WHERE (
		(
			(TblCompraNF.ID_CompraNF) IN (
				SELECT TOP 1 ID_CompraNF
				FROM TblCompraNF
				WHERE NumPed_CompraNF IS NULL
				ORDER BY ID_CompraNF
				)
			)
		);


-- qryUpdateNumPed_Contador
UPDATE tblParametros
SET tblParametros.ValorDoParametro = Format(IIf(IsNull(DLookup(" [ValorDoParametro] ", " [tblParametros] ", " [TipoDeParametro] = 'NumPed_CompraNF' ")), '000001', DLookup(" [ValorDoParametro] ", " [tblParametros] ", " [TipoDeParametro] = 'NumPed_CompraNF' ") + 1), '000000')
WHERE tblParametros.TipoDeParametro = " NumPed_CompraNF "


-- qryUpdate_NFe
UPDATE tblCompraNF
INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso
SET tblCompraNF.ModeloDoc_CompraNF = " 55 "
	,tblCompraNF.TPNF_CompraNF = " 1 "
	,tblCompraNF.DTEntd_CompraNF = DATE ()
	,tblCompraNF.VTotICMS_CompraNF = " 0 "
	,tblCompraNF.Sit_CompraNF = " 6 "
	,tblCompraNF.BaseCalcICMS_CompraNF = " 0 
	"
WHERE (
		((tblDadosConexaoNFeCTe.codMod) = 55)
		AND ((tblDadosConexaoNFeCTe.registroValido) = 1)
		AND ((tblDadosConexaoNFeCTe.registroProcessado) = 0)
		);


-- qryUpdate_CTe
UPDATE tblCompraNF
INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso
SET tblCompraNF.ModeloDoc_CompraNF = 57
	,tblCompraNF.TPNF_CompraNF = 1
	,tblCompraNF.Sit_CompraNF = 6
	,tblCompraNF.VTotIPI_CompraNF = 0
WHERE (
		((tblDadosConexaoNFeCTe.codMod) = 57)
		AND ((tblDadosConexaoNFeCTe.registroValido) = 1)
		AND ((tblDadosConexaoNFeCTe.registroProcessado) = 0)
		);


-- qryUpdate_CTeItens
UPDATE tblCompraNFItem
INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNFItem.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso
SET tblCompraNFItem.Item_CompraNFItem = 1
	,tblCompraNFItem.ID_Grade_CompraNFItem = 1
	,tblCompraNFItem.QtdFat_CompraNFItem = 1
	,tblCompraNFItem.IPI_CompraNFItem = 0
	,tblCompraNFItem.FlagEst_CompraNFItem = 0
	,tblCompraNFItem.BaseCalcICMS_CompraNFItem = 100
	,tblCompraNFItem.IseICMS_CompraNFItem = 0
	,tblCompraNFItem.BaseCalcIPI_CompraNFItem = 0
	,tblCompraNFItem.DebIPI_CompraNFItem = 0
	,tblCompraNFItem.IseIPI_CompraNFItem = 0
	,tblCompraNFItem.TxMLSubsTrib_CompraNFItem = 0
	,tblCompraNFItem.BaseCalcICMSSubsTrib_CompraNFItem = 0
	,tblCompraNFItem.VTotICMSSubsTrib_compranfitem = 0
	,tblCompraNFItem.VTotDesc_CompraNFItem = 0
	,tblCompraNFItem.VTotFrete_CompraNFItem = 0
	,tblCompraNFItem.VTotPIS_CompraNFItem = 0
	,tblCompraNFItem.VTotBaseCalcPIS_CompraNFItem = 0
	,tblCompraNFItem.VTotCOFINS_CompraNFItem = 0
	,tblCompraNFItem.VTotIseICMS_CompraNFItem = 0
	,tblCompraNFItem.VTotBaseCalcCOFINS_CompraNFItem = 0
	,tblCompraNFItem.VTotSNCredICMS_CompraNFItem = 0
	,tblCompraNFItem.VTotSeg_CompraNFItem = 0
	,tblCompraNFItem.VTotOutDesp_CompraNFItem = 0
WHERE (
		((tblDadosConexaoNFeCTe.codMod) = 57)
		AND ((tblDadosConexaoNFeCTe.registroValido) = 1)
		AND ((tblDadosConexaoNFeCTe.registroProcessado) = 0)
		);
