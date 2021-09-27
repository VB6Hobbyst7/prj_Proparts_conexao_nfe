--Segue alguns últimos detalhes:
-- - Quando importar cada XML, precisa recortar o arquivo da pasta da empresa e colar dentro de uma pasta chama “Processados”, porém dentro de cada pasta de cada empresa, pois não podemos misturar os XML´s de cada empresa.
-- - Não encontrei um formulário com os XML´s que não foram processados e o motivo.




-- '' #20210823
-- '' #tblDadosConexaoNFeCTe.registroValido
-- #qryUpdateFornecedoresValidos
UPDATE (
		SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF
		FROM tmpClientes
		) AS qryFornecedoresValidos
INNER JOIN tblDadosConexaoNFeCTe ON qryFornecedoresValidos.strCNPJ_CPF = tblDadosConexaoNFeCTe.CNPJ_emit

SET tblDadosConexaoNFeCTe.registroValido = 1;


-- '' #20210823
-- '' -- FiltroCFOP
-- '' #tblDadosConexaoNFeCTe.FiltroCFOP
-- '' #tmpDadosConexaoNFeCTe.ID_NatOp_CompraNF
-- #qryUpdateCFOP_PSC_PES 

UPDATE  ( SELECT 
           tmpNatOp.ID_NatOper, tmpNatOp.Fil_NatOper, tmpNatOp.CFOP_NatOper, qryPscPes.strXMLCFOP, qryPscPes.strEstado  
       FROM (SELECT  
               strSplit(ValorDoParametro,'|',0) AS strFil_NatOper,  strSplit(ValorDoParametro,'|',1) AS strEstado,  strSplit(ValorDoParametro,'|',2) AS strXMLCFOP,  strSplit(ValorDoParametro,'|',3) AS strCFOP_NatOper  
             FROM  
               tblParametros  
             WHERE  
               TipoDeParametro='FiltroFil' And strSplit(ValorDoParametro,'|',0) In ('PSC','PES'))  AS qryPscPes  
       INNER JOIN tmpNatOp ON (qryPscPes.strCFOP_NatOper = tmpNatOp.CFOP_NatOper) AND (qryPscPes.strFil_NatOper = tmpNatOp.Fil_NatOper) )  AS tmpPscPes  
INNER JOIN  
   (   SELECT  *  
       FROM  tblDadosConexaoNFeCTe  
       WHERE tblDadosConexaoNFeCTe.registroValido IN (SELECT TOP 1 cint(tblParametros.ValorDoParametro) FROM [tblParametros] WHERE TipoDeParametro = 'registroValido')  
       AND tblDadosConexaoNFeCTe.ID_NatOp_CompraNF IS NULL )  AS tmpDadosConexaoNFeCTe 
ON (tmpPscPes.strXMLCFOP = tmpDadosConexaoNFeCTe.CFOP) AND (tmpPscPes.Fil_NatOper = tmpDadosConexaoNFeCTe.ID_Empresa) 
SET  tmpDadosConexaoNFeCTe.ID_NatOp_CompraNF = [tmpPscPes].[ID_NatOper], tmpDadosConexaoNFeCTe.FiltroCFOP = [tmpPscPes].[CFOP_NatOper];


-- '' #20210823
-- '' #tblCompraNF.CFOP_CompraNF
-- '' #tblCompraNF.ID_NatOp_CompraNF
-- '' #tblCompraNF.Sit_CompraNF
-- #qryUpdateID_NatOp_CompraNF

UPDATE tblDadosConexaoNFeCTe
INNER JOIN tblCompraNF ON tblDadosConexaoNFeCTe.ChvAcesso = tblCompraNF.ChvAcesso_CompraNF

SET tblCompraNF.CFOP_CompraNF = [tblDadosConexaoNFeCTe].[FiltroCFOP]
	,tblCompraNF.ID_NatOp_CompraNF = [tblDadosConexaoNFeCTe].[ID_NatOp_CompraNF]
	,tblCompraNF.Sit_CompraNF = [tblDadosConexaoNFeCTe].[Sit_CompraNF]
	,tblCompraNF.CFOP_CompraNF = [tblDadosConexaoNFeCTe].[FiltroCFOP]
	,tblCompraNF.ModeloDoc_CompraNF = [tblDadosConexaoNFeCTe].[codMod]
	
WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) > 0));



-- '' #20210823
-- '' #tblCompraNF.Fil_CompraNF
-- #qryUpdateFilCompraNF As String = _

UPDATE (
		SELECT tmpEmpresa.ID_Empresa
			,STRPontos(tmpEmpresa.CNPJ_Empresa) AS strCNPJ_CPF
			,tmpEmpresa.CNPJ_Empresa
		FROM tmpEmpresa
		WHERE (((tmpEmpresa.CNPJ_Empresa) IS NOT NULL))
		) AS qryEmpresas
INNER JOIN tblCompraNF ON qryEmpresas.strCNPJ_CPF = tblCompraNF.CNPJ_CPF_CompraNF

SET tblCompraNF.Fil_CompraNF = qryEmpresas.ID_Empresa;



-- '' #20210823
-- '' #tblCompraNF.NumPed_CompraNF
-- #qryUpdateNumPed_CompraNF 

UPDATE TblCompraNF
SET TblCompraNF.NumPed_CompraNF = Format(IIf(IsNull(DMax('NumPed_CompraNF', 'TblCompraNF')), '000001', DMax('NumPed_CompraNF', 'TblCompraNF') + 1), '000000')
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



-- '' #20210823
-- '' #tblCompraNF.ID_Forn_CompraNF
-- #qryUpdateIdFornCompraNF

UPDATE (
		SELECT STRPontos(tmpClientes.CNPJ_CPF) AS strCNPJ_CPF
			,tmpClientes.CÓDIGOClientes
		FROM tmpClientes
		) AS qryClientesFornecedor
INNER JOIN tblCompraNF ON tblCompraNF.CNPJ_CPF_CompraNF = qryClientesFornecedor.strCNPJ_CPF

SET tblCompraNF.ID_Forn_CompraNF = qryClientesFornecedor.CÓDIGOClientes;



-- '' #20210823
-- '' #tblCompraNF.IDVD_CompraNF
-- #qryUpdate_IDVD

UPDATE tblCompraNF
SET tblCompraNF.IDVD_CompraNF = Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6)
WHERE (
		((Left([Obs_CompraNF], 6)) = 'PEDIDO ')
		AND ((tblCompraNF.CNPJ_CPF_CompraNF) = '12680452000302')
		AND ((Val(Left(Trim(Replace(Replace([tblCompraNF].[Obs_CompraNF], 'PEDIDO: ', ''), 'PEDIDO ', '')), 6))) > 0)
		);


UPDATE tblCompraNF
SET tblCompraNF.IDVD_CompraNF = NULL 
WHERE tblCompraNF.ModeloDoc_CompraNF = 57


-- ######################################################################

-- #tblOrigemDestino


-- '' DTEntd_CompraNF
If XMLDTEmi = "00:00:00" Then
	If Forms!frmCompraNF_ImpXML!Finalidade = 0 Then
		If objNode.ParentNode.nodeName = "dhEmi" Then
			XMLDTEmi = CDate(Replace(Mid(objNode.NodeValue, 1, 10), "-", "/"))
		End If
	Else
		If objNode.ParentNode.nodeName = "dhEmi" Then
			XMLDTEmi = CDate(Replace(Mid(objNode.NodeValue, 1, 10), "-", "/"))
		End If
	End If
End If


-- ''HoraEntd_CompraNF
If XMLdhSaiEnt = "00:00:00" Then
	If Forms!frmCompraNF_ImpXML!Finalidade = 4 Then
		If objNode.ParentNode.nodeName = "dhSaiEnt" Then
			XMLdhSaiEnt = (Replace(Mid(objNode.NodeValue, 12, 8), "-", "/"))
		End If
	End If
End If


'' #AILTON - AJUSTE 1.00
If Forms!frmCompraNF_ImpXML!Finalidade = 4 Then
	rsCompraNF!VTotProd_CompraNF = XMLvNF
	rsCompraNF!VTotNF_CompraNF = XMLvNF
	rsCompraNF!HoraEntd_CompraNF = XMLdhSaiEnt
	sCompraNF!Sit_CompraNF = 6
Else
	rsCompraNF!VTotProd_CompraNF = 0
	rsCompraNF!VTotNF_CompraNF = 0
	rsCompraNF!Sit_CompraNF = 5
End If



-- ######################################################################

-- #VTotProd_CompraNF / infCte/vPrest|vTPrest 

        If XMLValNF = 0 Then
            If objNode.ParentNode.nodeName = "vTPrest" Then
                XMLValNF = objNode.NodeValue / 100
            End If
        End If

-- ######################################################################

-- #VTotNF_CompraNF / cteProc/CTe/infCte/vPrest|vTPrest / nfeProc/NFe/infNFe/total|ICMSTot/vNF

        If XMLValNF = 0 Then
            If objNode.ParentNode.nodeName = "vTPrest" Then
                XMLValNF = objNode.NodeValue / 100
            End If
        End If



-- '' #20210823
-- '' #tblCompraNF.ModeloDoc_CompraNF
-- '' #tblCompraNF.CFOP_CompraNF
-- #qryUpdate_ModeloDoc_CFOP

-- UPDATE tblCompraNF
-- INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso

-- SET tblCompraNF.ModeloDoc_CompraNF = [tblDadosConexaoNFeCTe].[codMod]
	-- ,tblCompraNF.CFOP_CompraNF = [tblDadosConexaoNFeCTe].[FiltroCFOP]
-- WHERE (((tblDadosConexaoNFeCTe.ID_Tipo) > 0));



-- '' #20210823
-- '' #tblCompraNF.CFOP_CompraNF
-- '' #tblCompraNF.Fil_CompraNF
-- #qryUpdateCFOP_FilCompra


-- UPDATE tblCompraNF 
	-- SET 
	-- tblCompraNF.CFOP_CompraNF = DLookUp("[FiltroCFOP]","[tblDadosConexaoNFeCTe]","[ChvAcesso]='" & [tblCompraNF].[ChvAcesso_CompraNF] & "'")
	-- , tblCompraNF.Fil_CompraNF = DLookUp("[ID_EMPRESA]","[tblDadosConexaoNFeCTe]","[ChvAcesso]='" & [tblCompraNF].[ChvAcesso_CompraNF] & "'");


-- -X-X-X-X-X-X-X-X

-- SELECT
	-- tblCompraNF.CFOP_CompraNF 
	-- , tblCompraNF.Fil_CompraNF 
-- FROM 
	-- tblCompraNF 
-- INNER JOIN tblDadosConexaoNFeCTe ON tblCompraNF.ChvAcesso_CompraNF = tblDadosConexaoNFeCTe.ChvAcesso;

-- -X-X-X-X-X-X-X-X

-- SELECT
	-- FiltroCFOP
	-- ,ID_EMPRESA
-- FROM
	-- tblDadosConexaoNFeCTe
-- INNER JOIN tblCompraNF ON tblDadosConexaoNFeCTe.ChvAcesso = tblCompraNF.ChvAcesso_CompraNF;






-- ######################################################################







-- ######################################################################






-- '' #AILTON - qryUpdateFornecedoresValidos
-- '' #ENTENDIMENTO - BLOQUEIO POR FORNECEDOR - ( VAMOS TRABALHAR APENAS COM EMITENTE )
-- strSql = "SELECT Clientes.CNPJ_CPF, Clientes.CÓDIGOClientes AS ID_Cad FROM Clientes WHERE  (Clientes.CNPJ_CPF='" & Format(XMLCNPJEmi, "00\.000\.000/0000\-00") & "');"


-- '' IDCadFor
-- 'SaveSQLString strSQL
-- rsCad.CursorType = adOpenKeyset
-- rsCad.LockType = adLockOptimistic
-- rsCad.CursorLocation = adUseClient
-- rsCad.Open strSql, CNN, , , adCmdText
-- If rsCad.RecordCount = 0 Then
    -- MsgBox "Fornecedor não cadastrado !" & Chr(10) & "CNPJ: " & XMLCNPJEmi, vbInformation, "Atenção"
    -- rsCad.Close
    -- Exit Sub
-- Else
    -- '' #AILTON - ID_FORNECEDOR - qryUpdateFornecedoresValidos
    -- IDCadFor = rsCad!ID_Cad
-- End If







-- ######################################################################





-- ######################################################################


-- ######################################################################
-- $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
-- ######################################################################


-- '' #AILTON - qryExpurgoDeComprasJaCadastradas
-- '' #DUVIDA - QUAL O OBJETIVO ?
-- '' #ENTENDIMENTO - NÃO GERAR DUPLICIDADE NO CADASTRO DE COMPRA VERIFICANDO A EXISTENCIA DA MESMA PELOS SEGUINTES CAMPOS: XMLCNPJEmi E XMLNumNF
-- strSql = "SELECT Clientes.CNPJ_CPF, tblCompraNF.ID_Forn_CompraNF, tblCompraNF.NumNF_CompraNF, tblCompraNF.DTEntd_CompraNF FROM tblCompraNF INNER JOIN Clientes ON tblCompraNF.ID_Forn_CompraNF = Clientes.CÓDIGOClientes WHERE  (Clientes.CNPJ_CPF='" & Format(XMLCNPJEmi, "00\.000\.000/0000\-00") & "') AND (tblCompraNF.NumNF_CompraNF=" & XMLNumNF & ");"

-- 'SaveSQLString strSQL

-- rsCompraNF.CursorType = adOpenKeyset
-- rsCompraNF.LockType = adLockOptimistic
-- rsCompraNF.CursorLocation = adUseClient
-- rsCompraNF.Open strSql, CNN, , , adCmdText
-- If rsCompraNF.RecordCount > 0 Then
    -- 'MsgBox Me.Finalidade.Column(1) & " já importad" & IIf(Me.Finalidade.Column(0) = 0, "o", "a") & "  !", vbInformation, "Atenção"
    -- strXMLjaImp = strXMLjaImp & XMLNumNF & " - Data Entrada: " & rsCompraNF!DTEntd_CompraNF & Chr(10)
    -- 'Contador XML's já importados
    -- CountXMLImp = CountXMLImp + 1
    -- rsCompraNF.Close
    -- Exit Sub
-- End If
-- rsCompraNF.Close


-- ######################################################################
-- $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
-- ######################################################################





-- ######################################################################






