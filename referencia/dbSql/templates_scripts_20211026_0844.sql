<<<<<<< HEAD
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

INSERT INTO tblCompraNF (
	ChvAcesso_CompraNF
	,CNPJ_CPF_CompraNF
	,NomeCompleto_CompraNF
	,NumNF_CompraNF
	,DTEmi_CompraNF
	,HoraEntd_CompraNF
	,Obs_CompraNF
	,Serie_CompraNF
	,TPNF_CompraNF
	,BaseCalcICMS_CompraNF
	,VTotICMS_CompraNF
	,VTotNF_CompraNF
	,IDVD_CompraNF
	)
SELECT (
		'32210268365501000296550000000637741001351624'
		,'68365501000296'
		,'PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
		,63774
		,'2021/02/15'
		,'1899/12/30 09:44:00'
		,'PEDIDO 322295    TICKET TRANSF'
		,'0'
		,'1'
		,4527.48
		,181.10
		,4980.23
		,'322295'
		);

	
	
	

select 
BaseCalcICMS_CompraNF
	,ChvAcesso_CompraNF
	,CNPJ_CPF_CompraNF
	,DTEmi_CompraNF
	,HoraEntd_CompraNF
	,NomeCompleto_CompraNF
	,NumNF_CompraNF
	,Obs_CompraNF
	,Serie_CompraNF
	,TPNF_CompraNF
	,VTotICMS_CompraNF
	,VTotNF_CompraNF
from tblCompraNF;


INSERT INTO tblCompraNF (
	BaseCalcICMS_CompraNF
	,ChvAcesso_CompraNF
	,CNPJ_CPF_CompraNF
	,DTEmi_CompraNF
	,HoraEntd_CompraNF
	,NomeCompleto_CompraNF
	,NumNF_CompraNF
	,Obs_CompraNF
	,Serie_CompraNF
	,TPNF_CompraNF
	,VTotICMS_CompraNF
	,VTotNF_CompraNF
	)
SELECT 1972.30
	,'42210368365501000377550000000064281001362494'
	,'68365501000377'
	,'2021/03/02'
	,'1899/12/30 16:18:00'
	,'PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
	,6428
	,'PEDIDO 324862    TICKET -;'
	,'0'
	,'1'
	,78.89
	,2169.53;


SELECT DISTINCT * FROM (VALUES (1), (1), (1), (2), (5), (1), (6)) AS X(a);


INSERT INTO tblCompraNF (
	BaseCalcICMS_CompraNF
	,ChvAcesso_CompraNF
	,CNPJ_CPF_CompraNF
	,HoraEntd_CompraNF
	,NomeCompleto_CompraNF
	,NumNF_CompraNF
	,Serie_CompraNF
	,TPNF_CompraNF
	,VTotICMS_CompraNF
	,VTotNF_CompraNF
	)
SELECT strBaseCalcICMS_CompraNF
	,strChvAcesso_CompraNF
	,strCNPJ_CPF_CompraNF
	,strHoraEntd_CompraNF
	,strNomeCompleto_CompraNF
	,strNumNF_CompraNF
	,strSerie_CompraNF
	,strTPNF_CompraNF
	,strVTotICMS_CompraNF
	,strVTotNF_CompraNF
FROM (
	VALUES 89.28
		,'42210348740351012767570000021186731952977908'
		,'48740351012767'
		,'1899/12/30 20:14:07'
		,'BRASPRESS TRANSPORTES URGENTES LTDA'
		,2118673
		,'0'
		,'0'
		,10.71
		,89.28
	) AS TMP(strBaseCalcICMS_CompraNF, strChvAcesso_CompraNF, strCNPJ_CPF_CompraNF, strHoraEntd_CompraNF, strNomeCompleto_CompraNF, strNumNF_CompraNF, strSerie_CompraNF, strTPNF_CompraNF, strVTotICMS_CompraNF, strVTotNF_CompraNF)
LEFT JOIN tblCompraNF ON tblCompraNF.ChvAcesso_CompraNF = tmp.strChvAcesso_CompraNF
WHERE tblCompraNF.ChvAcesso_CompraNF IS NULL;




SELECT * FROM tblCompraNF;

INSERT INTO tblCompraNF (
	ChvAcesso_CompraNF
	,CNPJ_CPF_CompraNF
	,NomeCompleto_CompraNF
	,HoraEntd_CompraNF
	,NumNF_CompraNF
	,Serie_CompraNF
	,TPNF_CompraNF
	,BaseCalcICMS_CompraNF
	,VTotICMS_CompraNF
	,VTotNF_CompraNF
	)
SELECT strChvAcesso_CompraNF
	,strCNPJ_CPF_CompraNF
	,strNomeCompleto_CompraNF
	,strHoraEntd_CompraNF
	,strNumNF_CompraNF
	,strSerie_CompraNF
	,strTPNF_CompraNF
	,strBaseCalcICMS_CompraNF
	,strVTotICMS_CompraNF
	,strVTotNF_CompraNF
FROM (
	VALUES '42210348740351012767570000021186731952977908'
		,'48740351012767'
		,'BRASPRESS TRANSPORTES URGENTES LTDA'
		,'1899/12/30 20:14:07'
		,2118673
		,'0'
		,'0'
		,89.28
		,10.71
		,89.28
	) AS TMP(strChvAcesso_CompraNF, strCNPJ_CPF_CompraNF, strNomeCompleto_CompraNF, strHoraEntd_CompraNF, strNumNF_CompraNF, strSerie_CompraNF, strTPNF_CompraNF, strBaseCalcICMS_CompraNF, strVTotICMS_CompraNF, strVTotNF_CompraNF)

















INSERT INTO tblCompraNF (
ChvAcesso_CompraNF,CNPJ_CPF_CompraNF,NomeCompleto_CompraNF,HoraEntd_CompraNF,NumNF_CompraNF,Serie_CompraNF,TPNF_CompraNF,BaseCalcICMS_CompraNF,VTotICMS_CompraNF,VTotNF_CompraNF
)
SELECT ChvAcesso
	,CNPJ_CPF
	,NomeCompleto
	,HoraEntd
	,NumNF
	,Serie
	,TPNF
	,BaseCalcICMS
	,VTotICMS
	,VTotNF
FROM (
	VALUES '42210348740351012767570000021186731952977908'
		,'48740351012767'
		,'BRASPRESS TRANSPORTES URGENTES LTDA'
		,'1899/12/30 20:14:07'
		,2118673
		,'0'
		,'0'
		,89.28
		,10.71
		,89.28
	) AS TMP(ChvAcesso, CNPJ_CPF, NomeCompleto, HoraEntd, NumNF, Serie, TPNF, BaseCalcICMS, VTotICMS, VTotNF);


INSERT INTO tblCompraNF (
ChvAcesso_CompraNF
,CNPJ_CPF_CompraNF
,NomeCompleto_CompraNF
,HoraEntd_CompraNF
,NumNF_CompraNF
,Serie_CompraNF
,TPNF_CompraNF
,BaseCalcICMS_CompraNF
,VTotICMS_CompraNF
,VTotNF_CompraNF	
)
SELECT 
strChvAcesso_CompraNF
,strCNPJ_CPF_CompraNF
,strNomeCompleto_CompraNF
,strHoraEntd_CompraNF
,strNumNF_CompraNF
,strSerie_CompraNF
,strTPNF_CompraNF
,strBaseCalcICMS_CompraNF
,strVTotICMS_CompraNF
,strVTotNF_CompraNF
FROM (
	VALUES (
	'42210348740351012767570000021186731952977908'
	,'48740351012767'
	,'BRASPRESS TRANSPORTES URGENTES LTDA'
	,'20:14:07'
	,'2118673'
	,'0'
	,'0'
	,'89.28'
	,'10.71'
	,'89.28'
		)
	) AS TMP(
	strChvAcesso_CompraNF
	,strCNPJ_CPF_CompraNF
	,strNomeCompleto_CompraNF
	,strHoraEntd_CompraNF
	,strNumNF_CompraNF
	,strSerie_CompraNF
	,strTPNF_CompraNF
	,strBaseCalcICMS_CompraNF
	,strVTotICMS_CompraNF
	,strVTotNF_CompraNF	
	);



SELECT 
strChvAcesso_CompraNF
,strCPNJ_Dest
,stremit_UF
,strCNPJ_CPF_CompraNF
,strNomeCompleto_CompraNF
,strCFOP
,strHoraEntd_CompraNF
,strcodMod
,strNumNF_CompraNF
,strSerie_CompraNF
,strTPNF_CompraNF
,strBaseCalcICMS_CompraNF
,strVTotICMS_CompraNF
,strrem_UF
,strCNPJ_Rem_CompraNF
,strVTotNF_CompraNF
FROM (
	VALUES (
		'42210348740351012767570000021186731952977908'
		,'00011953000155'
		,'SC'
		,'48740351012767'
		,'BRASPRESS TRANSPORTES URGENTES LTDA'
		,'6353'
		,'20:14:07'
		,'57'
		,'2118673'
		,'0'
		,'0'
		,'89.28'
		,'10.71'
		,'SC'
		,'68365501000377'
		,'89.28'
		)
	) AS TMP(strChvAcesso_CompraNF, strCPNJ_Dest, stremit_UF, strCNPJ_CPF_CompraNF, strNomeCompleto_CompraNF, strCFOP, strHoraEntd_CompraNF, strcodMod, strNumNF_CompraNF, strSerie_CompraNF, strTPNF_CompraNF, strBaseCalcICMS_CompraNF, strVTotICMS_CompraNF, strrem_UF, strCNPJ_Rem_CompraNF, strVTotNF_CompraNF);


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



SELECT tmp.strArquivo
FROM (
	VALUES (
("C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210348740351012767570000021186701559009401-cteproc.xml")
,("C:\xmls\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\32210348740351015359570000000443691303812742-cteproc.xml")
,("C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210307872326000158550040001550831011035318-nfeproc.xml")
,("C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210312680452000302550020000902331810472980-nfeproc.xml")
,("C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210320147617000494570010009658691999034138-cteproc.xml")
,("C:\xmls\68.365.5010001-05 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\32210268365501000296550000000637741001351624-nfeproc.xml")
=======
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

INSERT INTO tblCompraNF (
	ChvAcesso_CompraNF
	,CNPJ_CPF_CompraNF
	,NomeCompleto_CompraNF
	,NumNF_CompraNF
	,DTEmi_CompraNF
	,HoraEntd_CompraNF
	,Obs_CompraNF
	,Serie_CompraNF
	,TPNF_CompraNF
	,BaseCalcICMS_CompraNF
	,VTotICMS_CompraNF
	,VTotNF_CompraNF
	,IDVD_CompraNF
	)
SELECT (
		'32210268365501000296550000000637741001351624'
		,'68365501000296'
		,'PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
		,63774
		,'2021/02/15'
		,'1899/12/30 09:44:00'
		,'PEDIDO 322295    TICKET TRANSF'
		,'0'
		,'1'
		,4527.48
		,181.10
		,4980.23
		,'322295'
		);

	
	
	

select 
BaseCalcICMS_CompraNF
	,ChvAcesso_CompraNF
	,CNPJ_CPF_CompraNF
	,DTEmi_CompraNF
	,HoraEntd_CompraNF
	,NomeCompleto_CompraNF
	,NumNF_CompraNF
	,Obs_CompraNF
	,Serie_CompraNF
	,TPNF_CompraNF
	,VTotICMS_CompraNF
	,VTotNF_CompraNF
from tblCompraNF;


INSERT INTO tblCompraNF (
	BaseCalcICMS_CompraNF
	,ChvAcesso_CompraNF
	,CNPJ_CPF_CompraNF
	,DTEmi_CompraNF
	,HoraEntd_CompraNF
	,NomeCompleto_CompraNF
	,NumNF_CompraNF
	,Obs_CompraNF
	,Serie_CompraNF
	,TPNF_CompraNF
	,VTotICMS_CompraNF
	,VTotNF_CompraNF
	)
SELECT 1972.30
	,'42210368365501000377550000000064281001362494'
	,'68365501000377'
	,'2021/03/02'
	,'1899/12/30 16:18:00'
	,'PROPARTS COM ART ESPORTIVOS E TECN EIRELI'
	,6428
	,'PEDIDO 324862    TICKET -;'
	,'0'
	,'1'
	,78.89
	,2169.53;








SELECT DISTINCT * FROM (VALUES (1), (1), (1), (2), (5), (1), (6)) AS X(a);


INSERT INTO tblCompraNF (
	BaseCalcICMS_CompraNF
	,ChvAcesso_CompraNF
	,CNPJ_CPF_CompraNF
	,HoraEntd_CompraNF
	,NomeCompleto_CompraNF
	,NumNF_CompraNF
	,Serie_CompraNF
	,TPNF_CompraNF
	,VTotICMS_CompraNF
	,VTotNF_CompraNF
	)
SELECT strBaseCalcICMS_CompraNF
	,strChvAcesso_CompraNF
	,strCNPJ_CPF_CompraNF
	,strHoraEntd_CompraNF
	,strNomeCompleto_CompraNF
	,strNumNF_CompraNF
	,strSerie_CompraNF
	,strTPNF_CompraNF
	,strVTotICMS_CompraNF
	,strVTotNF_CompraNF
FROM (
	VALUES 89.28
		,'42210348740351012767570000021186731952977908'
		,'48740351012767'
		,'1899/12/30 20:14:07'
		,'BRASPRESS TRANSPORTES URGENTES LTDA'
		,2118673
		,'0'
		,'0'
		,10.71
		,89.28
	) AS TMP(strBaseCalcICMS_CompraNF, strChvAcesso_CompraNF, strCNPJ_CPF_CompraNF, strHoraEntd_CompraNF, strNomeCompleto_CompraNF, strNumNF_CompraNF, strSerie_CompraNF, strTPNF_CompraNF, strVTotICMS_CompraNF, strVTotNF_CompraNF)
LEFT JOIN tblCompraNF ON tblCompraNF.ChvAcesso_CompraNF = tmp.strChvAcesso_CompraNF
WHERE tblCompraNF.ChvAcesso_CompraNF IS NULL;




SELECT * FROM tblCompraNF;

INSERT INTO tblCompraNF (
	ChvAcesso_CompraNF
	,CNPJ_CPF_CompraNF
	,NomeCompleto_CompraNF
	,HoraEntd_CompraNF
	,NumNF_CompraNF
	,Serie_CompraNF
	,TPNF_CompraNF
	,BaseCalcICMS_CompraNF
	,VTotICMS_CompraNF
	,VTotNF_CompraNF
	)
SELECT strChvAcesso_CompraNF
	,strCNPJ_CPF_CompraNF
	,strNomeCompleto_CompraNF
	,strHoraEntd_CompraNF
	,strNumNF_CompraNF
	,strSerie_CompraNF
	,strTPNF_CompraNF
	,strBaseCalcICMS_CompraNF
	,strVTotICMS_CompraNF
	,strVTotNF_CompraNF
FROM (
	VALUES '42210348740351012767570000021186731952977908'
		,'48740351012767'
		,'BRASPRESS TRANSPORTES URGENTES LTDA'
		,'1899/12/30 20:14:07'
		,2118673
		,'0'
		,'0'
		,89.28
		,10.71
		,89.28
	) AS TMP(strChvAcesso_CompraNF, strCNPJ_CPF_CompraNF, strNomeCompleto_CompraNF, strHoraEntd_CompraNF, strNumNF_CompraNF, strSerie_CompraNF, strTPNF_CompraNF, strBaseCalcICMS_CompraNF, strVTotICMS_CompraNF, strVTotNF_CompraNF)

















INSERT INTO tblCompraNF (
ChvAcesso_CompraNF,CNPJ_CPF_CompraNF,NomeCompleto_CompraNF,HoraEntd_CompraNF,NumNF_CompraNF,Serie_CompraNF,TPNF_CompraNF,BaseCalcICMS_CompraNF,VTotICMS_CompraNF,VTotNF_CompraNF
)
SELECT ChvAcesso
	,CNPJ_CPF
	,NomeCompleto
	,HoraEntd
	,NumNF
	,Serie
	,TPNF
	,BaseCalcICMS
	,VTotICMS
	,VTotNF
FROM (
	VALUES '42210348740351012767570000021186731952977908'
		,'48740351012767'
		,'BRASPRESS TRANSPORTES URGENTES LTDA'
		,'1899/12/30 20:14:07'
		,2118673
		,'0'
		,'0'
		,89.28
		,10.71
		,89.28
	) AS TMP(ChvAcesso, CNPJ_CPF, NomeCompleto, HoraEntd, NumNF, Serie, TPNF, BaseCalcICMS, VTotICMS, VTotNF);


INSERT INTO tblCompraNF (
ChvAcesso_CompraNF
,CNPJ_CPF_CompraNF
,NomeCompleto_CompraNF
,HoraEntd_CompraNF
,NumNF_CompraNF
,Serie_CompraNF
,TPNF_CompraNF
,BaseCalcICMS_CompraNF
,VTotICMS_CompraNF
,VTotNF_CompraNF	
)
SELECT 
strChvAcesso_CompraNF
,strCNPJ_CPF_CompraNF
,strNomeCompleto_CompraNF
,strHoraEntd_CompraNF
,strNumNF_CompraNF
,strSerie_CompraNF
,strTPNF_CompraNF
,strBaseCalcICMS_CompraNF
,strVTotICMS_CompraNF
,strVTotNF_CompraNF
FROM (
	VALUES (
	'42210348740351012767570000021186731952977908'
	,'48740351012767'
	,'BRASPRESS TRANSPORTES URGENTES LTDA'
	,'20:14:07'
	,'2118673'
	,'0'
	,'0'
	,'89.28'
	,'10.71'
	,'89.28'
		)
	) AS TMP(
	strChvAcesso_CompraNF
	,strCNPJ_CPF_CompraNF
	,strNomeCompleto_CompraNF
	,strHoraEntd_CompraNF
	,strNumNF_CompraNF
	,strSerie_CompraNF
	,strTPNF_CompraNF
	,strBaseCalcICMS_CompraNF
	,strVTotICMS_CompraNF
	,strVTotNF_CompraNF	
	);



SELECT 
strChvAcesso_CompraNF
,strCPNJ_Dest
,stremit_UF
,strCNPJ_CPF_CompraNF
,strNomeCompleto_CompraNF
,strCFOP
,strHoraEntd_CompraNF
,strcodMod
,strNumNF_CompraNF
,strSerie_CompraNF
,strTPNF_CompraNF
,strBaseCalcICMS_CompraNF
,strVTotICMS_CompraNF
,strrem_UF
,strCNPJ_Rem_CompraNF
,strVTotNF_CompraNF
FROM (
	VALUES (
		'42210348740351012767570000021186731952977908'
		,'00011953000155'
		,'SC'
		,'48740351012767'
		,'BRASPRESS TRANSPORTES URGENTES LTDA'
		,'6353'
		,'20:14:07'
		,'57'
		,'2118673'
		,'0'
		,'0'
		,'89.28'
		,'10.71'
		,'SC'
		,'68365501000377'
		,'89.28'
		)
	) AS TMP(strChvAcesso_CompraNF, strCPNJ_Dest, stremit_UF, strCNPJ_CPF_CompraNF, strNomeCompleto_CompraNF, strCFOP, strHoraEntd_CompraNF, strcodMod, strNumNF_CompraNF, strSerie_CompraNF, strTPNF_CompraNF, strBaseCalcICMS_CompraNF, strVTotICMS_CompraNF, strrem_UF, strCNPJ_Rem_CompraNF, strVTotNF_CompraNF);


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



SELECT tmp.strArquivo
FROM (
	VALUES (
("C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210348740351012767570000021186701559009401-cteproc.xml")
,("C:\xmls\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\32210348740351015359570000000443691303812742-cteproc.xml")
,("C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210307872326000158550040001550831011035318-nfeproc.xml")
,("C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210312680452000302550020000902331810472980-nfeproc.xml")
,("C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210320147617000494570010009658691999034138-cteproc.xml")
,("C:\xmls\68.365.5010001-05 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\32210268365501000296550000000637741001351624-nfeproc.xml")
>>>>>>> f4084cb29d769387d25e7b837853d2119e0da429
) AS TMP(strArquivo);