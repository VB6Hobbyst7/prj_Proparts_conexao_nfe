
INSERT INTO tblNatOp (Fil_NatOper,IDTP_NatOper,CFOP_NatOper,Descr_NatOper) 
SELECT * FROM ( VALUES
('PES','26','1.932','Transporte Rodoviário')
,('PSC','26','1.932','Transporte Rodoviário')
,('PSP','26','1.932','Transporte Rodoviário')
,('PES','26','2.932','Transporte Rodoviário')
,('PSC','26','2.932','Transporte Rodoviário')
,('PSP','26','2.932','Transporte Rodoviário')
,('PSP','','3.353','Transporte Rodoviário')
,('PSC','26','3.353','Transporte Rodoviário')
,('PES','','3.353','Transporte Rodoviário')) AS tmp(str_Fil_NatOper,str_IDTP_NatOper,str_CFOP_NatOper,str_Descr_NatOper);


SELECT Fil_NatOper,IDTP_NatOper,CFOP_NatOper,Descr_NatOper FROM tblNatOp;