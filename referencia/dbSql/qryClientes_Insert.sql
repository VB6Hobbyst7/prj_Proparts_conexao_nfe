SET IDENTITY_INSERT Clientes ON

INSERT INTO Clientes (CÓDIGOClientes,NomeCompleto,CNPJ_CPF,Estado,envia,opMTB,opAv,opTri,opEstr,opDH,opCami,opOut,OptSimples_Cad,vdRS,vdAV,vdSP,vdSC,vdSR,vdTR,vd661,FlagSel_Cad,FlagAtivo,FlagSemRest,FlagBloq,FlagMsgCob,FlagSitTransp_Cad) 
SELECT * FROM ( VALUES
(23306,'Proparts Com Art Esp e Tec - Filial Es','68.365.501/0002-96','ES',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
,(46036,'Proparts Com Art Esp e Tec - Filial Sc','68.365.501/0003-77','SC',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
,(1025,'Proparts Com Art Esp e Tec - Matriz Sp','68.365.501/0001-05','SP',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)
) AS tmp(str_CODIGOClientes,str_NomeCompleto,str_CNPJ_CPF,str_Estado,envia,opMTB,opAv,opTri,opEstr,opDH,opCami,opOut,OptSimples_Cad,vdRS,vdAV,vdSP,vdSC,vdSR,vdTR,vd661,FlagSel_Cad,FlagAtivo,FlagSemRest,FlagBloq,FlagMsgCob,FlagSitTransp_Cad);


SET IDENTITY_INSERT Clientes OFF