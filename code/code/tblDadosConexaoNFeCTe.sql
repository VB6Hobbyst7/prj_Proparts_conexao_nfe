INSERT INTO tblDadosConexaoNFeCTe (ID_Empresa,ID_Tipo,codMod,codIntegrado,dhEmi,CNPJ_emit,Razao_emit,CNPJ_Rem,CPNJ_Dest,CaminhoDoArquivo,Chave,Comando)
Select [ID_Empresa] as strID_Empresa,[ID_Tipo] as strID_Tipo,[codMod] as strcodMod,[codIntegrado] as strcodIntegrado,[dhEmi] as strdhEmi,[CNPJ_emit] as strCNPJ_emit,[Razao_emit] as strRazao_emit,[CNPJ_Rem] as strCNPJ_Rem,[CPNJ_Dest] as strCPNJ_Dest,[CaminhoDoArquivo] as strCaminhoDoArquivo,[Chave] as strChave,[Comando] as strComando;

