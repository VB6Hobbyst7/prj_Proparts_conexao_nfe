Attribute VB_Name = "modConexaoNfeCte_01_CriacaoDeAmbiente"
Option Compare Database

Private Comando As Variant
Private con As ADODB.Connection


''' CONTROLES DO SISTEMA

'' PARAMETROS
Private Const deleteParametros As String = "DROP TABLE tblParametros"
Private Const createParametros As String = "CREATE TABLE tblParametros (ID AutoIncrement CONSTRAINT PrimaryKey PRIMARY KEY,TipoDeParametro TEXT(50), ValorDoParametro TEXT (255));"
Private Const qryParametro As String = "SELECT tblParametros.ValorDoParametro FROM tblParametros WHERE (((tblParametros.TipoDeParametro) = 'strParametro'))"

'' ORIGEM_DESTINO
Private Const deleteOrigemDestino As String = "DROP TABLE tblOrigemDestino"
Private Const createOrigemDestino As String = "CREATE TABLE tblOrigemDestino (ID AutoIncrement CONSTRAINT PrimaryKey PRIMARY KEY , Origem TEXT(255), Destino TEXT(255), Tipo TEXT(255), Tag TEXT(255), TagOrigem Integer, tabela TEXT(255), campo TEXT(255))"
Private Const qryOrigemDestino As String = "INSERT INTO tblOrigemDestino (Destino,Tipo) VALUES('strDestino','strTipo')"
Private Const qryOrigemDestinoSplit As String = "UPDATE tblOrigemDestino SET tblOrigemDestino.tabela = strSplit([Destino],""."",0), tblOrigemDestino.campo = Replace(Replace(strSplit([Destino],""."",1),""_CompraNFItem"",""""),""_CompraNF"","""");"

'' PROCESSAMENTO ( LEITURA DOS ARQUIVOS XML )
Private Const deleteProcessamento As String = "DROP TABLE tblProcessamento"
Private Const Processamento As String = "CREATE TABLE tblProcessamento(ID AutoIncrement CONSTRAINT PrimaryKey PRIMARY KEY,pk TEXT (50),chave TEXT (255),valor TEXT (255),NomeTabela TEXT (255),NomeCampo TEXT (255))"

'' TIPOS ( IDENTIFICAÇÃO DE DADOS )
Private Const deleteTipos As String = "DROP TABLE tblTipos"
Private Const createTipos As String = "CREATE TABLE tblTipos (ID AutoIncrement CONSTRAINT PrimaryKey PRIMARY KEY,codMod Integer,Descricao TEXT (255));"


''' SISTEMA

'' TABELA DE DADOS GERAIS
Private Const deleteDados As String = "DROP TABLE tblDadosConexaoNFeCTe"
Private Const createDados As String = "CREATE TABLE tblDadosConexaoNFeCTe(ID AutoIncrement CONSTRAINT PrimaryKey PRIMARY KEY,ID_Empresa TEXT (3),ID_Tipo Integer,codMod Integer,codIntegrado Integer,dhEmi TEXT (50),CNPJ_emit TEXT (50),Razao_emit TEXT (255),CNPJ_Rem TEXT (50),CPNJ_Dest TEXT (50),CaminhoDoArquivo TEXT (255),Chave TEXT (255),Comando TEXT (255),codTipoEvento TEXT (255), registroValido Integer, registroProcessado Integer);"

'' TABELA TEMPORARIA PARA SIMULAÇÃO DE COMPRAS
Private Const deleteCompras As String = "DROP TABLE tblCompraNF"
Private Const createCompras As String = "CREATE TABLE tblCompraNF (ID_CompraNF  AutoIncrement CONSTRAINT PrimaryKey PRIMARY KEY,IDOLD_CompraNF Integer,Fil_CompraNF  TEXT (255),NumNF_CompraNF  TEXT (255),NumPed_CompraNF  TEXT (255),NumOrc_CompraNF  TEXT (255),Esp_CompraNF  TEXT (255),Serie_CompraNF  TEXT (255),TPNF_CompraNF  TEXT (255),ID_NatOp_CompraNF  Integer,ID_NatOpOLD_CompraNF  Integer,CFOP_CompraNF  TEXT (255),IESubsTrib_CompraNF  TEXT (255),DTEmi_CompraNF  date,DTEntd_CompraNF  date,HoraEntd_CompraNF  TEXT (255),ID_Forn_CompraNF  Integer,ID_FornOld_CompraNF  Integer,ID_Compr_CompraNF  Integer,ID_Transp_CompraNF  Integer,ID_CondPgto_CompraNF  Integer,BaseCalcICMSSubsTrib_CompraNF  double,VTotICMSSubsTrib_CompraNF  double,VTotFrete_CompraNF  double,VTotSeguro_CompraNF  double,VTotOutDesp_CompraNF  double,BaseCalcICMS_CompraNF  double,VTotICMS_CompraNF  double,VTotIPI_CompraNF  double,VTotISS_CompraNF  double,VTotProd_CompraNF  double,VTotServ_CompraNF  double," & _
                                        "VTotNF_CompraNF  double,TxDesc_CompraNF  double,VTotDesc_CompraNF  double,TPFrete_CompraNF  double,Placa_CompraNF  TEXT (255),UFVeic_CompraNF  TEXT (255),QtdVol_CompraNF  Integer,EspVol_CompraNF TEXT (255),MarcaVol_CompraNF  TEXT (255),NumVol_CompraNF  TEXT (255),PesoBrt_CompraNF  double,PesoLiq_CompraNF  double,DdAdic_CompraNF  TEXT (255),Obs_CompraNF  TEXT (255),Sit_CompraNF  TEXT (255),IDCli_Depto_CompraNF  Integer,IDCli_Contato_CompraNF  Integer,IDCli_Email_CompraNF  Integer,IDCli_Fone_CompraNF  Integer,CondEsp_CompraNF  TEXT (255),Validade_CompraNF  TEXT (255),PzEntg_CompraNF TEXT (255)," & _
                                        "Garantia_CompraNF  TEXT (255),FlagSimples_CompraNF  TEXT (255),FlagDescBaseICMS_CompraNF  TEXT (255),FlagExp_CompraNF  TEXT (255),ModeloDoc_CompraNF  TEXT (255),ChvAcesso_CompraNF  TEXT (255),VTotPIS_CompraNF  double,VTotCOFINS_CompraNF  double,VTotPISRet_CompraNF  double,VTotCOFINSRet_CompraNF  double,VTotCSLLRet_CompraNF  double,VTotIRRFRet_CompraNF  double,FlagSomaST_CompraNF  TEXT(255),FlagCalculado_CompraNF  TEXT (255),VTotISSRet_CompraNF  double,DTExt_CompraNF  date,CNPJ_CPF_CompraNF  TEXT (255),NomeCompleto_CompraNF  TEXT (255),NomeCompleto_VendaNF  TEXT (255),ID_Imp_CompraNF  Integer,SitOR_CompraNF  TEXT (255),NumOR_CompraNF  TEXT (255),FlagNEnvWMAS_CompraNF  TEXT (255),IDVD_CompraNF  TEXT (255),IDVendaNF_CompraNF  TEXT (255),FlagTransf_CompraNF  TEXT(255));"

'' TABELA TEMPORARIA PARA SIMULAÇÃO DE ITENS DAS COMPRAS
Private Const deleteComprasItens As String = "DROP TABLE tblCompraNFItem"
Private Const createComprasItens As String = "CREATE TABLE tblCompraNFItem (ID_CompraNFItem  AutoIncrement CONSTRAINT PrimaryKey PRIMARY KEY , " & _
                                            "IDOLD_CompraNFItem Integer , ID_CompraNF_CompraNFItem Integer , ID_CompraNFOLD_CompraNFItem Integer , Item_CompraNFItem Integer , ID_Prod_CompraNFItem Integer , ID_ProdOld_CompraNFItem Integer , ID_Grade_CompraNFItem Integer , Almox_CompraNFItem Integer , QtdFat_CompraNFItem Integer , VUnt_CompraNFItem double , TxDesc_CompraNFItem double , VUntDesc_CompraNFItem double , ICMS_CompraNFItem double , ISS_CompraNFItem double , IPI_CompraNFItem double , ID_NatOp_CompraNFItem Integer , ID_NatOpOLD_CompraNFItem Integer , CFOP_CompraNFItem TEXT (255) , ST_CompraNFItem TEXT (255) , FlagEst_CompraNFItem Byte , EstDe_CompraNFItem TEXT (255) , EstPara_CompraNFItem TEXT (255) , DTEmi_CompraNFItem date , Esp_CompraNFItem TEXT (255) , Série_CompraNFItem TEXT (255) , Num_CompraNFItem TEXT (255) , Dia_CompraNFItem TEXT (255) , UF_CompraNFItem TEXT (255) , VTot_CompraNFItem double , " & _
                                            "VCntb_CompraNFItem double , BaseCalcICMS_CompraNFItem double , VTotBaseCalcICMS_CompraNFItem double , DebICMS_CompraNFItem double , IseICMS_CompraNFItem double , OutICMS_CompraNFItem double , BaseCalcIPI_CompraNFItem double , DebIPI_CompraNFItem double , IseIPI_CompraNFItem double , OutIPI_CompraNFItem double , Obs_CompraNFItem TEXT (255) , TxMLSubsTrib_CompraNFItem double , TxIntSubsTrib_CompraNFItem double , TxExtSubsTrib_CompraNFItem double , BaseCalcICMSSubsTrib_CompraNFItem double , VTotICMSSubsTrib_compranfitem double , VTotDesc_CompraNFItem double , VTotFrete_CompraNFItem double ,  VTotSeg_CompraNFItem double , STIPI_CompraNFItem TEXT (255) , STPIS_CompraNFItem TEXT (255) , STCOFINS_CompraNFItem TEXT (255) , nID_CompraNFItem TEXT (255) , PIS_CompraNFItem double ,  COFINS_CompraNFItem double , VTotBaseCalcPIS_CompraNFItem double , VTotBaseCalcCOFINS_CompraNFItem double , VTotPIS_CompraNFItem double , VTotCOFINS_CompraNFItem double , " & _
                                            "VTotOutDesp_CompraNFItem double ,  VUntCustoSI_CompraNFItem double , VTotDebISSRet_CompraNFItem double , VTotIseICMS_CompraNFItem double , VTotOutICMS_CompraNFItem double , SNCredICMS_CompraNFItem double , VTotSNCredICMS_CompraNFItem double)"

''==============================================================================================================='
'' OBJETIVO          : CONSTRUÇÃO UNITARIA
''==============================================================================================================='

' Sub teste_CriarTipos(): Dim arr() As Variant: arr = Array(deleteTipos, createTipos): executarComandos arr: CadastroDeItens ItensDeTipos: End Sub
' Sub teste_CriarDados(): Dim arr() As Variant: arr = Array(deleteDados, createDados): executarComandos arr: End Sub

''==============================================================================================================='
'' OBJETIVO          : CONSTRUÇÃO UNITARIA
''==============================================================================================================='


Sub main_criacao()
''==============================================================================================================='
'' OBJETIVO          : Criacao De Ambiente para uso da nova aplicação ( Conexao NF-e e CT-e )
''==============================================================================================================='

On Error Resume Next
    Dim arr() As Variant
    
    
    '' X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X
    '' #EXCLUIR - USAR APENAS EM AMBIENTE DE DESENVOLVIMENTO
    '' X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X
    
'    arr = Array(deleteCompras, deleteComprasItens)
'    executarComandos arr
    
    
    ''#ExclusaoTabelasAuxiliares - Exclusao de tabelas auxiliares caso existam
'    arr = Array(deleteProcessamento, deleteOrigemDestino, deleteParametros, deleteTipos, deleteDados)
'    executarComandos arr
    
On Error GoTo 0


    '' X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X
    '' #EXCLUIR - USAR APENAS EM AMBIENTE DE DESENVOLVIMENTO
    '' X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X
'    arr = Array(createCompras, createComprasItens)
'    executarComandos arr

    ''#CriacaoTabelasAuxiliares - Criação de tabelas auxiliares para uso no processamento de arquivos xmls e json
    arr = Array(createTipos, createParametros, createOrigemDestino, Processamento, createDados)
    executarComandos arr

    ''#CadastroDeTipos - Cadastro de tipos para classificação de registros
    CadastroDeItens ItensDeTipos

    ''#CadastroDeParametros - Cadastro de parametros ex: ( Caminhos, Valores padrões e outros )
    CadastroDeItens ItensDeParametros
            
    ''#ItensDeOrigemDestino - Relacionamento entre campos dos arquivos (nfe,cte) das tabela (tblCompraNF,tblCompraNFItem)
'    CadastroDeItens ItensDeOrigemDestino
    
    
    MsgBox "Concluido!", vbOKOnly + vbInformation, "main_criacao"


End Sub


'' #####################################################################
'' ### #Libs - USADAS APENAS NESTE MÓDULO PARA CRIAÇÃO
'' #####################################################################

Private Sub CadastroDeItens(Itens As Collection)
Dim con As ADODB.Connection: Set con = CurrentProject.Connection
Dim i As Variant

    For Each i In Itens
        con.Execute i
    Next i

Set con = Nothing

End Sub

Private Sub executarComandos(comandos() As Variant)
Dim Comando As Variant

    For Each Comando In comandos
        Application.CurrentDb.Execute Comando
    Next Comando

End Sub

Private Function getTypeText(ID As Integer) As String
Dim myData As Object: Set myData = CreateObject("Scripting.Dictionary")

    myData.add 1, "dbBoolean"
    myData.add 2, "dbByte"
    myData.add 3, "dbInteger"
    myData.add 4, "dbLong"
    myData.add 5, "dbCurrency"
    myData.add 6, "dbSingle"
    myData.add 7, "dbDouble"
    myData.add 8, "dbDate"
    myData.add 9, "dbBinary"
    myData.add 10, "dbText"
    myData.add 11, "dbLongBinary"
    myData.add 12, "dbMemo"
    myData.add 15, "dbGUID"
    myData.add 16, "dbBigInt"
    myData.add 17, "dbVarBinary"
    myData.add 18, "dbChar"
    myData.add 19, "dbNumeric"
    myData.add 20, "dbDecimal"
    myData.add 21, "dbFloat"
    myData.add 22, "dbTime"
    myData.add 23, "dbTimeStamp"

getTypeText = myData(ID)

End Function

'' #####################################################################
'' ### #Repositorios - CARGA DE DADOS INICIAIS
'' #####################################################################


Private Function ItensDeTipos() As Collection
Set ItensDeTipos = New Collection

    '' TIPOS DE CADASTROS
    ItensDeTipos.add "INSERT INTO tblTipos (codMod,Descricao) VALUES(57,'0 - CT-e')"
    ItensDeTipos.add "INSERT INTO tblTipos (codMod,Descricao) VALUES(0,'1 - NF-e Importação')"
    ItensDeTipos.add "INSERT INTO tblTipos (codMod,Descricao) VALUES(0,'2 - NF-e Consumo')"
    ItensDeTipos.add "INSERT INTO tblTipos (codMod,Descricao) VALUES(0,'3 - NF-e com código Sisparts')"
    ItensDeTipos.add "INSERT INTO tblTipos (codMod,Descricao) VALUES(55,'4 - NF-e Retorno Armazém')"
    ItensDeTipos.add "INSERT INTO tblTipos (codMod,Descricao) VALUES(0,'5 - NF-e')"
    ItensDeTipos.add "INSERT INTO tblTipos (codMod,Descricao) VALUES(55,'6 - NF-e Transferência com código Sisparts')"
    ItensDeTipos.add "INSERT INTO tblTipos (codMod,Descricao) VALUES(0,'7 - NF-e Transferência Uso/Consumo com código Sisparts')"

End Function

Private Function ItensDeParametros() As Collection
Set ItensDeParametros = New Collection

    '' ORIGEM/DESTINO
    '' #AILTON - TagOrigem [0,null,1,xml,2,tbl,3,Duvida]
    
    '' CAMINHOS
    ItensDeParametros.add "INSERT INTO tblParametros (TipoDeParametro,ValorDoParametro) VALUES('caminhoDeColeta','C:\temp\Coleta\')"
    ItensDeParametros.add "INSERT INTO tblParametros (TipoDeParametro,ValorDoParametro) VALUES('caminhoDeProcessados','C:\temp\Processados\')"
    
    '' USUARIO
    ItensDeParametros.add "INSERT INTO tblParametros (TipoDeParametro,ValorDoParametro) VALUES('UsuarioErpCod','000001')"
    ItensDeParametros.add "INSERT INTO tblParametros (TipoDeParametro,ValorDoParametro) VALUES('UsuarioErpNome','RoboProparts')"
    
    '' TABELAS AUXILIARES
    ItensDeParametros.add "INSERT INTO tblParametros (TipoDeParametro,ValorDoParametro) VALUES('tabelaAuxiliar','tblParametros')"
    ItensDeParametros.add "INSERT INTO tblParametros (TipoDeParametro,ValorDoParametro) VALUES('tabelaAuxiliar','tblTipos')"
    ItensDeParametros.add "INSERT INTO tblParametros (TipoDeParametro,ValorDoParametro) VALUES('tabelaAuxiliar','tblOrigemDestino')"
    ItensDeParametros.add "INSERT INTO tblParametros (TipoDeParametro,ValorDoParametro) VALUES('tabelaAuxiliar','tblDadosConexaoNFeCTe')"
    
    '' TABELAS TEMPORARIAS
    ItensDeParametros.add "INSERT INTO tblParametros (TipoDeParametro,ValorDoParametro) VALUES('tabelaProcessamento','tblProcessamento')"
        
    '' TABELAS MAPEAMENTO ( ORIGEM / DESTINO )
    ItensDeParametros.add "INSERT INTO tblParametros (TipoDeParametro,ValorDoParametro) VALUES('tblOrigemDestino','tblDadosConexaoNFeCTe')"
    ItensDeParametros.add "INSERT INTO tblParametros (TipoDeParametro,ValorDoParametro) VALUES('tblOrigemDestino','tblCompraNF')"
    ItensDeParametros.add "INSERT INTO tblParametros (TipoDeParametro,ValorDoParametro) VALUES('tblOrigemDestino','tblCompraNFItem')"
        
        
        
    '' X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X
    '' #EXCLUIR - USAR APENAS EM AMBIENTE DE DESENVOLVIMENTO
    '' X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X
    ItensDeParametros.add "INSERT INTO tblParametros (TipoDeParametro,ValorDoParametro) VALUES('tabelaAuxiliar','tblCompraNF')"
    ItensDeParametros.add "INSERT INTO tblParametros (TipoDeParametro,ValorDoParametro) VALUES('tabelaAuxiliar','tblCompraNFItem')"
        
        
    
End Function



Private Function ItensDeOrigemDestino() As Collection
Set ItensDeOrigemDestino = New Collection

    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblDadosConexaoNFeCTe.ID_Empresa' , 'dbLong' , '' , cInt('0') , 'tblDadosConexaoNFeCTe' , 'ID_Empresa')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblDadosConexaoNFeCTe.ID_Tipo' , 'dbLong' , '' , cInt('0') , 'tblDadosConexaoNFeCTe' , 'ID_Tipo')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblDadosConexaoNFeCTe.codMod' , 'dbLong' , 'ide/mod' , cInt('1') , 'tblDadosConexaoNFeCTe' , 'codMod')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblDadosConexaoNFeCTe.codIntegrado' , 'dbLong' , '' , cInt('0') , 'tblDadosConexaoNFeCTe' , 'codIntegrado')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblDadosConexaoNFeCTe.dhEmi' , 'dbText' , 'ide/dhEmi' , cInt('1') , 'tblDadosConexaoNFeCTe' , 'dhEmi')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblDadosConexaoNFeCTe.CNPJ_emit' , 'dbText' , 'emit/CNPJ' , cInt('1') , 'tblDadosConexaoNFeCTe' , 'CNPJ_emit')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblDadosConexaoNFeCTe.Razao_emit' , 'dbText' , 'emit/xNome' , cInt('1') , 'tblDadosConexaoNFeCTe' , 'Razao_emit')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblDadosConexaoNFeCTe.CNPJ_Rem' , 'dbText' , 'rem/CNPJ' , cInt('1') , 'tblDadosConexaoNFeCTe' , 'CNPJ_Rem')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblDadosConexaoNFeCTe.CPNJ_Dest' , 'dbText' , 'dest/CNPJ' , cInt('1') , 'tblDadosConexaoNFeCTe' , 'CPNJ_Dest')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblDadosConexaoNFeCTe.CaminhoDoArquivo' , 'dbText' , 'CaminhoDoArquivo' , cInt('2') , 'tblDadosConexaoNFeCTe' , 'CaminhoDoArquivo')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblDadosConexaoNFeCTe.Chave' , 'dbText' , 'infCTeNorm/infDoc/infNFe/chave' , cInt('1') , 'tblDadosConexaoNFeCTe' , 'Chave')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblDadosConexaoNFeCTe.Comando' , 'dbText' , 'Comando' , cInt('2') , 'tblDadosConexaoNFeCTe' , 'Comando')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.Fil_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'Fil')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.NumNF_CompraNF' , 'dbText' , 'ide/cNF' , cInt('1') , 'tblCompraNF' , 'NumNF')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.NumPed_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'NumPed')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.NumOrc_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'NumOrc')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.Esp_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'Esp')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.Serie_CompraNF' , 'dbText' , 'ide/serie' , cInt('1') , 'tblCompraNF' , 'Serie')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.TPNF_CompraNF' , 'dbText' , 'ide/tpNF' , cInt('1') , 'tblCompraNF' , 'TPNF')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.ID_NatOp_CompraNF' , 'dbLong' , 'ide/natOp' , cInt('1') , 'tblCompraNF' , 'ID_NatOp')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.ID_NatOpOLD_CompraNF' , 'dbLong' , 'xxx' , cInt('0') , 'tblCompraNF' , 'ID_NatOpOLD')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.CFOP_CompraNF' , 'dbText' , 'ide/CFOP' , cInt('1') , 'tblCompraNF' , 'CFOP')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.IESubsTrib_CompraNF' , 'dbText' , 'emit/IE' , cInt('1') , 'tblCompraNF' , 'IESubsTrib')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.DTEmi_CompraNF' , 'dbDate' , 'ide/dhEmi' , cInt('1') , 'tblCompraNF' , 'DTEmi')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.DTEntd_CompraNF' , 'dbDate' , 'xxx' , cInt('0') , 'tblCompraNF' , 'DTEntd')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.HoraEntd_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'HoraEntd')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.ID_Forn_CompraNF' , 'dbLong' , 'xxx' , cInt('0') , 'tblCompraNF' , 'ID_Forn')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.ID_FornOld_CompraNF' , 'dbLong' , 'xxx' , cInt('0') , 'tblCompraNF' , 'ID_FornOld')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.ID_Compr_CompraNF' , 'dbLong' , 'transp/modFrete' , cInt('3') , 'tblCompraNF' , 'ID_Compr')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.ID_Transp_CompraNF' , 'dbLong' , 'transp/modFrete' , cInt('3') , 'tblCompraNF' , 'ID_Transp')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.ID_CondPgto_CompraNF' , 'dbLong' , 'xxx' , cInt('0') , 'tblCompraNF' , 'ID_CondPgto')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.BaseCalcICMSSubsTrib_CompraNF' , 'dbDouble' , 'xxx' , cInt('0') , 'tblCompraNF' , 'BaseCalcICMSSubsTrib')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.VTotICMSSubsTrib_CompraNF' , 'dbDouble' , 'xxx' , cInt('0') , 'tblCompraNF' , 'VTotICMSSubsTrib')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.VTotFrete_CompraNF' , 'dbDouble' , 'xxx' , cInt('0') , 'tblCompraNF' , 'VTotFrete')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.VTotSeguro_CompraNF' , 'dbDouble' , 'total/ICMSTot/vSeg' , cInt('3') , 'tblCompraNF' , 'VTotSeguro')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.VTotOutDesp_CompraNF' , 'dbDouble' , 'total/ICMSTot/vOutro' , cInt('1') , 'tblCompraNF' , 'VTotOutDesp')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.BaseCalcICMS_CompraNF' , 'dbDouble' , 'xxx' , cInt('0') , 'tblCompraNF' , 'BaseCalcICMS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.VTotICMS_CompraNF' , 'dbDouble' , 'total/ICMSTot/vICMS' , cInt('1') , 'tblCompraNF' , 'VTotICMS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.VTotIPI_CompraNF' , 'dbDouble' , 'total/ICMSTot/vIPI' , cInt('3') , 'tblCompraNF' , 'VTotIPI')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.VTotISS_CompraNF' , 'dbDouble' , 'xxx' , cInt('0') , 'tblCompraNF' , 'VTotISS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.VTotProd_CompraNF' , 'dbDouble' , 'total/ICMSTot/vProd' , cInt('1') , 'tblCompraNF' , 'VTotProd')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.VTotServ_CompraNF' , 'dbDouble' , 'xxx' , cInt('0') , 'tblCompraNF' , 'VTotServ')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.VTotNF_CompraNF' , 'dbDouble' , 'total/ICMSTot/vNF' , cInt('1') , 'tblCompraNF' , 'VTotNF')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.TxDesc_CompraNF' , 'dbDouble' , 'xxx' , cInt('0') , 'tblCompraNF' , 'TxDesc')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.VTotDesc_CompraNF' , 'dbDouble' , 'total/ICMSTot/vDesc' , cInt('3') , 'tblCompraNF' , 'VTotDesc')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.TPFrete_CompraNF' , 'dbDouble' , 'total/ICMSTot/vFrete' , cInt('3') , 'tblCompraNF' , 'TPFrete')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.Placa_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'Placa')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.UFVeic_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'UFVeic')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.QtdVol_CompraNF' , 'dbLong' , 'xxx' , cInt('0') , 'tblCompraNF' , 'QtdVol')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.EspVol_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'EspVol')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.MarcaVol_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'MarcaVol')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.NumVol_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'NumVol')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.PesoBrt_CompraNF' , 'dbDouble' , 'xxx' , cInt('0') , 'tblCompraNF' , 'PesoBrt')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.PesoLiq_CompraNF' , 'dbDouble' , 'xxx' , cInt('0') , 'tblCompraNF' , 'PesoLiq')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.DdAdic_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'DdAdic')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.Obs_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'Obs')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.Sit_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'Sit')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.IDCli_Depto_CompraNF' , 'dbLong' , 'xxx' , cInt('0') , 'tblCompraNF' , 'IDCli_Depto')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.IDCli_Contato_CompraNF' , 'dbLong' , 'xxx' , cInt('0') , 'tblCompraNF' , 'IDCli_Contato')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.IDCli_Email_CompraNF' , 'dbLong' , 'xxx' , cInt('0') , 'tblCompraNF' , 'IDCli_Email')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.IDCli_Fone_CompraNF' , 'dbLong' , 'xxx' , cInt('0') , 'tblCompraNF' , 'IDCli_Fone')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.CondEsp_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'CondEsp')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.Validade_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'Validade')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.PzEntg_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'PzEntg')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.Garantia_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'Garantia')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.FlagSimples_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'FlagSimples')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.FlagDescBaseICMS_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'FlagDescBaseICMS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.FlagExp_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'FlagExp')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.ModeloDoc_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'ModeloDoc')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.ChvAcesso_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'ChvAcesso')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.VTotPIS_CompraNF' , 'dbDouble' , 'total/ICMSTot/vPIS' , cInt('0') , 'tblCompraNF' , 'VTotPIS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.VTotCOFINS_CompraNF' , 'dbDouble' , 'total/ICMSTot/vCOFINS' , cInt('0') , 'tblCompraNF' , 'VTotCOFINS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.VTotPISRet_CompraNF' , 'dbDouble' , 'xxx' , cInt('0') , 'tblCompraNF' , 'VTotPISRet')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.VTotCOFINSRet_CompraNF' , 'dbDouble' , 'xxx' , cInt('0') , 'tblCompraNF' , 'VTotCOFINSRet')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.VTotCSLLRet_CompraNF' , 'dbDouble' , 'xxx' , cInt('0') , 'tblCompraNF' , 'VTotCSLLRet')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.VTotIRRFRet_CompraNF' , 'dbDouble' , 'xxx' , cInt('0') , 'tblCompraNF' , 'VTotIRRFRet')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.FlagSomaST_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'FlagSomaST')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.FlagCalculado_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'FlagCalculado')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.VTotISSRet_CompraNF' , 'dbDouble' , 'xxx' , cInt('0') , 'tblCompraNF' , 'VTotISSRet')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.DTExt_CompraNF' , 'dbDate' , 'xxx' , cInt('0') , 'tblCompraNF' , 'DTExt')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.CNPJ_CPF_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'CNPJ_CPF')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.NomeCompleto_CompraNF' , 'dbText' , 'emit/xNome' , cInt('1') , 'tblCompraNF' , 'NomeCompleto')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.NomeCompleto_VendaNF' , 'dbText' , 'emit/xNome' , cInt('3') , 'tblCompraNF' , 'NomeCompleto_VendaNF')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.ID_Imp_CompraNF' , 'dbLong' , 'emit/tpImp' , cInt('3') , 'tblCompraNF' , 'ID_Imp')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.SitOR_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'SitOR')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.NumOR_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'NumOR')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.FlagNEnvWMAS_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'FlagNEnvWMAS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.IDVD_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'IDVD')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.IDVendaNF_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'IDVendaNF')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNF.FlagTransf_CompraNF' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNF' , 'FlagTransf')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.ID_CompraNFItem' , 'dbLong' , 'xxx' , cInt('0') , 'tblCompraNFItem' , 'ID')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.IDOLD_CompraNFItem' , 'dbLong' , 'xxx' , cInt('0') , 'tblCompraNFItem' , 'IDOLD')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.ID_CompraNF_CompraNFItem' , 'dbLong' , 'xxx' , cInt('0') , 'tblCompraNFItem' , 'ID')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.ID_CompraNFOLD_CompraNFItem' , 'dbLong' , 'xxx' , cInt('0') , 'tblCompraNFItem' , 'IDOLD')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.Item_CompraNFItem' , 'dbLong' , 'xxx' , cInt('0') , 'tblCompraNFItem' , 'Item')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.ID_Prod_CompraNFItem' , 'dbLong' , 'det nItem=ContadorX /prod/cProd' , cInt('1') , 'tblCompraNFItem' , 'ID_Prod')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.ID_ProdOld_CompraNFItem' , 'dbLong' , 'xxx' , cInt('0') , 'tblCompraNFItem' , 'ID_ProdOld')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.ID_Grade_CompraNFItem' , 'dbLong' , 'xxx' , cInt('0') , 'tblCompraNFItem' , 'ID_Grade')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.Almox_CompraNFItem' , 'dbLong' , 'xxx' , cInt('0') , 'tblCompraNFItem' , 'Almox')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.QtdFat_CompraNFItem' , 'dbLong' , 'det nItem=ContadorX /prod/qCom' , cInt('3') , 'tblCompraNFItem' , 'QtdFat')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VUnt_CompraNFItem' , 'dbDouble' , 'det nItem=ContadorX /prod/vUnCom' , cInt('3') , 'tblCompraNFItem' , 'VUnt')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.TxDesc_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'TxDesc')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VUntDesc_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'VUntDesc')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.ICMS_CompraNFItem' , 'dbDouble' , 'det nItem=ContadorX /impost/ICMS/CSOSN' , cInt('3') , 'tblCompraNFItem' , 'ICMS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.ISS_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'ISS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.IPI_CompraNFItem' , 'dbDouble' , 'det nItem=ContadorX /impost/IPI/CST' , cInt('3') , 'tblCompraNFItem' , 'IPI')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.ID_NatOp_CompraNFItem' , 'dbLong' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'ID_NatOp')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.ID_NatOpOLD_CompraNFItem' , 'dbLong' , 'xxx' , cInt('0') , 'tblCompraNFItem' , 'ID_NatOpOLD')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.CFOP_CompraNFItem' , 'dbText' , 'det nItem=ContadorX /prod/CFOP' , cInt('1') , 'tblCompraNFItem' , 'CFOP')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.ST_CompraNFItem' , 'dbText' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'ST')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.FlagEst_CompraNFItem' , 'dbByte' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'FlagEst')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.EstDe_CompraNFItem' , 'dbText' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'EstDe')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.EstPara_CompraNFItem' , 'dbText' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'EstPara')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.DTEmi_CompraNFItem' , 'dbDate' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'DTEmi')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.Esp_CompraNFItem' , 'dbText' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'Esp')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.Série_CompraNFItem' , 'dbText' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'Série')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.Num_CompraNFItem' , 'dbText' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'Num')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.Dia_CompraNFItem' , 'dbText' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'Dia')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.UF_CompraNFItem' , 'dbText' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'UF')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VTot_CompraNFItem' , 'dbDouble' , 'det nItem=ContadorX /prod/vProd' , cInt('3') , 'tblCompraNFItem' , 'VTot')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VCntb_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'VCntb')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.BaseCalcICMS_CompraNFItem' , 'dbDouble' , 'det nItem=ContadorX /impost/ICMS/CSOSN' , cInt('3') , 'tblCompraNFItem' , 'BaseCalcICMS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VTotBaseCalcICMS_CompraNFItem' , 'dbDouble' , 'det nItem=ContadorX /impost/ICMS/CSOSN' , cInt('3') , 'tblCompraNFItem' , 'VTotBaseCalcICMS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.DebICMS_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'DebICMS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.IseICMS_CompraNFItem' , 'dbDouble' , 'det nItem=ContadorX /impost/ICMS/CSOSN' , cInt('3') , 'tblCompraNFItem' , 'IseICMS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.OutICMS_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'OutICMS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.BaseCalcIPI_CompraNFItem' , 'dbDouble' , 'det nItem=ContadorX /impost/IPI/IPINT/CST' , cInt('0') , 'tblCompraNFItem' , 'BaseCalcIPI')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.DebIPI_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'DebIPI')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.IseIPI_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'IseIPI')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.OutIPI_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'OutIPI')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.Obs_CompraNFItem' , 'dbText' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'Obs')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.TxMLSubsTrib_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'TxMLSubsTrib')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.TxIntSubsTrib_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'TxIntSubsTrib')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.TxExtSubsTrib_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'TxExtSubsTrib')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.BaseCalcICMSSubsTrib_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'BaseCalcICMSSubsTrib')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VTotICMSSubsTrib_compranfitem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'VTotICMSSubsTrib')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VTotDesc_CompraNFItem' , 'dbDouble' , 'total/ICMSTot/vDesc' , cInt('3') , 'tblCompraNFItem' , 'VTotDesc')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VTotFrete_CompraNFItem' , 'dbDouble' , 'total/ICMSTot/vFrete' , cInt('3') , 'tblCompraNFItem' , 'VTotFrete')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VTotSeg_CompraNFItem' , 'dbDouble' , 'total/ICMSTot/vSeg' , cInt('3') , 'tblCompraNFItem' , 'VTotSeg')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.STIPI_CompraNFItem' , 'dbText' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'STIPI')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.STPIS_CompraNFItem' , 'dbText' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'STPIS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.STCOFINS_CompraNFItem' , 'dbText' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'STCOFINS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.nID_CompraNFItem' , 'dbText' , 'xxx' , cInt('0') , 'tblCompraNFItem' , 'nID')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.PIS_CompraNFItem' , 'dbDouble' , 'total/ICMSTot/vPIS' , cInt('3') , 'tblCompraNFItem' , 'PIS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.COFINS_CompraNFItem' , 'dbDouble' , 'total/ICMSTot/vCOFINS' , cInt('0') , 'tblCompraNFItem' , 'COFINS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VTotBaseCalcPIS_CompraNFItem' , 'dbDouble' , 'imposto/PIS/PISAliq/vBC' , cInt('1') , 'tblCompraNFItem' , 'VTotBaseCalcPIS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VTotBaseCalcCOFINS_CompraNFItem' , 'dbDouble' , 'imposto/COFINS/COFINSAliq/vBC' , cInt('1') , 'tblCompraNFItem' , 'VTotBaseCalcCOFINS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VTotPIS_CompraNFItem' , 'dbDouble' , 'imposto/PIS/PISAliq/vPIS' , cInt('1') , 'tblCompraNFItem' , 'VTotPIS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VTotCOFINS_CompraNFItem' , 'dbDouble' , 'imposto/COFINS/COFINSAliq/vCOFINS' , cInt('1') , 'tblCompraNFItem' , 'VTotCOFINS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VTotOutDesp_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'VTotOutDesp')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VUntCustoSI_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'VUntCustoSI')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VTotDebISSRet_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'VTotDebISSRet')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VTotIseICMS_CompraNFItem' , 'dbDouble' , 'imposto/ICMS/ICMSSN102/CSOSN' , cInt('1') , 'tblCompraNFItem' , 'VTotIseICMS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VTotOutICMS_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'VTotOutICMS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.SNCredICMS_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'SNCredICMS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblCompraNFItem.VTotSNCredICMS_CompraNFItem' , 'dbDouble' , 'xxx' , cInt('3') , 'tblCompraNFItem' , 'VTotSNCredICMS')"
    ItensDeOrigemDestino.add "INSERT INTO tblOrigemDestino (Destino,Tipo,Tag,TagOrigem,tabela,campo) VALUES ('tblDadosConexaoNFeCTe.CPNJ_Dest' , 'dbText' , 'dest/CPF' , cInt('1') , 'tblDadosConexaoNFeCTe' , 'CPNJ_Dest')"


End Function



''#######################################################################################
''### LIMBO
''#######################################################################################


'' #DESCONTINUADO - FOI SUBSTITUIDO POR cadastrar_selecaoDeCampos()
'Private Sub CadastroOrigemDestino() ''DESATIVADO
''' RELACIONAR DE TAGs COM CAMPOS DAS TABELAS USADOS NO PROCESSAMENTO DOS ARQUIVOS
'Dim db As Database
'Dim tdf As TableDef
'Dim x As Integer
'
'    Set db = CurrentDb
'    Set con = CurrentProject.Connection
'
'    For Each Comando In carregarParametros("tblOrigemDestino", qryParametro)
'        For Each tdf In db.TableDefs
'           If left(tdf.Name, 4) <> "MSys" And (tdf.Name = Comando) Then
'              For x = 0 To tdf.Fields.count - 1
'                con.Execute Replace(Replace(qryOrigemDestino, "strDestino", tdf.Name & "." & tdf.Fields(x).Name), "strTipo", getTypeText(tdf.Fields(x).Type))
'              Next x
'           End If
'        Next tdf
'    Next Comando
'
'    Application.CurrentDb.Execute qryOrigemDestinoSplit
'
'Set con = Nothing
'
'End Sub



